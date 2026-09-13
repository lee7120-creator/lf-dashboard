"""리런마다 전체를 다시 만지지 않는지 — 두 대시보드의 '상시 비용' 검사.

Streamlit은 리런마다 스크립트 본문을 **통째로 다시 실행**한다. 그래서 한 번만 하면 되는
일을 모듈 최상단이나 위젯 인자에 그냥 두면, 어느 페이지를 보든 필터를 뭘 만지든 매번
그 값을 낸다. 증상이 '전체적으로 좀 느리네'로만 보여서 원인이 안 드러나는 종류다.

실제로 냈던 비용 (합성 1.2만 건 · 13페이지 리런 합계 19.2초 → 7.6초):

| 무엇 | 얼마 | 어디 |
|------|------|------|
| 브랜드 사전 CSV 두 개를 매 리런 파싱 | 0.41초 | 모듈 최상단 (모든 페이지) |
| 리포트 HTML을 받지도 않는데 매번 굽기 | 0.26초 | 페이지 하단 (모든 페이지) |
| 백업 CSV 5개를 접힌 expander 안에서 매번 직렬화 | 0.08초+ | 사이드바 (모든 페이지) |
| 백업 신선도 해시를 쓰지도 않는데 매번 계산 | 0.13초 | 사이드바 (모든 페이지) |
| 리더보드 1.2만 행 Styler 번역 | 2.60초 | 3번 페이지 |

`st.download_button`은 **data를 미리 받는 API**라 누르지 않아도 만들어진다. 대신
`data=`에 **인자 없는 함수**를 주면 누른 뒤에 만든다. 접힌 `st.expander` 안이라고
건너뛰지 않는다 — 본문은 그대로 실행된다.

로컬 실행:
    python tests/test_rerun_cost.py
"""
import ast
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

APPS = ["send_perf_dashboard.py", "weekly_report.py"]

# data= 안에 이게 있으면 '미리 만들었다'는 뜻. 람다 안에 있으면 눌러야 만들어지니 괜찮다.
EAGER_CALLS = {"to_csv", "to_excel", "to_json", "to_parquet", "getvalue", "to_html"}
EAGER_FUNCS = {"build_report_html", "build_report_excel", "df_to_xlsx_bytes", "xlsx_bytes",
               "notes_dict_to_df"}

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def _tree(name):
    return ast.parse((ROOT / name).read_text(encoding="utf-8"), filename=name)


def _download_calls(tree):
    """`st.download_button(...)` / `col.download_button(...)` 호출 전부."""
    for node in ast.walk(tree):
        if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                and node.func.attr == "download_button"):
            yield node


def _data_arg(call):
    for kw in call.keywords:
        if kw.arg == "data":
            return kw.value
    return call.args[1] if len(call.args) > 1 else None


def _eager_inside(expr):
    """람다 **밖**에서 직렬화하는 호출을 찾는다 (람다 안은 눌러야 실행된다)."""
    hits = []

    def walk(n):
        if isinstance(n, ast.Lambda):
            return                                        # 지연 생성 — 여기서 멈춘다
        if isinstance(n, ast.Call):
            f = n.func
            if isinstance(f, ast.Attribute) and f.attr in EAGER_CALLS:
                hits.append(f.attr)
            elif isinstance(f, ast.Name) and f.id in EAGER_FUNCS:
                hits.append(f.id)
        for c in ast.iter_child_nodes(n):
            walk(c)

    walk(expr)
    return hits


@case
def t_download_data_is_made_on_click_not_every_rerun():
    bad = []
    for name in APPS:
        for call in _download_calls(_tree(name)):
            arg = _data_arg(call)
            if arg is None:
                continue
            hits = _eager_inside(arg)
            if hits:
                bad.append(f"{name}:{call.lineno} — data= 안에서 {'/'.join(sorted(set(hits)))}()")
    assert not bad, (
        "누르지도 않은 다운로드를 매 리런 만들고 있어요. `data=`에 인자 없는 람다를 주세요:\n  "
        + "\n  ".join(bad))


@case
def t_brand_dicts_are_read_once_not_every_rerun():
    """사전 CSV는 세션 안에서 안 바뀐다 — 캐시에 얹혀 있어야 한다."""
    import send_perf_dashboard as S
    fn = getattr(S, "_build_brand_dicts", None)
    assert fn is not None, "_build_brand_dicts가 없어요 — 사전 로드가 캐시 밖으로 나갔나요?"
    assert hasattr(fn, "clear"), (
        "_build_brand_dicts가 st.cache_resource로 감싸여 있지 않아요. "
        "모듈 최상단은 리런마다 다시 실행되니 사전 CSV를 매번 파싱하게 돼요.")
    assert len(S.BRAND_MAP) > 1000 and len(S.BRAND_CODE) > 1000


@case
def t_brand_dicts_are_not_touched_after_the_builder():
    """캐시가 돌려주는 건 **같은 객체**다. 밖에서 고치면 두 번째 리런부터 이미 고쳐진
    사전을 또 고친다. 지금 규칙은 우연히 멱등이지만 규칙이 하나 늘면 조용히 깨진다."""
    tree = _tree("send_perf_dashboard.py")
    names = {"BRAND_MAP", "BRAND_EXACT", "BRAND_CODE", "_BRAND_CANON"}
    bad = []
    for node in tree.body:                                # 모듈 최상단만 본다
        for sub in ast.walk(node):
            tgt = None
            if isinstance(sub, ast.Subscript) and isinstance(sub.ctx, ast.Store):
                tgt = sub.value
            elif (isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute)
                  and sub.func.attr in {"setdefault", "update", "pop", "clear"}):
                tgt = sub.func.value
            if isinstance(tgt, ast.Name) and tgt.id in names:
                bad.append(f"line {sub.lineno} — {tgt.id}")
    assert not bad, ("빌더 밖에서 브랜드 사전을 고치고 있어요. 후처리는 "
                     "_build_brand_dicts 안에서 끝내세요:\n  " + "\n  ".join(bad))


@case
def t_leaderboard_formats_in_the_browser_not_in_python():
    """리더보드는 필터 전체(실백업 1.2만 행)를 담는다. Styler를 넘기면 Streamlit이
    **모든 행을 미리 문자열로 번역**해(`Styler._translate`) 1.2만 행에서 2.6초가 든다.
    화면 서식은 column_config(브라우저가 그린다)로, 엑셀만 Styler로 내보낸다."""
    sys.path.insert(0, str(ROOT / "tests"))
    from streamlit.testing.v1 import AppTest
    from smoke_pages import synth_store

    at = AppTest.from_file(str(ROOT / "send_perf_dashboard.py"), default_timeout=600)
    at.session_state["camp_store"] = synth_store(weeks=8, per_day=2)
    at.run()
    at.sidebar.radio[0].set_value("3. 캠페인 리더보드").run()
    assert not at.exception, at.exception[0].value

    dfs = [e for e in at._tree if type(e).__name__ == "Dataframe"]
    assert dfs, "리더보드 표를 못 찾았어요"
    cfg = str(getattr(dfs[0].proto, "columns", ""))
    for want in ("%,.0f", "%.2f%%"):
        assert want in cfg, f"column_config에 {want} 서식이 없어요 — Styler로 되돌아갔나요?\n{cfg}"

    # 진짜 증거는 여기다 — Styler를 넘기면 Streamlit이 행별 표시문자열(display_values)을
    # 같이 실어 보낸다. 그 payload가 없다는 건 1.2만 행 번역을 안 했다는 뜻이다.
    ad = dfs[0].proto.arrow_data
    assert not ad.HasField("styler"), (
        f"행별 표시문자열 {len(ad.styler.display_values)}칸이 실려 있어요 — "
        "리더보드가 Styler 렌더로 되돌아갔어요.")

    # 엑셀 내보내기(Styler 경로)는 살아 있어야 한다 — 숫자+표시형식 규칙이 거기 걸려 있다.
    labels = [b.label for b in at.get("download_button")]
    assert "⬇️ 엑셀" in labels, f"엑셀 버튼이 사라졌어요: {labels}"


def main():
    fails = []
    for fn in CASES:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {e}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"리런 비용 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
