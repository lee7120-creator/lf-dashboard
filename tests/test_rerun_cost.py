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
import re
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


_FUNC = (ast.FunctionDef, ast.AsyncFunctionDef)


def _scopes(tree):
    """함수 하나(와 모듈 최상단)를 한 범위로 본다.

    모듈 전체를 한 통으로 보면 **이름만 같은** 다른 함수의 지역변수가 걸려 오탐이 난다 —
    실제로 `lazy_download`의 `data`(이미 만들어 둔 바이트)가 파서 함수의
    `data = f.getvalue()`에 걸렸다.
    """
    for node in ast.walk(tree):
        if isinstance(node, (*_FUNC, ast.Module)):
            yield node


def _own_nodes(scope):
    """그 범위가 **직접** 가진 노드들 — 중첩 함수 안으로는 안 들어간다."""
    stack = list(ast.iter_child_nodes(scope))
    while stack:
        n = stack.pop()
        yield n
        if isinstance(n, _FUNC):
            continue                                      # 중첩 함수는 자기 범위에서 본다
        stack += list(ast.iter_child_nodes(n))


def _eager_locals(scope):
    """그 함수 안에서 `이름 = <직렬화 호출>` 로 담아 둔 것 → 어떤 호출이었나.

    `data=`만 봐서는 **한 줄 위에서 미리 만든 것**을 못 잡는다:

        csv = df.to_csv(index=False).encode("utf-8-sig")   # ← 리런마다 여기서 만든다
        st.download_button("...", csv, ...)                # ← data=는 그냥 이름

    실제로 「통합 데이터·다운로드」가 이 모양이라 검사를 통과한 채로 남아 있었다.
    """
    out = {}
    for node in _own_nodes(scope):
        if not isinstance(node, ast.Assign) or len(node.targets) != 1:
            continue
        tgt = node.targets[0]
        if not isinstance(tgt, ast.Name):
            continue
        hits = _eager_inside(node.value)
        if hits:
            out.setdefault(tgt.id, []).append((node.lineno, hits))
    return out


def _download_calls(scope):
    """그 범위가 직접 부르는 `st.download_button(...)` / `col.download_button(...)`."""
    for node in _own_nodes(scope):
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
        for scope in _scopes(_tree(name)):
            eager = _eager_locals(scope)
            for call in _download_calls(scope):
                arg = _data_arg(call)
                if arg is None:
                    continue
                hits = _eager_inside(arg)
                if hits:
                    bad.append(f"{name}:{call.lineno} — data= 안에서 "
                               f"{'/'.join(sorted(set(hits)))}()")
                elif isinstance(arg, ast.Name) and arg.id in eager:
                    # 같은 함수 안 한 줄 위에서 미리 만든 것도 결국 매 리런 만든다
                    _ln, _h = eager[arg.id][0]
                    bad.append(f"{name}:{call.lineno} — data={arg.id} 인데 "
                               f"{_ln}번 줄에서 {'/'.join(sorted(set(_h)))}()로 미리 만들었어요")
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


@case
def t_pages_are_not_referenced_by_number():
    """페이지를 **번호로** 가리키지 말 것 — 재정렬하면 조용히 끊긴다.

    실제로 `page.startswith("09.")`로 박아 둔 '마스터 없이도 조직·카테고리는 열린다'
    예외가, 페이지를 재정렬하자 안 열리게 됐다. 증상이 그냥 **빈 화면**이라 안 드러난다.
    이름(`PAGE_ORGCAT` 같은 상수)으로 가리키면 번호가 바뀌어도 산다.
    """
    bad = []
    for name in APPS:
        for node in ast.walk(_tree(name)):
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                    and node.func.attr in ("startswith", "endswith")
                    and isinstance(node.func.value, ast.Name)
                    and node.func.value.id == "page"):
                for a in node.args:
                    if isinstance(a, ast.Constant) and re.match(r"^\d{2}\.", str(a.value)):
                        bad.append(f"{name}:{node.lineno} — page.{node.func.attr}({a.value!r})")
    assert not bad, ("페이지를 번호로 가리키고 있어요. 이름 상수를 쓰세요:\n  "
                     + "\n  ".join(bad))


def _page_map(name):
    """앱이 실제로 내거는 **페이지 → 하위탭** 표를 모은다.

    주간보고는 `PAGES` 리스트(하위탭 없음), 발송성과는 그룹 dict(키=페이지, 값=하위탭)다.
    어느 쪽이든 **번호로 시작하는 문자열**이라 모양으로 찾는다 — 변수명을 박으면 이름을
    바꿀 때 검사가 조용히 빈 표를 보고 통과한다.
    """
    pages = {}
    for node in ast.walk(_tree(name)):
        if isinstance(node, ast.Dict):
            got = {k.value: {c.value for c in v.elts
                             if isinstance(c, ast.Constant) and isinstance(c.value, str)}
                   for k, v in zip(node.keys, node.values)
                   if (isinstance(k, ast.Constant) and isinstance(k.value, str)
                       and re.match(r"^\d+\.\s", k.value) and isinstance(v, ast.List))}
        elif isinstance(node, ast.List):
            got = {c.value: set() for c in node.elts
                   if isinstance(c, ast.Constant) and isinstance(c.value, str)
                   and re.match(r"^\d+\.\s", c.value)}
        else:
            continue
        if len(got) >= 3:                # 페이지 목록이라고 볼 만한 덩어리만
            for k, v in got.items():
                pages.setdefault(k, set()).update(v)
    return pages


@case
def t_page_references_in_prose_match_the_real_pages():
    """안내 문구가 가리키는 「NN. 페이지 › 하위탭」이 **실제로 있어야** 한다.

    번호 참조를 금지하면 정확한 참조(`「12. 회원UV·거래액」`)까지 죽는다. 금지가 아니라
    **대조**라야 썩은 것만 잡힌다.

    실제로 페이지를 9→7로 줄인 뒤 「09. 조직·카테고리별 실적」이라고 안내하는 문구가 둘
    남아, **없는 페이지 번호로 사람을 보내고 있었다**. `page.startswith` 검사는 제어
    흐름만 봐서 이걸 못 봤고, 증상은 화면이 멀쩡히 떠서 눈으로도 안 잡힌다.
    """
    _norm = lambda t: re.sub(r"\s+", " ", t).strip()
    bad = []
    for name in APPS:
        pages = _page_map(name)
        assert pages, f"{name}에서 페이지 목록을 못 찾았어요 — 검사가 헛돌고 있어요."
        for node in ast.walk(_tree(name)):
            if not (isinstance(node, ast.Constant) and isinstance(node.value, str)):
                continue
            for ref in re.findall(r"「(\d+\.\s*[^」]+)」", node.value):
                head, *rest = [_norm(x) for x in ref.split("›")]
                if head not in pages:
                    bad.append(f"{name}:{node.lineno} — 「{ref}」 (그런 페이지가 없어요)")
                elif rest and pages[head] and rest[0] not in pages[head]:
                    bad.append(f"{name}:{node.lineno} — 「{ref}」 "
                               f"({head}에 그런 하위탭이 없어요)")
    assert not bad, ("없는 페이지를 가리키는 안내 문구예요. 번호를 고치거나 이름 상수를 "
                     "쓰세요:\n  " + "\n  ".join(bad))


@case
def t_derived_periods_never_reach_the_store():
    """파생한 주·월은 **저장·백업에 안 실린다** — 담기면 다음 세션엔 '파일 값'이 된다.

    파생은 파일이 있는 기간을 비켜 가므로, 한 번 저장되면 그 기간은 영영 안 갱신된다.
    일별을 고쳐 다시 올려도 옛 파생이 그 자리를 차지한 채 남는다 — 증상이 '숫자가 안
    바뀌네'로만 보여 원인이 안 드러난다.
    """
    tree = _tree("weekly_report.py")
    fn = next((n for n in ast.walk(tree)
               if isinstance(n, ast.FunctionDef) and n.name == "main"), None)
    assert fn is not None, "main()을 못 찾았어요"
    src = ast.unparse(fn)
    assert "odf_raw = odf" in src and "orgcat_derive_periods(odf_raw)" in src, \
        "파생 전 프레임을 따로 안 들고 있어요"
    for call in ("save_orgcat_store(", "make_backup_zip("):
        for node in ast.walk(fn):
            if (isinstance(node, ast.Call)
                    and ast.unparse(node.func).endswith(call.rstrip("("))):
                args = " ".join(ast.unparse(a) for a in node.args)
                assert "odf_raw" in args or "odf" not in args.split(), \
                    f"{call} 에 파생본(odf)을 넘겨요 — odf_raw여야 해요: {args[:80]}"


@case
def t_digit_parses_from_labels_are_guarded():
    """선택지 라벨에서 숫자를 뽑을 땐 **None을 가드**한다.

    `re.search(...).group()`은 매치가 없으면 `AttributeError`로 죽는다 — 주간보고
    7페이지를 통째로 날린 사고가 정확히 이 모양이었다. 지금 라벨엔 다 숫자가 있어
    안 터지지만, 라벨 문구는 자주 바뀌는 자리라 한 번 바꾸면 그 페이지가 죽는다.
    """
    bad = []
    for name in APPS:
        for node in ast.walk(_tree(name)):
            if not (isinstance(node, ast.Attribute) and node.attr == "group"):
                continue
            inner = node.value
            if (isinstance(inner, ast.Call) and isinstance(inner.func, ast.Attribute)
                    and inner.func.attr in ("search", "match")
                    and ast.unparse(inner.func.value) == "re"):
                bad.append(f"{name}:{node.lineno} — {ast.unparse(node)[:60]}")
    assert not bad, ("`re.search(...).group()`을 그대로 부르고 있어요. 매치가 없으면 "
                     "페이지가 죽어요:\n  " + "\n  ".join(bad))


@case
def t_weekly_store_csvs_are_cached_by_file_signature():
    """주간보고의 저장소 CSV는 **캐시 함수 안에서만** 읽는다.

    주간보고는 저장소를 세션이 아니라 디스크에서 매번 읽는다(발송성과는 세션에 들고
    있어 이 문제가 없다). 조직×카테고리 실파일이 70만 행이라, 캐시가 없으면 CSV 파싱
    0.42초 + `orgcat_fill` 0.26초를 **조직×카테고리를 안 쓰는 페이지까지** 매 리런 낸다.
    증상이 '전체적으로 좀 느리네'로만 보여 원인이 안 드러나는 종류다.

    `pd.read_csv(<모듈 상수>)`만 본다 — 업로드 파일 파서는 `io.BytesIO(...)`라 안 걸린다.
    """
    tree = _tree("weekly_report.py")
    def _cached(fn):
        for d in fn.decorator_list:
            if "cache_data" in ast.unparse(d) or "cache_resource" in ast.unparse(d):
                return True
        return False
    bad = []
    for fn in ast.walk(tree):
        if not isinstance(fn, _FUNC) or _cached(fn):
            continue
        for n in _own_nodes(fn):
            if (isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                    and n.func.attr == "read_csv" and n.args
                    and isinstance(n.args[0], ast.Name)):
                bad.append(f"{fn.name}():{n.lineno} — pd.read_csv({n.args[0].id})")
    assert not bad, ("저장소 CSV를 캐시 밖에서 읽고 있어요. 파일 서명"
                     "(`os.stat`의 mtime_ns·size)을 인자로 받는 `@st.cache_data` 함수로 "
                     "옮겨 주세요:\n  " + "\n  ".join(bad))


@case
def t_caches_are_bounded():
    """`@st.cache_data`엔 **`max_entries`(또는 `ttl`)를 반드시 준다.**

    기본이 **무제한**이라, 키가 하나 늘 때마다 결과 프레임이 그대로 남는다. 저장소
    읽기는 `(path, mtime, size)`가 키라서 **저장할 때마다 새 항목이 생기고 옛 70만 행·
    120만 행 프레임이 안 버려진다** — 실측으로 저장 한 번에 평균 +273MB씩 올라가
    564MB → 1,657MB가 됐다. 한도를 넘으면 Streamlit이 컨테이너를 죽이고 화면엔
    **「Oh no. Error running app.」** 만 뜬다. 리런하면 살아나서 원인이 안 드러난다.

    상한은 쓰임새에 맞춰 준다 — 파일 서명 캐시는 1(가장 최근 한 벌만 쓸모가 있다),
    업로드 묶음은 2, 파일별 파서는 4쯤.
    """
    bad = []
    for name in APPS:
        for node in ast.walk(_tree(name)):
            if not isinstance(node, _FUNC):
                continue
            for d in node.decorator_list:
                if not (isinstance(d, ast.Call)
                        and "cache_data" in ast.unparse(d.func)):
                    continue
                kw = {k.arg for k in d.keywords}
                if not ({"max_entries", "ttl"} & kw):
                    bad.append(f"{name}:{node.lineno} — {node.name}()")
    assert not bad, ("한도 없는 `@st.cache_data`가 있어요. 키가 늘수록 결과가 그대로 "
                     "쌓여 OOM으로 앱이 죽어요:\n  " + "\n  ".join(bad))


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
