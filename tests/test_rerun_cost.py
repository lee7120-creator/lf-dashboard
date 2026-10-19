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
import collections
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
    from smoke_pages import group_of, synth_store

    at = AppTest.from_file(str(ROOT / "send_perf_dashboard.py"), default_timeout=600)
    at.session_state["camp_store"] = synth_store(weeks=8, per_day=2)
    at.run()
    at.sidebar.radio[0].set_value(group_of("캠페인 리더보드")).run()
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
def t_planning_memo_is_folded_not_mixed_with_the_numbers():
    """「앱푸시 수신동의 모수」의 **기획 문서**는 접어 두고 분석과 섞지 않는다.

    이 화면은 740줄에 블록이 열셋이었다. 그중 ①상태 매트릭스 ②상태 전환 흐름
    ③KPI 구조 ④실행 로드맵은 숫자가 아니라 **운영 설계 메모**인데, 분석 블록과
    같은 높이로 펼쳐져 있어서 이 화면이 무엇을 묻는 화면인지 안 읽혔다.

    **정적으로 본다.** 이 화면은 앱푸시 원천 + 기간 선택이 맞아떨어져야 본문 끝까지
    가는데, 그 픽스처를 맞추는 비용에 비해 렌더로 더 잡히는 게 「라벨이 바뀌었다」
    정도다. 되돌아가는 길은 둘 뿐이고 둘 다 AST로 보인다 — 메모를 다시 `main()`
    안에 펼쳐 놓거나, 호출을 빼 **메모가 통째로 사라지는데 화면은 멀쩡히 뜨거나.**
    """
    tree = _tree("send_perf_dashboard.py")
    HEADS = ("① 상태 매트릭스", "② 상태 전환 흐름", "③ KPI 구조", "④ 실행 로드맵")

    fn = next((n for n in tree.body
               if isinstance(n, ast.FunctionDef) and n.name == "push_consent_playbook"), None)
    assert fn, "push_consent_playbook() 가 모듈 레벨에 없어요 — 메모를 다시 main() 안에 폈나요?"
    inside = "\n".join(str(getattr(n, "value", "")) for n in ast.walk(fn)
                       if isinstance(n, ast.Constant) and isinstance(n.value, str))
    missing = [h for h in HEADS if h not in inside]
    assert not missing, f"메모 함수 안에서 사라진 블록이 있어요: {missing}"

    # main() 안에 같은 제목이 남아 있으면 둘로 갈린 것이다
    main_fn = next(n for n in tree.body
                   if isinstance(n, ast.FunctionDef) and n.name == "main")
    dup = [h for h in HEADS
           if any(isinstance(n, ast.Constant) and isinstance(n.value, str) and h in n.value
                  for n in ast.walk(main_fn))]
    assert not dup, f"main() 안에도 같은 블록이 남아 있어요: {dup}"

    # 호출은 반드시 **접이식 안**에서 — 그냥 부르면 접은 의미가 없다
    folded = False
    for n in ast.walk(main_fn):
        if not isinstance(n, ast.With):
            continue
        it = " ".join(ast.unparse(w.context_expr) for w in n.items)
        if "st.expander" not in it:
            continue
        if any(isinstance(c, ast.Call) and getattr(c.func, "id", "") == "push_consent_playbook"
               for c in ast.walk(n)):
            folded = True
            assert "expanded=False" in it, f"메모가 펼쳐진 채로 열려요 — {it}"
    assert folded, ("push_consent_playbook() 을 `with st.expander(...)` 안에서 안 불러요 — "
                    "호출이 빠지면 메모가 통째로 사라지는데 화면은 멀쩡히 떠요.")


@case
def t_pages_are_not_referenced_by_number():
    """페이지를 **번호로** 가리키지 말 것 — 재정렬하면 조용히 끊긴다.

    실제로 `page.startswith("09.")`로 박아 둔 '마스터 없이도 조직·카테고리는 열린다'
    예외가, 페이지를 재정렬하자 안 열리게 됐다. 증상이 그냥 **빈 화면**이라 안 드러난다.
    이름(`PAGE_ORGCAT` 같은 상수)으로 가리키면 번호가 바뀌어도 산다.

    변수 이름을 `page`로 박아 두면 같은 실수가 다른 이름으로 돌아온다 — 실제로
    발송성과의 그룹 안내가 `_grp.startswith("6.")`로 남아 있었다(재정렬하면 「전사 MTD를
    올리면 볼 수 있어요」가 조용히 안 뜬다). 받는 쪽을 안 가리고 **번호로 시작하는
    리터럴**을 다 본다. 한 자리 번호도 마찬가지다.
    """
    bad = []
    for name in APPS:
        for node in ast.walk(_tree(name)):
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                    and node.func.attr in ("startswith", "endswith")):
                tgt = ast.unparse(node.func.value)
                for a in node.args:
                    for c in (a.elts if isinstance(a, ast.Tuple) else [a]):
                        if isinstance(c, ast.Constant) and re.match(r"^\d+\.", str(c.value)):
                            bad.append(f"{name}:{node.lineno} — "
                                       f"{tgt}.{node.func.attr}({c.value!r})")
            # 번호까지 박은 라벨과의 동등 비교도 같은 자리다 — 주간보고 분기가 그랬다.
            if isinstance(node, ast.Compare):
                for c in node.comparators:
                    if (isinstance(c, ast.Constant) and isinstance(c.value, str)
                            and re.match(r"^\d+\.\s", c.value)):
                        bad.append(f"{name}:{node.lineno} — {ast.unparse(node)[:60]}")
    assert not bad, ("페이지를 번호로 가리키고 있어요. 이름 상수를 쓰세요:\n  "
                     + "\n  ".join(bad))


def _page_map(name):
    """앱이 실제로 내거는 **페이지 → 하위탭** 표를 모은다.

    주간보고는 `PAGES` 리스트(하위탭 없음), 발송성과는 그룹 dict(키=페이지, 값=하위탭)다.
    어느 쪽이든 **번호로 시작하는 문자열**이라 모양으로 찾는다 — 변수명을 박으면 이름을
    바꿀 때 검사가 조용히 빈 표를 보고 통과한다.

    주간보고는 번호를 목록에서 붙이는 형태(`PAGE_ORDER` → `f"{i:02d}. {이름}"`)라
    리터럴이 없다. **이름 상수를 풀어** 같은 규칙으로 번호를 매겨 본다 — 앱이 번호 꼴을
    바꾸면 여기서 어긋나 검사가 시끄럽게 실패한다(조용히 통과하는 것보다 낫다).
    """
    tree = _tree(name)
    consts = {n.targets[0].id: n.value.value
              for n in tree.body
              if isinstance(n, ast.Assign) and len(n.targets) == 1
              and isinstance(n.targets[0], ast.Name)
              and isinstance(n.value, ast.Constant) and isinstance(n.value.value, str)}
    pages = {}
    for node in ast.walk(tree):
        if (isinstance(node, ast.List) and len(node.elts) >= 3
                and all(isinstance(e, ast.Name) and e.id in consts for e in node.elts)):
            for i, e in enumerate(node.elts, 1):
                pages.setdefault(f"{i:02d}. {consts[e.id]}", set())
            continue
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


@case
def t_ui_prose_is_haeyoche():
    """화면에 뜨는 글은 **해요체**다 — 합쇼체가 섞이면 여러 사람이 이어붙인 인상을 준다.

    두 대시보드 모두 해요체로 통일돼 있는데, 문구를 고치다 보면 한두 줄이 합쇼체로
    남는다. 「…하나로 못박습니다」·「…정상입니다 — 」가 실제로 그렇게 남아 있었다.
    화면은 멀쩡히 떠서 눈으로만 잡히는 자리라 검사로 못 박는다.

    **AI 프롬프트는 안 본다** — `system`/`user` 문자열은 출력 서식을 지시하는 내용이라
    말투를 바꾸면 생성 결과가 깨진다. 그 문자열은 `st.*`를 안 거치니 자연히 빠진다.

    **위젯의 `help=`·라벨도 화면에 뜨는 글이다.** 본문만 보면 거기 남은 합쇼체를 놓친다 —
    실제로 「…운영 차이를 보게 됩니다」·「…급락처럼 보입니다」 둘이 `help=`에 있었다.
    """
    show = {"caption", "info", "markdown", "warning", "success", "error", "write",
            "subheader", "header", "title", "metric", "toast"}
    widget = {"selectbox", "multiselect", "radio", "checkbox", "slider", "text_input",
              "number_input", "button", "download_button", "toggle", "pills",
              "segmented_control", "select_slider", "date_input", "text_area",
              "expander", "file_uploader", "form_submit_button", "color_picker"}
    # 종결 어미로 쓰인 것만 — 문장 끝(끝·문장부호·태그·줄표) 앞에 있을 때다.
    hap = re.compile(r"(합니다|습니다|입니다|됩니다|드립니다|바랍니다|하십시오|십시오)"
                     r"\s*(?:[.!?]|<|—|$)")
    bad = []
    for name in APPS:
        for node in ast.walk(_tree(name)):
            if not isinstance(node, ast.Call):
                continue
            attr = getattr(node.func, "attr", None) or ""
            if attr in show:
                args = list(node.args)
            elif attr in widget:
                args = list(node.args[:1])          # 첫 인자가 라벨이다
            else:
                continue
            args += [k.value for k in node.keywords
                     if k.arg in ("help", "label", "body", "placeholder")]
            for arg in args:
                for sub in ast.walk(arg):
                    if isinstance(sub, ast.Constant) and isinstance(sub.value, str):
                        m = hap.search(sub.value)
                        if m:
                            bad.append(f"{name}:{sub.lineno} — [{attr} · {m.group(1)}] "
                                       f"{sub.value.strip()[:70]}")
    assert not bad, ("화면 문구에 합쇼체가 섞였어요. 해요체로 맞춰 주세요:\n  "
                     + "\n  ".join(bad))


# ──────────────────────────────────────────────────────────────────────
# 차트 레이아웃 — `update_layout`은 리런마다 재귀 병합을 다시 돈다
# ──────────────────────────────────────────────────────────────────────
# `fig.update_layout(**lay)`는 플로틀리가 기존 레이아웃에 속성을 하나씩 얹으면서 매직
# 밑줄 경로를 다시 파싱하고 검증한다 — 한 장에 12.3ms다. `set_layout`은 레이아웃을
# 통째로 갈아 끼워 3.1ms에 끝낸다(4배). 실측으로 「1. 주간보고」 리런 1.15초 중
# 0.29초 · 「9. 사이트·앱 지표」 0.73초 중 0.21초가 여기였다.
#
# 빠른 길의 전제는 **기본 템플릿을 새 레이아웃에 안 싣는 것**이다(그걸 같이 넘기면
# 12ms가 그대로 돌아온다). 안 실어도 직렬화할 때 기본값이 다시 붙어서 결과 JSON이
# 같지만, **앱이 템플릿을 직접 고르면 그 순간 깨진다** — 아래 둘이 같이 못 박는다.


def _fig_json(fig):
    import json
    return json.loads(fig.to_json())


@case
def t_set_layout_matches_update_layout():
    """`set_layout`과 `update_layout`의 결과 figure가 한 글자도 다르지 않아야 한다."""
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    import weekly_report as W

    x, y = list(range(12)), [i * 1.5 for i in range(12)]

    def plain():
        f = go.Figure()
        f.add_trace(go.Scatter(x=x, y=y, name="올해"))
        return f

    def sub():
        # 패널 제목은 layout.annotations 로 들어간다 — 갈아 끼울 때 잃기 쉬운 자리다
        f = make_subplots(rows=2, cols=2, subplot_titles=["가", "나", "다", "라"],
                          shared_xaxes=True, vertical_spacing=0.14)
        for r in (1, 2):
            for c in (1, 2):
                f.add_trace(go.Bar(x=x, y=y), row=r, col=c)
        return f

    def _slay():
        l = W.base_layout(430, title="제목")
        l.pop("xaxis", None)
        l.pop("yaxis", None)
        l["showlegend"] = True
        return l

    cases = [("단일", plain, lambda: W.base_layout(320, ysuffix="%", title="제목")),
             ("서브플롯", sub, _slay)]
    for nm, mk, mklay in cases:
        a, b = mk(), mk()
        a.update_layout(**mklay())
        W.set_layout(b, mklay())
        ja, jb = _fig_json(a), _fig_json(b)
        assert ja["layout"] == jb["layout"], f"{nm}: layout이 달라요"
        assert ja["data"] == jb["data"], f"{nm}: data가 달라요"

    # 여러 번 겹쳐 입히는 길 (base → 추가 조정) 도 같아야 한다
    a, b = plain(), plain()
    a.update_layout(**W.base_layout(320, title="제목"))
    a.update_layout(yaxis2=dict(overlaying="y", side="right"))
    a.update_layout(legend=dict(x=0.5))
    W.set_layout(b, W.base_layout(320, title="제목"))
    W.set_layout(b, dict(yaxis2=dict(overlaying="y", side="right")))
    W.set_layout(b, dict(legend=dict(x=0.5)))
    assert _fig_json(a)["layout"] == _fig_json(b)["layout"], "겹쳐 입히면 달라져요"

    # 레이아웃을 입힌 뒤 축을 손봐도(`update_xaxes`) 같아야 한다 — stacked_panels 모양
    a, b = sub(), sub()
    a.update_layout(**_slay())
    a.update_xaxes(linecolor="#e2e8f0", tickfont=dict(size=9))
    W.set_layout(b, _slay())
    b.update_xaxes(linecolor="#e2e8f0", tickfont=dict(size=9))
    assert _fig_json(a)["layout"] == _fig_json(b)["layout"], "축을 손보면 달라져요"

    # 패널 제목이 살아 있어야 한다 (안 살면 위 대조가 둘 다 비어서 통과할 수 있다)
    assert len(_fig_json(b)["layout"].get("annotations", [])) == 4, \
        "패널 제목(annotations)이 사라졌어요"


@case
def t_set_layout_takes_the_fast_path():
    """빠른 길을 타는지 — 느린 길로 되돌리면 잡힌다.

    위 대조 검사만 두면 `set_layout`의 본문을 `fig.update_layout(**lay)` 한 줄로
    바꿔도 **값이 같으니 통과한다**. 느려진 건 값으로 안 보여서, 셋을 같이 본다.

    ③이 핵심이다 — `go.Figure()`가 꽂아 둔 기본 템플릿을 새 레이아웃에 같이 넘기면
    그 큰 dict를 다시 검증하느라 12ms가 그대로 돌아오는데, **결과 JSON은 똑같아서**
    대조로는 영영 안 잡힌다. 실제로 그렇게 짜 보고서야 드러났다.
    """
    import time
    import plotly.graph_objects as go
    import chart_layout as C
    import weekly_report as W

    def mk():
        f = go.Figure()
        f.add_trace(go.Scatter(x=list(range(30)), y=list(range(30))))
        return f

    # ① 병합 API를 안 불렀는지
    ul = []
    _oul = go.Figure.update_layout
    go.Figure.update_layout = lambda self, *a, **k: (ul.append(1), _oul(self, *a, **k))[1]
    # ③ go.Layout 에 넘기는 dict에 템플릿이 없는지 — lay_merge 결과가 곧 그 dict다
    merged = []
    _omg = C.lay_merge

    def _mg(a, b):
        out = _omg(a, b)
        merged.append(out)
        return out

    C.lay_merge = _mg
    try:
        C.set_layout(mk(), W.base_layout(320, title="제목"))
    finally:
        go.Figure.update_layout = _oul
        C.lay_merge = _omg

    assert not ul, "set_layout이 update_layout을 불렀어요 (느린 길)"
    assert merged, "lay_merge를 안 거쳤어요 — 검사가 헛돌고 있어요"
    assert "template" not in merged[-1], \
        "기본 템플릿을 go.Layout에 넘겼어요 — 그러면 12ms가 그대로 돌아와요"

    # ② 실제로 더 빠른지. 실측은 12.3ms 대 3.1ms(4배)라 2배 여유를 두고 본다.
    def _best(fn, n=12):
        fn()
        return min((lambda: (lambda t0: (fn(), time.perf_counter() - t0)[1])(
            time.perf_counter()))() for _ in range(n))

    slow = _best(lambda: mk().update_layout(**W.base_layout(320, title="제목")))
    fast = _best(lambda: C.set_layout(mk(), W.base_layout(320, title="제목")))
    assert fast * 2 < slow, (f"빠르지 않아요 — set_layout {fast * 1000:.1f}ms vs "
                             f"update_layout {slow * 1000:.1f}ms (2배는 돼야 해요)")


@case
def t_charts_do_not_call_update_layout_directly():
    """두 앱에 `update_layout`이 남아 있으면 안 된다 — 느린 길이 다시 스며든다.

    그리고 **템플릿을 직접 고르지 않는 것**이 `set_layout`의 전제다. 누가
    `template=`을 쓰기 시작하면 `set_layout`이 그걸 버려서 화면이 조용히 달라진다.
    """
    # **줄 문자열이 아니라 호출을 본다.** 줄로 보면 임포트 위의 설명 주석
    # (`fig.update_layout(**lay)`는 느리다…)에 걸려 오탐이 나고, 그걸 피하려고 예외를
    # 두면 그 예외가 진짜 호출을 통과시키는 구멍이 된다.
    def _calls(tree, attr):
        return [n.lineno for n in ast.walk(tree)
                if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                and n.func.attr == attr]

    def _kwargs(tree, name):
        return [n.lineno for n in ast.walk(tree) if isinstance(n, ast.Call)
                for k in n.keywords if k.arg == name]

    for name in APPS:
        tree = _tree(name)
        bad = _calls(tree, "update_layout")
        assert not bad, f"{name}: update_layout 직접 호출이 남았어요 — 줄 {bad}"
        tpl = _kwargs(tree, "template")
        assert not tpl, f"{name}: 플로틀리 템플릿을 직접 골랐어요 — 줄 {tpl}"
    # 공용 모듈엔 폴백 하나만 (`_props`가 없어지는 미래 판 대비)
    _cl = ast.parse((ROOT / "chart_layout.py").read_text(encoding="utf-8"))
    _n = _calls(_cl, "update_layout")
    assert len(_n) == 1, f"chart_layout.py의 update_layout이 {len(_n)}곳이에요 — 폴백 하나여야 해요"


# ──────────────────────────────────────────────────────────────────────
# 페이지 리포트 — 몽키패치를 걷어낸 자리
# ──────────────────────────────────────────────────────────────────────

@case
def t_streamlit_is_not_monkeypatched():
    """**`st`의 함수를 갈아 끼우지 않는다.**

    예전엔 페이지 리포트를 모으려고 `st.plotly_chart`·`st.dataframe`을 패치했다.
    `st`는 **프로세스 공유**라 동시 접속한 두 세션이 패치 블록을 같이 지나가면 뒤
    세션이 앞 세션의 래퍼를 '원본'으로 집어 가고, 그 뒤로 래퍼가 자기를 호출해
    `RecursionError`로 **앱 전체가 내려갔다.** 운영에서 두 번 났다 — 두 번째는
    «그걸 고친 배포»에서 났는데, 고친 코드가 배포가 남긴 옛 래퍼를 못 알아봤다.
    리런하면 프로세스가 새로 떠서 살아나니 증상이 '가끔 죽는다'로만 보였다.

    지금은 가로채기를 그만두고 호출부를 모았다(`chart()`·`table()`). 패치가 없으니
    세션 사이 경합도, 배포가 남긴 옛 래퍼도 없다 — **다시 들이지 않는 게 이 검사다.**
    """
    import streamlit as _st
    for mod in ("send_perf_dashboard.py", "weekly_report.py"):
        tree = ast.parse((ROOT / mod).read_text(encoding="utf-8"))
        bad = [(mod, n.lineno, t.attr) for n in ast.walk(tree)
               if isinstance(n, ast.Assign) for t in n.targets
               if isinstance(t, ast.Attribute) and isinstance(t.value, ast.Name)
               and t.value.id == "st"]
        assert not bad, f"`st`에 대입하고 있어요(몽키패치): {bad}"

    sys.path.insert(0, str(ROOT / "tests"))
    from streamlit.testing.v1 import AppTest
    from smoke_pages import group_of, synth_store
    at = AppTest.from_file(str(ROOT / "send_perf_dashboard.py"), default_timeout=600)
    at.session_state["camp_store"] = synth_store(weeks=40, seed=3)
    at.run()
    assert not at.exception, at.exception[0].value
    at.run()                                   # 같은 프로세스에서 한 번 더
    assert not at.exception, at.exception[0].value
    for nm in ("plotly_chart", "dataframe"):
        assert not hasattr(getattr(_st, nm), "_lf_orig"), f"st.{nm}이 래퍼로 덮였어요"

    # **기능을 죽이고 패치만 없애면 안 된다** — 리포트가 여전히 차야 한다.
    # **종류별로 본다.** 「차 있는가」만 보면 차트 기록이 끊겨도 표가 남아 통과한다
    # (실제로 그 주입이 안 잡혔다).
    at.sidebar.radio[0].set_value(group_of("종합 요약"))
    at.run()
    assert not at.exception, at.exception[0].value
    kinds = collections.Counter(k for k, _ in (at.session_state.get("_page_report") or []))
    assert kinds.get("fig"), f"차트를 하나도 못 모았어요 — {dict(kinds)}"
    assert kinds.get("table"), f"표를 하나도 못 모았어요 — {dict(kinds)}"


@case
def t_charts_go_through_the_chart_helper():
    """차트는 **`chart()`를 거친다** — `st.plotly_chart`를 직접 부르지 않는다.

    직접 부르면 **그 차트만 리포트에서 조용히 빠진다.** 화면은 멀쩡히 떠서 눈으로도
    코드 리뷰로도 안 잡힌다 — 가로채기를 그만둔 대가로 생긴 자리라 규칙으로 막는다.
    """
    tree = ast.parse((ROOT / "send_perf_dashboard.py").read_text(encoding="utf-8"))
    fn = next(n for n in tree.body
              if isinstance(n, ast.FunctionDef) and n.name == "chart")
    ok = set(range(fn.lineno, fn.end_lineno + 1))
    bad = [(n.lineno, ast.unparse(n.func)) for n in ast.walk(tree)
           if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
           and n.func.attr == "plotly_chart" and n.lineno not in ok]
    assert not bad, f"`chart()`를 안 거친 차트 호출: {bad}"


@case
def t_a_leftover_wrapper_does_not_double_record():
    """옛 배포가 남긴 래퍼가 `st.plotly_chart`에 앉아 있어도 **두 번 안 담는다.**

    Streamlit Cloud는 배포할 때 프로세스를 재시작하지 않고 스크립트만 갈아 끼우기도
    한다. 패치를 걷어낸 코드를 올려도 **옛 래퍼는 모듈 속성에 그대로 남는다** — 그
    래퍼도 리포트에 담으니 차트가 두 벌이 된다. 리포트를 받아 보면 같은 차트가 두
    번 나오는데 화면은 멀쩡해서, 증상이 '리포트가 좀 이상하네'로만 보인다.
    (재귀로 앱이 내려간 두 번째 사고가 정확히 «프로세스가 안 죽는다»는 이 자리였다.)

    `chart()`·`table()`이 `st`가 아니라 `st._main`에 넘기므로 모듈 속성을 무엇으로
    덮어써도 안 걸린다.
    """
    sys.path.insert(0, str(ROOT / "tests"))
    import streamlit as _st
    from streamlit.testing.v1 import AppTest
    from smoke_pages import group_of, synth_store

    _seen = []
    _save_f, _save_d = _st.plotly_chart, _st.dataframe

    def _legacy(kind, orig):
        def _cap(*a, **k):                     # 옛 배포의 가로채기 흉내
            _seen.append(kind)
            try:
                _st.session_state.setdefault("_page_report", []).append((kind, a[0]))
            except Exception:                  # noqa: BLE001
                pass
            return orig(*a, **k)
        return _cap

    _st.plotly_chart = _legacy("fig", _save_f)
    _st.dataframe = _legacy("table", _save_d)
    try:
        at = AppTest.from_file(str(ROOT / "send_perf_dashboard.py"), default_timeout=600)
        at.session_state["camp_store"] = synth_store(weeks=40, seed=3)
        at.run()
        at.sidebar.radio[0].set_value(group_of("종합 요약"))
        at.run()
        assert not at.exception, at.exception[0].value
    finally:
        _st.plotly_chart, _st.dataframe = _save_f, _save_d

    assert not _seen, (f"옛 래퍼를 그대로 탔어요 — {collections.Counter(_seen)}. "
                       "`st`가 아니라 `st._main`에 넘겨야 해요.")
    kinds = collections.Counter(k for k, _ in (at.session_state.get("_page_report") or []))
    assert kinds.get("fig") and kinds.get("table"), f"리포트가 안 찼어요 — {dict(kinds)}"


@case
def t_column_charts_keep_their_container():
    """컬럼 안 차트는 **컬럼 안에** 남아야 한다.

    `col.plotly_chart(fig)`를 `chart(fig)`로 바꾸면 그 차트가 **컬럼 밖 본문**에
    그려진다. 예외가 안 나고 레이아웃만 달라져 화면을 안 보면 안 드러난다 — 옮길 때
    `_to=`를 빠뜨리기 딱 좋은 자리다. 소스에서 `_to=`를 세지 않고 **화면에서** 본다.
    """
    sys.path.insert(0, str(ROOT / "tests"))
    from streamlit.testing.v1 import AppTest
    from smoke_pages import group_of, synth_store

    at = AppTest.from_file(str(ROOT / "send_perf_dashboard.py"), default_timeout=600)
    at.session_state["camp_store"] = synth_store(weeks=60, seed=5)
    at.run()
    at.sidebar.radio[0].set_value(group_of("전년 동요일 비교"))
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    subs[0].set_value("전년 동요일 비교")
    at.run()
    assert not at.exception, at.exception[0].value

    def _under_columns(block, inside=False):
        n = 0
        for ch in block.children.values():
            t = getattr(ch, "type", "")
            if t == "plotly_chart":
                n += int(inside)
            elif hasattr(ch, "children"):
                n += _under_columns(ch, inside or t == "column")
        return n

    total, inside = len(at.get("plotly_chart")), _under_columns(at.main)
    assert total >= 2, f"이 화면에 차트가 {total}개뿐이라 규칙을 못 봐요"
    assert inside >= 2, (f"컬럼 안에 있어야 할 차트가 본문으로 빠졌어요 — "
                         f"전체 {total}개 중 컬럼 안 {inside}개")


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
