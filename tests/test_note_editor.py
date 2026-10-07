"""보고란 에디터 테스트 — 모듈 교체(streamlit-quill → streamlit-lexical-extended).

에디터가 HTML을 쓰다가 마크다운을 쓰게 되면서, 이미 저장된 글을 **잃지 않고**
읽어 들이는 게 이 교체의 전부다. 여기서 조용히 틀리는 방식이 정해져 있다.

| 겉으로 보이는 것 | 실제로 일어난 일 |
|---|---|
| 「편집을 열었더니 `<p>` 가 보인다」 | HTML을 마크다운 에디터에 그냥 넣었다 |
| 「계층이 한 단계씩 사라진다」 | 다시 읽을 때 중첩 불릿의 들여쓰기를 strip했다 |
| 「표가 안 그려진다」 | 표 줄에 강제 개행(공백 2칸)을 붙였다 |
| 「방금 친 글이 되돌아간다」 | 리런마다 `value=`를 다시 넣었다 |
| 「취소했는데 글이 남아 있다」 | 편집을 닫을 때 에디터 상태를 안 비웠다 |

전부 **화면은 멀쩡히 뜬다.** 렌더 스모크로는 안 잡힌다.

로컬 실행:
    python tests/test_note_editor.py
"""
import html
import pathlib
import re
import sys
import tempfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import note_markdown as NM  # noqa: E402

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# ── 실제로 저장돼 있는 두 가지 꼴 ────────────────────────────────────
# ① Quill 1.3.7이 내보낸 HTML. 중첩은 <ul> 중첩이 아니라 평면 ql-indent-N이다.
QUILL_HTML = (
    '<p>9월 4주차 요약</p>'
    '<ul><li>발송량 <strong>증가</strong></li>'
    '<li class="ql-indent-1">발송 1,234건, 전주비 +5.2%</li>'
    '<li>효율 <em>개선</em></li>'
    '<li class="ql-indent-1">CTR △0.3%p</li></ul>'
    '<p><span style="color: rgb(230, 0, 0);">빨간 글씨</span> F&amp;C 브랜드</p>'
)
# ② 「자동 생성」·「AI 생성」이 쓴 플레인 텍스트. 프롬프트가 '-'와 'ㄴ'를 지시한다.
AI_PLAIN = (
    "- 발송량 증가\n"
    "ㄴ 발송 1,234건, 전주비 +5.2%\n"
    "- 효율 개선\n"
    "ㄴ CTR △0.3%p"
)


def _visible(text):
    """마크다운·HTML 문법을 뺀 가시 글자만. 글자가 사라졌는지 보는 데 쓴다.

    엔티티는 **먼저 푼다** — 원본의 `&amp;`와 변환본의 `&`를 그냥 맞대면
    없어지지도 않은 `a`·`m`·`;`가 사라진 글자로 잡힌다.
    """
    t = re.sub(r"<[^>]+>", "", str(text))
    t = html.unescape(t)
    t = re.sub(r"\]\([^)]*\)", "", t)              # 링크 주소
    return re.sub(r"[\s*_~`#>|\[\].\-]", "", t)


@case
def t_quill_html_becomes_markdown():
    """Quill HTML이 마크다운으로 읽힌다 — 날 태그가 남으면 에디터에 그대로 보인다."""
    md = NM.to_markdown(QUILL_HTML)
    assert "<" not in md and ">" not in md, md
    assert "- 발송량 **증가**" in md, md
    assert "    - 발송 1,234건, 전주비 +5.2%" in md, md   # ql-indent-1 → 중첩(4칸)
    assert "*개선*" in md, md


@case
def t_no_text_is_lost():
    """변환이 글자를 지우지 않는다.

    색·정렬처럼 **서식**은 잃어도 되지만 **글자**는 안 된다. 'ㄴ'만은 예외 —
    계층 표식이라 진짜 중첩 불릿으로 올라가며 글자 자리를 떠난다.
    """
    for name, src in (("HTML", QUILL_HTML), ("플레인", AI_PLAIN)):
        want = _visible(src).replace("ㄴ", "")
        got = _visible(NM.to_markdown(src))
        missing = [c for c in want if c not in got]
        assert not missing, f"{name}: 사라진 글자 {missing}\n  원본 {src!r}\n  변환 {got!r}"


@case
def t_entities_are_unescaped_not_doubled():
    """`F&amp;C`는 `F&C`로 읽는다 — 실데이터에 & 들어간 브랜드가 있다."""
    md = NM.to_markdown(QUILL_HTML)
    assert "F&C" in md, md
    assert "&amp;" not in md, md


@case
def t_ai_submarker_becomes_a_nested_bullet():
    """'ㄴ' 줄은 중첩 불릿이 된다.

    그냥 두면 마크다운이 앞 불릿의 '이어진 줄'로 흡수해 한 줄로 붙여 버린다.
    AI 프롬프트는 못 고친다 — 출력 서식을 지시하는 문자열이다.
    """
    md = NM.to_markdown(AI_PLAIN)
    assert md.splitlines()[1] == "    - 발송 1,234건, 전주비 +5.2%", md.splitlines()


@case
def t_nesting_survives_a_second_read():
    """다시 읽어도 계층이 그대로다.

    `to_markdown`은 편집을 **열 때마다** 돈다. 들여쓰기를 strip하면 열고 닫기만
    해도 한 단계씩 평평해져, 아무도 안 고친 글이 조용히 망가진다.
    """
    for src in (QUILL_HTML, AI_PLAIN, "- 가\n    - 가-1\n        - 가-1-1\n- 나"):
        a = NM.to_markdown(src)
        b = NM.to_markdown(a)
        assert a == b, f"두 번째 읽기에서 달라졌어요\n  1회 {a!r}\n  2회 {b!r}"


@case
def t_markdown_passes_through_unchanged():
    """이미 마크다운인 글은 손대지 않는다 — 특히 표.

    같은 문단 안 개행을 지키려고 앞 줄에 공백 2칸을 붙이는데, 그걸 표 줄에
    붙이면 구분선이 깨져 표가 통째로 안 그려진다.
    """
    src = ("## 제목\n\n- 가\n- 나\n\n"
           "| 지표 | 값 |\n| --- | --- |\n| 발송 | 1,234 |\n"
           "위 표는 전주 기준이에요.")
    assert NM.to_markdown(src) == src, repr(NM.to_markdown(src))


@case
def t_plain_line_breaks_are_kept():
    """평문 줄바꿈이 살아 있다 — 마크다운은 그냥 개행을 공백으로 먹는다."""
    md = NM.to_markdown("첫째 줄\n둘째 줄")
    assert md == "첫째 줄  \n둘째 줄", repr(md)


@case
def t_html_table_becomes_a_markdown_table():
    """붙여 넣은 표는 마크다운 표로 — 구분선이 없으면 표로 안 읽힌다."""
    md = NM.to_markdown("<table><tr><th>지표</th><th>값</th></tr>"
                        "<tr><td>발송</td><td>1,234</td></tr></table>")
    assert md.splitlines() == ["| 지표 | 값 |", "| --- | --- |", "| 발송 | 1,234 |"], md


@case
def t_plain_text_is_not_mistaken_for_html():
    """부등호가 든 평문을 HTML로 오인하지 않는다."""
    for s in (AI_PLAIN, "매출 < 목표", "a > b 인 경우", "- 3 < 5"):
        assert not NM.looks_like_html(s), s
    for s in (QUILL_HTML, "<p>가</p>", "<ul><li>가</li></ul>"):
        assert NM.looks_like_html(s), s


@case
def t_empty_input_is_empty_output():
    """빈 값에 군더더기를 만들지 않는다."""
    for s in (None, "", "   ", "\n\n"):
        assert NM.to_markdown(s) == "", repr(NM.to_markdown(s))


# ══════════════════════════════════════════════════════════════════
# 에디터 마운트 — 값 주입 타이밍
#
# 리런마다 `value=`를 다시 주면 래퍼가 그걸 '바깥에서 바뀐 값'으로 보고 에디터를
# 되돌린다(방금 친 글이 사라진다). 그래서 **저장소가 실제로 바뀐 때만** 심는다.
#
# 컴포넌트를 기록용 가짜로 바꿔 «무엇을 넘겼는지»를 직접 본다. AppTest는 진짜
# 컴포넌트의 입력을 흉내 내지 못해서(프런트가 리런마다 값을 다시 보내 주는 일을
# 대신해 주지 않는다) 실모듈로는 이 규칙을 반증할 수가 없다. 대신 아래
# `t_real_module_receives_the_markdown`이 실모듈이 뜨는지를 따로 본다.
# ══════════════════════════════════════════════════════════════════
_HARNESS = """
import sys
sys.path.insert(0, %(root)r)
import streamlit as st
import note_editor as NE

REC = st.session_state.setdefault("rec", [])

def _fake(value=None, placeholder="", min_height=0, key=None, toolbar=None):
    REC.append(value)
    st.session_state.setdefault(key, {"value": value if value is not None else ""})
    return st.session_state[key].get("value", "")

NE._lexical = _fake
NE.HAS_LEXICAL = True

STORE = st.session_state.setdefault("store", {"n": %(src)r})
out = NE.note_editor(STORE["n"], key="ed", min_height=200)
st.text("OUT=" + repr(out))
if st.button("닫기"):
    NE.note_editor_reset("ed")
if st.button("자동생성"):
    STORE["n"] = "- 새로 만든 글"
    NE.note_editor_reset("ed")
"""


def _harness(src=QUILL_HTML):
    from streamlit.testing.v1 import AppTest
    f = pathlib.Path(tempfile.mkdtemp()) / "harness.py"
    f.write_text(_HARNESS % {"root": str(ROOT), "src": src}, encoding="utf-8")
    return AppTest.from_file(str(f), default_timeout=90).run()


def _rec(at):
    return list(at.session_state["rec"])


def _out(at):
    return [t.value for t in at.main.text if t.value.startswith("OUT=")][0]


@case
def t_editor_is_seeded_with_the_converted_markdown():
    """편집을 열면 변환된 «마크다운»이 들어간다 — 날 HTML이 들어가면 태그가 보인다."""
    at = _harness()
    assert not at.exception, [e.value for e in at.exception]
    assert _rec(at) == [NM.to_markdown(QUILL_HTML)], _rec(at)


@case
def t_typing_survives_a_rerun():
    """리런에서 값을 다시 심지 않는다 — 심으면 방금 친 글이 되돌아간다."""
    at = _harness()
    at.run()                                        # 아무것도 안 바뀐 리런
    assert _rec(at)[1] is None, _rec(at)

    at.session_state["ed"] = {"value": "- 사용자가 친 글"}
    at.run()
    assert _rec(at)[2] is None, _rec(at)
    assert "사용자가 친 글" in _out(at), _out(at)


@case
def t_closing_the_editor_discards_unsaved_text():
    """「보기」로 닫으면 안 저장한 글을 버린다 — 다시 열 때 저장소 값을 심는다."""
    at = _harness()
    at.session_state["ed"] = {"value": "- 저장 안 한 글"}
    at.run()
    at.button[0].click().run()                      # 「닫기」 → note_editor_reset
    at.run()
    assert _rec(at)[-1] == NM.to_markdown(QUILL_HTML), _rec(at)


@case
def t_regenerated_text_reaches_the_editor():
    """「자동 생성」이 저장소를 바꾸면 에디터에도 새 글이 들어간다."""
    at = _harness()
    at.run()
    at.button[1].click().run()                      # 「자동생성」 → 저장소 교체 + reset
    at.run()
    assert _rec(at)[-1] == "- 새로 만든 글", _rec(at)


@case
def t_real_module_receives_the_markdown():
    """진짜 streamlit-lexical-extended로도 변환된 마크다운이 들어간다.

    가짜로만 보면 '모듈이 실제로 뜨는가'를 못 본다 — 임포트·마운트가 깨지면
    보고란이 통째로 비는데 증상은 그냥 **빈 화면**이다.
    """
    from streamlit.testing.v1 import AppTest

    import note_editor as NE
    if not NE.HAS_LEXICAL:
        raise AssertionError("streamlit-lexical-extended를 못 불러왔어요 "
                             "(requirements.txt · Python 3.11+ 확인)")
    f = pathlib.Path(tempfile.mkdtemp()) / "real.py"
    f.write_text(
        "import sys\n"
        f"sys.path.insert(0, {str(ROOT)!r})\n"
        "import streamlit as st\n"
        "from note_editor import note_editor\n"
        f"st.text('OUT=' + repr(note_editor({QUILL_HTML!r}, key='ed', min_height=200)))\n",
        encoding="utf-8")
    at = AppTest.from_file(str(f), default_timeout=90).run()
    assert not at.exception, [e.value for e in at.exception]
    raw = at.session_state["ed"]
    got = raw.get("value") if isinstance(raw, dict) else getattr(raw, "value", raw)
    assert got == NM.to_markdown(QUILL_HTML), repr(got)


@case
def t_reseeded_value_wins_over_the_stale_session():
    """값을 다시 심은 리런에서는 «래퍼가 준 값»을 쓴다.

    세션값은 프런트가 마지막으로 보낸 값이라 그 시점엔 아직 옛 글이다. 그쪽을
    먼저 보면 「자동 생성」 직후 저장이 옛 글을 되살린다. 앞 하네스는 가짜가
    세션값을 그대로 돌려줘 둘이 같아지므로 이 규칙을 반증하지 못한다 —
    여기서는 **일부러 다르게** 돌려준다.
    """
    from streamlit.testing.v1 import AppTest
    f = pathlib.Path(tempfile.mkdtemp()) / "stale.py"
    f.write_text(
        "import sys\n"
        f"sys.path.insert(0, {str(ROOT)!r})\n"
        "import streamlit as st\n"
        "import note_editor as NE\n"
        "st.session_state['ed'] = {'value': '옛 글(프런트가 아직 안 보낸 상태)'}\n"
        "NE._lexical = lambda **k: '새로 심은 글'\n"
        "NE.HAS_LEXICAL = True\n"
        "st.text('OUT=' + repr(NE.note_editor('- 아무거나', key='ed')))\n",
        encoding="utf-8")
    at = AppTest.from_file(str(f), default_timeout=90).run()
    assert not at.exception, [e.value for e in at.exception]
    out = [t.value for t in at.main.text if t.value.startswith("OUT=")][0]
    assert "새로 심은 글" in out, out


@case
def t_nesting_uses_four_spaces_not_two():
    """중첩 한 단계는 **공백 4칸**이다.

    Lexical의 마크다운 임포터는 4칸·탭일 때만 중첩으로 읽는다. 2칸으로 내보내면
    에디터에서 계층이 통째로 평평해지고, 그 상태로 저장하면 중첩이 **영영 사라진다**.
    표시 쪽(Streamlit 마크다운)은 2칸도 중첩으로 읽어 주기 때문에, 화면의 글 상자와
    에디터가 **서로 다르게** 보인다 — 눈으로 둘을 나란히 놓고서야 드러났다.

    근거(브라우저에서 중첩 `<ul>` 개수로 확인): 2칸 0 · 3칸 0 · 4칸 1 · 탭 1.
    """
    md = NM.to_markdown(AI_PLAIN)
    sub = [l for l in md.splitlines() if l.lstrip().startswith("- ") and l != l.lstrip()]
    assert sub, md
    for line in sub:
        ind = len(line) - len(line.lstrip(" "))
        assert ind % 4 == 0, f"중첩 들여쓰기가 4칸 배수가 아니에요: {ind}칸 — {line!r}"

    deep = NM.to_markdown("- 가\n    - 나\n        - 다")
    assert [len(l) - len(l.lstrip(" ")) for l in deep.splitlines()] == [0, 4, 8], deep


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
    print(f"보고란 에디터 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
