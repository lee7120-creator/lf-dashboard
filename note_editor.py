"""보고란 리치 에디터 — 발송성과·주간보고 공용.

`streamlit-quill`(PyPI 최종 배포 **2020-09-27** · Quill **1.3.7**)에서
`streamlit-lexical-extended`로 갈았다. 바뀐 것:

| | streamlit-quill | streamlit-lexical-extended |
|---|---|---|
| 배포 | 2020-09-27 | 2026-07-29 |
| 마운트 | iframe (Components v1) | **인라인 (Components v2)** |
| 글꼴 | iframe 안 Helvetica + bootstrap 197KB | 앱 테마(Pretendard) 그대로 |
| 입력 | **타이핑 한 글자마다 리런** | `debounce`(0.5초) |
| 높이 | 고정 | `min_height`로 내용 따라 늘어남 |
| 서식 | 크기·색·배경·정렬 | **제목·인용·표·되돌리기** |
| 저장 꼴 | HTML | 마크다운 |

**잃은 것** — 글자색·형광펜·정렬·글자크기 버튼. 보고서가 실제로 쓰는 색은
△ 빨강 / + 초록인데 그건 `_note_render`가 자동으로 입히므로 영향이 없다.
글자크기 자리는 제목(`block_type`)이 대신한다.

**모듈을 못 불러오면 평범한 텍스트 박스로 내려간다.** 조용히 내려가면
'왜 에디터가 안 뜨지'만 남으므로 화면에 이유를 밝힌다. 이 모듈은 Python 3.11
이상을 요구한다 — `streamlit>=1.58`이 이미 3.10을 깔고 있어 턱이 한 칸이다.
"""

from __future__ import annotations

import streamlit as st

from note_markdown import to_markdown

try:
    from streamlit_lexical_extended import streamlit_lexical_extended as _lexical
    HAS_LEXICAL = True
except Exception:                                         # noqa: BLE001
    HAS_LEXICAL = False

__all__ = ["HAS_LEXICAL", "NOTE_TOOLBAR", "note_editor", "note_editor_reset"]

# 보고란에 실제로 쓰는 것만 올린다. 전체(None)로 두면 쓰지 않는 버튼까지 떠
# 좁은 칼럼에서 줄바꿈이 생긴다.
NOTE_TOOLBAR = [
    "undo", "redo", "block_type",
    "bold", "italic", "underline", "strikethrough",
    "bullet_list", "numbered_list", "quote", "table",
]

_NO_MODULE = (
    "리치 에디터 모듈을 못 불러와 기본 입력창으로 띄웠어요. 서식 버튼은 안 보이지만 "
    "마크다운 문법(`- 불릿` · `**굵게**` · `## 제목`)은 그대로 적용돼요."
)


def _seed_key(key: str) -> str:
    return f"_noteseed_{key}"


def _unwrap(v):
    """래퍼가 str·{"value": str}·객체 중 아무거나 줄 수 있어 한 겹 벗긴다."""
    if isinstance(v, str):
        return v
    inner = v.get("value") if isinstance(v, dict) else getattr(v, "value", None)
    return inner if isinstance(inner, str) else None


def _current(key: str, returned):
    """에디터의 현재 내용.

    **래퍼의 반환값을 먼저 본다.** 세션값(`st.session_state[key]`)은 프런트가
    마지막으로 보낸 값이라, 값을 다시 심은 리런에서는 아직 «옛» 내용이다 —
    그쪽을 먼저 보면 「자동 생성」 직후에 옛 글을 저장하게 된다. 래퍼는 두 경우를
    이미 갈라 준다: 다시 심었으면 심은 값을, 아니면 지금 편집 중인 값을 돌려준다.
    세션값은 반환이 비었을 때의 보조 수단으로만 쓴다.
    """
    val = _unwrap(returned)
    if val is not None:
        return val
    val = _unwrap(st.session_state.get(key))
    return val if val is not None else ""


def note_editor(value, key: str, min_height: int = 260, placeholder: str = "") -> str:
    """저장된 보고란 문구를 띄우고 **지금 편집 중인 마크다운**을 돌려준다.

    `value`는 HTML(옛 Quill 저장분)이든 플레인 텍스트(자동·AI 생성)든 상관없다 —
    `note_markdown.to_markdown`이 가른다.

    **값은 처음 한 번만 넣는다.** 리런마다 `value=`를 다시 주면 래퍼가 그걸
    '바깥에서 바뀐 값'으로 보고 에디터를 되돌려, 방금 친 글이 사라진다. 저장소가
    실제로 바뀐 때(첫 진입·자동 생성·AI 생성)만 다시 심는다
    (`t_typing_survives_a_rerun`).
    """
    md = to_markdown(value)
    if not HAS_LEXICAL:
        st.caption(_NO_MODULE)
        return st.text_area("내용", md, key=f"ta_{key}", height=max(120, min_height),
                            label_visibility="collapsed")

    seed = _seed_key(key)
    if st.session_state.get(seed) != md:
        st.session_state[seed] = md
        ret = _lexical(value=md, placeholder=placeholder, min_height=min_height,
                       key=key, toolbar=NOTE_TOOLBAR)
    else:
        ret = _lexical(value=None, placeholder=placeholder, min_height=min_height,
                       key=key, toolbar=NOTE_TOOLBAR)
    return _current(key, ret)


def note_editor_reset(key: str) -> None:
    """편집을 닫을 때 부른다 — 다음에 열면 저장소 값을 다시 심는다.

    안 부르면 「보기」로 닫아도 에디터가 안 저장한 글을 그대로 들고 있어,
    다시 열었을 때 취소가 안 먹은 것처럼 보인다
    (`t_closing_the_editor_discards_unsaved_text`).
    """
    st.session_state.pop(_seed_key(key), None)
