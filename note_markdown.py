"""보고란 문구를 마크다운으로 읽어 들이는 변환기 (발송성과·주간보고 공용).

에디터를 streamlit-quill(HTML)에서 streamlit-lexical-extended(마크다운)로 바꾸면서
필요해졌다. 저장소에는 두 가지 꼴이 섞여 있다.

  · **HTML** — Quill 에디터로 직접 고쳐 저장한 글 (`<p>`·`<ul><li>`·`<strong>` …)
  · **플레인 텍스트** — 「자동 생성」·「AI 생성」이 쓴 글. AI 프롬프트가 '- '(요약)과
    'ㄴ '(세부)의 계층형 불릿을 지시하므로 사실상 마크다운에 가깝다.

둘 다 마크다운으로 읽어야 에디터가 원문을 그대로 보여 준다. 한쪽만 다루면 나머지
꼴이 에디터에서 `<p>` 같은 날 태그로 보이거나 줄이 통째로 뭉개진다.

**두 앱이 같은 변환기를 쓴다.** 금액 파서(`_num`·`_promo_num`)를 양쪽에 따로 둬서
한쪽만 고치면 조용히 갈리던 전례가 있어, 여기는 처음부터 모듈 하나로 둔다.

**이스케이프는 하지 않는다.** 저장된 글의 지배적인 꼴이 이미 '- ' 불릿이라,
마크다운 특수문자를 막으면 정작 불릿으로 읽혀야 하는 줄이 `\\- `로 보인다.
대신 **글자를 잃지 않는 것**을 규칙으로 두고 검사가 그걸 본다
(`tests/test_note_editor.py`의 `t_no_text_is_lost`).
"""

from __future__ import annotations

import re
from html.parser import HTMLParser

__all__ = ["looks_like_html", "html_to_markdown", "plain_to_markdown", "to_markdown"]

# 저장된 글이 HTML인지 가르는 힌트. Quill이 실제로 내보내는 태그만 본다 —
# 한글 산문에 '<p>' 같은 꼴이 글자로 들어올 일은 없다.
_HTML_HINT = re.compile(
    r"</?(?:p|div|ul|ol|li|br|h[1-6]|strong|b|em|i|u|s|strike|del|a|span"
    r"|blockquote|pre|code|table|tr|td|th)\b[^>]*>",
    re.I,
)

# 인라인 서식 → 마크다운 울타리. 밑줄(u)은 마크다운에 대응이 없어 글자만 남긴다.
_WRAP = {
    "strong": "**", "b": "**",
    "em": "*", "i": "*",
    "s": "~~", "strike": "~~", "del": "~~",
    "code": "`",
}
_SKIP = {"script", "style", "head", "title"}
_BR = "\x00"            # <br> 자리표시 — flush에서 마크다운 강제 개행으로 바꾼다
_SUBMARK = "\u3134\u2514\u21b3"      # ㄴ · └ · ↳ — AI 프롬프트가 쓰는 '세부' 표식
# 중첩 한 단계 = **공백 4칸**. 2칸이 아니다 — Lexical의 마크다운 임포터는 4칸이나
# 탭일 때만 중첩으로 읽고, 2~3칸은 통째로 평평하게 만든다(브라우저에서 2·3·4칸·탭을
# 띄워 중첩 <ul>이 생기는지로 확인했다: 2칸 0개 · 3칸 0개 · 4칸 1개 · 탭 1개).
# 표시 쪽(Streamlit 마크다운)은 2칸·4칸 둘 다 중첩으로 읽으므로 4칸이 양쪽을 만족한다.
_IND = "    "
# 마크다운 블록 문법으로 시작하는 줄. 이 줄엔 강제 개행(공백 2칸)을 붙이지 않는다.
_BLOCKISH = re.compile(r"^(?:#{1,6}\s|>|\||```|~~~|-{3,}$|\*{3,}$|\d+[.)]\s|[-*+]\s)")


def looks_like_html(text) -> bool:
    """Quill이 저장한 HTML인지. 플레인 텍스트면 False."""
    return bool(_HTML_HINT.search(str(text or "")))


class _ToMarkdown(HTMLParser):
    """Quill 1.3.7이 내보내는 HTML 부분집합을 마크다운 줄 목록으로 바꾼다.

    Quill은 중첩 목록을 `<ul>` 안쪽 중첩이 아니라 **평면 `<li class="ql-indent-N">`**
    으로 낸다. 다른 데서 붙여 넣은 진짜 중첩도 받아야 하니 둘을 더해서 깊이를 센다.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.lines: list[str] = []
        self.buf: list[str] = []
        self.lists: list[dict] = []     # [{"kind": "ul"|"ol", "i": 번호}]
        self.qlindent = 0               # 현재 <li>의 ql-indent-N
        self.skip = 0
        self.pre = 0
        self.heading = 0
        self.quote = 0
        self.href: str | None = None
        self.row: list[str] | None = None   # 표 한 줄을 모으는 칸
        self.tablerows = 0

    # ── 줄 만들기 ───────────────────────────────────────────────
    def _flush(self, prefix: str = "", cont: str = "") -> None:
        """버퍼를 한 줄로 떨어뜨린다. <br>은 마크다운 강제 개행(공백 2칸)으로."""
        raw = "".join(self.buf)
        self.buf = []
        chunks = [c.strip() for c in raw.split(_BR)]
        chunks = [c for c in chunks if c]
        if not chunks:
            return
        self.lines.append(prefix + chunks[0])
        for c in chunks[1:]:
            # 이어지는 줄은 앞 줄 끝에 공백 2칸을 붙여 같은 문단 안 개행으로 만든다
            self.lines[-1] += "  "
            self.lines.append(cont + c)

    def _blank(self) -> None:
        if self.lines and self.lines[-1] != "":
            self.lines.append("")

    # ── 태그 ────────────────────────────────────────────────────
    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in _SKIP:
            self.skip += 1
            return
        if self.skip:
            return
        a = dict(attrs)

        if tag == "br":
            self.buf.append(_BR)
        elif tag in _WRAP:
            self.buf.append(_WRAP[tag])
        elif tag == "a":
            self.href = a.get("href") or ""
            self.buf.append("[")
        elif tag in ("ul", "ol"):
            self.lists.append({"kind": tag, "i": 0})
        elif tag == "li":
            self.qlindent = 0
            m = re.search(r"ql-indent-(\d+)", a.get("class") or "")
            if m:
                self.qlindent = int(m.group(1))
        elif re.fullmatch(r"h[1-6]", tag):
            self._flush()
            self._blank()
            self.heading = int(tag[1])
        elif tag == "blockquote":
            self._flush()
            self._blank()
            self.quote += 1
        elif tag == "pre":
            self._flush()
            self._blank()
            self.pre += 1
            self.lines.append("```")
        elif tag in ("td", "th"):
            self.buf = []
        elif tag == "tr":
            self.row = []
        elif tag == "table":
            self._flush()
            self._blank()
            self.tablerows = 0

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in _SKIP:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return

        if tag in _WRAP:
            self.buf.append(_WRAP[tag])
        elif tag == "a":
            self.buf.append(f"]({self.href})" if self.href else "]")
            self.href = None
        elif tag in ("ul", "ol"):
            if self.lists:
                self.lists.pop()
            if not self.lists:
                self._blank()
        elif tag == "li":
            ctx = self.lists[-1] if self.lists else {"kind": "ul", "i": 0}
            depth = max(0, len(self.lists) - 1) + self.qlindent
            pad = _IND * depth
            if ctx["kind"] == "ol":
                ctx["i"] += 1
                mark = f"{ctx['i']}. "
            else:
                mark = "- "
            self._flush(pad + mark, pad + "  ")
            self.qlindent = 0
        elif tag == "p" or tag == "div":
            self._flush()
            self._blank()
        elif re.fullmatch(r"h[1-6]", tag):
            self._flush("#" * self.heading + " ")
            self.heading = 0
            self._blank()
        elif tag == "blockquote":
            self._flush("> ")
            self.quote = max(0, self.quote - 1)
            self._blank()
        elif tag == "pre":
            self._flush()
            self.lines.append("```")
            self.pre = max(0, self.pre - 1)
            self._blank()
        elif tag in ("td", "th"):
            if self.row is not None:
                self.row.append("".join(self.buf).strip().replace("|", r"\|"))
                self.buf = []
        elif tag == "tr":
            if self.row:
                self.lines.append("| " + " | ".join(self.row) + " |")
                self.tablerows += 1
                if self.tablerows == 1:          # 머리글 구분선 — 없으면 표로 안 읽힌다
                    self.lines.append("| " + " | ".join("---" for _ in self.row) + " |")
            self.row = None
        elif tag == "table":
            self._blank()

    def handle_data(self, data):
        if self.skip:
            return
        if self.pre:
            self.buf.append(data)
            return
        # Quill은 빈 줄을 <p><br></p>로 내고 들여쓰기에 \n을 섞는다 — 공백을 눌러 둔다
        self.buf.append(re.sub(r"\s+", " ", data))

    def result(self) -> str:
        self._flush()
        out: list[str] = []
        for ln in self.lines:
            if ln == "" and (not out or out[-1] == ""):
                continue
            out.append(ln.rstrip() if ln.strip() else "")
        return "\n".join(out).strip("\n")


def html_to_markdown(text) -> str:
    """Quill이 저장한 HTML을 마크다운으로."""
    p = _ToMarkdown()
    p.feed(str(text or ""))
    p.close()
    return p.result()


def plain_to_markdown(text) -> str:
    """플레인 텍스트(자동 생성·AI 생성 결과)를 마크다운으로.

    AI 프롬프트가 지시하는 'ㄴ '(세부) 줄을 **중첩 불릿**으로 올린다. 그냥 두면
    마크다운이 앞 불릿의 이어진 줄로 흡수해 한 줄로 붙여 버린다. 프롬프트는
    손대지 않는다 — 출력 서식을 지시하는 문자열이라 말투를 바꾸면 결과가 깨진다.

    **이미 마크다운인 글도 그대로 통과해야 한다.** 같은 문단 안 개행을 지키려고
    앞 줄에 공백 2칸을 붙이는데, 그걸 표 줄(`| a | b |`)에 붙이면 구분선이
    깨져 표가 통째로 안 그려진다 — 블록 문법 줄에는 붙이지 않는다
    (`t_markdown_passes_through_unchanged`).
    """
    out: list[str] = []
    for raw in str(text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        ln = raw.replace("\t", "  ").rstrip()
        t = ln.strip()
        if not t:
            if out and out[-1] != "":
                out.append("")
            continue
        # 이미 중첩된 불릿의 들여쓰기를 지킨다. 왼쪽을 그냥 strip하면 다시 읽을
        # 때마다 한 단계씩 평평해져, 편집을 열고 닫기만 해도 계층이 사라진다
        # (`t_nesting_survives_a_second_read`).
        # (n+2)//4 — 4칸이 한 단계이면서 2칸으로 적어 온 글도 한 단계로 받는다.
        # 우리 출력(4·8·12칸)을 다시 넣어도 같은 단계가 나와야 한다(멱등).
        pad = _IND * ((len(ln) - len(ln.lstrip(" ")) + 2) // 4)
        if t[0] in _SUBMARK:
            out.append(pad + _IND + "- " + t[1:].strip())
        elif t[:2] in ("- ", "* ", "+ "):
            out.append(pad + "- " + t[2:].strip())
        elif _BLOCKISH.match(t):
            out.append(pad + t)
        else:
            # 같은 문단 안의 개행을 지키려면 앞 줄 끝에 공백 2칸이 필요하다.
            # 단 앞 줄이 블록 문법이면 붙이지 않는다 (표·제목·울타리가 깨진다).
            if out and out[-1] and not _BLOCKISH.match(out[-1].strip()):
                out[-1] += "  "
            out.append(t)
    return "\n".join(out).strip("\n")


def to_markdown(text) -> str:
    """저장된 보고란 문구를 에디터가 읽을 마크다운으로. 꼴은 알아서 가른다."""
    s = str(text or "")
    if not s.strip():
        return ""
    return html_to_markdown(s) if looks_like_html(s) else plain_to_markdown(s)
