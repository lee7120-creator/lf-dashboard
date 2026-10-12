"""차트 레이아웃을 figure에 싸게 입힌다 — 두 대시보드가 같이 쓴다.

**`fig.update_layout(**lay)`를 쓰지 말고 `set_layout(fig, lay)`를 쓴다.**

`update_layout`은 플로틀리가 기존 레이아웃에 속성을 하나씩 얹으면서 매직 밑줄 경로
(`xaxis_tickfont_size` 같은 표기)를 매번 다시 파싱하고 검증한다. 레이아웃 하나가
40개 남짓한 말단 속성을 들고 있어서 **차트 한 장에 12.3ms**가 든다. 차트가 열 장
넘는 화면이 여럿이라 리런마다 그대로 쌓인다.

실측(합성 캠페인 12,275건 · 사이트 22,400행) — `update_layout`을 무력화하고 같은
페이지를 다시 재서 얻은 값이다:

| 발송성과 페이지 | 리런 | 레이아웃을 끄면 |
|---|---|---|
| 주간보고 | 1.15초 | 0.86초 (−0.29) |
| 사이트·앱 지표 | 0.73초 | 0.52초 (−0.21) |
| 타이밍·카테고리 | 0.53초 | 0.43초 (−0.10) |

어디가 비싼지는 한 호출을 통째로 프로파일해서 좁혔다 — `go.Layout(**lay)` **생성**은
0.79ms고 `fig.layout`에 **꽂는** 것도 1.8ms다. 비싼 건 둘 다 아니라 **병합**이었다.
그래서 병합을 평범한 dict로 먼저 해 두고 레이아웃을 통째로 갈아 끼운다 — 합쳐
3.1ms(**4배**).

마지막 함정: `go.Figure()`가 꽂아 둔 **기본 템플릿을 새 레이아웃에 같이 넘기면
12ms가 그대로 돌아온다**(그 큰 dict를 다시 검증한다). 그게 정확히 범인이었다. 안
실어도 직렬화할 때 `plotly.io.templates.default`가 다시 붙어서 **결과 JSON이 한
글자도 다르지 않다** — 전제는 앱이 템플릿을 직접 고르지 않는 것이고,
`tests/test_rerun_cost.py`가 그 전제와 동치를 같이 못 박는다.
"""
import plotly.graph_objects as go


def lay_merge(a, b):
    """b를 a 위에 깊게 얹는다 — 플로틀리를 안 거치니 검증이 없다."""
    out = dict(a)
    for k, v in b.items():
        out[k] = (lay_merge(out[k], v)
                  if isinstance(v, dict) and isinstance(out.get(k), dict) else v)
    return out


def set_layout(fig, lay):
    """`lay`를 figure에 입힌다. 이미 명시된 속성은 안 잃는다.

    서브플롯 축 도메인·패널 제목 주석(`layout.annotations`)처럼 `make_subplots`가
    심어 둔 것들이 그 자리다 — dict로 먼저 합치므로 그대로 남는다.

    **레이아웃을 입힌 뒤에 `add_trace(row=, col=)`를 부르지 말 것.** 트레이스를 먼저
    다 넣고 마지막에 입힌다(지금 두 앱이 다 그 순서다).
    """
    cur = getattr(fig.layout, "_props", None)
    if cur is None:                   # 플로틀리가 사설 속성을 안 주면 예전 길로
        fig.update_layout(**lay)
        return fig
    fig.layout = go.Layout(**lay_merge(
        {k: v for k, v in cur.items() if k != "template"}, lay))
    return fig
