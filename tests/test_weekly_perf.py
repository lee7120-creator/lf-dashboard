"""발송성과 「0. 주간보고」 — 기간 단위(일/주/월)·값 기준·비교·추이 테스트.

이 페이지는 오래 **주 고정**이었다. 기간 단위를 갖게 되면서 조용히 틀릴 자리가 셋 생겼다.

1. 비교 이름이 주에 박혀 있으면 일·월 단위에서 **화면이 거짓말을 한다** — 「전주비」 칸이
   실은 전일·전월을 재고 있어도 값은 멀쩡히 찍힌다.
2. 달마다 일수가 달라(28~31일) 누계로 맞대면 '2월이 10% 적다'가 **날짜 수만으로** 나온다.
   「일평균」이 그걸 없애는 자리라, 가산 지표만 나누고 비율은 그대로 둬야 한다.
3. 부분 기간을 안 자르면 2일치가 7일치와 맞붙어 △70%대 가짜 급락이 뜬다. 예전엔 주에서만
   잘랐다 — 월 단위에선 월초에 열 때마다 같은 착시가 난다.

전부 **화면은 멀쩡히 뜨고 숫자만 틀리는** 모양이라 스모크로는 안 잡힌다.

로컬 실행:
    python tests/test_weekly_perf.py
"""
import pathlib
import sys

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest        # noqa: E402

from smoke_pages import synth_store             # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 400
PAGE = "0. 주간보고"

# 전년 비교가 살아 있으려면 1년을 넘겨야 한다. 70주면 최근 넉 달만 전년 짝이 있어서
# '전년이 있는 기간'과 '없는 기간'이 한 픽스처에 같이 들어온다.
STORE = synth_store(weeks=70)


def _open(unit=None, camp=None, **ss):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = STORE if camp is None else camp
    at.run()
    assert not at.exception, at.exception[0].value
    at.sidebar.radio[0].set_value(PAGE)
    at.run()
    assert not at.exception, at.exception[0].value
    if unit:
        at.session_state["wr_unit"] = unit
    for k, v in ss.items():
        at.session_state[k] = v
    if unit or ss:
        at.run()
        assert not at.exception, at.exception[0].value
    return at


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success):
        out += [str(e.value) for e in coll]
    return out


def _kpi(at):
    """「주요 지표 현황」 표 — 지표 칼럼이 있고 칼럼이 MultiIndex가 아닌 첫 표."""
    for d in at.dataframe:
        cols = list(getattr(d.value, "columns", []))
        if cols and not isinstance(cols[0], tuple) and "지표" in [str(c) for c in cols]:
            return d.value
    return None


def _trend(at):
    """「주요 지표 추이」 표 — 칼럼이 (블록, 기간) 2단인 표."""
    for d in at.dataframe:
        cols = list(getattr(d.value, "columns", []))
        if cols and isinstance(cols[0], tuple):
            return d.value
    return None


def _num(s):
    """화면 문자열에서 숫자만 — '1,234' · '3.85%' · '1,234원' · '1,621만' · '4.35억'.

    `won()`이 큰 금액을 만·억으로 줄여 찍는다. 그걸 모르고 float()에 넘기면 테스트가
    '값이 틀렸다'가 아니라 ValueError로 죽어 원인이 안 드러난다."""
    t = str(s).replace(",", "").replace("원", "").replace("%", "").strip()
    mul = 1.0
    if t.endswith("억"):
        t, mul = t[:-1], 1e8
    elif t.endswith("만"):
        t, mul = t[:-1], 1e4
    return float(t) * mul


def _elapsed_days(unit):
    """화면이 쓰는 경과 일수 — 기준 기간 첫날부터 실적이 있는 마지막 날까지.

    최신 기간은 진행 중이라 달·주 전체가 아니다. '달 길이로 나눴겠지'라고 적으면
    테스트가 규칙이 아니라 달력을 재게 된다."""
    d = STORE.copy()
    d["dt"] = pd.to_datetime(d["date"], format="%Y%m%d")
    last = d["dt"].max()
    if unit == "일":
        return 1
    ps = (last - pd.Timedelta(days=int(last.weekday())) if unit == "주"
          else last.replace(day=1))
    return int((last.normalize() - ps).days) + 1


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_comparison_names_follow_the_unit():
    """비교 칼럼 이름이 단위를 따라가야 한다.

    주에 박아 두면 일 단위에서 「전주비」 칸이 실은 전일을 재게 되는데, 값은 멀쩡히
    찍혀서 눈으로는 안 잡힌다."""
    want = {"일별": ["전일비", "전주비", "전년비"],
            "주별": ["전주비", "전월비", "전년비"],
            "월별": ["전월비", "전전월비", "전년비"]}
    for unit, cols in want.items():
        t = _kpi(_open(unit))
        assert t is not None, f"{unit}: 주요 지표 현황 표를 못 찾았어요"
        got = [str(c) for c in t.columns]
        assert got[0] == "지표", got
        assert [c for c in got if c in cols] == cols, f"{unit}: {got}"
        # 다른 단위의 이름이 섞이면 안 된다 (전주비는 일·주 둘 다 쓰므로 제외)
        stray = [c for c in got
                 if c.endswith("비") and c not in cols]
        assert not stray, f"{unit}에 엉뚱한 비교 칼럼이 남았어요 — {stray}"


@case
def t_reference_period_label_follows_the_unit():
    """기준 기간 칼럼 이름도 단위를 따라간다."""
    for unit, head in (("일별", "기준 일자"), ("주별", "기준 주차"), ("월별", "기준 월")):
        t = _kpi(_open(unit))
        cur = [str(c) for c in t.columns if str(c).startswith("기준 ")]
        assert cur and cur[0].startswith(head), f"{unit}: {cur}"


@case
def t_daily_average_divides_additive_metrics_only():
    """「일평균」은 가산 지표만 기간 일수로 나눈다 — 비율은 그대로다.

    달마다 일수가 다른 월 단위에서 특히 중요하다. CTR까지 나누면 3.85%가 0.12%가 된다."""
    a = _kpi(_open("월별"))
    b = _kpi(_open("월별", wr_valmode="일평균"))
    cur_a = [c for c in a.columns if str(c).startswith("기준 ")][0]
    cur_b = [c for c in b.columns if str(c).startswith("기준 ")][0]

    def _v(t, col, met):
        return _num(t[t["지표"] == met][col].iloc[0])

    want = _elapsed_days("월")
    assert want > 1, "픽스처의 기준 달이 하루뿐이라 규칙을 반증하지 못해요"
    for met in ("발송", "UV", "캠페인수", "주문건수"):
        s, m = _v(a, cur_a, met), _v(b, cur_b, met)
        assert m < s, f"{met}: 일평균({m})이 누계({s})보다 작아야 해요"
        _days = s / m
        assert abs(_days - want) / want < 0.02, \
            f"{met}: {_days:.2f}일로 나눴어요 (기대 {want}일 — 경과분까지여야 해요)"
    for met in ("CTR", "주문CR", "RPS", "객단가"):
        assert abs(_v(a, cur_a, met) - _v(b, cur_b, met)) < 1e-6, \
            f"{met}는 이미 나눈 값이라 일평균에서도 그대로여야 해요"


@case
def t_daily_average_uses_the_same_divisor_for_every_metric():
    """가산 지표 다섯이 **같은 일수**로 나뉘어야 한다 — 하나만 다른 창을 쓰면
    같은 화면에서 지표끼리 단위가 어긋난다."""
    a, b = _kpi(_open("주별")), _kpi(_open("주별", wr_valmode="일평균"))
    ca = [c for c in a.columns if str(c).startswith("기준 ")][0]
    cb = [c for c in b.columns if str(c).startswith("기준 ")][0]
    want = _elapsed_days("주")
    got = {}
    for met in ("발송", "UV", "주문건수", "캠페인수"):
        s = _num(a[a["지표"] == met][ca].iloc[0])
        m = _num(b[b["지표"] == met][cb].iloc[0])
        got[met] = s / m
    # 화면 값이 정수로 반올림돼 찍히니 소수점까지 같기를 요구하면 작은 지표에서 깨진다.
    # 규칙을 깨면(한 지표만 다른 창) 몇 배씩 벌어지므로 상대 오차로 충분히 잡힌다.
    off = {k: v for k, v in got.items() if abs(v - want) / want >= 0.02}
    assert not off, f"기대 {want}일과 다른 창을 쓴 지표가 있어요 — {off}"


@case
def t_comparison_picker_scopes_cards_and_table():
    """「비교」에서 뺀 비교는 표에서도 사라져야 한다 (카드와 표가 같은 목록을 본다)."""
    at = _open("주별", wr_cmp_주=["전년비"])
    got = [str(c) for c in _kpi(at).columns]
    assert "전년비" in got, got
    assert "전주비" not in got and "전월비" not in got, got
    # 하나도 안 고르면 비교가 통째로 사라져 화면이 무의미해진다 — 전체로 되돌린다
    at2 = _open("주별", wr_cmp_주=[])
    got2 = [str(c) for c in _kpi(at2).columns]
    assert {"전주비", "전월비", "전년비"} <= set(got2), got2


@case
def t_trend_table_puts_actuals_left_and_yoy_right():
    """추이 표는 **왼쪽 실적 · 오른쪽 전년비**다.

    증감만 가로로 훑어야 '어느 기간부터 꺾였나'가 보인다. 값·증감을 기간마다 번갈아
    끼우면 그게 안 된다."""
    t = _trend(_open("주별"))
    assert t is not None, "추이 표를 못 찾았어요"
    blocks = [c[0] for c in t.columns]
    assert set(blocks) == {"실적", "전년비"}, sorted(set(blocks))
    # 실적이 다 나온 뒤에 전년비가 나와야 한다
    assert blocks == sorted(blocks, key=lambda b: 0 if b == "실적" else 1), blocks
    # 행은 지표 전부
    assert list(t.index) == ["캠페인수", "발송", "UV", "주문건수", "거래액",
                             "CTR", "주문CR", "RPS", "객단가"], list(t.index)


@case
def t_trend_window_follows_the_control():
    """「추이 구간」이 차트·표가 보는 기간 수를 정한다."""
    n13 = len({c[1] for c in _trend(_open("주별")).columns})
    assert n13 == 13, f"기본은 최근 13기간이어야 해요 — {n13}"
    tall = _trend(_open("주별", **{"wr_twin_주": "전체"}))
    nall = len({c[1] for c in tall.columns})
    assert nall > n13, f"'전체'가 13개 이하예요 — {nall}"
    tyr = _trend(_open("주별", **{"wr_twin_주": "올해 전체"}))
    years = {lb.split("년")[0] for _b, lb in tyr.columns}
    assert len(years) == 1, f"'올해 전체'에 다른 해가 섞였어요 — {sorted(years)}"


@case
def t_trend_periods_are_the_unit_not_always_weeks():
    """추이 기간 라벨도 단위를 따라간다 — 월로 바꿨는데 주차가 남으면 안 된다."""
    for unit, mark in (("일별", "("), ("주별", "주차"), ("월별", "월")):
        labs = {c[1] for c in _trend(_open(unit)).columns}
        assert all(mark in lb for lb in labs), f"{unit}: {sorted(labs)[:3]}"
    assert not any("주차" in lb for lb in {c[1] for c in _trend(_open("월별")).columns})


@case
def t_month_unit_drops_the_duplicate_mtd_table():
    """월 단위에선 위 「주요 지표 현황」이 곧 월 누계 비교라 MTD 표를 또 그리지 않는다."""
    def _has_mtd(at):
        for d in at.dataframe:
            # 「월말 마감 예상」 표에도 '당월 MTD' 칼럼이 있다 — 기간 표기가 붙은
            # 월 누계 표만 골라야 엉뚱한 표를 세지 않는다
            if any(str(c).startswith("당월 MTD (") for c in getattr(d.value, "columns", [])):
                return True
        return False
    assert _has_mtd(_open("주별")), "주 단위엔 MTD 표가 있어야 해요"
    assert _has_mtd(_open("일별")), "일 단위엔 MTD 표가 있어야 해요"
    at = _open("월별")
    assert not _has_mtd(at), "월 단위에서 같은 비교를 두 번 보여 주고 있어요"
    assert any("월 누계 비교라" in t for t in _texts(at)), "왜 없는지 말해 주지 않았어요"


@case
def t_partial_period_clamps_every_comparison():
    """부분 기간이면 비교 기간도 **같은 경과분까지만** 집계해야 한다.

    안 자르면 2일치가 7일치와 맞붙어 △70%대 가짜 급락이 뜬다. 화면에 찍힌 전주 값을
    픽스처에서 직접 센 값과 대조한다."""
    at = _open("주별")
    t = _kpi(at)
    # 기준주는 오늘이 속한 주 = 진행 중 → 부분 기간
    cur_col = [str(c) for c in t.columns if str(c).startswith("기준 주차")][0]
    assert any("동요일 누계" in x for x in _texts(at)), "부분 기간 안내가 없어요"

    d = STORE.copy()
    d["dt"] = pd.to_datetime(d["date"], format="%Y%m%d")
    last = d["dt"].max()
    ws = last - pd.Timedelta(days=int(last.weekday()))
    elapsed = int((last.normalize() - ws).days)
    prev = ws - pd.Timedelta(days=7)
    want = float(d[(d["dt"] >= prev)
                   & (d["dt"] <= prev + pd.Timedelta(days=elapsed))]["send"].sum())
    full = float(d[(d["dt"] >= prev)
                   & (d["dt"] <= prev + pd.Timedelta(days=6))]["send"].sum())
    assert want < full, "픽스처가 부분 주를 못 만들어 규칙을 반증하지 못해요"
    # 전주 실적 열은 접혀 있으니 펼쳐서 읽는다
    at2 = _open("주별", wr_sum_cols_주={"selection": {"columns": ["전주비"]}})
    t2 = _kpi(at2)
    pcol = [str(c) for c in t2.columns if str(c).startswith("전주 (")]
    assert pcol, [str(c) for c in t2.columns]
    got = _num(t2[t2["지표"] == "발송"][pcol[0]].iloc[0])
    assert abs(got - want) < 1.0, f"화면 {got:,.0f} vs 같은 요일까지 {want:,.0f} (전체 {full:,.0f})"
    assert cur_col


@case
def t_screen_value_matches_a_direct_sum():
    """기준 기간 값이 픽스처를 직접 센 값과 같아야 한다.

    추이·카드가 기간별 groupby로 값을 읽는데, 그 인덱스가 슬라이스와 한 칸이라도
    어긋나면 화면 숫자가 통째로 틀어진다."""
    d = STORE.copy()
    d["dt"] = pd.to_datetime(d["date"], format="%Y%m%d")
    last = d["dt"].max()
    m0 = last.replace(day=1)
    want = float(d[(d["dt"] >= m0) & (d["dt"] <= last)]["send"].sum())
    t = _kpi(_open("월별"))
    cur = [c for c in t.columns if str(c).startswith("기준 월")][0]
    got = _num(t[t["지표"] == "발송"][cur].iloc[0])
    assert abs(got - want) < 1.0, f"화면 {got:,.0f} vs 직접 합 {want:,.0f}"


@case
def t_period_without_data_reads_as_missing_not_zero():
    """발송이 아예 없던 비교 기간은 '–'여야 한다 — 0이면 '보냈는데 성과가 0'으로 읽힌다."""
    d = STORE.copy()
    d["dt"] = pd.to_datetime(d["date"], format="%Y%m%d")
    last = d["dt"].max()
    ws = last - pd.Timedelta(days=int(last.weekday()))
    hole0, hole1 = ws - pd.Timedelta(days=14), ws - pd.Timedelta(days=8)
    camp = d[~((d["dt"] >= hole0) & (d["dt"] <= hole1))].drop(columns=["dt"])
    # 기준을 그 구멍의 다음 주로 잡으면 전주가 통째로 빈다
    at = _open(camp=camp)
    at.session_state["wr_unit"] = "주별"
    at.session_state["wr_period_주"] = None
    at.run()
    labs = [s for s in at.selectbox if str(s.label).startswith("기준 ")]
    assert labs, [s.label for s in at.selectbox]
    target = [o for o in labs[0].options
              if f"{(hole1 + pd.Timedelta(days=1)).strftime('%m/%d')}~" in o]
    assert target, f"구멍 다음 주를 못 찾았어요 — {labs[0].options[:4]}"
    at.session_state["wr_period_주"] = target[0]
    at.session_state["wr_sum_cols_주"] = {"selection": {"columns": ["전주비"]}}
    at.run()
    assert not at.exception, at.exception[0].value
    t = _kpi(at)
    pcol = [str(c) for c in t.columns if str(c).startswith("전주 (")]
    assert pcol, [str(c) for c in t.columns]
    got = str(t[t["지표"] == "발송"][pcol[0]].iloc[0])
    assert got == "–", f"빈 기간이 '{got}'로 찍혔어요 — 0이 아니라 '–'여야 해요"


@case
def t_recent_button_jumps_to_the_latest_period():
    """「↩ 최근으로」는 가장 최근 기간으로 되돌린다 (옛 기간을 보다가 한 번에 돌아오는 자리)."""
    at = _open("주별")
    sel = [s for s in at.selectbox if str(s.label).startswith("기준 ")][0]
    newest, older = sel.options[0], sel.options[3]
    at.session_state["wr_period_주"] = older
    at.run()
    assert not at.exception, at.exception[0].value
    btn = [b for b in at.button if "최근으로" in str(b.label)]
    assert btn, [b.label for b in at.button][:5]
    btn[0].click()
    at.run()
    assert not at.exception, at.exception[0].value
    assert at.session_state["wr_period_주"] == newest, \
        f"{at.session_state['wr_period_주']} (기대 {newest})"


@case
def t_decomposition_names_the_right_previous_period():
    """증감 기여 분해도 단위를 따라간다 — 월 단위에서 「전주 대비」가 남으면 거짓말이다."""
    for unit, pvn, nope in (("월별", "전월", "전주"), ("일별", "전일", "전월")):
        body = " ".join(_texts(_open(unit)))
        assert f"거래액 {pvn} 대비" in body, f"{unit}: '거래액 {pvn} 대비'가 없어요"
        assert f"거래액 {nope} 대비" not in body, f"{unit}: '거래액 {nope} 대비'가 남았어요"


def main():
    # 인자를 주면 그 이름이 든 케이스만 돈다 — 규칙을 일부러 깨서 FAIL이 뜨는지 볼 때
    # 전체(케이스마다 앱을 새로 렌더)를 다시 돌릴 이유가 없다.
    want = sys.argv[1:]
    fails = []
    for fn in [f for f in CASES if not want or any(w in f.__name__ for w in want)]:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {str(e)[:400]}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"주간보고 페이지 테스트 {len(CASES) if not want else '선택'}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
