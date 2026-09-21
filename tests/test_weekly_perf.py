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
import re
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest        # noqa: E402

from smoke_pages import group_of, synth_store   # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 400
PAGE = group_of("주간보고")

# 전년 비교가 살아 있으려면 1년을 넘겨야 한다. 70주면 최근 넉 달만 전년 짝이 있어서
# '전년이 있는 기간'과 '없는 기간'이 한 픽스처에 같이 들어온다.
STORE = synth_store(weeks=70)


def _open(unit=None, camp=None, site=None, **ss):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = STORE if camp is None else camp
    if site is not None:
        at.session_state["site_store_df"] = site
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


def _push_fixture():
    """앱푸시 수신동의 원천 — 「앱푸시 수신동의 요약」 블록을 띄우는 최소 프레임.

    이 블록은 `push_consent_df`가 있어야만 렌더된다. 없으면 칼럼 이름 규칙을 깨도
    표 자체가 안 떠서 검사가 조용히 통과한다."""
    d = pd.to_datetime(STORE["date"], format="%Y%m%d")
    days = pd.date_range(d.min(), d.max(), freq="D")
    rows = []
    for i, g in enumerate(("Total", "기존", "신규")):
        for j, dt in enumerate(days):
            rows.append({"date": dt, "group": g,
                         "consent": 100000 + i * 1000 + j * 7,
                         "added": 300 + i * 10 + (j % 5),
                         "removed": 120 + i * 5 + (j % 3),
                         "diff": 180 + i * 5, "is_outlier": False})
    return pd.DataFrame(rows)


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
    """「추이 구간」이 차트·표가 보는 기간 수를 정한다.

    「최근 N」은 **단위마다 다르다** — 13으로 고정하면 일 단위가 13일뿐이라 추세가
    안 보인다. 라벨에 적힌 개수와 실제로 그린 개수가 같아야 한다."""
    for unit, u, want in (("일별", "일", 30), ("주별", "주", 13), ("월별", "월", 13)):
        at = _open(unit)
        got = len({c[1] for c in _trend(at).columns})
        assert got == want, f"{unit}: 최근 구간이 {got}개예요 (기대 {want})"
        chips = [b for b in at.get("button_group")
                 if str(getattr(b, "label", "")) == "추이 구간"]
        assert chips and str(chips[0].value).startswith(f"최근 {want}"), \
            f"{unit}: 라벨과 개수가 달라요 — {chips[0].value if chips else None}"
    n13 = len({c[1] for c in _trend(_open("주별")).columns})
    tall = _trend(_open("주별", **{"wr_twin_주": "전체"}))
    nall = len({c[1] for c in tall.columns})
    assert nall > n13, f"'전체'가 13개 이하예요 — {nall}"
    tyr = _trend(_open("주별", **{"wr_twin_주": "올해 전체"}))
    years = {lb.split("년")[0] for _b, lb in tyr.columns}
    assert len(years) == 1, f"'올해 전체'에 다른 해가 섞였어요 — {sorted(years)}"


@case
def t_year_window_uses_the_iso_year_of_the_label():
    """「올해 전체」는 **라벨에 찍힌 연도**로 거른다 — 주는 ISO 기준연도다.

    달력 연도로 거르면 연초 주가 조용히 빠진다: 2026년 1주차의 월요일은 2025-12-29라
    `p.year`가 2025다. 증상은 '추이가 2주차부터 시작하네'뿐이라 눈으로는 안 잡히고,
    9월에서 끝나는 픽스처로는 이 경계를 아예 안 밟는다."""
    _END = "2026-01-18"                       # 연말·연초를 품도록 끝을 1월로 민다
    d = STORE.copy()
    dt = pd.to_datetime(d["date"], format="%Y%m%d")
    d["date"] = (dt + (pd.Timestamp(_END) - dt.max())).dt.strftime("%Y%m%d")
    at = _open(camp=d)
    at.session_state["wr_unit"] = "주별"
    at.run()
    labs = [s_.label for s_ in at.selectbox if str(s_.label).startswith("기준 ")]
    assert labs, [s_.label for s_ in at.selectbox]
    # 2026년 1주차(월요일 2025-12-29)가 목록에 있어야 규칙을 반증할 수 있다
    opts = [s_ for s_ in at.selectbox if str(s_.label).startswith("기준 ")][0].options
    assert any(o.startswith("2026년 1주차") for o in opts), \
        f"픽스처가 연초 주를 안 품어 규칙을 반증하지 못해요 — {opts[:4]}"
    at.session_state["wr_twin_주"] = "올해 전체"
    at.run()
    assert not at.exception, at.exception[0].value
    got = {lb for _b, lb in _trend(at).columns}
    assert any(lb.startswith("2026년 1주차") for lb in got), \
        f"2026년 1주차가 「올해 전체」에서 빠졌어요 — {sorted(got)[:5]}"
    assert all(lb.startswith("2026년") for lb in got), \
        f"다른 해가 섞였어요 — {sorted(got)[:5]}"


@case
def t_report_note_is_keyed_to_the_week_whatever_the_unit():
    """보고란은 **늘 주 단위**로 저장한다.

    기준 기간의 첫날을 키로 쓰면 단위마다 칸이 갈려, 주로 써 둔 글이 일·월로 바꾸는
    순간 화면에서 사라진다(파일엔 남는다). 왜 없어졌는지가 안 드러나는 모양이다."""
    keys = {}
    for unit in ("일별", "주별", "월별"):
        at = _open(unit)
        # 자동 생성 버튼의 키가 곧 저장 키다 (btn_r_kpi_<YYYYMMDD>)
        hit = [b for b in at.button if "btn_r_kpi_" in str(b.proto.id)]
        assert hit, f"{unit}: 보고란 자동 생성 버튼을 못 찾았어요"
        keys[unit] = str(hit[0].proto.id).split("btn_r_kpi_")[1][:8]
    for unit, k in keys.items():
        assert pd.Timestamp(k).weekday() == 0, \
            f"{unit}: 저장 키 {k}가 월요일이 아니에요 (주 단위 앵커여야 해요)"
    assert keys["일별"] == keys["주별"], \
        f"일·주가 다른 칸을 열어요 — {keys} (주로 써 둔 글이 일에서 안 보여요)"


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
    픽스처에서 직접 센 값과 대조한다.

    **픽스처를 수요일에서 끊는다.** 기본 `STORE`는 '오늘'까지라, 일요일에 돌리면 기준주가
    통째로 차서 부분 기간 경로를 아예 안 밟는다 — 검사가 규칙이 아니라 **달력을 재게**
    된다(실제로 일요일→월요일로 넘어가는 자정에 이 검사만 빨개졌다). 요일과 무관하게
    같은 경로를 타도록 마지막 수요일까지만 남긴다."""
    d = STORE.copy()
    d["dt"] = pd.to_datetime(d["date"], format="%Y%m%d")
    _last = d["dt"].max()
    _cut = _last - pd.Timedelta(days=(int(_last.weekday()) - 2) % 7)   # 마지막 수요일
    d = d[d["dt"] <= _cut].copy()
    camp = d.drop(columns=["dt"])

    at = _open("주별", camp=camp)
    t = _kpi(at)
    # 기준주는 수요일에서 끊겨 있으니 늘 진행 중 → 부분 기간
    cur_col = [str(c) for c in t.columns if str(c).startswith("기준 주차")][0]
    assert any("동요일 누계" in x for x in _texts(at)), "부분 기간 안내가 없어요"

    last = d["dt"].max()
    ws = last - pd.Timedelta(days=int(last.weekday()))
    elapsed = int((last.normalize() - ws).days)
    assert elapsed == 2, f"수요일에서 끊었는데 경과가 {elapsed}일이에요"
    prev = ws - pd.Timedelta(days=7)
    want = float(d[(d["dt"] >= prev)
                   & (d["dt"] <= prev + pd.Timedelta(days=elapsed))]["send"].sum())
    full = float(d[(d["dt"] >= prev)
                   & (d["dt"] <= prev + pd.Timedelta(days=6))]["send"].sum())
    assert want < full, "픽스처가 부분 주를 못 만들어 규칙을 반증하지 못해요"
    # 전주 실적 열은 접혀 있으니 펼쳐서 읽는다
    at2 = _open("주별", camp=camp, wr_sum_cols_주={"selection": {"columns": ["전주비"]}})
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
def t_trend_metric_picker_is_chips_not_a_tag_box():
    """「추이에 올릴 지표」는 **칩 버튼**이다 — 태그를 넣고 빼는 multiselect가 아니다.

    multiselect는 **고른 것만** 보여서 뭘 더 켤 수 있는지가 화면에 안 남는다. 칩은
    전 목록이 늘 떠 있고 안 고른 건 회색으로 남아, 한 번에 보인다. 위 「비교」와도
    같은 모양이라 화면이 한 가지로 읽힌다."""
    at = _open()
    LB = "추이에 올릴 지표"
    assert LB not in [m.label for m in at.multiselect], \
        "태그 상자(multiselect)로 되돌아갔어요 — 칩이어야 해요"
    chips = [b for b in at.get("button_group") if str(getattr(b, "label", "")) == LB]
    assert chips, f"칩을 못 찾았어요 — {[getattr(b, 'label', None) for b in at.get('button_group')]}"
    c = chips[0]
    # 전 지표가 칩으로 다 떠 있어야 '뭘 더 켤 수 있나'가 보인다
    assert list(c.options) == ["캠페인수", "발송", "UV", "주문건수", "거래액",
                               "CTR", "주문CR", "RPS", "객단가"], list(c.options)
    assert set(c.value) == {"발송", "UV", "주문건수", "거래액", "CTR", "주문CR"}, \
        f"기본으로 켜 둘 지표가 달라요 — {list(c.value)}"


def _fig_traces(at, idx=0):
    """추이 차트 한 장의 트레이스 목록 (figure JSON에서 읽는다)."""
    import json
    return json.loads(at.get("plotly_chart")[idx].proto.spec).get("data", [])


def _trend_traces(at, name="올해"):
    """추이 차트의 트레이스 — AppTest는 plotly에 `.value`를 안 주므로 figure JSON을 읽는다."""
    import json
    out = []
    for e in at.get("plotly_chart"):
        f = json.loads(e.proto.spec)
        ttl = (f.get("layout", {}).get("title") or {}).get("text", "")
        for tr in f.get("data", []):
            if tr.get("name") == name:
                out.append((ttl, tr))
    return out


@case
def t_chart_tooltip_carries_the_deltas():
    """점에 커서를 대면 값만이 아니라 **증감까지** 뜬다.

    값만 뜨면 '그래서 얼마나 늘었나'를 눈으로 재거나 아래 표로 내려가야 한다.
    이름은 단위를 따라간다 — 일 단위에서 「전주 대비」가 뜨면 화면이 거짓말이다."""
    for unit, pvn in (("일별", "전일"), ("주별", "전주"), ("월별", "전월")):
        at = _open(unit)
        trs = _trend_traces(at)
        assert trs, f"{unit}: 올해 트레이스를 못 찾았어요"
        for ttl, tr in trs:
            ht = tr.get("hovertemplate") or ""
            assert f"{pvn} 대비" in ht, f"{unit}/{ttl}: 툴팁에 '{pvn} 대비'가 없어요 — {ht}"
            assert "전년 대비" in ht, f"{unit}/{ttl}: 툴팁에 전년 대비가 없어요 — {ht}"
            cd = tr.get("customdata")
            assert cd and len(cd) == len(tr.get("y") or []), \
                f"{unit}/{ttl}: customdata가 점 수와 안 맞아요"


@case
def t_tooltip_delta_unit_follows_the_metric():
    """비율 지표는 %p, 나머지는 % — 툴팁도 표와 같은 단위여야 한다."""
    at = _open("주별", wr_trend_pills=["CTR", "거래액"])
    seen = {}
    for ttl, tr in _trend_traces(at):
        vals = [d[0] for d in (tr.get("customdata") or []) if d and d[0] != "–"]
        if vals:
            seen[ttl] = vals
    assert "CTR" in seen and "거래액" in seen, f"두 차트를 못 찾았어요 — {list(seen)}"
    assert all(v.endswith("%p") for v in seen["CTR"]), f"CTR이 %p가 아니에요 — {seen['CTR'][:3]}"
    assert not any(v.endswith("%p") for v in seen["거래액"]), \
        f"거래액이 %p로 찍혔어요 — {seen['거래액'][:3]}"


@case
def t_tooltip_delta_compares_the_real_previous_period():
    """'전일 대비'는 **달력상 직전 기간**과 맞댄 값이어야 한다.

    선 위의 앞 점으로 재면, 발송이 없어 점이 안 생긴 날이 끼었을 때 「전일 대비」라고
    써 놓고 실은 며칠 전과 비교하게 된다. 그런 날을 일부러 지워 두고 확인한다."""
    d = STORE.copy()
    dt = pd.to_datetime(d["date"], format="%Y%m%d")
    gap = dt.max() - pd.Timedelta(days=3)          # 하루를 통째로 비운다
    d = d[dt != gap]
    at = _open("일별", camp=d)
    lab_gone = f"{gap.year}년 {gap.month}/{gap.day}"
    trs = _trend_traces(at)
    assert trs, "올해 트레이스를 못 찾았어요"
    ttl, tr = trs[0]
    xs = list(tr.get("x") or [])
    assert not any(str(x).startswith(lab_gone) for x in xs), \
        f"픽스처가 빈 날을 못 만들어 규칙을 반증하지 못해요 — {lab_gone}"
    nxt = [i for i, x in enumerate(xs)
           if str(x).startswith(f"{(gap + pd.Timedelta(days=1)).month}/"
                                f"{(gap + pd.Timedelta(days=1)).day}")
           or str(x).startswith(f"{gap.year}년 {(gap + pd.Timedelta(days=1)).month}/"
                                f"{(gap + pd.Timedelta(days=1)).day}")]
    assert nxt, f"빈 날 다음 점을 못 찾았어요 — {xs[-6:]}"
    got = (tr.get("customdata") or [])[nxt[0]][0]
    assert got == "–", \
        f"빈 날 다음 점의 '전일 대비'가 '{got}'예요 — 직전 날이 없으니 '–'여야 해요"


@case
def t_tooltip_deltas_are_coloured_like_the_tables():
    """툴팁 증감도 표·KPI 카드와 **같은 색**이어야 한다 — △ 빨강 · + 초록.

    값만 검게 뜨면 같은 숫자를 표에선 색으로, 차트에선 부호로 읽게 된다.
    plotly 호버는 태그를 tspan으로 바꾸며 `style`을 그대로 입히므로 `<span style>`이면
    된다. 색은 **점마다 다르니 customdata로 같이 싣는다** — 템플릿 문자열 하나로는
    못 가른다. 값 칸(0·1)은 그대로 두고 CSS만 2·3에 붙였다.
    `–`는 hoverlabel 글자색(`DELTA_FG`)을 줘서 안 칠한 것처럼 보이게 한다."""
    import send_perf_dashboard as S
    at = _open("주별", wr_trend_pills=["거래액", "CTR"])
    trs = _trend_traces(at)
    assert trs, "올해 트레이스를 못 찾았어요"
    seen = set()
    for ttl, tr in trs:
        ht = tr.get("hovertemplate") or ""
        for i in (0, 1):
            tag = '<span style="%%{customdata[%d]}">%%{customdata[%d]}</span>' % (i + 2, i)
            assert tag in ht, f"{ttl}: 증감에 색이 안 붙었어요 — {ht}"
        for cd in tr.get("customdata") or []:
            assert len(cd) == 4, f"{ttl}: customdata가 값 2 + 색 2가 아니에요 — {cd}"
            for val, css in ((cd[0], cd[2]), (cd[1], cd[3])):
                want = (S.DELTA_UP if str(val).startswith("+")
                        else S.DELTA_DN if str(val).startswith("△")
                        else S.DELTA_FG)
                assert str(css).startswith(f"color:{want}"), \
                    f"{ttl}: '{val}'에 «{css}»가 붙었어요 (기대 {want})"
                seen.add(want)
    assert {S.DELTA_UP, S.DELTA_DN} <= seen, \
        f"오름·내림이 둘 다 안 나와 색 규칙을 반증하지 못해요 — {seen}"

    # '–'(직전 기간이 없는 점)는 안 칠한 것처럼 보여야 한다 — 하루를 비워 만든다.
    d = STORE.copy()
    dt = pd.to_datetime(d["date"], format="%Y%m%d")
    d = d[dt != dt.max() - pd.Timedelta(days=3)]
    at2 = _open("일별", camp=d, wr_trend_pills=["거래액"])
    dash = [(cd[0], cd[2]) for _t, tr in _trend_traces(at2)
            for cd in (tr.get("customdata") or []) if cd[0] == "–"]
    assert dash, "픽스처가 '–'를 못 만들어 중립 색 규칙을 반증하지 못해요"
    for val, css in dash:
        assert css == f"color:{S.DELTA_FG}", f"'–'에 «{css}»가 붙었어요"


@case
def t_prior_year_line_runs_to_the_end_of_the_year():
    """「올해 전체」에선 **전년 선을 그 해 끝까지** 그린다.

    올해가 아직 안 온 칸도 전년 값이 있으면 x자리를 세운다 — '남은 기간에 전년은
    어땠나'(계절성)를 보려고 여는 화면이라서다. 데이터에 있는 기간만 모으면 전년 선이
    올해와 같은 지점에서 잘려 그 뒤를 못 본다.

    표는 **올해 실적이 있는 기간만** 담는다. 뒷칸은 실적·전년비가 둘 다 '–'라 넣으면
    빈 칼럼만 늘어난다."""
    at = _open("주별", **{"wr_twin_주": "올해 전체"})
    trs = {n: t for n, t in
           ((tr.get("name"), tr) for tr in _fig_traces(at))}
    assert "올해" in trs and "전년" in trs, sorted(trs)
    cy, py = trs["올해"], trs["전년"]
    xs = list(cy["x"])
    _last_cur = max(i for i, v in enumerate(cy["y"]) if v is not None)
    _last_py = max(i for i, v in enumerate(py["y"]) if v is not None)
    assert _last_py > _last_cur, (
        f"전년 선이 올해와 같은 지점에서 끊겼어요 — 올해 {xs[_last_cur]} · "
        f"전년 {xs[_last_py]} (픽스처가 전년 뒷기간을 안 품었을 수도 있어요)")
    assert len(xs) > _last_cur + 1, "올해 뒤로 x자리가 안 생겼어요"
    # 표는 올해 실적이 있는 기간까지만
    tb = {c[1] for c in _trend(at).columns}
    assert xs[_last_cur] in tb, f"표에 마지막 실적 기간이 없어요 — {sorted(tb)[-2:]}"
    assert xs[_last_py] not in tb, \
        f"실적이 없는 기간이 표에 들어왔어요 — {xs[_last_py]} ('–'만 늘어선 칼럼)"


@case
def t_recent_window_is_not_stretched_for_the_prior_year():
    """「최근 N」은 사용자가 일부러 좁힌 창이라 전년 때문에 늘리지 않는다."""
    at = _open("주별")
    for tr in _fig_traces(at):
        if tr.get("name") == "올해":
            assert len(tr["x"]) == 13, f"최근 13주가 {len(tr['x'])}칸으로 늘었어요"
            assert all(v is not None for v in tr["y"]), \
                "최근 창에 올해 값이 빈 칸이 생겼어요"
            break
    else:
        raise AssertionError("올해 트레이스를 못 찾았어요")


SITE_MET = "앱푸시 회원UV(일평균·천명)"


@case
def t_site_metric_joins_the_trend_when_the_source_is_there():
    """앱푸시 회원UV도 추이에 올린다 — 칩·차트·표 셋 다.

    사이트 원천이라 캠페인 집계(`_gsum`)에 없다. 값은 이미 그 기간의 일평균이라
    「값 기준」이 또 나누면 안 되고, 천명 단위라 소수 한 자리로 찍어야 위 표들과
    같은 서식이 된다."""
    from test_site_metrics import _site_store
    at = _open("주별", site=_site_store(days=500))
    chips = [b for b in at.get("button_group")
             if str(getattr(b, "label", "")) == "추이에 올릴 지표"]
    assert chips, "칩을 못 찾았어요"
    assert SITE_MET in list(chips[0].options), \
        f"선택지에 없어요 — {list(chips[0].options)}"
    assert SITE_MET in list(chips[0].value), \
        f"기본으로 안 켜져 있어요 — {list(chips[0].value)}"
    # 차트 한 장이 실제로 그려지고 값이 있어야 한다
    hit = [(t, tr) for t, tr in _trend_traces(at) if t == SITE_MET]
    assert hit, f"차트를 못 찾았어요 — {[t for t, _ in _trend_traces(at)]}"
    tr = hit[0][1]
    assert any(v is not None for v in tr["y"]), "차트에 값이 하나도 없어요"
    assert "%{y:,.1f}" in (tr.get("hovertemplate") or ""), \
        f"천명이라 소수 한 자리여야 해요 — {tr.get('hovertemplate')}"
    # 표에도 한 줄
    t = _trend(at)
    assert SITE_MET in list(t.index), f"표에 없어요 — {list(t.index)}"
    # **실적 칸만** 본다 — 한 줄 전체를 보면 전년비('+3.7%')에 소수점이 있어서
    # 서식을 정수로 깨도 통과한다(실제로 그렇게 심어 보고 확인했다).
    _vals = [str(t.loc[SITE_MET, c]) for c in t.columns
             if c[0] == "실적" and str(t.loc[SITE_MET, c]) != "–"]
    assert _vals, "실적 칸이 전부 비었어요"
    assert all("." in v for v in _vals), \
        f"천명이라 소수 한 자리여야 해요 — {_vals[:3]}"


@case
def t_site_metric_is_absent_without_the_source():
    """사이트 데이터를 안 올렸으면 선택지에 안 띄운다.

    눌러도 아무 선이 안 생기는 선택지는 '왜 안 그려지지'만 남긴다."""
    at = _open("주별")
    chips = [b for b in at.get("button_group")
             if str(getattr(b, "label", "")) == "추이에 올릴 지표"][0]
    assert SITE_MET not in list(chips.options), \
        f"원천이 없는데 선택지에 있어요 — {list(chips.options)}"
    assert SITE_MET not in list(_trend(at).index), "원천이 없는데 표에 줄이 있어요"


@case
def t_site_metric_is_not_divided_again_by_the_value_mode():
    """「일평균」으로 바꿔도 앱푸시 회원UV는 그대로다 — 이미 일평균이라서."""
    from test_site_metrics import _site_store
    site = _site_store(days=500)
    a = _trend(_open("주별", site=site))
    b = _trend(_open("주별", site=site, wr_valmode="일평균"))
    ca = [c for c in a.columns if c[0] == "실적"][-1]
    cb = [c for c in b.columns if c[0] == "실적"][-1]
    assert ca == cb, f"비교할 기간이 달라요 — {ca} vs {cb}"
    assert a.loc[SITE_MET, ca] == b.loc[SITE_MET, cb], \
        f"값 기준이 사이트 지표까지 나눴어요 — {a.loc[SITE_MET, ca]} → {b.loc[SITE_MET, cb]}"
    # 가산 지표는 반대로 줄어야 한다 (픽스처가 규칙을 반증하는지 확인)
    assert _num(a.loc["발송", ca]) > _num(b.loc["발송", cb]), \
        "발송이 안 줄었어요 — 일평균 모드가 안 걸린 픽스처예요"


@case
def t_prose_reads_right_in_every_unit():
    """단위를 바꿔도 화면 글이 말이 돼야 한다.

    기간 단위가 일·주·월로 갈리면서 조사와 세는 말이 어긋났다 —
    「기준 **월가** 부분 기간이라」·「**7개 월**」. 산출식 캡션엔 단위와 무관하게
    「기준주·전주」가 박혀 있어 월로 봐도 주라고 말했다. 전부 화면은 멀쩡히 뜨고
    글만 틀리는 자리라 눈으로만 잡힌다."""
    import re
    for unit, uname, peradj in (("일별", "일자", "일간"), ("주별", "주차", "주간"),
                                ("월별", "월", "월간")):
        at = _open(unit, push_consent_df=_push_fixture())
        txt = " ".join(re.sub(r"\s+", " ", str(e.value))
                       for e in list(at.caption) + list(at.info) + list(at.markdown))
        tex = " ".join(str(e.value) for e in at.latex)
        cols = " ".join(str(c) for d in at.dataframe
                        for c in getattr(d.value, "columns", []))
        # ① 조사 — 받침 있는 단위에 '가'가 붙으면 안 된다
        assert f"기준 {uname}가 " not in txt or uname[-1] in "자차", \
            f"{unit}: '기준 {uname}가'로 찍혔어요 (받침이 있으면 '이')"
        if uname == "월":
            assert "기준 월이 " in txt, f"{unit}: '기준 월이'가 안 보여요"
            assert "기준 월가" not in txt, f"{unit}: '기준 월가'가 남았어요"
        # ② 세는 말 — '7개 월'이 아니라 '7개월'
        assert not re.search(r"\d+개 월\b", txt), f"{unit}: '개 월'로 띄어 썼어요"
        # ③ 산출식은 캡션도 수식도 단위를 따라가야 한다
        assert "기준주·전주 거래액" not in txt, f"{unit}: LMDI 캡션에 '기준주·전주'가 남았어요"
        assert "0=전주·1=기준주" not in txt, f"{unit}: 믹스 캡션에 '전주·기준주'가 남았어요"
        assert r"\text{기준주}" not in tex, f"{unit}: 수식 첨자에 '기준주'가 남았어요"
        assert r"\text{전주}" not in tex or unit == "주별", \
            f"{unit}: 수식 첨자에 '전주'가 남았어요"
        # ④ 단위를 형용사로 쓰는 칼럼 — '주간 신규추가'가 일·월에서도 '주간'이었다
        assert "신규추가" in cols, f"{unit}: 앱푸시 수신동의 표가 안 떴어요 (픽스처 문제)"
        assert f"{peradj} 신규추가" in cols, \
            f"{unit}: 칼럼이 '{peradj} 신규추가'가 아니에요 — {cols[:200]}"
        # ⑤ 믹스 분해 주석의 '이번 주'도 단위를 따라간다
        assert "이번 주 새로 시작" not in txt, f"{unit}: 믹스 주석에 '이번 주'가 남았어요"


@case
def t_trend_values_match_a_direct_sum():
    """추이 표의 값도 그 기간을 직접 센 합과 같아야 한다.

    추이는 카드·표와 **다른 경로**로 값을 읽는다 — 기간별 `groupby`를 한 번 만들어
    거기서 집는다. 그 인덱스가 한 칸만 어긋나도 추이 숫자가 통째로 밀리는데, 위 카드는
    멀쩡하니 눈으로는 안 잡힌다. 기준 기간만 대조하는 검사로는 실제로 못 잡았다."""
    t = _trend(_open("일별"))
    labs = [c[1] for c in t.columns if c[0] == "실적"]
    assert len(labs) >= 3, labs
    d = STORE.copy()
    d["dt"] = pd.to_datetime(d["date"], format="%Y%m%d")
    for lb in (labs[0], labs[len(labs) // 2], labs[-1]):
        m = re.match(r"(\d{4})년 (\d{1,2})/(\d{1,2})", lb)
        assert m, f"일별 라벨 모양이 바뀌었어요 — {lb}"
        day = pd.Timestamp(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        want = float(d[d["dt"] == day]["send"].sum())
        assert want > 0, f"{lb}: 픽스처에 그날 발송이 없어 규칙을 반증하지 못해요"
        got = _num(t.loc["발송", ("실적", lb)])
        assert abs(got - want) < 1.0, f"{lb}: 화면 {got:,.0f} vs 직접 합 {want:,.0f}"


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
