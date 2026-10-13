"""주간보고 「주간보고 요약」 상단 — 필터 바 · 실적 요약 타일 · KPI 목표 · 메모.

공동 작업 저장소(`ryunj/first-dashboard`)의 화면을 옮긴 블록이다. 조용히 틀릴 자리가
넷이라 화면에 찍힌 값을 읽어 대조한다.

1. 「누계」는 가산 지표만 일수를 곱한다. 비율까지 곱하면 가입율이 몇 배로 뛰는데
   화면은 멀쩡히 뜬다.
2. 진행 중인 기간을 누계로 맞대면 비교 기간도 **같은 일수까지만** 센다. 안 자르면
   4일치가 7일치와 맞붙어 가짜 급락이 뜬다.
3. KPI 누적은 월별 일평균 × 일수다. 진행 중인 달은 지난 날까지만 곱한다.
4. 메모는 **날짜가 겹치는** 기간에서 보인다. 주로 써 둔 메모가 월 단위에서 사라지면
   같은 기간을 보는데 메모가 안 보인다.

픽스처는 날짜를 **고정**한다(마지막 일자 2026-09-10). 오늘을 끝으로 잡으면 월말·
일요일마다 «진행 중» 여부가 바뀌어 같은 코드가 날마다 다른 답을 낸다.

로컬 실행:
    python tests/test_summary_top.py
"""
import datetime
import json
import os
import pathlib
import re
import shutil
import sys
import tempfile

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
APP = ROOT / "weekly_report.py"
TIMEOUT = 300
PAGE = "01. 주간보고 요약"
LAST = datetime.date(2026, 9, 10)            # 목요일 · 월말 아님 — 진행 중인 주·달이 생긴다

STORE_COLS = ["gran", "metric", "segment", "year", "label", "close", "sortkey", "value"]
# 연도별 일평균 — 두 해를 일부러 다르게 심어야 전년비가 «값»으로 보인다
VAL = {
    2026: {"첫구매 거래액": 2e8, "첫구매 고객수": 2000, "비회원트래픽": 100_000,
           "가입자수": 1000, "첫구매 객단가": 100_000, "가입율": 0.5, "당일가입CR": 0.05},
    2025: {"첫구매 거래액": 1e8, "첫구매 고객수": 1000, "비회원트래픽": 100_000,
           "가입자수": 500, "첫구매 객단가": 100_000, "가입율": 0.5, "당일가입CR": 0.04},
}
AUG26_REV = 1.5e8                            # 2026-08만 다르게 — 「직전 기간」이 값으로 갈린다


def _v(year, met, mo=None):
    if year == 2026 and mo == 8 and met == "첫구매 거래액":
        return AUG26_REV
    return VAL[year][met]


def _wk(d):
    """목요일이 속한 달의 N번째 주 — 앱과 «다른 경로»로 만든다(같은 함수를 쓰면 대조가 안 된다)."""
    thu = d + datetime.timedelta(days=3 - d.weekday())
    return thu.year, f"{thu.month:02d}월 {(thu.day - 1) // 7 + 1}주차", \
        thu.year * 10000 + thu.month * 100 + (thu.day - 1) // 7 + 1


def synth():
    rows = []
    for y in (2025, 2026):
        for mo in range(1, 13 if y == 2025 else 10):
            cl = "mtd" if (y == 2026 and mo == 9) else "final"
            for met in VAL[y]:
                rows.append(dict(gran="월", metric=met, segment="*TOTAL", year=y,
                                 label=f"{mo}월", close=cl, sortkey=y * 10000 + mo * 100,
                                 value=_v(y, met, mo)))
    d = datetime.date(2025, 1, 6)
    while d <= LAST:
        y, lb, sk = _wk(d)
        cl = "mtd" if d + datetime.timedelta(days=6) > LAST else "final"
        for met in VAL[y]:
            rows.append(dict(gran="주", metric=met, segment="*TOTAL", year=y, label=lb,
                             close=cl, sortkey=sk, value=VAL[y][met]))
        d += datetime.timedelta(days=7)
    for y in (2025, 2026):
        d = datetime.date(y, 8, 1)
        while d <= LAST.replace(year=y):
            for met in VAL[y]:
                rows.append(dict(gran="일", metric=met, segment="*TOTAL", year=y,
                                 label=f"{d.month}/{d.day}", close="final",
                                 sortkey=y * 10000 + d.month * 100 + d.day,
                                 value=_v(y, met, d.month)))
            d += datetime.timedelta(days=1)
    return pd.DataFrame(rows)[STORE_COLS]


def _open(insights=None, **ss):
    from streamlit.testing.v1 import AppTest
    tmp = tempfile.mkdtemp()
    shutil.copy(APP, os.path.join(tmp, "weekly_report.py"))
    for mod in ("table_export.py", "chart_layout.py", "note_markdown.py", "note_editor.py"):
        if (ROOT / mod).exists():
            shutil.copy(ROOT / mod, os.path.join(tmp, mod))
    synth().to_csv(os.path.join(tmp, "wr_data_store.csv"), index=False, encoding="utf-8-sig")
    if insights:
        with open(os.path.join(tmp, "wr_insights.json"), "w", encoding="utf-8") as f:
            json.dump(insights, f, ensure_ascii=False)
    # 스토어가 cwd 기준 CSV라 테스트가 끝날 때까지 임시 폴더에 머문다(러너가 되돌린다)
    os.chdir(tmp)
    at = AppTest.from_file(os.path.join(tmp, "weekly_report.py"), default_timeout=TIMEOUT)
    for k, v in ss.items():
        at.session_state[k] = v
    at.run()
    assert not at.exception, at.exception[0].value
    rad = [r for r in at.sidebar.radio if r.label == "페이지"]
    assert rad and rad[0].value == PAGE, "첫 화면이 요약이 아니에요"
    return at


def _html(at):
    return "\n".join(str(m.value) for m in at.markdown)


def _tile(at, met):
    """(값, [(비교 이름, 증감)…], (기준 이름, 기준 값)) — 화면에 찍힌 그대로."""
    h = _html(at)
    m = re.search(r'<div class="st-l">' + re.escape(met) + r'</div><div class="st-v">([^<]*)</div>'
                  r'((?:<div class="st-r">.*?</div>)+)</div>', h)
    assert m, f"{met} 타일이 없어요"
    rs = re.findall(r'<div class="st-r"><span>([^<]*)</span><(?:b|span)[^>]*>([^<]*)</(?:b|span)>',
                    m.group(2))
    return m.group(1), rs[:-1], rs[-1]


def _caps(at):
    return " ".join(str(c.value) for c in at.caption)


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_tiles_render_every_metric_in_the_sidebar_week():
    """기본은 주별 · 사이드바 기준 주차다(진행 중이면 직전 완결 주)."""
    at = _open()
    for met in VAL[2026]:
        _tile(at, met)
    v, deltas, base = _tile(at, "첫구매 거래액")
    assert v == "200.0백만원", v
    assert deltas[0] == ("전년비", "+100.0%"), deltas
    # 직전 주(2026-08 4주차)도 2억이라 +0.0%다 — 전년 주와 맞대면 +100%로 갈린다
    assert deltas[1] == ("전주비", "+0.0%"), deltas
    assert base == ("전년", "100.0백만원"), base
    hs = [str(m.value) for m in at.markdown if str(m.value).startswith("#### 실적 요약")]
    assert hs and "2026년 9월 1주차" in hs[0], hs


@case
def t_sum_mode_multiplies_only_additive_metrics():
    """누계는 가산 지표만 일수를 곱한다 — 가입율·객단가는 그대로여야 한다."""
    a = _open(wr_refw="09월 1주차")
    b = _open(wr_refw="09월 1주차", wr_sum_mode="누계")
    assert _tile(a, "첫구매 거래액")[0] == "200.0백만원"
    assert _tile(b, "첫구매 거래액")[0] == "1,400.0백만원", _tile(b, "첫구매 거래액")[0]
    assert _tile(b, "가입자수")[0] == "7,000명", _tile(b, "가입자수")[0]
    for met in ("가입율", "첫구매 객단가", "당일가입CR"):
        assert _tile(a, met)[0] == _tile(b, met)[0], (met, _tile(a, met)[0], _tile(b, met)[0])
    # 가입율은 계산값(가입자수 ÷ 트래픽)이다 — 파일 값(50%)이면 규칙이 갈린 것이다
    assert _tile(a, "가입율")[0] == "1.00%", _tile(a, "가입율")[0]


@case
def t_partial_week_compares_the_same_number_of_days():
    """진행 중인 주(9/7~9/10, 4일)를 누계로 보면 전년도 4일치까지만 센다.

    안 자르면 2억×4 ↔ 1억×7로 맞붙어 +100%가 +14.3%로 줄어든다.
    """
    at = _open(wr_refw="09월 2주차", wr_sum_mode="누계")
    v, deltas, base = _tile(at, "첫구매 거래액")
    assert v == "800.0백만원", v
    assert deltas[0] == ("전년비", "+100.0%"), deltas
    assert base == ("전년", "400.0백만원"), base
    assert "4일치까지 집계했어요" in _caps(at), "진행 중이라는 걸 화면이 안 밝혀요"


@case
def t_partial_month_previous_period_is_capped():
    """9월(10일치) 누계를 직전 달과 맞대면 8월도 10일치까지만 센다(2억×10 ↔ 1.5억×10)."""
    at = _open(wr_sum_unit="월", wr_sum_mode="누계", wr_sum_cmp="직전 기간")
    v, deltas, base = _tile(at, "첫구매 거래액")
    assert v == "2,000.0백만원", v
    assert deltas[0] == ("전월비", "+33.3%"), deltas
    assert deltas[1][0] == "전년비", deltas
    assert base == ("전월", "1,500.0백만원"), base


@case
def t_comparison_names_follow_the_unit():
    at = _open(wr_sum_unit="일", wr_sum_cmp="직전 기간")
    _v0, deltas, base = _tile(at, "첫구매 거래액")
    assert deltas[0][0] == "전일비" and base[0] == "전일", (deltas, base)
    assert "364일 전 같은 요일" not in _caps(at)


def _kpi_rows(at, title):
    h = _html(at)
    m = re.search(r'<div class="kp-t">' + re.escape(title) + r'</div>(.*?)</div></div></div>', h)
    assert m, f"{title} KPI 카드가 없어요"
    return dict(re.findall(r'<span>([^<]*)</span><b[^>]*>([^<]*)</b>', m.group(1)))


@case
def t_kpi_accumulates_monthly_daily_means():
    """누적 = 1~7월 2억×212일 + 8월 1.5억×31일 + 9월 2억×10일 = 490.5억.

    목표를 981억으로 두면 달성률이 정확히 50.0%다. 9월을 30일치로 곱하면 51.8%,
    8월 값을 안 갈라 읽으면 51.6%로 갈린다.
    """
    at = _open(insights={"wr_kpi_targets::2026": json.dumps({"첫구매 거래액": 9.81e10})})
    r = _kpi_rows(at, "첫구매 거래액")
    assert r.get("달성률") == "50.0%", r
    assert r.get("연간 목표") == "98,100백만원", r
    assert "253일 경과" in _caps(at), "경과일이 안 맞아요"
    # 목표 없는 지표는 '미입력'이다 — 0으로 나누거나 지어내지 않는다
    assert _kpi_rows(at, "가입자수").get("연간 목표") == "미입력"


def _memo(gran, year, label, review):
    return json.dumps(dict(gran=gran, year=year, label=label, review=review, plan="",
                           created="2026-09-01T09:00", updated="2026-09-01T09:00"),
                      ensure_ascii=False)


MEMOS = {"wr_memo2::a": _memo("주", 2026, "09월 1주차", "주 메모"),
         "wr_memo2::b": _memo("월", 2026, "3월", "봄 메모")}


@case
def t_memo_shows_where_the_dates_overlap():
    at = _open(insights=MEMOS, wr_refw="09월 1주차")
    h = _html(at)
    assert "주 메모" in h and "봄 메모" not in h, "이 화면 조건에 겹치는 메모만 보여야 해요"
    # 같은 기간을 월 단위로 보면 주 메모도 보여야 한다(8/31~9/6이 9월과 겹친다)
    at = _open(insights=MEMOS, wr_refw="09월 1주차", wr_sum_unit="월")
    h = _html(at)
    assert "주 메모" in h and "봄 메모" not in h, "월 단위에서 주 메모가 사라졌어요"
    at = _open(insights=MEMOS, wr_memo2_filter="모든 메모")
    h = _html(at)
    assert "주 메모" in h and "봄 메모" in h


@case
def t_pin_moves_the_summary_to_the_memo_period():
    at = _open(insights=MEMOS, wr_memo2_filter="모든 메모")
    at.button(key="wr_memo2_go_b").click().run()
    assert not at.exception, at.exception[0].value
    hs = [str(m.value) for m in at.markdown if str(m.value).startswith("#### 실적 요약")]
    assert hs and "2026년 3월" in hs[0], hs


@case
def t_memo_save_writes_the_insight_file():
    at = _open(wr_refw="09월 1주차")
    at.button(key="wr_memo2_new").click().run()
    assert not at.exception, at.exception[0].value
    at.text_area(key="wr_memo2_rv_new").input("트래픽이 빠졌어요").run()
    at.text_area(key="wr_memo2_pl_new").input("다음 주 쿠폰 재발송").run()
    at.button(key="wr_memo2_save_new").click().run()
    assert not at.exception, at.exception[0].value
    with open("wr_insights.json", encoding="utf-8") as f:
        saved = json.load(f)
    got = [json.loads(v) for k, v in saved.items() if k.startswith("wr_memo2::")]
    assert len(got) == 1, saved
    assert got[0]["review"] == "트래픽이 빠졌어요" and got[0]["plan"] == "다음 주 쿠폰 재발송"
    assert (got[0]["gran"], got[0]["label"]) == ("주", "09월 1주차"), got[0]
    assert "트래픽이 빠졌어요" in " ".join(str(m.value) for m in at.markdown)


def main():
    cwd = os.getcwd()
    fails = 0
    for fn in CASES:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                                   # noqa: BLE001
            fails += 1
            print(f"  FAIL {fn.__name__}: {e}")
        finally:
            os.chdir(cwd)
    if fails:
        print(f"\n요약 상단 테스트 {fails}건 실패 ❌")
        sys.exit(1)
    print(f"\n요약 상단 테스트 {len(CASES)}건 통과 ✅")


if __name__ == "__main__":
    main()
