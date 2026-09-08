"""「02. 첫구매 퍼널별 상세 실적」 렌더·계산 테스트.

이 화면은 **비율 두 칸의 출처가 다르다**. 가입율은 두 카운트에서 계산하고, 당일가입CR은
파일 값을 그대로 쓴다(분자인 당일가입 첫구매 고객수가 데이터에 없다). 둘을 한쪽으로
통일해 버리면 화면 숫자가 통째로 바뀌는데, 증상이 '값이 좀 다르네'로만 보여서 스모크로는
절대 안 잡힌다. 그래서 파일 값과 계산 값을 **일부러 다르게** 심어 두고 어느 쪽이 나오는지
본다.

로컬 실행:
    python tests/test_funnel_page.py
"""
import os
import pathlib
import re
import shutil
import sys
import tempfile

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import weekly_report as W                        # noqa: E402

APP = ROOT / "weekly_report.py"
TIMEOUT = 300
PAGE = "02. 첫구매 퍼널별 상세 실적"

STORE_COLS = ["gran", "metric", "segment", "year", "label", "close", "sortkey", "value"]
YEARS = (2025, 2026)
MONTHS = range(1, 10)
WEEKS = range(1, 5)

# 연도별 원값 — 채널 두 개(직접·광고)가 전체와 다른 비율을 갖도록 일부러 갈라 놨다.
#   전체 가입율 2026 = 1000/100000 = 1.00% · 직접 = 800/40000 = 2.00% · 광고 = 200/60000 = 0.33%
# 파일이 주는 「가입율」 지표는 50%로 심어 둔다 — 화면이 파일 값을 쓰면 대번에 드러난다.
VALS = {
    2026: {"*TOTAL": dict(traffic=100_000, join=1_000, cust=500, rev=50_000_000),
           "직접":   dict(traffic=40_000, join=800, cust=300, rev=30_000_000),
           "광고":   dict(traffic=60_000, join=200, cust=200, rev=20_000_000)},
    2025: {"*TOTAL": dict(traffic=80_000, join=500, cust=400, rev=32_000_000),
           "직접":   dict(traffic=30_000, join=400, cust=250, rev=20_000_000),
           "광고":   dict(traffic=50_000, join=100, cust=150, rev=12_000_000)},
}
FILE_RATE = 0.5           # 파일이 주는 가입율 — 계산값(1.00%)과 확연히 다르게
DAILY_CR = 0.0725         # 당일가입CR — 역산이 불가능한 값이라 파일 값이 그대로 나와야
FILE_AOV = 111_111        # 파일이 주는 객단가 — 거래액/고객수(100,000원)와 다르게
PUSH_DAILY = {1: 100.0, 8: 200.0, 15: 300.0, 22: 400.0}    # 일평균 250


def _rows_for(gran, label, sortkey, year):
    out = []

    def add(met, seg, v):
        out.append(dict(gran=gran, metric=met, segment=seg, year=year, label=label,
                        close="final", sortkey=sortkey, value=float(v)))
    for seg, d in VALS[year].items():
        add("비회원트래픽", seg, d["traffic"])
        add("가입자수", seg, d["join"])
        add("첫구매 고객수", seg, d["cust"])
        add("첫구매 거래액", seg, d["rev"])
        add("가입율", seg, FILE_RATE)
        add("당일가입CR", seg, DAILY_CR)
        add("첫구매 객단가", seg, FILE_AOV)
    return out


def synth_store(with_push=True):
    rows = []
    for y in YEARS:
        for mo in MONTHS:
            rows += _rows_for("월", f"{mo}월", y * 10000 + mo * 100, y)
            for wk in WEEKS:
                rows += _rows_for("주", f"{mo:02d}월 {wk}주차",
                                  y * 10000 + mo * 100 + wk, y)
            if with_push:
                # 앱푸시 수신동의는 원천이 일자 헤더 표라 **일별로만** 쌓인다
                for dd, v in PUSH_DAILY.items():
                    rows.append(dict(gran="일", metric="앱푸시수신동의", segment="*TOTAL",
                                     year=y, label=f"{mo}/{dd}", close="final",
                                     sortkey=y * 10000 + mo * 100 + dd, value=v))
    return pd.DataFrame(rows)[STORE_COLS]


TREE = {"e-영업1": ["가방", "지갑"], "e-영업2": ["슈즈"]}


def synth_orgcat():
    """조직 > 카테고리 2단 — 브랜드·상품 칸은 지금 실제 export처럼 비워 둔다."""
    rows = []

    def add(gran, label, sortkey, year, org, cat, met, v):
        rows.append({"gran": gran, "metric": met, "org": org, "cat": cat,
                     "brand": "", "item": "", "lfms": "N", "year": year,
                     "label": label, "close": "final", "sortkey": sortkey,
                     "value": float(v)})
    for y in YEARS:
        k = 1.0 if y == 2026 else 0.8
        for gran, labels in (("월", [(f"{m}월", m * 100) for m in MONTHS]),
                             ("주", [(f"{m:02d}월 {w}주차", m * 100 + w)
                                     for m in MONTHS for w in WEEKS])):
            for label, sk in labels:
                s = y * 10000 + sk
                add(gran, label, s, y, "*TOTAL", "*TOTAL", "첫구매 거래액", 50e6 * k)
                add(gran, label, s, y, "*TOTAL", "*TOTAL", "첫구매 고객수", 500 * k)
                base = 30e6
                for org, cats in TREE.items():
                    add(gran, label, s, y, org, "*TOTAL", "첫구매 거래액", base * k)
                    add(gran, label, s, y, org, "*TOTAL", "첫구매 고객수", 300 * k)
                    part = base / len(cats)
                    for cat in cats:
                        add(gran, label, s, y, org, cat, "첫구매 거래액", part * k)
                        add(gran, label, s, y, org, cat, "첫구매 고객수", 150 * k)
                    base = 20e6
    return pd.DataFrame(rows)[W.ORGCAT_COLS]


def _open(store=None, orgcat=None, mode=None):
    """임시 폴더에 스토어를 깔고 앱을 띄운 뒤 새 페이지로 이동한다."""
    from streamlit.testing.v1 import AppTest
    tmp = tempfile.mkdtemp()
    shutil.copy(APP, os.path.join(tmp, "weekly_report.py"))
    _te = ROOT / "table_export.py"
    if _te.exists():
        shutil.copy(_te, os.path.join(tmp, "table_export.py"))
    (synth_store() if store is None else store).to_csv(
        os.path.join(tmp, W.DATA_STORE), index=False, encoding="utf-8-sig")
    oc = synth_orgcat() if orgcat is None else orgcat
    if oc is not None and not oc.empty:
        oc.to_csv(os.path.join(tmp, W.ORGCAT_STORE), index=False, encoding="utf-8-sig")
    # 스토어가 **cwd 기준 CSV**라 여기서 나오면 안 된다. 테스트가 끝날 때까지 임시 폴더에
    # 머물고, 되돌리는 건 러너(main)가 맡는다 — 중간에 돌아오면 다음 at.run()부터
    # 스토어가 통째로 사라져 '누적 데이터가 없어요' 화면을 보고 통과해 버린다.
    os.chdir(tmp)
    at = AppTest.from_file(os.path.join(tmp, "weekly_report.py"), default_timeout=TIMEOUT)
    at.run()
    assert not at.exception, at.exception[0].value
    rad = [r for r in at.sidebar.radio if r.label == "페이지"]
    assert rad, "페이지 라디오를 못 찾았어요"
    rad[0].set_value(PAGE); at.run()
    assert not at.exception, at.exception[0].value
    if mode:
        cmp_r = [r for r in at.radio if r.label == "비교 기준"]
        assert cmp_r, f"비교 기준 라디오가 없어요 — {[r.label for r in at.radio]}"
        cmp_r[0].set_value(mode); at.run()
        assert not at.exception, at.exception[0].value
    return at


def _html(at):
    return "\n".join(str(m.value) for m in at.markdown)


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success,
                 at.title, at.header, at.subheader):
        out += [str(e.value) for e in coll]
    return out


def _kpi(at, label):
    m = re.search(r'kpi-label">' + re.escape(label) + r'</div><div class="kpi-value">([^<]*)</div>',
                  _html(at))
    return m.group(1) if m else None


def _rate(at, label):
    m = re.search(r'>' + re.escape(label) + r'</div>\s*<div[^>]*font-weight:700[^>]*>([^<]*)</div>',
                  _html(at))
    return m.group(1) if m else None


def _tail(at, label):
    m = re.search(r'>' + re.escape(label) + r' <b[^>]*>([^<]*)</b>', _html(at))
    return m.group(1) if m else None


def _frames(at):
    out = []
    for t in at.dataframe:
        v = getattr(t.value, "data", t.value)
        if hasattr(v, "columns"):
            out.append(v)
    return out


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# ── 등록·렌더 ───────────────────────────────────────────────────────
@case
def t_page_is_registered_second():
    """요약 바로 다음 자리여야 한다 — 개괄에서 이 화면으로 내려오는 흐름이라서."""
    at = _open()
    opts = list([r for r in at.sidebar.radio if r.label == "페이지"][0].options)
    assert PAGE in opts, f"페이지가 없어요 — {opts}"
    assert opts.index(PAGE) == 1, f"두 번째 자리가 아니에요 — {opts}"


@case
def t_all_blocks_render():
    at = _open()
    txt = " ".join(_texts(at))
    for mark in ("① ", "② ", "③ ", "④ ", "⑤ "):
        assert mark in txt, f"{mark} 블록이 안 보여요"


@case
def t_mtd_mode_renders():
    at = _open(mode="월누적(MTD) — 전년 동월")
    txt = " ".join(_texts(at))
    assert "① " in txt and "④ " in txt, "MTD 모드에서 블록이 빠졌어요"
    assert _kpi(at, "가입자수"), "MTD 모드에서 퍼널 카드가 안 그려졌어요"


@case
def t_funnel_left_page_01():
    """01에 그대로 남아 있으면 같은 화면이 두 벌이 된다."""
    at = _open()
    [r for r in at.sidebar.radio if r.label == "페이지"][0].set_value("01. 주간보고 요약")
    at.run()
    assert not at.exception, at.exception[0].value
    txt = " ".join(_texts(at))
    assert "전환 퍼널" not in txt, "01에 전환 퍼널이 남아 있어요"
    assert "채널 기여 분해" not in txt, "01에 채널 기여 분해가 남아 있어요"


# ── 비율 두 칸의 출처 ───────────────────────────────────────────────
@case
def t_join_rate_is_computed_not_read_from_file():
    """가입율은 `가입자수 ÷ 비회원트래픽`이다.

    파일의 「가입율」 지표를 50%로 심어 뒀다 — 그 값이 화면에 뜨면 계산이 아니라 파일을
    읽고 있다는 뜻이고, 채널로 쪼갤 때 합계와 규칙이 갈린다.
    """
    at = _open()
    got = _rate(at, "가입율")
    assert got == "1.00%", f"가입율이 계산값이 아니에요 — {got} (파일 값이면 50.00%)"


@case
def t_daily_cr_comes_from_the_file():
    """당일가입CR은 분자가 데이터에 없어 역산이 불가능하다 — 파일 값 그대로여야 한다."""
    at = _open()
    got = _rate(at, "당일가입 첫구매율")
    assert got == f"{DAILY_CR * 100:.2f}%", f"당일가입CR이 파일 값이 아니에요 — {got}"


@case
def t_aov_keeps_the_file_value():
    """객단가는 「01」 KPI 카드와 같은 숫자여야 한다 — 거래액÷고객수로 다시 만들지 않는다."""
    at = _open()
    got = _tail(at, "객단가")
    assert got == f"{FILE_AOV:,}원", f"객단가가 파일 값이 아니에요 — {got} (역산이면 100,000원)"


@case
def t_funnel_does_not_pretend_to_multiply_out():
    """`가입자수 × 당일가입CR ≠ 첫구매 고객수`라는 걸 화면이 밝혀야 한다.

    합성값으로 1,000 × 7.25% = 72.5명인데 첫구매 고객수는 500명이다. 안 밝히면
    보는 사람이 산식이 맞는 줄 알고 읽는다.
    """
    at = _open()
    assert _kpi(at, "첫구매 고객수") == "500명", _kpi(at, "첫구매 고객수")
    txt = " ".join(_texts(at))
    assert "전체 첫구매" in txt and "안 맞아요" in txt, "곱셈이 안 닫힌다는 안내가 없어요"


@case
def t_counts_and_yoy_are_right():
    at = _open()
    assert _kpi(at, "비회원트래픽") == "100,000명", _kpi(at, "비회원트래픽")
    assert _kpi(at, "가입자수") == "1,000명", _kpi(at, "가입자수")
    # 가입자수 500 → 1,000 = +100.0%
    assert "+100.0% (전년동주)" in _html(at), "가입자수 전년비가 틀려요"


# ── ② 채널별 상세 ───────────────────────────────────────────────────
@case
def t_channel_table_covers_every_step():
    at = _open()
    fr = [f for f in _frames(at) if "비회원트래픽" in f.columns]
    assert fr, f"채널 표가 없어요 — {[list(f.columns) for f in _frames(at)]}"
    tbl = fr[0]
    for stp in W.FUNNEL_STEPS:
        assert stp in tbl.columns, f"{stp} 칸이 없어요"
        assert f"{stp} 전년비" in tbl.columns, f"{stp} 전년비 칸이 없어요"
    assert list(tbl.index)[:3] == ["전체", "직접", "광고"], list(tbl.index)


@case
def t_channel_rates_are_per_channel():
    """채널별 가입율은 **그 채널의 두 카운트**로 만들어야 한다.

    전체 값을 그대로 복사하면 세 줄이 전부 1.00%가 되어 '어느 채널이 새는지'를 못 본다.
    """
    at = _open()
    tbl = [f for f in _frames(at) if "비회원트래픽" in f.columns][0]
    got = {ix: tbl.loc[ix, "가입율"] for ix in ("전체", "직접", "광고")}
    assert got == {"전체": "1.00%", "직접": "2.00%", "광고": "0.33%"}, got


@case
def t_channel_table_carries_the_rate_caveat():
    at = _open()
    assert any("채널 합이 전체와 다른 게 정상" in t for t in _texts(at)), \
        "비율 지표 주의 문구가 없어요"


# ── ③ 채널 기여 분해 ────────────────────────────────────────────────
@case
def t_decomposition_offers_additive_only():
    """비율 지표는 채널 합 ≠ 전체라 분해가 성립하지 않는다 — 선택지에 있으면 안 된다."""
    at = _open()
    box = [s for s in at.selectbox if s.label == "분해할 단계"]
    assert box, f"분해 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
    opts = list(box[0].options)
    assert opts, "분해 지표가 하나도 없어요"
    for bad in ("가입율", "당일가입CR", "첫구매 객단가"):
        assert bad not in opts, f"비율 지표 «{bad}»가 분해 선택지에 있어요 — {opts}"
    for good in ("비회원트래픽", "가입자수", "첫구매 고객수", "첫구매 거래액"):
        assert good in opts, f"«{good}»가 빠졌어요 — {opts}"


@case
def t_decomposition_draws_a_waterfall():
    at = _open()
    assert at.get("plotly_chart"), "워터폴 차트가 없어요"


# ── ④ 조직 > 카테고리 ───────────────────────────────────────────────
@case
def t_orgcat_lists_orgs_then_categories():
    at = _open()
    box = [s for s in at.selectbox if s.label == "1. 조직"]
    assert box, f"조직 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
    assert set(TREE) <= set(box[0].options), list(box[0].options)
    fr = [f for f in _frames(at) if f.index.name == "조직"]
    assert fr, "조직 표가 없어요"
    assert set(TREE) <= set(fr[0].index), list(fr[0].index)

    box[0].set_value("e-영업1"); at.run()
    assert not at.exception, at.exception[0].value
    fr = [f for f in _frames(at) if f.index.name == "카테고리"]
    assert fr, "조직을 골랐는데 카테고리 표가 안 나와요"
    assert set(TREE["e-영업1"]) == set(fr[0].index), list(fr[0].index)


@case
def t_orgcat_share_only_for_additive():
    """하위 합이 상위와 맞는 건 거래액뿐이다 — 고객수에 비중을 붙이면 거짓말이 된다."""
    at = _open()
    box = [s for s in at.selectbox if s.label == "지표"]
    assert box, f"지표 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
    assert box[0].value == "첫구매 고객수" or "첫구매 거래액" in box[0].options
    box[0].set_value("첫구매 거래액"); at.run()
    fr = [f for f in _frames(at) if f.index.name == "조직"][0]
    assert "비중" in fr.columns, f"거래액엔 비중이 있어야 해요 — {list(fr.columns)}"
    box = [s for s in at.selectbox if s.label == "지표"][0]
    box.set_value("첫구매 고객수"); at.run()
    fr = [f for f in _frames(at) if f.index.name == "조직"][0]
    assert "비중" not in fr.columns, f"고객수엔 비중이 없어야 해요 — {list(fr.columns)}"
    assert any("하위 합이 상위와 안 맞는" in t for t in _texts(at)), "왜 뺐는지 안 밝혔어요"


@case
def t_orgcat_missing_source_explains_itself():
    at = _open(orgcat=pd.DataFrame())
    assert any("조직×카테고리 데이터가 없어요" in t for t in _texts(at)), \
        "원천이 없을 때 안내가 없어요"
    assert not at.exception, at.exception[0].value


# ── ⑤ 앱 수신동의 ───────────────────────────────────────────────────
@case
def t_push_period_avg_reads_daily_rows():
    """앱푸시는 일별로만 쌓인다 — 월은 정확히, 주는 근사로 묶는다."""
    d = synth_store()
    v, approx = W.push_period_avg(d, 2026, "9월")
    assert abs(v - np.mean(list(PUSH_DAILY.values()))) < 1e-6, v
    assert approx is False, "월은 근사가 아니에요"
    v2, approx2 = W.push_period_avg(d, 2026, "09월 1주차")
    assert approx2 is True, "주는 근사라고 알려야 해요"
    assert abs(v2 - PUSH_DAILY[1]) < 1e-6, v2      # 1~7일 = 1일치뿐
    assert pd.isna(W.push_period_avg(d, 2026, "없는라벨")[0])


@case
def t_app_block_shows_consent_rate():
    at = _open()
    fr = [f for f in _frames(at) if f.index.name == "비율"]
    assert fr, f"앱 비율 표가 없어요 — {[f.index.name for f in _frames(at)]}"
    tbl = fr[0]
    assert "신규회원 앱 수신동의율" in tbl.index, list(tbl.index)
    # 주간 기본 모드: 그 주 앱푸시 일평균 ÷ 가입자수
    assert tbl.loc["신규회원 앱 수신동의율", "2026년"] != "–", tbl.to_dict()


@case
def t_missing_app_install_says_why():
    """앱설치 원천이 없으면 빈 줄만 두지 말고 왜 비었는지 말해야 한다."""
    at = _open()
    assert any("앱설치" in t and "안 올라와서" in t for t in _texts(at)), \
        "앱설치 부재 안내가 없어요"


@case
def t_app_install_file_is_recognized():
    """올릴 예정인 원천 — 파일명만으로 지표가 잡혀야 한다."""
    for nm in ("월_앱설치(일평균).xlsx", "주_신규앱설치.xlsx", "일_앱다운로드.csv"):
        kind, gran, met = W.detect_file(nm)
        assert (kind, met) == ("metric", "앱설치"), f"{nm} → {(kind, gran, met)}"
    assert W.METRIC_UNIT["앱설치"] == ("명", 1)


def main():
    fails, cwd = [], os.getcwd()
    for fn in CASES:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {str(e)[:300]}")
            fails.append(fn.__name__)
        finally:
            os.chdir(cwd)              # _open이 임시 폴더에 머무르므로 여기서 되돌린다
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"첫구매 퍼널 페이지 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
