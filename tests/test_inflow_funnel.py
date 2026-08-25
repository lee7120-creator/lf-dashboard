"""「6. 효율·피로도 › 유입 퍼널」 렌더·계산 테스트.

이 탭은 전사 MTD가 있어야 열린다 — 스모크는 캠페인 데이터만 넣으므로 이 경로를
통째로 안 밟는다. 단계 비율은 나눗셈 덩어리라 0 분모·전부 결측에서 조용히 죽거나
inf를 화면에 그대로 뿌리기 쉽고, 유니크유입 파생 지표는 `prepare_raw` 때와 같은
캐시 표식 문제(=옛 프레임이 남아 화면이 통째로 빈다)를 그대로 반복할 수 있다.

로컬 실행:
    python tests/test_inflow_funnel.py
"""
import io
import pathlib
import sys

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest        # noqa: E402

import send_perf_dashboard as S                 # noqa: E402
from smoke_pages import synth_store             # noqa: E402
from test_send_volume_band import synth_mtd     # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300
TAB = "유입 퍼널"


def _open_tab(camp=None, mtd=None):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=10) if camp is None else camp
    if mtd is not None:
        at.session_state["mtd_store_df"] = mtd
    at.run()
    assert not at.exception, at.exception[0].value
    at.sidebar.radio[0].set_value("6. 효율·피로도")
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    assert subs, "하위탭 라디오를 찾지 못했어요"
    assert TAB in subs[0].options, f"{TAB} 탭이 없어요 — {subs[0].options}"
    subs[0].set_value(TAB)
    at.run()
    assert not at.exception, at.exception[0].value
    return at


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success):
        out += [str(e.value) for e in coll]
    return out


def _cmp_table(at):
    """앞 구간 비교 표 — '단계' 칼럼을 가진 표. (컬럼 2번이 '이번 구간' 값)"""
    for t in at.dataframe:
        cols = list(getattr(t.value, "columns", []))
        if "단계" in cols and "변화" in cols:
            return t.value
    return None


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_tab_exists_in_group():
    """페이지 그룹에 등록돼 있어야 스모크 커버리지에도 자동으로 들어간다."""
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=10)
    at.run()
    at.sidebar.radio[0].set_value("6. 효율·피로도")
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    assert subs and TAB in subs[0].options, "탭이 하위탭 목록에 없어요"


@case
def t_renders_without_mtd():
    """MTD가 없으면 죽지 말고 올려 달라고 안내해야 한다."""
    at = _open_tab()
    assert any("MTD" in t for t in _texts(at)), "MTD 안내 문구가 없어요"


@case
def t_renders_with_mtd():
    """퍼널·추이·비교 표가 다 나와야 한다."""
    at = _open_tab(mtd=synth_mtd())
    txt = " ".join(_texts(at))
    assert "발송 고객" in txt and "유니크 유입" in txt, f"퍼널 카드가 없어요 — {txt[:300]}"
    assert "단계별 전환율 추이" in txt, "추이 블록이 없어요"
    tbl = _cmp_table(at)
    assert tbl is not None, "앞 구간 비교 표가 없어요"
    assert list(tbl.columns)[1:3] == ["전반", "후반"], \
        f"전체 기간이면 전반↔후반으로 대체해야 해요 — {list(tbl.columns)}"
    # AppTest는 plotly 요소를 노출하지 않는다 — 예외 없이 끝까지 돌았는지로 대신 본다
    assert "유입 품질" in txt, "유입 품질 블록이 없어요"


@case
def t_step_rates_use_period_sums():
    """단계 비율은 '기간 합계끼리 나눈 값'이어야 한다 (일별 비율 평균 금지).

    발송량이 들쭉날쭉하면 일별 비율 평균은 소량 발송일을 과대 대표한다.
    두 값이 눈에 띄게 다르도록 발송량을 크게 흔든 데이터로 확인한다."""
    mtd = synth_mtd(days=120, seed=11).copy()
    n = len(mtd)
    # 절반은 소량·고효율, 절반은 대량·저효율 — 합계비율과 일평균비율이 갈리게
    small = np.arange(n) % 2 == 0
    mtd.loc[small, "customers"] = 50_000.0
    mtd.loc[~small, "customers"] = 500_000.0
    mtd.loc[small, "uniqueInflow"] = 50_000.0 * 0.10
    mtd.loc[~small, "uniqueInflow"] = 500_000.0 * 0.02
    exp_sum = mtd["uniqueInflow"].sum() / mtd["customers"].sum()
    exp_mean = (mtd["uniqueInflow"] / mtd["customers"]).mean()
    assert abs(exp_sum - exp_mean) > 0.005, "테스트 데이터가 두 방식을 구분 못 해요"

    at = _open_tab(mtd=mtd)
    # 카드에 찍히는 '발송 고객의 N%'가 곧 선택 기간 전체의 단계 비율이다
    txt = " ".join(_texts(at))
    import re
    m = re.search(r"발송 고객의 (\d+\.\d+)%", txt)
    assert m, f"유입 단계 비율 문구가 없어요 — {txt[:400]}"
    got = float(m.group(1))
    assert abs(got - exp_sum * 100) < 0.05, \
        f"합계 기준({exp_sum*100:.2f}%)이 아니라 {got:.2f}% (일평균은 {exp_mean*100:.2f}%)"


@case
def t_prior_window_does_not_overlap():
    """기간을 좁히면 '직전 같은 기간'과 비교하되, 창이 겹치면 안 된다."""
    mtd = synth_mtd(days=180)
    at = _open_tab(mtd=mtd)
    dts = [d for d in at.date_input if d.label == "조회 기간"]
    assert dts, f"조회 기간 위젯이 없어요 — {[d.label for d in at.date_input]}"
    hi = pd.to_datetime(mtd["date"]).max().date()
    lo = (pd.Timestamp(hi) - pd.Timedelta(days=29)).date()
    dts[0].set_value((lo, hi))
    at.run()
    assert not at.exception, at.exception[0].value
    caps = [c for c in _texts(at) if "직전 같은 기간:" in c]
    assert caps, f"직전 기간 캡션이 없어요 — {_texts(at)[-6:]}"
    import re
    ds = re.findall(r"\d{4}-\d{2}-\d{2}", caps[0])
    assert len(ds) == 2, f"기간을 못 읽었어요 — {caps[0]}"
    p_lo, p_hi = pd.Timestamp(ds[0]), pd.Timestamp(ds[1])
    assert p_hi == pd.Timestamp(lo) - pd.Timedelta(days=1), \
        f"비교 창이 선택 창 바로 앞에서 끝나야 해요 — {p_hi.date()} vs {lo}"
    assert (p_hi - p_lo).days == 29, f"비교 창 길이가 달라요 — {(p_hi - p_lo).days + 1}일"
    tbl = _cmp_table(at)
    assert tbl is not None and list(tbl.columns)[1:3] == ["직전", "선택 기간"], \
        f"직전 비교 컬럼이 아니에요 — {None if tbl is None else list(tbl.columns)}"


@case
def t_missing_unique_inflow_is_announced():
    """유니크유입이 비어 있으면 빈 화면 대신 왜 못 그리는지 말해야 한다."""
    mtd = synth_mtd(days=90).copy()
    mtd["uniqueInflow"] = np.nan
    at = _open_tab(mtd=mtd)
    txt = " ".join(_texts(at))
    assert "유니크 유입" in txt and "다시 올려" in txt, f"안내 문구가 없어요 — {txt[:400]}"


@case
def t_zero_denominator_is_safe():
    """분모가 0이어도 죽거나 inf/nan%를 뿌리면 안 된다."""
    mtd = synth_mtd(days=90).copy()
    mtd["customers"] = 0.0
    at = _open_tab(mtd=mtd)
    txt = " ".join(_texts(at))
    assert "inf" not in txt.lower(), f"inf가 화면에 나왔어요 — {txt[:300]}"
    assert "nan%" not in txt.lower(), f"nan%가 화면에 나왔어요 — {txt[:300]}"


@case
def t_granularity_switch_renders():
    """집계 단위를 바꿔도 정상 렌더돼야 한다."""
    at = _open_tab(mtd=synth_mtd(days=150))
    grans = [r for r in at.radio if r.label == "집계"]
    assert grans, f"집계 라디오가 없어요 — {[r.label for r in at.radio]}"
    for opt in list(grans[0].options):
        tgt = [r for r in at.radio if r.label == "집계"][0]
        tgt.set_value(opt)
        at.run()
        assert not at.exception, f"집계={opt}에서 실패: {at.exception[0].value}"


@case
def t_derived_metrics_formula():
    """유니크유입 분모 파생 3종이 정의대로 계산돼야 한다."""
    mtd = synth_mtd(days=60)
    out = S.compute_mtd(mtd.copy())
    d = out["df"]
    for c in ("uniq_cr", "rev_per_uniq", "inflow_dup"):
        assert c in d.columns, f"{c} 파생이 없어요"
    exp_cr = (d["purchaseCust"] / d["uniqueInflow"]).clip(0, 1)
    assert np.allclose(d["uniq_cr"], exp_cr, equal_nan=True), "uniq_cr 공식이 달라요"
    assert np.allclose(d["rev_per_uniq"], d["revenue"] / d["uniqueInflow"], equal_nan=True), \
        "rev_per_uniq 공식이 달라요"
    assert np.allclose(d["inflow_dup"], d["totalInflow"] / d["uniqueInflow"], equal_nan=True), \
        "inflow_dup 공식이 달라요"


@case
def t_zero_unique_inflow_becomes_nan():
    """유니크유입 0인 날은 inf가 아니라 결측이어야 한다."""
    mtd = synth_mtd(days=40).copy()
    mtd.loc[mtd.index[:5], "uniqueInflow"] = 0.0
    d = S.compute_mtd(mtd)["df"]
    head = d.head(5)
    assert head["rev_per_uniq"].isna().all(), "0 분모가 inf로 남았어요"
    assert head["inflow_dup"].isna().all(), "0 분모가 inf로 남았어요"
    assert np.isfinite(d["rev_per_uniq"].tail(10)).all(), "정상 구간까지 결측이 됐어요"


@case
def t_cache_marker_mentions_uniq():
    """`compute_mtd`에 파생을 추가하면 캐시 표식도 같이 올려야 한다.

    `cached_compute_mtd`는 @st.cache_data라 자기 소스만 캐시 키로 본다 — 호출하는
    compute_mtd가 바뀐 걸 모른다. prio_g 때 이 표식을 빠뜨려 화면이 통째로 비었다."""
    src = io.open(ROOT / "send_perf_dashboard.py", encoding="utf-8").read()
    assert "MTDSET_VER" in src, "MTDSET_VER 표식이 없어요"
    assert "def cached_compute_mtd(mtd_df, ver=MTDSET_VER)" in src, \
        "cached_compute_mtd가 캐시 표식을 인자로 안 받아요"


@case
def t_new_metrics_selectable_in_mtd_tabs():
    """유니크유입 계열 지표를 다른 MTD 탭 셀렉트에서도 고를 수 있어야 한다."""
    for tab, want in (("발송 빈도·한계수익", "유입 대비 구매전환율"),
                      ("요일 패턴", "유니크 유입")):
        at = AppTest.from_file(APP, default_timeout=TIMEOUT)
        at.session_state["camp_store"] = synth_store(weeks=10)
        at.session_state["mtd_store_df"] = synth_mtd(days=200)
        at.run()
        at.sidebar.radio[0].set_value("6. 효율·피로도")
        at.run()
        subs = [r for r in at.radio if r.label != "페이지"][0]
        subs.set_value(tab)
        at.run()
        assert not at.exception, at.exception[0].value
        sels = [s for s in at.selectbox if s.label == "지표"]
        assert sels, f"[{tab}] 지표 셀렉트가 없어요"
        assert want in list(sels[0].options), \
            f"[{tab}] '{want}'가 선택지에 없어요 — {list(sels[0].options)}"
        sels[0].set_value(want)
        at.run()
        assert not at.exception, f"[{tab}] {want} 선택에서 실패: {at.exception[0].value}"


@case
def t_labels_exist_for_derived():
    """새 파생 지표에 라벨이 없으면 추세 표·툴팁이 KeyError로 죽는다."""
    for k in ("uniq_cr", "rev_per_uniq", "inflow_dup"):
        assert k in S.MTD_LABELS, f"{k} 라벨이 없어요"
        assert k in S.MTD_DERIVED, f"{k}가 MTD_DERIVED에 없어요"


def main():
    fails = []
    for fn in CASES:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {str(e)[:300]}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"유입 퍼널 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
