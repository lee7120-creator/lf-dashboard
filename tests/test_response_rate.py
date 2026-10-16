"""「6. 발송량·피로도 › 푸시 반응률」 — 월별 MTD 발송 상세 파싱·저장·화면 테스트.

월 반응률 = 유입 MAU ÷ 발송 MAU. 월 유니크는 일별 값을 더해 만들 수 없어서 월별 표를
따로 받아 **별도 저장소**에 쌓는다. 여기서 조용히 틀리는 자리는 넷이다.

1. 월별 표가 **일별 MTD로 오분류**되면 날짜를 못 읽어 빈 프레임이 되고, 증상은
   '올렸는데 아무 일도 없다'뿐이다.
2. **진행 중인 달**(`~10/15`)의 월 유니크를 완결 월과 맞대면 유입이 덜 쌓여 늘 낮게
   나와 가짜 하락이 뜬다.
3. 백업 ZIP 복원에서 `mtdmon.csv`가 이름의 `mtd` 때문에 **일별 갈래로 흘러간다**.
4. 주간 추적에서 진행 중인 주를 전년 7일과 맞대면 요일 구성이 달라진다.

로컬 실행:
    python tests/test_response_rate.py
"""
import ast
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
from smoke_pages import group_of, synth_store   # noqa: E402
from test_send_volume_band import synth_mtd     # noqa: E402
from test_inflow_funnel import _MTD_ROWS, _mtd_xlsx   # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300
TAB = "푸시 반응률"

# 실파일과 같은 모양 — 0행 `*TOTAL`, 1행 `2025년 1월`, 2행 `~1/31`, 3행부터 지표.
# 지표 행 순서·이름도 실파일 그대로 둔다(앞뒤 공백까지). `평균객단가`처럼 이름을
# «품기만 한» 행을 섞어야 접두어 매칭이 남의 행을 먹는지까지 드러난다.
_MON_ROWS = ["발송건당 거래액 ", "거래액 ", "총 발송 건수 ", "유니크 발송 고객수 ",
             "인당 발송 건수 ", "CTR ", "유니크 유입 ", "총 유입 ", "평균객단가",
             "구매 고객수 ", "객단가"]


def _months(start="2024-01", n=22):
    return list(pd.period_range(start, periods=n, freq="M"))


def _mon_xlsx(months, asof_last=15, smau=None, mau=None):
    """월별 MTD 발송 상세 — 마지막 달만 `~M/asof_last`로 진행 중."""
    n = len(months)
    smau = smau or [750_000 + 1_000 * i for i in range(n)]
    mau = mau or [170_000 - 2_000 * i for i in range(n)]
    hdr0 = [None, "*TOTAL"] + [None] * (n - 1)
    hdr1 = [None] + [f"{p.year}년 {p.month}월" for p in months]
    hdr2 = [None] + [f"~{p.month}/{p.days_in_month if i < n - 1 else asof_last}"
                     for i, p in enumerate(months)]
    rows = [hdr0, hdr1, hdr2]
    for lb in _MON_ROWS:
        key = lb.replace(" ", "")
        if key == "유니크발송고객수":
            vals = smau
        elif key == "유니크유입":
            vals = mau
        elif key == "평균객단가":
            vals = [999_999] * n
        else:
            vals = [100 + j for j in range(n)]
        rows.append([lb] + list(vals))
    b = io.BytesIO()
    pd.DataFrame(rows).to_excel(b, header=False, index=False)
    return b.getvalue()


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# ── 분류·파싱 ─────────────────────────────────────────────────────────
@case
def t_monthly_file_is_not_read_as_daily_mtd():
    """월별 표는 `mtdmon`, 일별 MTD는 그대로 `mtd`다 — 지표 행 이름이 같아서 갈린다."""
    assert S.classify_upload("표) 월 - MTD 발송 상세.xlsx", _mon_xlsx(_months())) == "mtdmon"
    assert S.classify_upload("MTD.xlsx", _mtd_xlsx(_MTD_ROWS, n=8)) == "mtd"


@case
def t_parser_reads_months_and_asof():
    ms = _months()
    out = S.parse_mtdmon_bytes(_mon_xlsx(ms, asof_last=15))
    assert len(out) == len(ms), out
    assert out["month"].iloc[0] == "202401" and out["asof"].iloc[0] == "20240131", out.head(2)
    last = out.iloc[-1]
    assert last["month"] == f"{ms[-1].year}{ms[-1].month:02d}"
    assert last["asof"] == f"{ms[-1].year}{ms[-1].month:02d}15", last["asof"]
    assert out["customers"].iloc[0] == 750_000 and out["uniqueInflow"].iloc[0] == 170_000
    # 이름을 «품기만 한» 행(평균객단가)이 객단가 자리를 먹지 않는다
    assert (out["avgOrderVal"] != 999_999).all(), out["avgOrderVal"].tolist()


@case
def t_missing_inflow_row_is_announced():
    b = _mon_xlsx(_months())
    df = pd.read_excel(io.BytesIO(b), header=None)
    df = df[df[0].astype(str).str.replace(" ", "") != "유니크유입"]
    bb = io.BytesIO(); df.to_excel(bb, header=False, index=False)
    out = S.parse_mtdmon_bytes(bb.getvalue())
    assert "유니크 유입" in out.attrs["mtdmon_missing"], out.attrs


# ── 반응률 규칙 ───────────────────────────────────────────────────────
@case
def t_partial_month_gets_no_yoy():
    """진행 중인 달은 전년비를 안 매긴다 — 15일치 유니크는 30일치보다 늘 작다."""
    ms = _months("2024-01", 22)                          # 2024-01 ~ 2025-10
    r = S.mtdmon_rates(S.parse_mtdmon_bytes(_mon_xlsx(ms)))
    last = r.iloc[-1]
    assert bool(last["partial"]), last
    assert pd.isna(last["rr_yoy"]), f"진행 중인 달에 전년비가 붙었어요 — {last['rr_yoy']}"
    sep = r[(r["y"] == 2025) & (r["m"] == 9)].iloc[0]
    ly = r[(r["y"] == 2024) & (r["m"] == 9)].iloc[0]
    assert not bool(sep["partial"])
    assert abs(sep["rr_yoy"] - (sep["rr"] - ly["rr"])) < 1e-12, (sep, ly)
    assert abs(sep["rr"] - sep["mau"] / sep["smau"]) < 1e-12


@case
def t_yoy_is_blank_when_the_prior_month_was_partial():
    """전년 같은 달이 진행 중이던 값만 있으면 그것과도 맞대지 않는다."""
    a = S.parse_mtdmon_bytes(_mon_xlsx(_months("2024-01", 10), asof_last=12))   # 2024-10 진행 중
    b = S.parse_mtdmon_bytes(_mon_xlsx(_months("2025-10", 2)))                 # 2025-10 완결
    r = S.mtdmon_rates(S.merge_mtdmon_store(a, b))
    oct25 = r[(r["y"] == 2025) & (r["m"] == 10)].iloc[0]
    assert not bool(oct25["partial"]) and pd.isna(oct25["rr_yoy"]), oct25


@case
def t_merge_keeps_the_newer_month():
    """같은 달을 다시 올리면 새 값이 이긴다(진행 중인 달은 매주 늘어난다)."""
    ms = _months("2025-01", 10)
    old = S.parse_mtdmon_bytes(_mon_xlsx(ms, asof_last=8))
    new = S.parse_mtdmon_bytes(_mon_xlsx(ms, asof_last=15,
                                         smau=[1] * 9 + [800_000], mau=[1] * 9 + [99_000]))
    m = S.merge_mtdmon_store(old, new)
    assert len(m) == 10 and m["month"].is_unique
    assert m.iloc[-1]["asof"].endswith("15") and m.iloc[-1]["uniqueInflow"] == 99_000, m.tail(1)


@case
def t_store_roundtrip_keeps_text_keys():
    """구글시트 라운드트립처럼 전부 문자열로 돌아와도 키가 안 깨진다."""
    d = S.parse_mtdmon_bytes(_mon_xlsx(_months()))
    back = S.finalize_mtdmon(d.astype(str))
    assert back["month"].tolist() == d["month"].tolist()
    assert np.allclose(back["customers"], d["customers"])


# ── 주간 추적 ─────────────────────────────────────────────────────────
def _daily(start, end, rr_fn):
    ds = pd.date_range(start, end, freq="D")
    cust = np.full(len(ds), 700_000.0)
    return pd.DataFrame({"date": ds, "customers": cust,
                         "uniqueInflow": [cust[0] * rr_fn(d) for d in ds]})


@case
def t_partial_week_compares_the_same_weekdays():
    """진행 중인 주(월·화)는 전년도 월·화만 본다. 전년 수~일엔 큰 값을 심어 둔다."""
    ly = _daily("2025-10-01", "2025-10-31",
                lambda d: 0.04 if d.isocalendar()[2] <= 2 else 0.20)
    ty = _daily("2026-10-05", "2026-10-13", lambda d: 0.05)   # 41주차 전체 + 42주차 월·화
    w = S.weekly_daily_rr(pd.concat([ly, ty], ignore_index=True))
    cur = w[(w["iy"] == 2026) & (w["iw"] == 42)].iloc[0]
    assert cur["days"] == 2, cur
    assert abs(cur["ly_drr"] - 0.04) < 1e-12, f"전년 7일을 다 맞댔어요 — {cur['ly_drr']}"
    full = w[(w["iy"] == 2026) & (w["iw"] == 41)].iloc[0]
    assert full["ly_drr"] > 0.1, full                       # 완결 주는 7일을 다 본다


@case
def t_prior_year_is_the_same_iso_week():
    """전년 같은 ISO 주차 — 날짜를 그대로 1년 빼면 요일이 어긋난다."""
    ly = _daily("2025-09-29", "2025-10-12", lambda d: 0.03 if d < pd.Timestamp("2025-10-06") else 0.09)
    ty = _daily("2026-10-05", "2026-10-11", lambda d: 0.05)   # 2026-W41 ↔ 2025-W41(10/6~12)
    w = S.weekly_daily_rr(pd.concat([ly, ty], ignore_index=True))
    cur = w[(w["iy"] == 2026) & (w["iw"] == 41)].iloc[0]
    assert abs(cur["ly_drr"] - 0.09) < 1e-12, cur


# ── 화면 ───────────────────────────────────────────────────────────────
def _open(mon=None, mtd=None):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=10)
    if mon is not None:
        at.session_state["mtdmon_store_df"] = mon
    if mtd is not None:
        at.session_state["mtd_store_df"] = mtd
    at.run()
    assert not at.exception, at.exception[0].value
    at.sidebar.radio[0].set_value(group_of(TAB)); at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    assert TAB in subs[0].options, subs[0].options
    subs[0].set_value(TAB); at.run()
    assert not at.exception, at.exception[0].value
    return at


def _texts(at):
    return [str(e.value) for coll in (at.markdown, at.caption, at.info, at.title)
            for e in coll]


@case
def t_empty_says_what_to_upload():
    at = _open()
    assert any("월별 MTD 발송 상세" in t for t in _texts(at)), _texts(at)


@case
def t_screen_values_match_the_store():
    """화면 표의 월 반응률을 «찍힌 값»으로 읽어 저장소와 대조한다."""
    ms = _months("2024-01", 22)
    mon = S.parse_mtdmon_bytes(_mon_xlsx(ms))
    at = _open(mon=mon, mtd=synth_mtd(days=420))
    tb = [t.value for t in at.dataframe if "월 반응률" in list(getattr(t.value, "columns", []))]
    assert tb, "월 반응률 표가 없어요"
    df = tb[0]
    r = S.mtdmon_rates(mon)
    sep = r[(r["y"] == 2025) & (r["m"] == 9)].iloc[0]
    row = df[df["월"] == "2025년 9월"].iloc[0]
    assert abs(row["월 반응률"] - sep["rr"] * 100) < 1e-9, (row, sep)
    assert any("(~10/15)" in str(x) for x in df["월"]), df["월"].tolist()
    part = df[df["월"].str.contains("~")].iloc[0]
    assert pd.isna(part["차이(%p)"]) and pd.isna(part["목표 대비(%p)"]), part
    wk = [t.value for t in at.dataframe if "일 반응률" in list(getattr(t.value, "columns", []))]
    assert wk and wk[0]["전년 같은 주"].notna().any(), "주간 표에 전년 값이 없어요"


@case
def t_partial_month_is_left_out_of_the_chart():
    """차트엔 완결 월만 — 진행 중인 달을 그리면 그 점만 뚝 떨어져 보인다."""
    import json
    ms = _months("2024-01", 22)
    mon = S.parse_mtdmon_bytes(_mon_xlsx(ms))
    at = _open(mon=mon)
    figs = [json.loads(e.proto.spec) for e in at.get("plotly_chart")]
    yr = [f for f in figs if any(str(tr.get("name", "")).endswith("년") for tr in f["data"])]
    assert yr, "월 반응률 차트가 없어요"
    t25 = [tr for tr in yr[0]["data"] if tr.get("name") == "2025년"][0]
    xs = t25["x"]
    if isinstance(xs, dict):                              # plotly가 bdata로 싣는 경우
        xs = np.frombuffer(__import__("base64").b64decode(xs["bdata"]),
                           dtype=np.dtype(xs["dtype"])).tolist()
    assert 10 not in [int(x) for x in xs], f"진행 중인 10월이 그려졌어요 — {xs}"


# ── 저장소 배선 ───────────────────────────────────────────────────────
@case
def t_backup_routes_monthly_before_daily_mtd():
    """`_apply_backup`에서 mtdmon 갈래가 `"mtd" in base`보다 앞이어야 한다."""
    src = (ROOT / "send_perf_dashboard.py").read_text(encoding="utf-8")
    i = src.index('base.startswith("mtdmon")')
    j = src.index('elif "mtd" in base:')
    assert i < j, "월별 MTD 복원 갈래가 일별 MTD 뒤에 있어요"


@case
def t_backup_zip_carries_the_monthly_store():
    src = (ROOT / "send_perf_dashboard.py").read_text(encoding="utf-8")
    assert '_z.writestr("mtdmon.csv"' in src
    assert "mtdmon_store_df" in src[src.index("def _bak_sig_now"):src.index('"bak_make"')]


@case
def t_page_is_in_the_menu_not_gated_by_daily_mtd():
    """월 표만 올려도 열린다 — MTD_TABS에 넣으면 일별 MTD가 없을 때 안내만 뜬다."""
    assert TAB in S.CAMPAIGN_GROUPS[group_of(TAB)]
    assert TAB not in S.MTD_TABS
    tree = ast.parse((ROOT / "send_perf_dashboard.py").read_text(encoding="utf-8"))
    assert any(isinstance(n, ast.Constant) and n.value == TAB for n in ast.walk(tree))


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
    print(f"푸시 반응률 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
