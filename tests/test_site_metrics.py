"""사이트 회원UV·거래액(채널×디바이스 와이드 리포트) 파싱·저장·화면 테스트.

두 리포트는 구조가 완전히 같아서(연·월·일 3단 헤더 + 채널/디바이스 행) 값으로 가른다.
잘못 가르면 회원UV 칸에 거래액이 들어가는데, 화면은 멀쩡히 떠서 눈으로는 안 잡힌다.

주간보고 MTD 표의 앱·PUSH 행은 **데이터가 없으면 통째로 빠져야** 한다 — 빈 행이 남으면
'0인가 없는 건가'를 구분할 수 없다.

로컬 실행:
    python tests/test_site_metrics.py
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

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300
PAGE = "12. 회원UV·거래액"


def synth_site_xlsx(kind="uv", days=120, seed=2):
    """태블로 export와 같은 모양의 합성 리포트.

    0행=연 · 1행=월 · 2행=일(연·월은 병합셀이라 첫 칸에만), A열=블록, B열=채널, C열=디바이스.
    '기준' 말고 '전주비' 블록도 넣어 둔다 — 그걸 실측치로 읽으면 안 되니까.
    """
    import openpyxl
    rng = np.random.default_rng(seed)
    end = pd.Timestamp.utcnow().tz_localize(None).normalize()
    dates = pd.date_range(end - pd.Timedelta(days=days - 1), end, freq="D")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet 1"
    r0, r1, r2 = [None] * 3, [None] * 3, [None] * 3
    _py, _pm = None, None
    for d in dates:
        r0.append(f"{d.year}년" if d.year != _py else None)
        r1.append(f"{d.month}월" if (d.year, d.month) != (_py, _pm) else None)
        r2.append(d.day)
        _py, _pm = d.year, d.month
    ws.append(r0); ws.append(r1); ws.append(r2)
    base = {"Total": 250_000, "직접": 90_000, "광고": 110_000, "PUSH": 38_000,
            "제휴": 5_000, "EP": 18_000, "미디어커머스": 800, "브랜드광고": 400}
    devr = {"Total": 1.0, "App": 0.38, "Mobile Web": 0.48, "PC Web": 0.14}
    for blk in ("기준", "전주비"):
        first = True
        for ch in S.SITE_CHANNELS:
            for dv in S.SITE_DEVICES:
                row = [blk if first else None, ch if dv == "Total" else None, dv]
                first = False
                for _ in dates:
                    if blk == "전주비":
                        row.append(round(float(rng.normal(0, 0.1)), 4))
                        continue
                    v = base[ch] * devr[dv] * float(rng.normal(1.0, 0.08))
                    # 회원UV는 사람 수라 정수, 거래액은 환산값이라 소수가 남는다
                    row.append(int(round(v)) if kind == "uv" else round(v * 3.9 / 1000.0, 6))
                ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _site_store(days=120):
    k1, d1 = S.parse_site_bytes(synth_site_xlsx("uv", days), "a.xlsx")
    k2, d2 = S.parse_site_bytes(synth_site_xlsx("amt", days, seed=3), "b.xlsx")
    assert (k1, k2) == ("uv", "amt"), (k1, k2)
    return S.merge_site_store(d1, d2)


def _open(page, site=None, camp=None):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=12) if camp is None else camp
    if site is not None:
        at.session_state["site_store_df"] = site
    at.run()
    assert not at.exception, at.exception[0].value
    at.sidebar.radio[0].set_value(page)
    at.run()
    assert not at.exception, at.exception[0].value
    return at


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success):
        out += [str(e.value) for e in coll]
    return out


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_kind_detection():
    """정수만 있으면 회원UV, 소수가 섞이면 거래액. 파일명이 말해 주면 그게 우선."""
    assert S.site_metric_kind(np.array([1.0, 2.0, 3.0])) == "uv"
    assert S.site_metric_kind(np.array([1.5, 2.25, 3.0])) == "amt"
    assert S.site_metric_kind(np.array([1.0, 2.0]), "일별 거래액.xlsx") == "amt"
    assert S.site_metric_kind(np.array([]), "무엇.xlsx") == "uv"


@case
def t_parse_shape_and_blocks():
    """'기준' 블록만 읽고, 채널×디바이스×날짜가 다 살아 있어야 한다."""
    k, d = S.parse_site_bytes(synth_site_xlsx("uv", days=30), "x.xlsx")
    assert k == "uv", k
    assert set(d["ch"]) == set(S.SITE_CHANNELS), sorted(set(d["ch"]))
    assert set(d["dev"]) == set(S.SITE_DEVICES), sorted(set(d["dev"]))
    assert d["date"].nunique() == 30, d["date"].nunique()
    assert len(d) == 30 * 8 * 4, len(d)
    # '전주비' 블록(±0.1 근처 소수)이 섞이면 1 미만 값이 대량으로 들어온다.
    # 행 수(30*32)가 두 배가 아닌 것과 함께 보면 '기준'만 읽었다는 증거가 된다.
    assert (d["uv"].abs() < 1).sum() == 0, f"전주비 블록이 섞였어요 — {(d['uv'].abs() < 1).sum()}건"


@case
def t_classify_upload_detects_site():
    """통합 업로더가 사이트 리포트를 mtd로 오분류하면 안 된다 (둘 다 날짜가 가로로 눕는다)."""
    got = S.classify_upload("회원UV.xlsx", synth_site_xlsx("uv", days=20))
    assert got == "site", f"site가 아니라 {got}로 분류됐어요"


@case
def t_merge_keeps_other_metric():
    """회원UV만 다시 올려도 이미 쌓인 거래액이 날아가면 안 된다."""
    k1, uv = S.parse_site_bytes(synth_site_xlsx("uv", 20), "a")
    k2, amt = S.parse_site_bytes(synth_site_xlsx("amt", 20, seed=9), "b")
    both = S.merge_site_store(uv, amt)
    assert both["uv"].notna().all() and both["amt"].notna().all()
    again = S.merge_site_store(both, uv)                 # 회원UV만 재업로드
    assert again["amt"].notna().all(), "거래액이 NaN으로 덮였어요"
    assert len(again) == len(both), (len(again), len(both))


@case
def t_finalize_recovers_string_roundtrip():
    """구글시트 왕복(전부 문자열)에서도 dtype이 살아나야 한다."""
    d = _site_store(days=10)
    rt = d.astype(str)                                   # 시트에서 읽으면 이 모양
    f = S.finalize_site(rt)
    assert len(f) == len(d), (len(f), len(d))
    assert pd.api.types.is_numeric_dtype(f["uv"]), f["uv"].dtype
    assert pd.api.types.is_numeric_dtype(f["amt"]), f["amt"].dtype


@case
def t_display_unit_is_divided():
    """화면 값은 원본 ÷1000 (회원UV 천명 · 거래액 백만원)."""
    d = _site_store(days=20)
    raw = d[(d["ch"] == "Total") & (d["dev"] == "Total")]["uv"].mean()
    got = S.site_pick(d, "Total", "Total")["uv"].mean()
    assert abs(got - raw / 1000.0) < 1e-6, (got, raw)


@case
def t_page_without_data_asks_for_upload():
    """데이터가 없으면 죽지 말고 올려 달라고 안내해야 한다."""
    at = _open(PAGE)
    assert any("올려 주세요" in t for t in _texts(at)), _texts(at)[:5]


@case
def t_page_renders_all_units():
    """일별·주차별·월별을 다 그려야 한다."""
    at = _open(PAGE, site=_site_store(days=200))
    rad = [r for r in at.radio if r.label == "집계 단위"]
    assert rad, f"집계 단위 라디오가 없어요 — {[r.label for r in at.radio]}"
    for opt in list(rad[0].options):
        tgt = [r for r in at.radio if r.label == "집계 단위"][0]
        tgt.set_value(opt)
        at.run()
        assert not at.exception, f"{opt}에서 실패: {at.exception[0].value}"
        assert any("추이" in t for t in _texts(at)), f"{opt}에서 추이 블록이 없어요"


@case
def t_channel_and_device_switch():
    """채널·디바이스를 바꿔도 정상 렌더된다 (PUSH·App 조합 포함)."""
    at = _open(PAGE, site=_site_store(days=120))
    for lab, val in (("채널", "PUSH"), ("디바이스", "App")):
        tgt = [s for s in at.selectbox if s.label == lab]
        assert tgt, f"{lab} 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
        tgt[0].set_value(val)
        at.run()
        assert not at.exception, f"{lab}={val}에서 실패: {at.exception[0].value}"
    assert any("PUSH · App" in t for t in _texts(at)), "선택이 제목에 안 반영됐어요"


@case
def t_weekly_mtd_rows_appear_with_data():
    """사이트 데이터가 있으면 주간보고 MTD 표에 앱·PUSH 행이 붙는다."""
    at = _open("0. 주간보고", site=_site_store(days=200))
    tbls = [t for t in at.dataframe if "지표" in list(getattr(t.value, "columns", []))]
    assert tbls, "MTD 표를 찾지 못했어요"
    names = set()
    for t in tbls:
        names |= set(t.value["지표"].astype(str))
    for want in ("앱 회원UV(천명)", "앱 거래액(백만원)", "PUSH 회원UV(천명)", "PUSH 거래액(백만원)"):
        assert want in names, f"{want} 행이 없어요 — {sorted(names)}"


@case
def t_weekly_mtd_rows_vanish_without_data():
    """사이트 데이터가 없으면 그 행만 통째로 빠져야 한다 (빈 행 금지)."""
    at = _open("0. 주간보고")
    names = set()
    for t in at.dataframe:
        if "지표" in list(getattr(t.value, "columns", [])):
            names |= set(t.value["지표"].astype(str))
    assert not [n for n in names if n.startswith(("앱 ", "PUSH "))], \
        f"데이터가 없는데 사이트 행이 남았어요 — {sorted(names)}"
    assert any("올리면" in t and "월 평균" in t for t in _texts(at)), "안내 문구가 없어요"


@case
def t_weekly_mtd_value_is_daily_mean():
    """MTD 칸 값은 그 구간의 일평균이어야 한다 (누계가 아니라)."""
    site = _site_store(days=200)
    at = _open("0. 주간보고", site=site)
    tbls = [t for t in at.dataframe if "지표" in list(getattr(t.value, "columns", []))]
    row, curcol = None, None
    for t in tbls:
        v = t.value
        hit = v[v["지표"].astype(str) == "앱 회원UV(천명)"]
        if len(hit):
            row = hit.iloc[0]
            curcol = [c for c in v.columns if str(c).startswith("당월 MTD")][0]
            break
    assert row is not None, "앱 회원UV 행을 못 찾았어요"
    shown = float(str(row[curcol]).replace(",", ""))
    # 화면이 쓴 것과 같은 창으로 직접 계산해 맞춰 본다
    import re
    m = re.search(r"\((\d+)/(\d+)~(\d+)/(\d+)\)", str(curcol))
    assert m, curcol
    _y = S.site_pick(site, "Total", "App")["dt"].max().year
    d0 = pd.Timestamp(_y, int(m.group(1)), int(m.group(2)))
    d1 = pd.Timestamp(_y, int(m.group(3)), int(m.group(4)))
    want = S.site_mean(site, "Total", "App", d0, d1)["uv"]
    assert abs(shown - want) < 0.15, f"화면 {shown} vs 계산 {want}"


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
    print(f"사이트 지표 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
