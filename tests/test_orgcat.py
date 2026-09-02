"""조직×카테고리(MICRO 대시보드 export) 파싱·저장·화면 테스트.

이 export는 마스터(전체관점)와 달리 세그먼트 축이 **둘**(구분06=조직, 구분07=카테고리)이고,
헤더 구성도 단위마다 다르다 — 월·주는 `연도/LFMS/기간/(구분+마감)` 4행, 일은 마감 행이 없어
3행이다. 행 번호를 박으면 한 단위만 조용히 깨진다.

LFMS 포함여부(Y/N)는 **모집단이 다른** 축이라 키에서 빠지면 같은 기간 값이 서로를 덮어쓴다.

로컬 실행:
    python tests/test_orgcat.py
"""
import contextlib
import io
import os
import pathlib
import shutil
import sys
import tempfile

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import weekly_report as W                        # noqa: E402
from smoke_weekly_report import synth_store      # noqa: E402

TIMEOUT = 300
ORGS = ["*TOTAL", "e-영업1", "e-영업2", "SPACE-R"]
CATS = {"*TOTAL": ["*TOTAL"], "e-영업1": ["*TOTAL", "골프", "남성"],
        "e-영업2": ["*TOTAL", "여성", "잡화"], "SPACE-R": ["*TOTAL", "-"]}
METS = ["일평균거래액", "거래액비중", "일평균고객수", "고객비중",
        "일평균객단가", "상품UV", "상품CR"]


def _v(met, org, cat, per):
    """지표·조직·카테고리·기간으로 결정되는 재현 가능한 값 (비율은 0~1)."""
    base = (abs(hash((org, cat))) % 90 + 10) * (per + 1)
    if met in ("거래액비중", "고객비중", "상품CR"):
        return round((base % 80 + 5) / 1000, 6)
    if met == "일평균거래액":
        return float(base * 1_000_000)
    return float(base * 7)


def synth_grid(gran="월", years=(2025,), lfms="N", periods=3, mtd_last=False):
    """실제 export와 같은 모양의 그리드(2차원 리스트).

    월·주는 마감 행이 따로 있고, 일은 없다. mtd_last=True면 마지막 기간에
    '일마감'(MTD) 칼럼을 하나 더 붙인다 — 기간 라벨이 비어 직전 라벨을 이어받는 자리.
    """
    def plabel(i):
        if gran == "월": return f"{i+1}월"
        if gran == "주": return f"{(i//5)+1:02d}월 {(i%5)+1}주차"
        return f"{(i//28)+1}/{(i%28)+1}"

    cols = []                                   # (year, label, close)
    for y in years:
        for i in range(periods):
            cols.append((y, plabel(i), "월마감" if gran == "월" else
                         ("주마감" if gran == "주" else "")))
        if mtd_last and gran != "일":
            cols.append((y, "", "일마감"))
    ncol = 7 + len(cols)

    def blank(): return [None] * ncol
    r_year, r_lfms, r_per, r_close = blank(), blank(), blank(), blank()
    prev_y = None
    for j, (y, lb, cl) in enumerate(cols):
        if y != prev_y:
            r_year[7 + j] = str(y); r_lfms[7 + j] = lfms; prev_y = y
        if lb: r_per[7 + j] = lb
        if cl: r_close[7 + j] = cl
    for j, nm in enumerate(["구분06", "구분07", "구분08", "구분09", "구분10", "구분11"]):
        (r_per if gran == "일" else r_close)[1 + j] = nm

    rows = [r_year, r_lfms, r_per] + ([] if gran == "일" else [r_close])
    for met in METS:
        first_of_metric = True
        for org in ORGS:
            first_of_org = True
            for cat in CATS[org]:
                r = blank()
                if first_of_metric: r[0] = met + " "      # 원본도 뒤에 공백이 붙어 온다
                if first_of_org: r[1] = org
                r[2] = cat
                for j in range(3, 7): r[3 + j - 3] = "-"
                last_lbl = None
                for j, (y, lb, cl) in enumerate(cols):
                    if lb: last_lbl = lb
                    per = [p for p in range(periods)
                           if plabel(p) == last_lbl][0] if last_lbl else 0
                    r[7 + j] = None if cat == "-" else _v(met, org, cat, per + y * 100)
                rows.append(r)
                first_of_metric = first_of_org = False
    return rows


def synth_orgcat_df(**kw):
    return W.parse_orgcat_grid(synth_grid(**kw))


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# ── 파싱 ────────────────────────────────────────────────────────────
@case
def t_detects_and_parses_all_grans():
    """월·주·일 세 헤더 모양을 다 읽어야 한다 (일은 마감 행이 없다)."""
    for gran in ("월", "주", "일"):
        rows = synth_grid(gran=gran)
        assert W.is_orgcat_grid(rows), f"[{gran}] 형식 인식 실패"
        d = W.parse_orgcat_grid(rows)
        assert not d.empty, f"[{gran}] 파싱 결과가 비었어요"
        assert set(d["gran"]) == {gran}, f"[{gran}] 단위 오판 — {set(d['gran'])}"
        assert set(d["metric"]) == {W.ORGCAT_MAP.get(m, m) for m in METS}, set(d["metric"])


@case
def t_metric_names_map_to_report_names():
    """일평균거래액·고객수·객단가는 마스터와 같은 이름이어야 fmt_value가 통한다."""
    d = synth_orgcat_df()
    for src, dst in W.ORGCAT_MAP.items():
        assert (d["metric"] == dst).any(), f"{src} → {dst} 매핑 실패"
        assert not (d["metric"] == src).any(), f"원 이름 {src}이 남아 있어요"
    assert "상품CR" in W.PCT_METRICS, "상품CR이 비율 지표로 등록되지 않았어요"


@case
def t_org_is_forward_filled():
    """조직은 병합셀이라 첫 행에만 있다 — 아래 카테고리 행까지 이어받아야 한다."""
    d = synth_orgcat_df()
    got = set(d[d["cat"] == "골프"]["org"])
    assert got == {"e-영업1"}, f"조직 ffill 실패 — {got}"


@case
def t_dash_category_rows_are_skipped():
    """카테고리 '-'는 카테고리 구분이 없는 조직의 자리표시라 *TOTAL과 겹친다."""
    d = synth_orgcat_df()
    assert "-" not in set(d["cat"]), f"'-' 카테고리가 남았어요 — {sorted(set(d['cat']))}"
    assert (d["org"] == "SPACE-R").any(), "SPACE-R가 통째로 빠졌어요(*TOTAL은 남아야 함)"


@case
def t_two_years_in_one_file():
    """한 파일에 2개년이 온다 — 연도는 병합셀이라 오른쪽으로 이어받아야 한다."""
    d = synth_orgcat_df(years=(2025, 2026))
    assert set(d["year"]) == {2025, 2026}, f"연도 ffill 실패 — {sorted(set(d['year']))}"
    a = W.opick(d, "월", "첫구매 거래액", "e-영업1", "골프", "N", 2025, "1월")
    b = W.opick(d, "월", "첫구매 거래액", "e-영업1", "골프", "N", 2026, "1월")
    assert np.isfinite(a) and np.isfinite(b) and a != b, f"연도별 값이 같아요 — {a} vs {b}"


@case
def t_mtd_column_inherits_label():
    """'일마감' 칼럼은 기간 라벨이 비어 직전 라벨을 이어받고 close=mtd가 된다."""
    d = synth_orgcat_df(mtd_last=True)
    mtd = d[d["close"] == "mtd"]
    assert not mtd.empty, "mtd 행이 없어요"
    assert set(mtd["label"]) == {"3월"}, f"직전 라벨을 못 이어받았어요 — {set(mtd['label'])}"
    # final이 있으면 final이 이긴다 (마스터 pick과 같은 규칙)
    assert W.opick(d, "월", "첫구매 거래액", "*TOTAL", "*TOTAL", "N", 2025, "3월") == \
        d[(d["close"] == "final") & (d["label"] == "3월") & (d["org"] == "*TOTAL") &
          (d["metric"] == "첫구매 거래액")]["value"].iloc[0]


@case
def t_lfms_is_part_of_the_key():
    """LFMS 포함/미포함은 모집단이 달라 서로 덮어쓰면 안 된다."""
    n = synth_orgcat_df(lfms="N")
    y = synth_orgcat_df(lfms="Y")
    assert set(n["lfms"]) == {"N"} and set(y["lfms"]) == {"Y"}, "LFMS 파싱 실패"
    m = W.merge_orgcat(n, y)
    assert len(m) == len(n) + len(y), \
        f"LFMS가 키에 없어 덮어썼어요 — {len(m)} (기대 {len(n) + len(y)})"
    assert "lfms" in W.ORGCAT_KEY, "ORGCAT_KEY에 lfms가 없어요"


# ── 라우팅 · 저장 ────────────────────────────────────────────────────
def _xlsx(rows):
    import openpyxl
    wb = openpyxl.Workbook(); ws = wb.active
    for r in rows: ws.append(r)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


@case
def t_not_routed_into_master_store():
    """조직×카테고리를 마스터 파서가 집어삼키면 두 축이 뭉개진 채 누적된다."""
    b = _xlsx(synth_grid())
    assert W.parse_file("조직카테고리.xlsx", b).empty, "마스터 파서가 이 파일을 먹었어요"
    pf, d = W.route_push("조직카테고리.xlsx", b)
    assert pf is None and d is None, f"route_push가 넘겨 버렸어요 — push={pf is not None}, df={d is not None}"
    assert not W.parse_orgcat_file("조직카테고리.xlsx", b).empty, "전용 파서가 못 읽어요"


@case
def t_shown_in_upload_recognition_list():
    """미인식으로 표시되면 사용자가 '왜 안 올라가지'로 헤맨다."""
    b = _xlsx(synth_grid())
    out = W.classify_uploads.__wrapped__((("조직카테고리.xlsx", b),))
    assert out and "조직×카테고리" in out[0][1], f"인식 목록 표기가 없어요 — {out}"


@case
def t_merge_prefers_new_and_keeps_history():
    """같은 키는 새 값이 이기고, 겹치지 않는 과거는 남아야 한다."""
    old = synth_orgcat_df(years=(2025,))
    new = synth_orgcat_df(years=(2026,))
    m = W.merge_orgcat(old, new)
    assert set(m["year"]) == {2025, 2026}, "과거 연도가 사라졌어요"
    bumped = old.copy(); bumped["value"] = bumped["value"] * 2
    m2 = W.merge_orgcat(old, bumped)
    assert len(m2) == len(old), f"같은 키가 중복됐어요 — {len(m2)} vs {len(old)}"
    k = W.ORGCAT_KEY
    j = m2.set_index(k)["value"].sort_index()
    assert np.allclose(j.values, bumped.set_index(k)["value"].sort_index().values), \
        "새 값이 이기지 않았어요"


@case
def t_backup_zip_roundtrip():
    """통합 백업 ZIP에 담기고, 그 ZIP으로 그대로 되살아나야 한다."""
    import zipfile
    d = synth_orgcat_df()
    z = W.make_backup_zip(pd.DataFrame(columns=W.STORE_COLS), {}, d)
    names = zipfile.ZipFile(io.BytesIO(z)).namelist()
    assert "wr_orgcat_store.csv" in names, f"백업에 안 담겼어요 — {names}"
    back = pd.read_csv(io.BytesIO(zipfile.ZipFile(io.BytesIO(z)).read("wr_orgcat_store.csv")),
                       encoding="utf-8-sig")
    assert set(W.ORGCAT_COLS) <= set(back.columns), f"컬럼 유실 — {list(back.columns)}"
    assert len(back) == len(d), f"행 수가 달라요 — {len(back)} vs {len(d)}"


@case
def t_store_csv_roundtrip():
    """CSV 왕복에서 lfms='N'이 문자열로 살아남아야 한다 (Y/N은 bool로 안 읽혀야)."""
    d = synth_orgcat_df(lfms="N")
    tmp = tempfile.mkdtemp(); cwd = os.getcwd(); os.chdir(tmp)
    try:
        W.save_orgcat_store(d)
        back = W.load_orgcat_store()
        assert len(back) == len(d), f"{len(back)} vs {len(d)}"
        assert set(back["lfms"].astype(str)) == {"N"}, set(back["lfms"])
        assert W.opick(back, "월", "첫구매 거래액", "e-영업1", "골프", "N", 2025, "1월") == \
            W.opick(d, "월", "첫구매 거래액", "e-영업1", "골프", "N", 2025, "1월")
    finally:
        os.chdir(cwd); shutil.rmtree(tmp, ignore_errors=True)


# ── 화면 ────────────────────────────────────────────────────────────
@contextlib.contextmanager
def _run(page="08. 조직·카테고리별 실적", orgcat=None, master=True):
    """페이지를 연 AppTest를 넘겨준다. 블록 안에서 at.run()을 더 불러도 되도록
    (위젯을 바꿔 다시 그리는 테스트가 있다) 임시 앱 디렉터리를 블록이 끝날 때 치운다."""
    from streamlit.testing.v1 import AppTest
    tmp = tempfile.mkdtemp()
    app = os.path.join(tmp, "weekly_report.py")
    shutil.copy(ROOT / "weekly_report.py", app)
    for extra in ("table_export.py",):
        if (ROOT / extra).exists():
            shutil.copy(ROOT / extra, os.path.join(tmp, extra))
    if master:
        synth_store().to_csv(os.path.join(tmp, "wr_data_store.csv"),
                             index=False, encoding="utf-8-sig")
    if orgcat is not None and not orgcat.empty:
        orgcat[W.ORGCAT_COLS].to_csv(os.path.join(tmp, "wr_orgcat_store.csv"),
                                     index=False, encoding="utf-8-sig")
    cwd = os.getcwd(); os.chdir(tmp)
    try:
        at = AppTest.from_file(app, default_timeout=TIMEOUT)
        at.run()
        assert not at.exception, at.exception[0].value
        rs = ([r for r in at.radio if r.label == "페이지"] or
              [r for r in at.sidebar.radio if r.label == "페이지"])
        assert rs, "페이지 라디오를 찾지 못했어요"
        assert page in rs[0].options, f"{page}가 메뉴에 없어요 — {list(rs[0].options)}"
        rs[0].set_value(page); at.run()
        assert not at.exception, at.exception[0].value
        yield at
    finally:
        os.chdir(cwd); shutil.rmtree(tmp, ignore_errors=True)


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success, at.subheader):
        out += [str(e.value) for e in coll]
    return out


@case
def t_page_is_in_menu_and_renders():
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        txt = " ".join(_texts(at))
        for need in ("① 조직별", "② 카테고리별", "③ 조직별 추이", "④ 조직 × 카테고리"):
            assert need in txt, f"{need} 블록이 없어요"
        assert len(at.dataframe) >= 3, f"표가 모자라요 — {len(at.dataframe)}개"


@case
def t_page_without_data_asks_for_upload():
    """데이터가 없으면 빈 화면 대신 무엇을 올려야 하는지 말해야 한다."""
    with _run() as at:
        txt = " ".join(_texts(at))
        assert "구분06" in txt and "올려" in txt, f"안내 문구가 없어요 — {txt[:300]}"


@case
def t_page_opens_without_master_data():
    """조직×카테고리만 올린 상태에서도 이 화면은 열려야 한다 (자체 기간 선택이라)."""
    with _run(orgcat=synth_orgcat_df(), master=False) as at:
        txt = " ".join(_texts(at))
        assert "① 조직별" in txt, f"마스터 없이 안 열려요 — {txt[:400]}"


@case
def t_table_value_matches_source():
    """화면 숫자가 파일 값 그대로여야 한다 (일평균을 다시 나누거나 하지 않는다)."""
    d = synth_orgcat_df(years=(2025, 2026))
    with _run(orgcat=d) as at:
        tbls = [t for t in at.dataframe
                if "2026년 3월" in [str(c) for c in
                                   getattr(getattr(t.value, "data", t.value), "columns", [])]]
        assert tbls, "기준 기간 칼럼을 가진 표를 못 찾았어요"
        v = getattr(tbls[0].value, "data", tbls[0].value)
        got = str(v.loc["e-영업1", "2026년 3월"])
        want = W.fmt_value("첫구매 거래액",
                           W.opick(d, "월", "첫구매 거래액", "e-영업1", "*TOTAL", "N", 2026, "3월"))
        assert got == want, f"화면 {got} vs 원본 {want}"


@case
def t_lfms_options_follow_granularity():
    """단위마다 받아 온 export가 달라도 막다른 화면이 되면 안 된다.

    월은 LFMS=N만, 일은 Y만 있는 상태에서 '일'로 바꾸면, 전역 LFMS 목록을 쓰면
    이전 선택 N이 남아 데이터가 있는데도 '없어요'가 뜬다."""
    d = pd.concat([synth_orgcat_df(gran="월", lfms="N"),
                   synth_orgcat_df(gran="일", lfms="Y")], ignore_index=True)
    with _run(orgcat=d) as at:
        g = [r for r in at.radio if r.label == "집계 단위"]
        assert g and "일" in g[0].options, f"집계 단위 라디오가 없어요 — {[r.label for r in at.radio]}"
        g[0].set_value("일"); at.run()
        assert not at.exception, at.exception[0].value
        txt = " ".join(_texts(at))
        assert "① 조직별" in txt, f"'일'로 바꾸니 빈 화면이 됐어요 — {txt[:400]}"
        assert "고른 조건에 데이터가 없어요" not in txt, "LFMS 선택이 단위를 따라가지 않았어요"


@case
def t_in_progress_period_is_flagged():
    """진행 중(일마감만 있는) 기간을 완결 기간처럼 보여 주면 며칠치를 한 달로 읽는다."""
    d = synth_orgcat_df(years=(2025, 2026), mtd_last=True)
    # 2026년 4월은 '일마감'만 있는 진행 중 기간이 되도록 final을 지운다
    d = d[~((d["year"] == 2026) & (d["label"] == "3월") & (d["close"] == "final"))]
    d.loc[(d["year"] == 2026) & (d["close"] == "mtd"), "label"] = "3월"
    with _run(orgcat=d) as at:
        pers = [s for s in at.selectbox if s.label == "기준 기간"]
        assert pers, "기준 기간 셀렉트가 없어요"
        shown = [str(o) for o in pers[0].options]   # AppTest는 이미 포맷된 문자열을 준다
        assert any("진행 중" in x for x in shown), f"진행 중 표기가 없어요 — {shown[:4]}"
        txt = " ".join(_texts(at))
        assert "진행 중이에요" in txt, f"진행 중 안내가 없어요 — {txt[:400]}"


@case
def t_metric_switch_keeps_rendering():
    """지표를 바꿔도(비율 지표 포함) 정상 렌더돼야 한다."""
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        sels = [s for s in at.selectbox if s.label == "지표"]
        assert sels, f"지표 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
        for opt in list(sels[0].options):
            tgt = [s for s in at.selectbox if s.label == "지표"][0]
            tgt.set_value(opt); at.run()
            assert not at.exception, f"지표={opt}에서 실패: {at.exception[0].value}"


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
    print(f"조직×카테고리 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
