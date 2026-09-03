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
import zlib
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
# 실제 뎁스: 영업조직 > 카테고리 > 브랜드 > 상품. 브랜드·상품은 아직 값이 안 오지만,
# 값이 들어오는 날 화면이 알아서 열리는지 확인하려면 4단짜리 합성본이 필요하다.
TREE = {
    "e-영업1": {"골프": {"헤지스": ["티셔츠", "바지"], "닥스": ["니트"]},
              "남성": {"질스튜어트": ["코트"]},
              "잡화": {}},          # 브랜드로 안 쪼개지는 카테고리 (그 칸이 빈 칸으로 온다)
    "e-영업2": {"여성": {"모그": ["원피스"]}, "잡화": {"루이까또즈": ["가방"]}},
    "SPACE-R": {},                       # 카테고리가 없는 조직 (원본에서 '-'로 온다)
    "미매칭": {},
}
# 실파일의 미매칭·기타·SPACE-R처럼 **전 기간 합쳐도 수천 원**인 잡음 조직.
# 살아 있는 조직과 수만 배 차이가 나야 숨김 규칙이 의미 있게 검증된다.
NOISE_ORGS = {"SPACE-R", "미매칭"}
METS = ["일평균거래액", "거래액비중", "일평균고객수", "고객비중",
        "일평균객단가", "상품UV", "상품CR"]
_MET_KEY = {"일평균거래액": "rev", "거래액비중": "rev_share", "일평균고객수": "cust",
            "고객비중": "cust_share", "일평균객단가": "aov", "상품UV": "uv", "상품CR": "cr"}


def _leaf(path, per):
    """말단 한 칸의 유입·전환·객단가.

    `hash()`는 프로세스마다 값이 달라(PYTHONHASHSEED) 실패가 재현되지 않는다 —
    crc32로 고정한다."""
    h = zlib.crc32("|".join(map(str, path)).encode()) % 90 + 10
    if path and path[0] in NOISE_ORGS:            # 실적이 사실상 없는 조직
        return 1.0 + per * 0.0, 0.001, 100.0
    return (float(h * 12 + per * 3),
            round(0.04 + (h % 17) / 300, 6),
            float(70_000 + (h % 23) * 4_000 + per * 250))


def _cell(uv, cr, aov):
    return dict(uv=uv, cr=cr, aov=aov, cust=uv * cr, rev=uv * cr * aov)


def _roll(kids):
    """상위 = 하위 합. 단 **거래액만 정확히 가산**이고 고객수·상품UV는 유니크라 덜 합쳐진다
    (실파일에서 조직 합이 전체보다 고객수 +2.4%·UV +33% 많았다).
    객단가·CR은 합이 아니라 합계끼리 나눠 다시 만든다 — 퍼널 항등식을 유지하려고."""
    rev = sum(k["rev"] for k in kids)
    cust = sum(k["cust"] for k in kids) * 0.976
    uv = sum(k["uv"] for k in kids) * 0.75
    return dict(uv=uv, cr=(cust / uv if uv else np.nan),
                aov=(rev / cust if cust else np.nan), cust=cust, rev=rev)


def _tree_cells(per, depth):
    """{(org, cat, brand, item): cell} — depth 단계까지만 펼친다. 상위는 하위 합."""
    cells = {}

    def walk(node, path):
        """path 아래를 재귀로 채우고 그 합계를 돌려준다."""
        if len(path) >= depth or not node:
            c = _cell(*_leaf(path, per))
            cells[tuple(path)] = c
            return c
        kids = []
        for name, child in (node.items() if isinstance(node, dict)
                            else {k: None for k in node}.items()):
            kids.append(walk(child, path + [name]))
        agg = _roll(kids)
        cells[tuple(path)] = agg
        return agg

    kids = [walk(v, [k]) for k, v in TREE.items()]
    cells[()] = _roll(kids)
    return cells


def _grid_values(years, periods, plabel, depth):
    out = {}
    for y in years:
        for i in range(periods):
            per = i + (y - min(years)) * 100
            cells = _tree_cells(per, depth)
            gr = cells[()]["rev"] or np.nan
            gc = cells[()]["cust"] or np.nan
            for c in cells.values():
                c["rev_share"] = c["rev"] / gr
                c["cust_share"] = c["cust"] / gc
            out[(y, plabel(i))] = cells
    return out


def _rows_for(depth):
    """(구분06~09 값 4개, cells 키) 목록 — 원본과 같은 순서·병합 규칙으로 낸다."""
    out = [(["*TOTAL", "*TOTAL", "-", "-"], ())]
    for org, cats in TREE.items():
        out.append(([org, "*TOTAL", "-", "-"], (org,)))
        if not cats:
            out.append((["", "-", "-", "-"], (org,)))      # 카테고리 없는 조직
            continue
        for cat, brands in cats.items():
            deeper = depth > 2 and brands
            # 브랜드가 없으면 그 칸은 병합셀이라 **빈 칸**으로 온다. '-'로 채우면
            # 파서의 '상위가 바뀌면 하위 초기화' 경로를 안 밟아 버그가 숨는다.
            out.append((["", cat, "*TOTAL" if deeper else None, None], (org, cat)))
            if not deeper:
                continue
            for brand, items in brands.items():
                deep2 = depth > 3 and items
                out.append((["", "", brand, "*TOTAL" if deep2 else "-"],
                            (org, cat, brand)))
                if not deep2:
                    continue
                for it in items:
                    out.append((["", "", "", it], (org, cat, brand, it)))
    return out


def synth_grid(gran="월", years=(2025,), lfms="N", periods=3, mtd_last=False, depth=2):
    """실제 export와 같은 모양의 그리드(2차원 리스트).

    월·주는 마감 행이 따로 있고 일은 없다. mtd_last=True면 마지막 기간에 '일마감'(MTD)
    칼럼을 하나 더 붙인다 — 기간 라벨이 비어 직전 라벨을 이어받는 자리.
    depth=2면 지금 실파일처럼 브랜드·상품이 전부 '-'다.
    """
    def plabel(i):
        if gran == "월": return f"{i+1}월"
        if gran == "주": return f"{(i//5)+1:02d}월 {(i%5)+1}주차"
        return f"{(i//28)+1}/{(i%28)+1}"

    VALS = _grid_values(years, periods, plabel, depth)
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
    layout = _rows_for(depth)
    for met in METS:
        first = True
        for lv, key in layout:
            r = blank()
            if first: r[0] = met + " "          # 원본도 뒤에 공백이 붙어 온다
            first = False
            for j, v in enumerate(lv):
                r[1 + j] = v or None
            r[5] = r[6] = "-"                   # 구분10·11은 늘 비어 있다
            last_lbl = None
            for j, (y, lb, cl) in enumerate(cols):
                if lb: last_lbl = lb
                cell = VALS.get((y, last_lbl), {}).get(key)
                r[7 + j] = None if cell is None else cell[_MET_KEY[met]]
            rows.append(r)
    return rows


def synth_orgcat_df(**kw):
    return W.parse_orgcat_grid(synth_grid(**kw))


def gv(d, path, metric, year, label="1월", gran="월", lfms="N"):
    """테스트용 값 조회 — 화면과 같은 경로로 노드 하나를 집는다."""
    sub = d[(d["gran"] == gran) & (d["lfms"] == lfms)]
    return W.opick(sub, tuple(path), metric, year, label)


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
    # 깊은 레벨: 새 카테고리가 시작되면 그 아래(브랜드·상품)는 초기화돼야 한다.
    # 안 그러면 앞 카테고리의 브랜드가 따라붙어 값이 엉뚱한 자리에 쌓인다.
    d4 = synth_orgcat_df(depth=4)
    stale = d4[(d4["cat"] == "남성") & (d4["brand"] == "헤지스")]
    assert stale.empty, f"앞 카테고리의 브랜드가 따라붙었어요 — {len(stale)}행"


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
    a = gv(d, ("e-영업1", "골프"), "첫구매 거래액", 2025)
    b = gv(d, ("e-영업1", "골프"), "첫구매 거래액", 2026)
    assert np.isfinite(a) and np.isfinite(b) and a != b, f"연도별 값이 같아요 — {a} vs {b}"


@case
def t_mtd_column_inherits_label():
    """'일마감' 칼럼은 기간 라벨이 비어 직전 라벨을 이어받고 close=mtd가 된다."""
    d = synth_orgcat_df(mtd_last=True)
    mtd = d[d["close"] == "mtd"]
    assert not mtd.empty, "mtd 행이 없어요"
    assert set(mtd["label"]) == {"3월"}, f"직전 라벨을 못 이어받았어요 — {set(mtd['label'])}"
    # final이 있으면 final이 이긴다 (마스터 pick과 같은 규칙)
    assert gv(d, (), "첫구매 거래액", 2025, "3월") == \
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
        # 빈 레벨은 CSV에서 NaN으로 돌아온다 — 문자열로 안 되돌리면 키가 다 어긋난다
        assert set(back["brand"]) == {""}, f"빈 레벨이 NaN으로 남았어요 — {set(back['brand'])}"
        a = gv(back, ("e-영업1", "골프"), "첫구매 거래액", 2025)
        b = gv(d, ("e-영업1", "골프"), "첫구매 거래액", 2025)
        assert np.isfinite(a) and np.isfinite(b) and abs(a - b) <= abs(b) * 1e-12, \
            f"CSV 왕복에서 값이 변했어요 — {a!r} vs {b!r}"
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


def _plain(at):
    """조사 검사는 **강조**·<태그>를 걷어낸 평문으로 해야 한다.
    마크다운 원문은 '**객단가**이'라 '객단가이'로는 절대 안 걸린다."""
    import re as _re
    return _re.sub(r"\s+", " ",
                   _re.sub(r"[*_`]", "", _re.sub(r"<[^>]+>", " ", " ".join(_texts(at))))).strip()


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success, at.subheader):
        out += [str(e.value) for e in coll]
    return out


@case
def t_page_is_in_menu_and_renders():
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        txt = " ".join(_texts(at))
        for need in ("① 어디가 움직였나", "② 왜 그랬나", "③ 추이", "④ 한눈에"):
            assert need in txt, f"{need} 블록이 없어요"
        assert len(at.dataframe) >= 2, f"표가 모자라요 — {len(at.dataframe)}개"


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
        assert "① 어디가 움직였나" in txt, f"마스터 없이 안 열려요 — {txt[:400]}"


@case
def t_table_value_matches_source():
    """화면 숫자가 파일 값 그대로여야 한다 (일평균을 다시 나누거나 하지 않는다)."""
    d = synth_orgcat_df(years=(2025, 2026))
    with _run(orgcat=d) as at:
        tbls = [t for t in at.dataframe
                if "조직" == getattr(getattr(t.value, "data", t.value), "index", pd.Index([])).name]
        assert tbls, "조직별 표를 못 찾았어요"
        v = getattr(tbls[0].value, "data", tbls[0].value)
        got = str(v.loc["e-영업1", "2026년 3월"])
        want = W.fmt_value("첫구매 거래액",
                           gv(d, ("e-영업1",), "첫구매 거래액", 2026, "3월"))
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
        assert "① 어디가 움직였나" in txt, f"'일'로 바꾸니 빈 화면이 됐어요 — {txt[:400]}"
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
        assert any("부분 기간" in x for x in shown), f"부분 기간 표기가 없어요 — {shown[:4]}"
        txt = " ".join(_texts(at))
        assert "진행 중이에요" in txt, f"진행 중 안내가 없어요 — {txt[:400]}"


@case
def t_metric_switch_keeps_rendering():
    """지표를 바꿔도(비율 지표 포함) 정상 렌더돼야 한다."""
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        sels = [s for s in at.selectbox if s.label == "진단 지표"]
        assert sels, f"진단 지표 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
        for opt in list(sels[0].options):
            tgt = [s for s in at.selectbox if s.label == "진단 지표"][0]
            tgt.set_value(opt); at.run()
            assert not at.exception, f"지표={opt}에서 실패: {at.exception[0].value}"


@case
def t_factor_split_is_exact_and_order_free():
    """유입·전환·객단가 기여액을 더하면 실제 증감과 정확히 같아야 한다 (LMDI)."""
    d = synth_orgcat_df(years=(2025, 2026))

    for path in [(), ("e-영업1",), ("e-영업1", "골프")]:
        p = tuple(gv(d, path, m, 2025) for m, _ in W.ORGCAT_FACTORS)
        q = tuple(gv(d, path, m, 2026) for m, _ in W.ORGCAT_FACTORS)
        got = W.factor_split(p, q)
        assert got, f"{path} 분해 실패"
        parts, dv = got
        act = gv(d, path, "첫구매 거래액", 2026) - gv(d, path, "첫구매 거래액", 2025)
        assert abs(sum(parts) - act) < 1.0, f"{path}: 합 {sum(parts):,.2f} vs 실제 {act:,.2f}"
        assert abs(dv - act) < 1.0, f"{path}: 총증감 {dv:,.2f} vs {act:,.2f}"


@case
def t_factor_split_refuses_nonpositive():
    """0·음수·결측이 섞이면 로그가 정의되지 않는다 — 숫자를 지어내지 말고 None."""
    ok = (100.0, 0.05, 90_000.0)
    assert W.factor_split(ok, ok) is not None
    for bad in [(0.0, 0.05, 90_000.0), (100.0, -0.01, 90_000.0),
                (100.0, 0.05, float("nan")), (None, 0.05, 90_000.0)]:
        assert W.factor_split(bad, ok) is None, f"{bad}에서 None이 아니에요"
        assert W.factor_split(ok, bad) is None, f"{bad}에서 None이 아니에요"


@case
def t_only_additive_metric_gets_contribution():
    """거래액만 조직 합이 전체와 같다 — 유니크 지표에 기여도를 붙이면 거짓말이 된다."""
    assert W.ORGCAT_ADDITIVE == {"첫구매 거래액"}, W.ORGCAT_ADDITIVE
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        def _cols():
            for t in at.dataframe:
                v = getattr(t.value, "data", t.value)
                if getattr(getattr(v, "index", None), "name", "") == "조직":
                    return [str(c) for c in v.columns]
            return []
        assert "기여도" in _cols(), f"거래액인데 기여도 칼럼이 없어요 — {_cols()}"
        tgt = [s for s in at.selectbox if s.label == "진단 지표"][0]
        tgt.set_value("상품UV"); at.run()
        assert not at.exception, at.exception[0].value
        assert "기여도" not in _cols(), f"유니크 지표에 기여도가 붙었어요 — {_cols()}"
        assert any("유니크 값이라" in str(c.value) for c in at.caption), \
            "왜 기여도가 없는지 설명이 없어요"


@case
def t_contributions_sum_to_total_yoy():
    """기여도(%p)를 다 더하면 전체 전년비와 같아야 한다 (분모가 전년 전체라서)."""
    d = synth_orgcat_df(years=(2025, 2026))

    tp = gv(d, (), "첫구매 거래액", 2025)
    tc = gv(d, (), "첫구매 거래액", 2026)
    share = 0.0
    for o in W.orgcat_children(d[(d["gran"] == "월") & (d["lfms"] == "N")], ()):
        a, b = gv(d, (o,), "첫구매 거래액", 2025), gv(d, (o,), "첫구매 거래액", 2026)
        if np.isfinite(a) and np.isfinite(b):
            share += (b - a) / tp * 100
    assert abs(share - (tc / tp - 1) * 100) < 0.01, \
        f"기여도 합 {share:.3f}%p vs 전체 전년비 {(tc/tp-1)*100:.3f}%"


@case
def t_diagnosis_summary_names_the_driver():
    """요약 문장이 '어디가·왜'를 짚어야 한다 — 표만 있으면 원인을 못 읽는다."""
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        txt = " ".join(_texts(at))
        assert "가장 크게" in txt, f"최대 기여 대상 문장이 없어요 — {txt[:500]}"
        assert any(o in txt for o in ("e-영업1", "e-영업2", "SPACE-R")), "조직 이름이 없어요"
        assert "거래액 변화를 쪼개면" in txt and "유입" in txt, "요인 분해 문장이 없어요"


@case
def t_drilldown_switches_level_to_categories():
    """조직을 고르면 ①이 그 조직의 **카테고리별**로 바뀌어야 한다."""
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        def _idx_name():
            for t in at.dataframe:
                v = getattr(t.value, "data", t.value)
                nm = getattr(getattr(v, "index", None), "name", "")
                if nm in ("조직", "카테고리"):
                    return nm, list(v.index)
            return None, []
        assert _idx_name()[0] == "조직", f"기본이 조직별이 아니에요 — {_idx_name()}"
        o = [s for s in at.selectbox if s.label == "1. 조직"][0]
        assert "e-영업1" in o.options, f"조직 선택지가 없어요 — {list(o.options)}"
        o.set_value("e-영업1"); at.run()
        assert not at.exception, at.exception[0].value
        nm, idx = _idx_name()
        assert nm == "카테고리", f"카테고리 레벨로 안 내려갔어요 — {nm}"
        assert "골프" in idx, f"e-영업1의 카테고리가 아니에요 — {idx}"
        assert "e-영업1" in " ".join(_texts(at)), "브레드크럼에 조직이 안 보여요"
        c = [s for s in at.selectbox if s.label == "2. 카테고리"][0]
        assert "골프" in c.options, f"카테고리 선택지가 안 열렸어요 — {list(c.options)}"
        c.set_value("골프"); at.run()
        assert not at.exception, at.exception[0].value
        assert "② 왜 그랬나" in " ".join(_texts(at)), "말단에서 요인 분해가 없어요"


def hand_frame(spec):
    """(org, cat, year) → (UV, CR, AOV) 명세로 long 프레임을 직접 만든다.

    합성 그리드는 세 요인이 같이 움직여서 '객단가가 1위'인 상황이 안 나온다 —
    조사처럼 특정 분기에서만 드러나는 버그를 잡으려면 값을 손으로 박아야 한다.
    """
    rec = []
    for (org, cat, year), (uv, cr, aov) in spec.items():
        vals = {"상품UV": uv, "상품CR": cr, "첫구매 객단가": aov,
                "첫구매 고객수": uv * cr, "첫구매 거래액": uv * cr * aov}
        for met, v in vals.items():
            rec.append(dict(gran="월", metric=met, org=org, cat=cat,
                            brand="", item="", lfms="N",
                            year=year, label="1월", close="final",
                            sortkey=year * 10000 + 100, value=float(v)))
    return pd.DataFrame(rec)[W.ORGCAT_COLS]


@case
def t_josa_matches_final_consonant():
    """'객단가이 가장'·'조직는' 같은 문장이 나오면 안 된다 (발송성과에서 겪은 버그)."""
    for w, eun, i_ga in [("조직", "은", "이"), ("카테고리", "는", "가"),
                         ("유입", "은", "이"), ("전환", "은", "이"),
                         ("객단가", "는", "가"), ("슈즈", "는", "가"),
                         ("e-영업1", "은", "이"), ("SPACE-R", "은", "이")]:
        assert W.josa(w) == eun, f"{w}+{W.josa(w)} (기대 {eun})"
        assert W.josa(w, "이가") == i_ga, f"{w}+{W.josa(w, '이가')} (기대 {i_ga})"

    # 객단가만 움직여 '객단가가 가장 크게 움직였어요'가 실제로 렌더되는 데이터
    spec = {}
    for org in ("*TOTAL", "e-영업1", "e-영업2"):
        base = 1000.0 if org == "*TOTAL" else 500.0
        spec[(org, "*TOTAL", 2025)] = (base, 0.10, 100_000.0)
        spec[(org, "*TOTAL", 2026)] = (base, 0.10, 150_000.0)
    BAD = ("조직는", "카테고리은", "객단가이", "유입가", "전환가")
    with _run(orgcat=hand_frame(spec)) as at:
        txt = _plain(at)
        assert "객단가가 가장 크게" in txt, f"객단가 주범 문장이 안 나왔어요 — {txt[:500]}"
        for bad in BAD:
            assert bad not in txt, f"조사가 틀렸어요 — '{bad}'"
    # 조직·카테고리 레벨 문장도 함께 (기본 합성 데이터: 유입·전환이 주범)
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        for bad in BAD:
            assert bad not in _plain(at), f"조사가 틀렸어요 — '{bad}'"
        [s for s in at.selectbox if s.label == "1. 조직"][0].set_value("e-영업1")
        at.run()
        assert not at.exception, at.exception[0].value
        txt = _plain(at)
        for bad in BAD:
            assert bad not in txt, f"조직 레벨에서 조사가 틀렸어요 — '{bad}'"


@case
def t_org_names_are_escaped():
    """조직·카테고리 이름은 업로드 파일에서 온다 — &·< 가 태그로 새면 안 된다."""
    assert W.esc("A&B <b>") == "A&amp;B &lt;b&gt;", W.esc("A&B <b>")


@case
def t_depth_follows_the_data():
    """뎁스는 코드가 아니라 **데이터**가 정한다 — 브랜드·상품 칸이 채워지면 그만큼 열린다."""
    assert [c for c, _lb, _h in W.ORGCAT_LEVELS] == ["org", "cat", "brand", "item"], \
        W.ORGCAT_LEVELS
    for depth in (2, 3, 4):
        d = synth_orgcat_df(depth=depth)
        assert W.orgcat_depth(d) == depth, f"depth={depth}인데 {W.orgcat_depth(d)}로 읽었어요"
    d2 = synth_orgcat_df(depth=2)
    assert set(d2["brand"]) == {""} and set(d2["item"]) == {""}, \
        "브랜드·상품이 안 비었어요 (원본 '-'는 빈 칸이어야 해요)"


@case
def t_children_walk_the_hierarchy():
    """한 단계 아래만 나와야 한다 — 합계(*TOTAL)·빈 칸이 섞이면 값이 두 번 세진다."""
    d = synth_orgcat_df(depth=4)
    sub = d[(d["gran"] == "월") & (d["lfms"] == "N")]
    assert W.orgcat_children(sub, ()) == ["e-영업1", "e-영업2", "SPACE-R", "미매칭"], \
        W.orgcat_children(sub, ())
    assert W.orgcat_children(sub, ("e-영업1",)) == ["골프", "남성", "잡화"]
    # 브랜드로 안 쪼개지는 카테고리는 그 아래가 비어 있어야 한다
    assert W.orgcat_children(sub, ("e-영업1", "잡화")) == []
    assert W.orgcat_children(sub, ("e-영업1", "골프")) == ["헤지스", "닥스"]
    assert W.orgcat_children(sub, ("e-영업1", "골프", "헤지스")) == ["티셔츠", "바지"]
    assert W.orgcat_children(sub, ("e-영업1", "골프", "헤지스", "티셔츠")) == []
    for kids in (W.orgcat_children(sub, ()), W.orgcat_children(sub, ("e-영업1",))):
        assert "*TOTAL" not in kids and "" not in kids, kids


@case
def t_every_level_is_additive_for_revenue():
    """거래액은 어느 단계에서도 하위 합 = 상위여야 기여도 분해가 성립한다."""
    d = synth_orgcat_df(depth=4)
    sub = d[(d["gran"] == "월") & (d["lfms"] == "N")]
    for path in [(), ("e-영업1",), ("e-영업1", "골프")]:
        parent = gv(d, path, "첫구매 거래액", 2025)
        kids = W.orgcat_children(sub, path)
        tot = sum(gv(d, tuple(path) + (k,), "첫구매 거래액", 2025) for k in kids)
        assert abs(tot / parent - 1) < 1e-6, \
            f"{path}: 하위 합 {tot:,.0f} vs 상위 {parent:,.0f}"


@case
def t_deep_drill_opens_brand_and_item():
    """브랜드·상품 값이 들어오면 3·4단 셀렉트가 코드 수정 없이 생겨야 한다."""
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026), depth=4)) as at:
        def _labels():
            return [s.label for s in at.selectbox if s.label[:1].isdigit()]
        assert _labels() == ["1. 조직"], f"처음엔 1단계만 보여야 해요 — {_labels()}"
        [s for s in at.selectbox if s.label == "1. 조직"][0].set_value("e-영업1"); at.run()
        assert not at.exception, at.exception[0].value
        assert "2. 카테고리" in _labels(), _labels()
        [s for s in at.selectbox if s.label == "2. 카테고리"][0].set_value("골프"); at.run()
        assert not at.exception, at.exception[0].value
        assert "3. 브랜드" in _labels(), f"브랜드 단계가 안 열렸어요 — {_labels()}"
        b = [s for s in at.selectbox if s.label == "3. 브랜드"][0]
        assert "헤지스" in b.options, list(b.options)
        b.set_value("헤지스"); at.run()
        assert not at.exception, at.exception[0].value
        assert "4. 상품" in _labels(), f"상품 단계가 안 열렸어요 — {_labels()}"
        it = [s for s in at.selectbox if s.label == "4. 상품"][0]
        assert "티셔츠" in it.options, list(it.options)
        it.set_value("티셔츠"); at.run()
        assert not at.exception, at.exception[0].value
        txt = _plain(at)
        assert "전체 › e-영업1 › 골프 › 헤지스 › 티셔츠" in txt, \
            f"브레드크럼이 끝까지 안 갔어요 — {txt[:400]}"
        assert "더 들어갈 단계가 없어요" in txt, "말단인데 안내가 없어요"


@case
def t_two_level_data_says_deeper_levels_are_empty():
    """지금처럼 브랜드·상품이 비어 있으면 왜 단계가 없는지 화면이 말해 줘야 한다."""
    with _run(orgcat=synth_orgcat_df(years=(2025, 2026), depth=2)) as at:
        txt = _plain(at)
        assert "브랜드·상품" in txt and "안 열려요" in txt, f"안내가 없어요 — {txt[:400]}"
        assert not [s for s in at.selectbox if s.label.startswith("3.")], "빈 단계가 열렸어요"


@case
def t_hides_items_without_meaningful_data():
    """미매칭·기타처럼 실적이 사실상 없는 항목은 기본으로 숨긴다 (전 기간 0.1% 미만)."""
    d = synth_orgcat_df(years=(2025, 2026), depth=2)
    sub = d[(d["gran"] == "월") & (d["lfms"] == "N")]
    live, hidden = W.orgcat_live(sub, ())
    assert "SPACE-R" not in live, f"실적 없는 조직이 남았어요 — {live}"
    assert "e-영업1" in live and "e-영업2" in live, live
    with _run(orgcat=d) as at:
        o = [s for s in at.selectbox if s.label == "1. 조직"][0]
        assert "SPACE-R" not in o.options, f"숨겨야 할 조직이 선택지에 있어요 — {list(o.options)}"
        assert any("숨겼어요" in str(c.value) for c in at.caption), "숨긴 사실을 안 밝혔어요"
        box = [c for c in at.checkbox if "실적 없는" in c.label]
        assert box, f"보기 토글이 없어요 — {[c.label for c in at.checkbox]}"
        box[0].set_value(True); at.run()
        assert not at.exception, at.exception[0].value
        o = [s for s in at.selectbox if s.label == "1. 조직"][0]
        assert "SPACE-R" in o.options, f"토글을 켜도 안 나와요 — {list(o.options)}"


@case
def t_old_two_level_backup_still_restores():
    """브랜드·상품 레벨이 없던 시절 백업도 그대로 살아나야 한다."""
    d = synth_orgcat_df(depth=2)
    old = d.drop(columns=["brand", "item"])
    back = W.orgcat_fill(old)
    assert list(back.columns) == W.ORGCAT_COLS, list(back.columns)
    assert set(back["brand"]) == {""} and set(back["item"]) == {""}
    assert len(back) == len(d)
    # 빈 칸이 NaN으로 온 CSV도 마찬가지
    nanned = old.copy(); nanned["brand"] = np.nan
    assert set(W.orgcat_fill(nanned)["brand"]) == {""}


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
