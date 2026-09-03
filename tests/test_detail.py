"""브랜드·상품 결제 원장 파싱·저장·큐브·화면 테스트.

MICRO export(구분06~09)는 브랜드·상품 칸이 늘 비어 온다. 브랜드·상품은 **결제 원장**으로
따로 오는데, 모양이 아주 다르다 — 한 줄이 (결제일자 × 조직 × 유입채널 × 카테고리 × 브랜드 ×
상품) 하나고, 값은 그날의 **합계**다.

여기서 조용히 틀리는 방식이 정해져 있다:
  · 두 원천을 한 store에 섞으면 **카테고리 어휘가 달라** 축이 뭉개진다.
  · 합계를 그대로 쓰면 3일치 이번 달이 30일치 전년 달과 맞붙어 △90% 가짜 급락이 뜬다.
  · 상품까지 큐브를 전부 펼치면 (기간 × 채널 × 상품)이 수백만 행이라 화면이 멈춘다.

로컬 실행:
    python tests/test_detail.py
"""
import contextlib
import datetime
import io
import os
import pathlib
import shutil
import sys
import tempfile
import zlib

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import weekly_report as W                        # noqa: E402
from smoke_weekly_report import synth_store      # noqa: E402

TIMEOUT = 300
HDR = ["결제_일자(YYYYMMDD)", "BPU", "AF대분류명", "대카테고리명", "ADMIN브랜드명",
       "상품코드", "상품명", "거래액", "주문고객수"]
# 실파일의 어휘 — MICRO의 구분07(골프·남성·잡화)과 **겹치지 않는다**.
TREE = {
    "e-영업1": {"가방": {"닥스 액세서리": ["크로스백", "서류가방"],
                      "질 바이 질스튜어트": ["케미백"]},
              "남성의류": {"일꼬르소": ["발마칸 코트"], "티엔지티": ["다운 점퍼"]}},
    "e-영업2": {"여성의류": {"아떼 바네사브루노": ["캐시미어 롱코트"]}},
    "e-영업4": {"향수": {"샤넬": ["코코 마드모아젤"]}},
}
CHS = ["광고", "직접", "EP"]


def _amt(path, day, ch):
    """말단 한 칸의 (거래액, 고객수) — crc32로 고정해 프로세스마다 같게."""
    h = zlib.crc32("|".join(map(str, path)).encode() + ch.encode()) % 40 + 1
    cust = h % 4 + 1
    return float(h * 30_000 + day * 700 + cust * 1_000), float(cust)


def synth_rows(days, chs=CHS, header=HDR, sep_rows=True):
    """실제 원장과 같은 모양의 그리드 — 헤더 1행 + 거래 행들."""
    rows = [list(header)]
    for d in days:
        for org, cats in TREE.items():
            for cat, brands in cats.items():
                for brand, items in brands.items():
                    for it in items:
                        for ch in chs:
                            rev, cust = _amt((org, cat, brand, it), d.day, ch)
                            rows.append([d.strftime("%Y%m%d"), org, ch, cat, brand,
                                         f"C{abs(hash(it)) % 10000:04d}", it,
                                         f"{rev:,.0f}", f"{cust:.0f}"])
    return rows


def days_in(y, m, n=None):
    import calendar
    last = calendar.monthrange(y, m)[1]
    return [datetime.date(y, m, i + 1) for i in range(n or last)]


def synth_detail(days=None):
    return W.parse_detail_grid(synth_rows(days or days_in(2025, 1)))


def as_tsv(rows):
    return "\n".join("\t".join(str(c) for c in r) for r in rows).encode("utf-8")


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# ── 파싱 ────────────────────────────────────────────────────────────
@case
def t_detects_and_parses():
    rows = synth_rows(days_in(2025, 1, 3))
    assert W.is_detail_grid(rows), "원장 형식 인식 실패"
    d = W.parse_detail_grid(rows)
    assert not d.empty and list(d.columns) == W.DETAIL_COLS, list(d.columns)
    assert set(d["org"]) == set(TREE), set(d["org"])
    assert set(d["ch"]) == set(CHS), set(d["ch"])
    assert d["rev"].dtype.kind == "f" and d["cust"].dtype.kind == "f", d.dtypes.to_dict()


@case
def t_header_names_can_wobble():
    """헤더 표기가 흔들려도 조각 부분 일치로 찾아야 한다 (수기 export라 자주 바뀐다)."""
    alt = ["결제 일자", "BPU명", "AF 대분류", "카테고리(대)", "브랜드명(ADMIN)",
           "상품 코드", "상품명칭", "거래액(원)", "주문 고객수"]
    d = W.parse_detail_grid(synth_rows(days_in(2025, 1, 2), header=alt))
    base = W.parse_detail_grid(synth_rows(days_in(2025, 1, 2)))
    assert not d.empty and d.equals(base), "헤더 변형에서 결과가 달라졌어요"


@case
def t_utf8_file_is_not_read_as_utf16():
    """길이가 짝수인 UTF-8 파일은 UTF-16으로도 예외 없이 풀린다 — 먼저 재 보면
    줄바꿈이 사라진 한 줄짜리 쓰레기가 되고, 증상은 '미인식'으로만 보인다."""
    raw = as_tsv(synth_rows(days_in(2025, 1, 2)))
    if len(raw) % 2:                       # 짝수 길이여야 UTF-16 오디코딩이 재현된다
        raw += b"\n"
    got = W._detail_rows("x.tsv", raw)
    assert len(got) > 5, f"줄이 안 끊겼어요 — {len(got)}줄"
    assert not W.parse_detail_file("x.tsv", raw).empty, "UTF-8 원장을 못 읽었어요"
    # UTF-16(BOM)으로 온 파일도 그대로 읽혀야 한다 (태블로 export가 이 형식이다)
    r16 = raw.decode("utf-8").encode("utf-16")
    assert not W.parse_detail_file("x.tsv", r16).empty, "UTF-16 원장을 못 읽었어요"


@case
def t_comma_csv_also_parses():
    """read_grid는 탭만 끊는다 — 원장은 콤마 CSV로도 온다."""
    import csv
    rows = synth_rows(days_in(2025, 1, 2))
    buf = io.StringIO()
    for r in rows:
        csv.writer(buf).writerow(r)
    d = W.parse_detail_file("x.csv", buf.getvalue().encode("utf-8"))
    assert not d.empty, "콤마 CSV를 못 읽었어요"
    assert d.equals(W.parse_detail_grid(rows)), "콤마 CSV 결과가 탭과 달라요"


@case
def t_numbers_keep_sign_and_commas():
    """거래액엔 천단위 콤마가 붙고 반품은 음수다 — 문자열로 남으면 합계가 이어붙기가 된다."""
    rows = [list(HDR),
            ["20250101", "e-영업1", "광고", "패션소품", "질 바이 질스튜어트",
             "J1", "머플러", "-63,818", "-1"],
            ["20250101", "e-영업3", "EP", "아동의류", "마리떼 키즈",
             "E1", "데님 팬츠", "87,991", "0"]]
    d = W.parse_detail_grid(rows)
    assert list(d["rev"]) == [-63818.0, 87991.0], list(d["rev"])
    assert list(d["cust"]) == [-1.0, 0.0], list(d["cust"])


@case
def t_date_forms_and_junk_rows():
    """YYYYMMDD·YYYY-MM-DD·엑셀 날짜셀을 다 읽고, 날짜가 아닌 줄(합계행 등)은 버린다."""
    rows = [list(HDR),
            ["20250101", "o", "광고", "c", "b", "x", "i", "100", "1"],
            ["2025-01-02", "o", "광고", "c", "b", "x", "i", "100", "1"],
            [datetime.datetime(2025, 1, 3), "o", "광고", "c", "b", "x", "i", "100", "1"],
            ["합계", "o", "광고", "c", "b", "x", "i", "999999", "9"],
            ["", "", "", "", "", "", "", "", ""]]
    d = W.parse_detail_grid(rows)
    assert list(d["date"]) == ["2025-01-01", "2025-01-02", "2025-01-03"], list(d["date"])
    assert d["rev"].sum() == 300, d["rev"].sum()


@case
def t_blank_level_becomes_its_own_node():
    """빈 브랜드 칸을 지우면 그 매출이 조용히 사라지고, 상위에 흡수시키면 하위 합이 안 맞는다."""
    rows = [list(HDR),
            ["20250101", "e-영업1", "광고", "가방", "", "x", "이름없음", "1000", "1"],
            ["20250101", "e-영업1", "광고", "가방", "닥스", "y", "크로스백", "2000", "1"]]
    d = W.parse_detail_grid(rows)
    assert W.DETAIL_NA in set(d["brand"]), sorted(set(d["brand"]))
    assert d["rev"].sum() == 3000


@case
def t_same_key_rows_are_summed():
    """상품코드(색상)로 갈린 같은 상품명은 한 줄로 합친다 — 코드는 축이 아니라 안 쌓는다."""
    rows = [list(HDR),
            ["20250101", "o", "광고", "c", "b", "CODE-BK", "다운 점퍼", "100", "1"],
            ["20250101", "o", "광고", "c", "b", "CODE-W2", "다운 점퍼", "200", "2"]]
    d = W.parse_detail_grid(rows)
    assert len(d) == 1 and d["rev"].iloc[0] == 300 and d["cust"].iloc[0] == 3, d.to_dict("records")
    assert "code" not in d.columns, list(d.columns)


# ── 라우팅·저장 ─────────────────────────────────────────────────────
@case
def t_not_routed_into_other_stores():
    """마스터·조직×카테고리 파서가 원장을 집어삼키면 안 된다.

    파일명이 마스터로 보여도(`전체관점 - 일자별 …`) 원장으로 가야 한다 — 이름으로
    가르면 원장 한 벌이 통째로 누적 데이터에 섞여 들어간다."""
    raw = as_tsv(synth_rows(days_in(2025, 1, 2)))
    for nm in ("원장.tsv", "전체관점 - 일자별 실적 (기본).csv", "월_가입율(일평균).csv"):
        assert W.route_push(nm, raw) == (None, None), f"[{nm}] {W.route_push(nm, raw)}"
        assert W.parse_orgcat_file(nm, raw).empty, f"[{nm}] 조직×카테고리 파서가 먹었어요"
        assert not W.parse_detail_file(nm, raw).empty, f"[{nm}] 원장으로 안 읽혔어요"
        got = W.classify_uploads(((nm, raw),))
        assert "원장" in got[0][1], f"[{nm}] 인식 목록이 원장이라고 안 해요 — {got}"


@case
def t_shown_in_upload_recognition_list():
    """'왜 안 올라가지'를 없애려면 인식 목록에 따로 찍혀야 한다."""
    raw = as_tsv(synth_rows(days_in(2025, 1, 2)))
    got = W.classify_uploads((("원장.tsv", raw),))
    assert got and "원장" in got[0][1] and got[0][1].startswith("✅"), got


@case
def t_merge_prefers_new_and_keeps_history():
    old = synth_detail(days_in(2025, 1, 5))
    new = synth_detail(days_in(2025, 1, 3))
    new = new.copy(); new["rev"] = 1.0
    m = W.merge_detail(old, new)
    assert len(m) == len(old), f"행이 늘었어요 — {len(old)} → {len(m)}"
    assert (m[m["date"] == "2025-01-01"]["rev"] == 1.0).all(), "신규가 안 이겼어요"
    assert (m[m["date"] == "2025-01-05"]["rev"] != 1.0).all(), "안 겹친 날짜가 덮였어요"


@case
def t_store_csv_roundtrip():
    """gsheets/CSV 라운드트립이 숫자를 문자열로, 빈 칸을 NaN으로 돌려준다."""
    d = synth_detail(days_in(2025, 1, 3))
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    try:
        os.chdir(tmp)
        W.save_detail_store(d)
        assert os.path.exists(W.DETAIL_STORE), "저장 파일이 없어요"
        back = W.load_detail_store()
        assert len(back) == len(d), f"{len(d)} → {len(back)}"
        assert back["rev"].dtype.kind == "f", back["rev"].dtype
        a = d.sort_values(W.DETAIL_KEY).reset_index(drop=True)
        b = back.sort_values(W.DETAIL_KEY).reset_index(drop=True)
        assert np.allclose(a["rev"], b["rev"]) and np.allclose(a["cust"], b["cust"])
        assert (a[W.DETAIL_KEY].values == b[W.DETAIL_KEY].values).all(), "키가 어긋났어요"
    finally:
        os.chdir(cwd); shutil.rmtree(tmp, ignore_errors=True)


@case
def t_backup_zip_roundtrip():
    """통합 백업 ZIP에 원장이 들어가고 그대로 복원돼야 한다."""
    import zipfile
    d = synth_detail(days_in(2025, 1, 3))
    z = W.make_backup_zip(synth_store(), {"a": "b"}, None, d)
    with zipfile.ZipFile(io.BytesIO(z)) as zf:
        names = zf.namelist()
        assert "wr_detail_store.csv" in names, names
        back = W.detail_fill(pd.read_csv(io.BytesIO(zf.read("wr_detail_store.csv")),
                                         encoding="utf-8-sig"))
    assert len(back) == len(d) and np.allclose(sorted(back["rev"]), sorted(d["rev"]))


@case
def t_gzip_store_upload_restores():
    """`.csv.gz`로 받은 원장을 그대로 올려도 복원돼야 한다 — 버퍼로 주면 pandas가
    확장자를 못 봐서 압축을 추론하지 못한다."""
    import gzip
    d = synth_detail(days_in(2025, 1, 2))
    raw = gzip.compress(d.to_csv(index=False).encode("utf-8-sig"))

    class _U:
        name = "wr_detail_store.csv.gz"
        def getvalue(self): return raw

    tmp = tempfile.mkdtemp(); cwd = os.getcwd()
    try:
        os.chdir(tmp)
        msg, changed = W.restore_from_upload(_U())
        assert changed and "원장" in msg, msg
        assert len(W.load_detail_store()) == len(d)
    finally:
        os.chdir(cwd); shutil.rmtree(tmp, ignore_errors=True)


# ── 큐브 ────────────────────────────────────────────────────────────
@case
def t_values_are_daily_means_not_sums():
    """합계로 내면 3일치 이번 달이 30일치 전년 달과 맞붙어 △90% 가짜 급락이 뜬다."""
    full = synth_detail(days_in(2025, 1))                       # 31일 전부
    part = synth_detail(days_in(2025, 1, 3))                    # 3일치
    vals = []
    for d in (full, part):
        c = W.detail_provider(W.detail_stamp(d, "월"), "월", W.DETAIL_ALL)(())
        vals.append(W.opick(c, (), "첫구매 거래액", 2025, "1월"))
    ratio = vals[1] / vals[0]
    assert 0.9 < ratio < 1.1, f"일평균이 아니에요 — 3일치가 한 달의 {ratio:.2f}배"


@case
def t_partial_period_is_marked_mtd():
    """적재 범위가 달을 다 못 덮으면 close=mtd다 — 완결 기간과 무게가 다르다고 알려야 한다."""
    s = W.detail_stamp(synth_detail(days_in(2025, 1, 3)), "월")
    assert set(s["close"]) == {"mtd"}, set(s["close"])
    assert set(s["nd"]) == {3}, set(s["nd"])
    s2 = W.detail_stamp(synth_detail(days_in(2025, 1)), "월")
    assert set(s2["close"]) == {"final"} and set(s2["nd"]) == {31}, (set(s2["close"]), set(s2["nd"]))


@case
def t_missing_zero_sales_day_does_not_inflate():
    """매출이 0인 날은 원장에 줄이 없다. 줄이 있는 날만 세면 일평균이 부풀려진다 —
    나눌 일수는 **적재 범위와 겹치는 달력 일수**여야 한다."""
    days = [d for d in days_in(2025, 1) if d.day not in (10, 11, 12)]
    s = W.detail_stamp(synth_detail(days), "월")
    assert set(s["nd"]) == {31}, f"관측된 날짜 수로 나눴어요 — {set(s['nd'])}"


@case
def t_revenue_is_exactly_additive():
    """기여도 분해가 거짓말이 되지 않으려면 하위 합이 상위와 정확히 같아야 한다."""
    d = synth_detail(days_in(2025, 1, 5))
    cube = W.detail_provider(W.detail_stamp(d, "월"), "월", W.DETAIL_ALL)
    c = cube(())
    tot = W.opick(c, (), "첫구매 거래액", 2025, "1월")
    kids = W.orgcat_children(c, ())
    ksum = sum(W.opick(c, (k,), "첫구매 거래액", 2025, "1월") for k in kids)
    assert abs(tot - ksum) < 1e-6, f"조직 합 {ksum} ≠ 전체 {tot}"
    # 한 단계 더: 카테고리 합도 조직과 같아야 한다
    org = kids[0]
    c1 = cube((org,))
    o = W.opick(c1, (org,), "첫구매 거래액", 2025, "1월")
    csum = sum(W.opick(c1, (org, k), "첫구매 거래액", 2025, "1월")
               for k in W.orgcat_children(c1, (org,)))
    assert abs(o - csum) < 1e-6, f"카테고리 합 {csum} ≠ {org} {o}"


@case
def t_funnel_identity_holds():
    """거래액 = 고객수 × 객단가가 정확히 성립해야 LMDI 분해가 실제 증감과 딱 맞는다."""
    d = synth_detail(days_in(2025, 1, 5))
    c = W.detail_provider(W.detail_stamp(d, "월"), "월", W.DETAIL_ALL)(())
    for path in ((), ("e-영업1",), ("e-영업1", "가방")):
        rv = W.opick(c, path, "첫구매 거래액", 2025, "1월")
        cu = W.opick(c, path, "첫구매 고객수", 2025, "1월")
        ao = W.opick(c, path, "첫구매 객단가", 2025, "1월")
        assert abs(rv - cu * ao) < 1e-6, f"{path}: {rv} ≠ {cu}×{ao}"


@case
def t_aov_refuses_nonpositive_customers():
    """반품으로 고객수가 0 이하가 되면 객단가를 지어내지 않고 비운다."""
    rows = [list(HDR),
            ["20250101", "o", "광고", "c", "b", "x", "i", "87991", "0"],
            ["20250101", "o", "광고", "c", "b2", "y", "j", "-1000", "-1"]]
    c = W.detail_provider(W.detail_stamp(W.parse_detail_grid(rows), "일"),
                          "일", W.DETAIL_ALL)(())
    for path in ((), ("o",)):
        v = W.opick(c, path, "첫구매 객단가", 2025, "1/1")
        assert not np.isfinite(v), f"{path}: 객단가를 지어냈어요 — {v}"


@case
def t_channel_axis_splits_and_totals():
    """유입채널은 모집단이 다른 축이다 — 채널별 합이 '전체'와 같아야 한다."""
    d = synth_detail(days_in(2025, 1, 4))
    s = W.detail_stamp(d, "월")
    tot = W.opick(W.detail_provider(s, "월", W.DETAIL_ALL)(()), (),
                  "첫구매 거래액", 2025, "1월")
    per = [W.opick(W.detail_provider(s, "월", ch)(()), (), "첫구매 거래액", 2025, "1월")
           for ch in CHS]
    assert abs(tot - sum(per)) < 1e-6, f"채널 합 {sum(per)} ≠ 전체 {tot}"
    assert all(p > 0 for p in per), per


@case
def t_cube_stays_small_at_the_root():
    """상품까지 다 펼치면 (기간 × 채널 × 상품)이 수백만 행이라 화면이 멈춘다.
    루트에선 손자(카테고리)까지만 만들어야 한다."""
    d = synth_detail(days_in(2025, 1, 6))
    cube = W.detail_provider(W.detail_stamp(d, "일"), "일", W.DETAIL_ALL)
    c = cube(())
    assert set(c["brand"]) == {W.ORGCAT_TOTAL}, "루트에서 브랜드까지 펼쳤어요"
    assert set(c["item"]) == {W.ORGCAT_TOTAL}, "루트에서 상품까지 펼쳤어요"
    # 파고들면 그 아래만 열린다
    c2 = cube(("e-영업1", "가방"))
    assert set(c2["item"]) != {W.ORGCAT_TOTAL}, "파고들어도 상품이 안 열려요"
    got = set(c2[c2["item"] != W.ORGCAT_TOTAL]["brand"])
    assert got <= set(TREE["e-영업1"]["가방"]), f"경로 밖 브랜드가 섞였어요 — {got}"


@case
def t_cube_keeps_siblings_for_every_select():
    """경로를 파고들어도 위 단계 셀렉트의 형제 목록이 사라지면 안 된다."""
    d = synth_detail(days_in(2025, 1, 4))
    cube = W.detail_provider(W.detail_stamp(d, "월"), "월", W.DETAIL_ALL)
    c = cube(("e-영업1", "가방", "닥스 액세서리"))
    assert set(W.orgcat_children(c, ())) == set(TREE), W.orgcat_children(c, ())
    assert set(W.orgcat_children(c, ("e-영업1",))) == set(TREE["e-영업1"])
    assert set(W.orgcat_children(c, ("e-영업1", "가방"))) == set(TREE["e-영업1"]["가방"])
    assert set(W.orgcat_children(c, ("e-영업1", "가방", "닥스 액세서리"))) \
        == set(TREE["e-영업1"]["가방"]["닥스 액세서리"])
    assert W.orgcat_depth(c) == 4, W.orgcat_depth(c)


@case
def t_week_belongs_to_the_month_of_its_thursday():
    """월요일로 달을 정하면 한 주가 두 달에 걸릴 때 어느 달인지가 흔들린다."""
    d = synth_detail([datetime.date(2025, 2, 1)])       # 토 — 그 주 목요일은 1/30
    s = W.detail_stamp(d, "주")
    assert set(s["label"]) == {"01월 5주차"}, set(s["label"])
    assert set(s["year"]) == {2025}, set(s["year"])
    d2 = synth_detail([datetime.date(2025, 1, 6)])      # 월 — 그 주 목요일은 1/9
    assert set(W.detail_stamp(d2, "주")["label"]) == {"01월 2주차"}
    # 연말: 그 주 목요일이 1월이면 다음 해 1월 1주차다 (정렬키도 따라가야 한다)
    d3 = synth_detail([datetime.date(2025, 12, 30)])    # 화 — 그 주 목요일은 2026-01-01
    s3 = W.detail_stamp(d3, "주")
    assert set(s3["year"]) == {2026} and set(s3["label"]) == {"01월 1주차"}, \
        (set(s3["year"]), set(s3["label"]))
    assert set(s3["sortkey"]) == {2026 * 10000 + 100 + 1}, set(s3["sortkey"])


@case
def t_period_labels_match_the_master_format():
    """라벨이 마스터·MICRO와 같은 꼴이어야 두 소스를 오갈 때 눈이 안 흔들린다."""
    d = synth_detail(days_in(2025, 3, 5))
    assert set(W.detail_stamp(d, "월")["label"]) == {"3월"}
    assert set(W.detail_stamp(d, "일")["label"]) == {f"3/{i}" for i in range(1, 6)}
    for gran in ("월", "주", "일"):
        s = W.detail_stamp(d, gran)
        for y, lb, sk in set(zip(s["year"], s["label"], s["sortkey"])):
            assert W.period_parts(gran, int(y), str(lb)) is not None, (gran, lb)


# ── 요인 분해 ───────────────────────────────────────────────────────
@case
def t_factor_split_handles_two_factors():
    """원장엔 상품UV·CR이 없어 (고객수, 객단가) 둘로 쪼갠다 — 3요인 전용이면 터진다."""
    got = W.factor_split((100.0, 50_000.0), (120.0, 45_000.0))
    assert got is not None, "2요인 분해가 None을 돌려줬어요"
    parts, tot = got
    assert len(parts) == 2, parts
    assert abs(sum(parts) - tot) < 1e-6, f"합 {sum(parts)} ≠ 실제 증감 {tot}"
    assert abs(tot - (120 * 45_000 - 100 * 50_000)) < 1e-6


@case
def t_factor_split_is_order_free():
    """LMDI라 요인 순서를 바꿔도 각 기여액이 같아야 한다."""
    a, b = (100.0, 50_000.0), (120.0, 45_000.0)
    p1, _ = W.factor_split(a, b)
    p2, _ = W.factor_split(a[::-1], b[::-1])
    assert abs(p1[0] - p2[1]) < 1e-6 and abs(p1[1] - p2[0]) < 1e-6, (p1, p2)
    # 3요인(MICRO)도 그대로 돌아야 한다
    p3, t3 = W.factor_split((10.0, 0.05, 70_000.0), (12.0, 0.06, 68_000.0))
    assert len(p3) == 3 and abs(sum(p3) - t3) < 1e-6


@case
def t_factor_split_refuses_nonpositive():
    for prev, cur in (((0.0, 50_000.0), (120.0, 45_000.0)),
                      ((100.0, 50_000.0), (-1.0, 45_000.0)),
                      ((100.0, np.nan), (120.0, 45_000.0))):
        assert W.factor_split(prev, cur) is None, (prev, cur)
    assert W.factor_split((), ()) is None
    assert W.factor_split((1.0, 2.0), (1.0,)) is None


# ── 화면 ────────────────────────────────────────────────────────────
@contextlib.contextmanager
def _run(page="08. 조직·카테고리별 실적", detail=None, orgcat=None, master=True):
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
    if detail is not None and not detail.empty:
        W.detail_fill(detail).to_csv(os.path.join(tmp, W.DETAIL_STORE),
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


def _sel(at, label):
    got = [s for s in at.selectbox if s.label == label]
    assert got, f"{label} 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
    return got[0]


def two_years(n=8):
    return W.merge_detail(synth_detail(days_in(2025, 1, n)),
                          synth_detail(days_in(2026, 1, n)))


@case
def t_page_renders_with_detail_only():
    """MICRO가 없어도 원장만으로 08이 열려야 한다."""
    with _run(detail=two_years()) as at:
        txt = " ".join(_texts(at))
        assert "조직·카테고리별 첫구매 실적" in txt, txt[:300]
        assert "원장" in txt, txt[:400]
        assert not [s for s in at.radio if s.label == "데이터 소스"], \
            "소스가 하나뿐인데 소스 라디오가 떴어요"


@case
def t_source_switch_appears_with_both():
    """두 원천이 있으면 갈라 볼 수 있어야 한다 — 카테고리 어휘가 달라 섞으면 안 된다."""
    from test_orgcat import synth_orgcat_df
    with _run(detail=two_years(), orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        rs = [r for r in at.radio if r.label == "데이터 소스"]
        assert rs, f"데이터 소스 라디오가 없어요 — {[r.label for r in at.radio]}"
        assert len(rs[0].options) == 2, list(rs[0].options)
        # MICRO 쪽은 구분08·09가 비어 브랜드·상품 단계가 없다 — 왜 없는지 화면에 밝힌다
        rs[0].set_value([o for o in rs[0].options if "MICRO" in str(o)][0]); at.run()
        assert not at.exception, at.exception[0].value
        assert "LFMS 포함" in [r.label for r in at.radio] + [s.label for s in at.selectbox] \
            or "LFMS 포함여부" in " ".join(_texts(at)), "MICRO 축(LFMS)이 안 보여요"
        assert "브랜드·상품" in " ".join(_texts(at)), "왜 단계가 없는지 안 알려줘요"
        _sel(at, "1. 조직").set_value("e-영업1"); at.run()
        _sel(at, "2. 카테고리").set_value("골프"); at.run()
        assert not at.exception, at.exception[0].value
        labs = [s.label for s in at.selectbox]
        assert "3. 브랜드" not in labs, f"MICRO에 없는 단계가 열렸어요 — {labs}"
        # 원장 쪽으로 바꾸면 같은 자리에서 브랜드까지 열린다
        rs = [r for r in at.radio if r.label == "데이터 소스"][0]
        rs.set_value([o for o in rs.options if "원장" in str(o)][0]); at.run()
        assert not at.exception, at.exception[0].value
        _sel(at, "1. 조직").set_value("e-영업1"); at.run()
        _sel(at, "2. 카테고리").set_value("가방"); at.run()
        assert not at.exception, at.exception[0].value
        assert "3. 브랜드" in [s.label for s in at.selectbox], [s.label for s in at.selectbox]


@case
def t_drill_opens_brand_then_item():
    """조직 > 카테고리 > 브랜드 > 상품 — 고를 때마다 한 단계씩 열려야 한다."""
    with _run(detail=two_years()) as at:
        _sel(at, "1. 조직").set_value("e-영업1"); at.run()
        assert not at.exception, at.exception[0].value
        assert "가방" in list(_sel(at, "2. 카테고리").options), \
            list(_sel(at, "2. 카테고리").options)
        _sel(at, "2. 카테고리").set_value("가방"); at.run()
        assert not at.exception, at.exception[0].value
        assert "닥스 액세서리" in list(_sel(at, "3. 브랜드").options)
        _sel(at, "3. 브랜드").set_value("닥스 액세서리"); at.run()
        assert not at.exception, at.exception[0].value
        items = list(_sel(at, "4. 상품").options)
        assert "크로스백" in items, items
        _sel(at, "4. 상품").set_value("크로스백"); at.run()
        assert not at.exception, at.exception[0].value
        txt = " ".join(_texts(at))
        assert "크로스백" in txt and "e-영업1 › 가방 › 닥스 액세서리" in txt, txt[:500]


@case
def t_table_value_matches_source():
    """화면 값이 원장 집계와 같아야 한다 (일평균 환산 포함)."""
    d = two_years()
    with _run(detail=d) as at:
        cube = W.detail_provider(W.detail_stamp(d, "월"), "월", W.DETAIL_ALL)(())
        want = W.opick(cube, ("e-영업1",), "첫구매 거래액", 2026, "1월")
        assert np.isfinite(want)
        got = [str(x) for df in at.dataframe for x in np.asarray(df.value).ravel()]
        assert any(W.fmt_value("첫구매 거래액", want) == g for g in got), \
            f"{W.fmt_value('첫구매 거래액', want)}가 표에 없어요 — {got[:12]}"


@case
def t_channel_filter_changes_the_numbers():
    """채널을 고르면 그 채널만 봐야 한다 — 안 걸리면 필터가 장식이다."""
    d = two_years()
    with _run(detail=d) as at:
        box = _sel(at, "유입 채널")
        assert W.DETAIL_ALL in list(box.options) and "광고" in list(box.options), \
            list(box.options)
        box.set_value("광고"); at.run()
        assert not at.exception, at.exception[0].value
        cube = W.detail_provider(W.detail_stamp(d, "월"), "월", "광고")(())
        want = W.opick(cube, (), "첫구매 거래액", 2026, "1월")
        assert W.fmt_value("첫구매 거래액", want) in " ".join(_texts(at)), \
            f"광고 채널 값 {W.fmt_value('첫구매 거래액', want)}이 안 보여요"


@case
def t_factor_block_says_customers_and_aov():
    """원장엔 상품UV·CR이 없다 — ②가 유입·전환을 요구하면 통째로 빈다."""
    with _run(detail=two_years()) as at:
        txt = " ".join(_texts(at))
        assert "고객수·객단가" in txt, f"② 제목이 원장용이 아니에요 — {txt[:600]}"
        assert "요인 분해를 할 수 없어요" not in txt, "요인 분해가 통째로 비었어요"


@case
def t_no_crash_on_all_metrics_and_grans():
    """지표·집계 단위를 바꿔도 안 죽어야 한다 (일 단위는 라벨이 90개까지 간다)."""
    with _run(detail=two_years()) as at:
        for m in W.DETAIL_METS:
            _sel(at, "진단 지표").set_value(m); at.run()
            assert not at.exception, f"[{m}] {at.exception[0].value}"
        for g in ("월", "주", "일"):
            r = [x for x in at.radio if x.label == "집계 단위"][0]
            r.set_value(g); at.run()
            assert not at.exception, f"[{g}] {at.exception[0].value}"


@case
def t_stale_selection_does_not_kill_the_page():
    """소스·채널을 바꾸면 세션에 남은 옛 선택이 옵션 밖이 된다 — 가드가 없으면 예외로 죽는다."""
    from test_orgcat import synth_orgcat_df
    with _run(detail=two_years(), orgcat=synth_orgcat_df(years=(2025, 2026))) as at:
        rs = [r for r in at.radio if r.label == "데이터 소스"][0]
        rs.set_value([o for o in rs.options if "원장" in str(o)][0]); at.run()
        _sel(at, "1. 조직").set_value("e-영업1"); at.run()
        _sel(at, "2. 카테고리").set_value("가방"); at.run()
        assert not at.exception, at.exception[0].value
        rs = [r for r in at.radio if r.label == "데이터 소스"][0]
        rs.set_value([o for o in rs.options if "MICRO" in str(o)][0]); at.run()
        assert not at.exception, f"소스를 되돌리자 죽었어요 — {at.exception[0].value}"


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
    print(f"브랜드·상품 원장 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
