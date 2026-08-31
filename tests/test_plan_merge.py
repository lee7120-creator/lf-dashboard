"""실적 ↔ 기획 문구 조인 경로 테스트.

발송성과 대시보드는 실적 엑셀을 올리고 문구는 기획 구글시트에서 가져와
`(발송일, AF코드)` 정확 일치로 붙인다. 이 경로가 조용히 깨지면 화면은 멀쩡히
뜨는데 문구 칸만 전부 비어서 스모크로는 절대 안 잡힌다. 그래서 따로 둔다.

로컬 실행:
    python tests/test_plan_merge.py
"""
import datetime
import io
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import send_perf_dashboard as S  # noqa: E402

PERF_HEADER = ["날짜", "요일", "시간대", "타겟 구분", "발송유형", "BPU", "우선순위",
               "카테고리", "속성", "담당자", "브랜드", "AF코드", "기획전", "발송",
               "UV", "VISIT", "고객수", "주문건수", "주문금액",
               "유입전환율", "주문전환율", "효율"]


def perf_xlsx(rows):
    """일자별 '누적_소재별' 형태의 실적 엑셀을 메모리에 만든다."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "누적_소재별"
    ws.append(PERF_HEADER)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def row(date, af, hour=800, send=100000, uv=5000, oc=50, amt=5_000_000):
    return [date, "월", hour, "타겟", "기본발송", "편성", 1, "통합", "통합", "홍길동",
            "브랜드", af, 113126, send, uv, uv + 500, uv + 400, oc, amt,
            uv / send, oc / send, amt / send]


class FakeWS:
    def __init__(self, title, rows):
        self.title = title
        self._rows = rows

    def get_all_values(self):
        return self._rows


class FakeSH:
    def __init__(self, wss):
        self._wss = wss

    def worksheets(self):
        return self._wss


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_upload_classified_as_perf():
    """일자별 실적 파일이 통합 업로더에서 perf로 분류돼야 한다."""
    b = perf_xlsx([row("20260803", "AP03")])
    got = S.classify_upload("누적_소재별_0805.xlsx", b)
    assert got == "perf", f"perf가 아니라 {got}로 분류됨"


@case
def t_exact_key_match():
    """(날짜, AF코드) 정확 일치 건에만 문구가 붙는다."""
    b = perf_xlsx([row("20260803", "AP03"), row("20260803", "AP02"),
                   row("20260804", "AP03")])
    perf = S.parse_perf_bytes(b)
    lookup = {("20260803", "AP03"): ("위켄드세일", "5% 추가적립")}
    m = S.merge_perf_plan(perf, lookup, keep_unmatched=True)
    assert list(m["matched"]) == [True, False, False], list(m["matched"])
    assert m.loc[0, "title"] == "위켄드세일"
    # 같은 AF코드가 다음 날 재사용돼도 전날 문구가 새어 나가면 안 된다
    assert m.loc[2, "title"] == "", "AF코드만으로 폴백 매칭되고 있음"


@case
def t_af_case_insensitive():
    """실적/기획의 AF코드 대소문자가 달라도 붙어야 한다."""
    b = perf_xlsx([row("20260803", "ap03")])
    perf = S.parse_perf_bytes(b)
    m = S.merge_perf_plan(perf, {("20260803", "AP03"): ("제목", "내용")},
                          keep_unmatched=True)
    assert bool(m.loc[0, "matched"]), "소문자 AF코드가 매칭 안 됨"


@case
def t_six_digit_date_recovered():
    """기획 시트의 6자리 날짜 오타(260803)도 20260803으로 복구돼 매칭된다."""
    rows = [["일자", "AF코드", "제목", "내용"],
            ["260803", "AP03", "제목", "내용"]]
    lookup = {}
    S._parse_plan_sheet(rows, lookup)
    assert ("20260803", "AP03") in lookup, lookup


@case
def t_merged_date_cell_forward_fill():
    """날짜 병합셀(둘째 줄부터 빈칸)이면 직전 날짜를 이어받는다."""
    rows = [["일자", "AF코드", "제목", "내용"],
            ["20260803", "AP03", "제목1", "내용1"],
            ["", "AP02", "제목2", "내용2"],
            [None, "AP04", "제목3", "내용3"]]
    lookup = {}
    S._parse_plan_sheet(rows, lookup)
    for af in ("AP03", "AP02", "AP04"):
        assert ("20260803", af) in lookup, f"{af} 날짜 forward-fill 실패: {lookup}"


@case
def t_no_date_anywhere_is_skipped():
    """끝내 날짜를 못 구하면 (None, af) 키로 오염시키지 말고 버린다."""
    rows = [["AF코드", "제목", "내용"], ["AP03", "제목", "내용"]]
    lookup = {}
    S._parse_plan_sheet(rows, lookup)
    assert all(k[0] is not None for k in lookup), lookup


@case
def t_gsheet_week_window():
    """금주 주차 시트는 읽고, 아직 안 온 차주 시트는 건너뛴다."""
    t = S.today_kst()
    mon = t - datetime.timedelta(days=t.weekday())
    sun = mon + datetime.timedelta(days=6)
    nxt = sun + datetime.timedelta(days=7)

    def nm(a, b):
        return f"{a.month}월 {((a.day - 1) // 7) + 1}주차({a.month:02d}{a.day:02d}~{b.month:02d}{b.day:02d})"

    this_t, next_t = nm(mon, sun), nm(sun + datetime.timedelta(days=1), nxt)
    sh = FakeSH([
        FakeWS(this_t, [["일자", "AF코드", "제목", "내용"],
                        [mon.strftime("%Y%m%d"), "AP03", "금주제목", "금주내용"]]),
        FakeWS(next_t, [["일자", "AF코드", "제목", "내용"],
                        [nxt.strftime("%Y%m%d"), "AP09", "차주제목", "차주내용"]]),
    ])
    lookup, read = S.parse_plan_gsheet(sh, recent=12)
    assert this_t in read, f"금주 시트({this_t})를 안 읽음 — read={read}"
    assert next_t not in read, f"차주 시트({next_t})를 읽어 버림 — read={read}"
    assert (mon.strftime("%Y%m%d"), "AP03") in lookup, lookup


@case
def t_undated_sheet_survives_recent_cap():
    """제목에 기간이 없는 주차 시트도 읽어야 한다.

    실제로 이번 주 시트에 '(810~816)'이 아직 안 붙어 있으면, 예전 코드는
    recent(기본 12)가 설정된 기본 경로에서 그 시트를 통째로 버렸다.
    화면엔 '최신 주 매칭률 0%'로만 뜨고 원인이 안 보여서, 업로드가 실패한 것처럼 읽혔다.
    """
    t = S.today_kst()
    mon = t - datetime.timedelta(days=t.weekday())
    d = mon.strftime("%Y%m%d")
    sh = FakeSH([
        FakeWS("8월 2주차", [["일자", "AF코드", "제목", "내용"],      # ← 기간 없음
                            [d, "AP100", "이번주제목", "이번주내용"]]),
        FakeWS("7월 1주차(701~707)", [["일자", "AF코드", "제목", "내용"],
                                     ["20260701", "AP01", "지난제목", "지난내용"]]),
    ])
    lookup, read = S.parse_plan_gsheet(sh, recent=12)
    assert "8월 2주차" in read, f"기간 없는 시트를 버렸어요 — read={read}"
    assert (d, "AP100") in lookup, lookup


@case
def t_undated_sheets_are_capped():
    """기간 없는 시트가 많아도 recent 만큼만 읽는다 (API 호출 폭증 방지)."""
    sh = FakeSH([FakeWS(f"{i % 12 + 1}월 2주차",
                        [["일자", "AF코드", "제목", "내용"],
                         ["20260810", f"AP{i:03d}", "t", "b"]]) for i in range(30)])
    _lk, read = S.parse_plan_gsheet(sh, recent=5)
    assert len(read) == 5, f"{len(read)}개를 읽었어요 — recent=5로 캡돼야 해요"


@case
def t_end_to_end_upload_then_plan():
    """업로드 실적 + 기획시트 lookup → 문구가 붙은 저장 가능한 프레임이 나온다."""
    t = S.today_kst()
    mon = t - datetime.timedelta(days=t.weekday())
    d = mon.strftime("%Y%m%d")
    perf = S.parse_perf_bytes(perf_xlsx([row(d, "AP03"), row(d, "AP02", hour=1000)]))
    lookup = {}
    S._parse_plan_sheet([["일자", "AF코드", "제목", "내용"],
                         [d, "AP03", "제목A", "내용A"],
                         ["", "AP02", "제목B", "내용B"]], lookup)
    m = S.merge_perf_plan(perf, lookup, keep_unmatched=True)
    assert m["matched"].all(), m[["date", "af", "matched"]].to_dict("records")
    keep = [c for c in S.STORE_COLS if c in m]
    assert {"title", "body", "matched"} <= set(keep), keep
    # 시간대는 HHMM 그대로 살아 있어야 일자 비교 화면의 시간축이 맞는다
    assert sorted(int(h) for h in m["hour"]) == [800, 1000], list(m["hour"])


@case
def t_exec_note_reads_four_tuple_lookup():
    """주간보고 「금주 집행 내용 요약」의 자동 생성이 4-튜플 lookup을 읽어야 한다.

    lookup 값은 W·X열(문구변경·잔여모수)이 붙으면서 `(제목, 내용, 문구변경, 잔여모수)`
    4-튜플이 됐는데, 이 버튼만 `(title, body)` 2-튜플로 고정해 풀고 있어 운영에서
    ValueError로 죽었다. 세션에 lookup이 있어야만 밟는 경로라 스모크로는 안 잡힌다."""
    from streamlit.testing.v1 import AppTest
    sys.path.insert(0, str(ROOT / "tests"))
    from smoke_pages import synth_store

    t = S.today_kst()
    mon = t - datetime.timedelta(days=t.weekday())
    d = mon.strftime("%Y%m%d")
    lookup = {}
    S._parse_plan_sheet([["일자", "AF코드", "제목", "내용"],
                         [d, "AP03", "제목A", "내용A"],
                         [d, "AP04", "제목B", "내용B"]], lookup)
    assert lookup and all(len(v) == 4 for v in lookup.values()), \
        f"lookup 값이 4-튜플이 아니에요 — {list(lookup.values())[:2]}"

    at = AppTest.from_file(str(ROOT / "send_perf_dashboard.py"), default_timeout=300)
    at.session_state["camp_store"] = synth_store(weeks=6)
    at.session_state["plan_lookup_gs"] = lookup
    at.run()
    assert not at.exception, at.exception[0].value
    at.sidebar.radio[0].set_value("0. 주간보고")
    at.run()
    assert not at.exception, at.exception[0].value

    bkey = f"btn_r_exec_{d}"
    btns = [b for b in at.button if b.proto.id.endswith(bkey) or bkey in str(b.proto.id)]
    assert btns, f"'{bkey}' 자동 생성 버튼을 못 찾았어요"
    btns[0].click()
    at.run()
    assert not at.exception, f"자동 생성에서 죽었어요: {at.exception[0].value}"
    body = " ".join(str(e.value) for e in at.markdown)
    assert "집행 예정" in body and "제목A" in body, \
        f"이번 주 기획 문구가 안 채워졌어요 — {body[-400:]}"


def main():
    fails = []
    for fn in CASES:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {e}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"문구 조인 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
