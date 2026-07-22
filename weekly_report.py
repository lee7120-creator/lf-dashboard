"""주간보고 — 첫구매 통합 실적 대시보드
여러 원천 엑셀(전체관점 마스터 + 지표별 파일)을 업로드하면 하나의 통합 뷰로 종합하고,
전년(YoY)·전주(WoW) 증감현황과 보고란을 갖춘 주간보고 화면을 만든다.
통합 결과는 (월/주/첫구매_요약 + 차트) 엑셀 워크북으로 다운로드할 수 있다.
"""

import datetime
import io, json, os, re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    from streamlit_quill import st_quill
    HAS_QUILL = True
except Exception:
    HAS_QUILL = False

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="주간보고 — 첫구매 통합 실적", page_icon="📋",
                   layout="wide", initial_sidebar_state="expanded")

INSIGHT_FILE = "wr_insights.json"

KST = datetime.timezone(datetime.timedelta(hours=9))
def today_kst():
    """서버가 UTC라 date.today()는 KST 새벽(00~09시)에 하루 밀린다 — '오늘'은 반드시 이걸 쓸 것"""
    return datetime.datetime.now(KST).date()

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.sdiv{border-top:1px solid #e2e8f0;margin:22px 0}
.report-box{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin:8px 0;line-height:1.7;background:#ffffff}
.report-box p{margin:0 0 4px}
.kpi-card{background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
.kpi-label{color:#64748b;font-size:12px}
.kpi-value{color:#1e293b;font-size:21px;font-weight:600;margin:2px 0 8px}
.kpi-delta{display:block;width:fit-content;font-size:12px;border-radius:6px;padding:2px 8px;margin:4px 0 0;font-weight:500;white-space:nowrap}
.kpi-delta.up{background:#ecfdf5;color:#15803d}
.kpi-delta.down{background:#fef2f2;color:#dc2626}
.kpi-delta.na{background:#f1f5f9;color:#94a3b8}
@media print {
  @page { margin: 12mm; }
  [data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stToolbar"],
  header, .stButton, .no-print, iframe,
  [data-testid="stExpander"] { display:none !important; }
  [data-testid="stAppViewContainer"], .main, .block-container { background:#fff !important; }
  .block-container { max-width:100% !important; padding-top:0 !important; }

  /* 겹침 방지: 레이아웃 블록을 정적 배치하고 넘침을 그대로 노출 */
  [data-testid="stVerticalBlock"], [data-testid="stHorizontalBlock"],
  [data-testid="stVerticalBlockBorderWrapper"], [data-testid="column"],
  .element-container { position:static !important; transform:none !important;
    overflow:visible !important; }

  /* 인쇄 시 Plotly 차트 높이 붕괴 → 아래 요소가 제목 위로 밀려 겹치는 문제 차단 */
  .stPlotlyChart, .js-plotly-plot, [data-testid="stPlotlyChart"] {
    min-height:240px !important; break-inside:avoid; page-break-inside:avoid; }

  /* 제목이 투명 배경으로 다른 요소와 겹쳐 보이지 않도록 */
  h1, h2, h3, h4 { background:#fff !important; position:relative; z-index:1;
    page-break-after:avoid; break-after:avoid; }

  .stPlotlyChart, .report-box, table,
  [data-testid="stMetric"], [data-testid="column"] {
    break-inside:avoid; page-break-inside:avoid; }

  /* 증감 색상(빨강/초록) 인쇄에 유지 */
  * { -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 색상 팔레트 (기존 first_purchase 와 동일 계열)
# ══════════════════════════════════════════════════════
PALETTE = {
    "blue":   ("rgba(79,143,255,1)",  "rgba(79,143,255,0.15)"),
    "red":    ("rgba(245,101,101,1)", "rgba(245,101,101,0.15)"),
    "green":  ("rgba(72,187,120,1)",  "rgba(72,187,120,0.15)"),
    "amber":  ("rgba(237,137,54,1)",  "rgba(237,137,54,0.15)"),
    "purple": ("rgba(159,122,234,1)", "rgba(159,122,234,0.15)"),
    "teal":   ("rgba(56,178,172,1)",  "rgba(56,178,172,0.15)"),
    "orange": ("rgba(249,115,22,1)",  "rgba(249,115,22,0.15)"),
    "slate":  ("rgba(100,116,139,1)", "rgba(100,116,139,0.15)"),
}
def clr(n): return PALETTE.get(n, PALETTE["blue"])[0]
def cbg(n): return PALETTE.get(n, PALETTE["blue"])[1]

CHANNEL_PAL = {
    "직접": "blue", "광고": "amber", "EP": "green", "PUSH": "purple",
    "제휴": "red", "브랜드광고": "teal", "미디어커머스": "orange", "*TOTAL": "slate",
}
CHANNELS = ["직접", "광고", "EP", "PUSH", "제휴", "브랜드광고", "미디어커머스"]
YEAR_PAL = ["slate", "blue", "red", "green", "purple", "amber", "teal"]

# ══════════════════════════════════════════════════════
# 지표 정의
# ══════════════════════════════════════════════════════
METRICS7 = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가",
            "비회원트래픽", "가입자수", "가입율", "당일가입CR"]
PCT_METRICS = {"가입율", "당일가입CR", "진입률", "CR", "거래액비중", "고객비중", "동의율", "유입율"}
# 마스터 파일 지표 → 보고서 지표 매핑
MASTER_MAP = {"일평균거래액": "첫구매 거래액", "일평균고객수": "첫구매 고객수",
              "일평균객단가": "첫구매 객단가"}
# 지표별 파일명 → 보고서 지표 매핑 (공백 제거 후 매칭)
METRIC_FILE_MAP = {
    "가입율": "가입율", "가입률": "가입율",
    "가입자수": "가입자수",
    "당일가입첫구매율": "당일가입CR", "당일가입CR": "당일가입CR",
    "비회원트래픽": "비회원트래픽",
}

METRIC_UNIT = {
    "첫구매 거래액": ("백만원", 1e6), "첫구매 고객수": ("명", 1),
    "첫구매 객단가": ("원", 1), "비회원트래픽": ("명", 1),
    "가입자수": ("명", 1), "가입율": ("%", 1), "당일가입CR": ("%", 1),
    "앱푸시수신동의": ("명", 1), "앱푸시_동의자수": ("명", 1),
    "앱푸시_신규추가": ("명", 1), "앱푸시_이탈": ("명", 1),
}

def fmt_value(metric, v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "–"
    if metric in PCT_METRICS: return f"{v*100:.2f}%"
    if metric == "첫구매 거래액": return f"{v/1e6:,.1f}백만원"
    if metric == "첫구매 객단가": return f"{v:,.0f}원"
    return f"{int(v):,}명"  # 명 단위(고객수 등)는 소수점 버림

def fmt_delta(metric, cur, prev):
    """전년비/전주비 문자열: 비율 지표는 %p 차이, 그 외 증감율"""
    if cur is None or prev is None: return None
    if isinstance(cur, float) and np.isnan(cur): return None
    if isinstance(prev, float) and np.isnan(prev): return None
    if metric in PCT_METRICS:
        d = (cur - prev) * 100
        return f"△{abs(d):.2f}%p" if d < 0 else f"+{d:.2f}%p"
    if prev == 0: return None
    d = (cur - prev) / prev * 100
    return f"△{abs(d):.1f}%" if d < 0 else f"+{d:.1f}%"

def style_delta_cols(tbl):
    """증감 컬럼(△/+)에 빨강/초록 색상 적용한 Styler 반환"""
    delta_cols = [c for c in tbl.columns
                  if any(k in str(c) for k in ("전년비", "전주비", "전월비", "증감"))]
    def _color(v):
        s = str(v)
        if s.startswith("△"): return "color:#dc2626;font-weight:600"
        if s.startswith("+"): return "color:#16a34a;font-weight:600"
        return ""
    try:
        return tbl.style.map(_color, subset=delta_cols)
    except Exception:
        return tbl

# ══════════════════════════════════════════════════════
# 파싱 — 원천 파일 → 통합 long DataFrame
# ══════════════════════════════════════════════════════
YEAR_RE   = re.compile(r"^(20\d{2})(\.0)?$")
PERIOD_RE = re.compile(r"^\s*(\d{1,2}\s*월(\s*\d\s*주차)?|\d{1,2}/\d{1,2})\s*$")

def detect_file(name):
    """파일명에서 (kind, granularity 힌트, metric) 감지.
    단위(일/주/월)와 마스터 여부는 parse_file에서 내용으로 최종 판별하므로 여기선 힌트만."""
    base = os.path.basename(name)
    base = re.sub(r"\.(xlsx|xls|csv)$", "", base, flags=re.I)
    key = base.replace(" ", "")
    # 단위 힌트: 일자별/주별/월별 키워드 또는 일_/주_/월_ 접두사
    gran = None
    if "일자별" in key or "데일리" in key: gran = "일"
    elif "주별" in key or "주간" in key:   gran = "주"
    elif "월별" in key or "월간" in key:   gran = "월"
    else:
        m = re.match(r"^(일|주|월)[_\s]", base)
        if m: gran = m.group(1)
    if "전체관점" in key or "마스터" in key:
        return "master", gran, None
    # 지표 키워드는 파일명 어느 위치에 있어도 인식
    for k, v in METRIC_FILE_MAP.items():
        if k.replace(" ", "") in key:
            return "metric", gran, v
    # 접두사형(월_xxx)인데 모르는 지표면 정리된 이름 그대로 사용
    m = re.match(r"^(일|주|월)[_\s]+(.+)$", base)
    if m:
        rest = re.sub(r"\(.*?\)", "", m.group(2)).strip()
        return "metric", m.group(1), rest or "기타"
    return None, gran, None

def _decode_text(data: bytes) -> str:
    for enc in ("utf-16", "utf-8-sig", "cp949", "utf-8"):
        try: return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError): pass
    return data.decode("utf-8", "replace")

def read_grid(name, data: bytes):
    """csv(UTF-16/탭 포함)·xlsx 모두 2차원 셀 그리드로 읽는다"""
    if name.lower().endswith((".csv", ".txt", ".tsv")):
        text = _decode_text(data)
        return [line.split("\t") for line in text.splitlines()]
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    return [list(row) for row in ws.iter_rows(values_only=True)]

def _cell(v):
    return "" if v is None else str(v).strip()

def _num(v):
    """셀 값 → float (콤마·% 처리, %는 비율로 변환)"""
    if v is None: return np.nan
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "–"): return np.nan
    pct = s.endswith("%")
    if pct: s = s[:-1]
    try: f = float(s)
    except ValueError: return np.nan
    return f / 100 if pct else f

def period_parts(gran, year, plabel):
    """기간 라벨 → (표준라벨, 정렬키)"""
    p = plabel.replace(" ", "")
    if gran == "월":
        m = re.match(r"(\d{1,2})월", p)
        if not m: return None
        mo = int(m.group(1))
        return f"{mo}월", year * 10000 + mo * 100
    if gran == "주":
        m = re.match(r"(\d{1,2})월(\d)주차", p)
        if not m: return None
        mo, wk = int(m.group(1)), int(m.group(2))
        return f"{mo:02d}월 {wk}주차", year * 10000 + mo * 100 + wk
    m = re.match(r"(\d{1,2})/(\d{1,2})", p)
    if not m: return None
    mo, dd = int(m.group(1)), int(m.group(2))
    return f"{mo}/{dd}", year * 10000 + mo * 100 + dd

def parse_file(name, data: bytes) -> pd.DataFrame:
    # 누적 데이터 백업 CSV(long 포맷) 재업로드 시 그대로 복원
    if name.lower().endswith(".csv"):
        first = data[:400].decode("utf-8", "ignore").splitlines()[0] if data else ""
        if "gran" in first and "metric" in first and "sortkey" in first:
            try:
                d = pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
                if set(STORE_COLS) <= set(d.columns):
                    return d[STORE_COLS]
            except Exception:
                pass
    kind, gran, file_metric = detect_file(name)
    try:
        rows = read_grid(name, data)
    except Exception:
        return pd.DataFrame()
    if not rows: return pd.DataFrame()

    # 헤더 행 탐색 (앞 6행): 연도 / 기간라벨 / 마감구분
    year_row = period_row = close_row = None
    for ri in range(min(6, len(rows))):
        cells = [_cell(c) for c in rows[ri]]
        nY = sum(1 for c in cells if YEAR_RE.match(c))
        nP = sum(1 for c in cells if PERIOD_RE.match(c))
        nC = sum(1 for c in cells if "마감" in c)
        if year_row is None and nY >= 1 and nP == 0: year_row = ri
        if period_row is None and nP >= 2: period_row = ri
        if close_row is None and nC >= 2: close_row = ri
    if period_row is None: return pd.DataFrame()
    data_start = max(r for r in (year_row, period_row, close_row) if r is not None) + 1

    # ── 내용 기반 최종 판별 (파일명은 힌트일 뿐)
    # 단위: 기간 라벨 형태가 결정 (N월 N주차 → 주, N/N → 일, N월 → 월)
    plabels = [_cell(c) for c in rows[period_row] if PERIOD_RE.match(_cell(c))]
    if any("주차" in p for p in plabels):  gran = "주"
    elif any("/" in p for p in plabels):   gran = "일"
    else:                                   gran = "월"
    # 마스터 여부: 헤더에 '구분01'이 있으면 마스터(전체관점) 구조
    if any("구분01" in _cell(c) for ri in range(data_start) for c in rows[ri]):
        kind = "master"
    if kind is None or (kind == "metric" and not file_metric):
        return pd.DataFrame()

    ncols = max(len(r) for r in rows)
    def cell(ri, ci):
        return _cell(rows[ri][ci]) if ri is not None and ci < len(rows[ri]) else ""

    # 컬럼별 연도(좌측 ffill)·기간·마감
    col_year, cur_y = {}, None
    for ci in range(ncols):
        m = YEAR_RE.match(cell(year_row, ci)) if year_row is not None else None
        if m: cur_y = int(m.group(1))
        col_year[ci] = cur_y
    # 기간 라벨이 빈 셀(병합)인 일마감/MTD 컬럼은 직전 라벨을 이어받는다
    data_cols, col_label, last_lbl = [], {}, None
    for ci in range(ncols):
        lbl = cell(period_row, ci)
        if PERIOD_RE.match(lbl) and col_year[ci]:
            last_lbl = lbl
            col_label[ci] = lbl
            data_cols.append(ci)
        elif (close_row is not None and "마감" in cell(close_row, ci)
              and last_lbl and col_year[ci]):
            col_label[ci] = last_lbl
            data_cols.append(ci)

    seg_col = 1 if kind == "master" else 0
    records, cur_metric = [], None
    for ri in range(data_start, len(rows)):
        if kind == "master":
            m0 = cell(ri, 0)
            if m0 and m0 not in ("-", "–"): cur_metric = m0
            metric = MASTER_MAP.get(cur_metric, cur_metric)
        else:
            metric = file_metric
        seg = cell(ri, seg_col)
        # '채널'은 세그먼트 컬럼의 헤더 텍스트 (해당 행 값은 요일) — 데이터 행이 아니므로 스킵
        if not seg or seg in ("-", "–", "채널") or metric is None: continue
        for ci in data_cols:
            pp = period_parts(gran, col_year[ci], col_label[ci])
            if pp is None: continue
            label, sortkey = pp
            close = cell(close_row, ci) if close_row is not None else ""
            records.append({
                "gran": gran, "metric": metric, "segment": seg,
                "year": col_year[ci], "label": label, "sortkey": sortkey,
                "close": "mtd" if "일마감" in close and gran != "일" else "final",
                "value": _num(rows[ri][ci] if ci < len(rows[ri]) else None),
            })
    return pd.DataFrame(records)

def _zip_entry_name(info):
    """zip 내 한글 파일명 복원 (UTF-8 플래그 없으면 cp437→cp949 재해석)"""
    if info.flag_bits & 0x800:
        return info.filename
    for enc in ("cp949", "utf-8", "euc-kr"):
        try: return info.filename.encode("cp437").decode(enc)
        except (UnicodeDecodeError, UnicodeEncodeError): pass
    return info.filename

def expand_uploads(uploads):
    """업로드 목록 → (이름, bytes) 목록. zip은 풀어서 내부 엑셀/CSV를 꺼낸다"""
    out = []
    for f in uploads:
        data = f.getvalue()
        if f.name.lower().endswith(".zip"):
            import zipfile
            try:
                zf = zipfile.ZipFile(io.BytesIO(data))
            except zipfile.BadZipFile:
                continue
            with zf:
                for info in zf.infolist():
                    if info.is_dir(): continue
                    name = _zip_entry_name(info)
                    if "__MACOSX" in name: continue
                    base = os.path.basename(name)
                    if base.startswith((".", "~")): continue
                    if not base.lower().endswith((".xlsx", ".xls", ".csv")): continue
                    out.append((base, zf.read(info)))
        else:
            out.append((f.name, data))
    return out

# (지표·세그먼트·연도)별 일 |중앙값| 대비 이 배수 초과 시 일회성 스파이크로 간주.
# 실측: 정상 피크 최대 2.9배(빼빼로데이 등 실제 이벤트 2.4~2.9배) vs
#       인공 스파이크 최소 5.4배(7/28~30 이관성 이탈 5.4~8.7배, 재동의·재분류 40~711배)
#       → 4로 설정하면 실제 이벤트는 보존되고 인공 스파이크만 제거된다
PUSH_SPIKE_RATIO = 4
# 흐름(flow) 지표만 마스킹 대상 — 잔고(앱푸시_동의자수)는 수준값이라 별도 글리치 규칙 사용
PUSH_FLOW_METRICS = {"앱푸시수신동의", "앱푸시_신규추가", "앱푸시_이탈"}

def mask_push_spikes(d: pd.DataFrame) -> pd.DataFrame:
    """앱푸시 흐름 지표의 대량 이관/재동의/재분류 스파이크 마스킹 (NaN 처리).
    (지표,세그먼트,연도)별 일 |중앙값|의 PUSH_SPIKE_RATIO배 초과 값을 이상치로 본다.
    (실데이터: 2025-04-19 재동의 16,496건=중앙값 76배, 2026-04-24 15,441건=84배,
     2025-01-01 세그먼트 재분류 6만건대. 연도마다 날짜가 달라 하드코딩 대신 통계 규칙 사용)"""
    m = d["metric"].isin(PUSH_FLOW_METRICS)
    if not m.any():
        return d
    med = (d.loc[m].groupby(["metric", "segment", "year"])["value"]
           .transform(lambda s: s.abs().median()))
    spike = d.loc[m, "value"].abs() > med * PUSH_SPIKE_RATIO
    d.loc[spike[spike].index, "value"] = np.nan
    return d

def mask_level_glitches(d: pd.DataFrame, rel=0.3) -> pd.DataFrame:
    """잔고(수준) 지표의 하루짜리 급변 글리치 마스킹.
    양옆 이웃과 rel(30%) 이상 괴리하면서 이웃끼리는 서로 비슷(rel/2 이내)하면 그 날만 NaN.
    지속되는 계단(연초 신규→기존 재분류, 4/19 재이관 등)은 한쪽 이웃과 일치하므로 보존.
    (실데이터: 2025-04-18 이관일에 잔고가 0으로 기록된 글리치 — 3개 세그먼트 공통)
    d에는 'dt'(datetime) 컬럼이 있어야 한다."""
    m = d["metric"] == "앱푸시_동의자수"
    if not m.any():
        return d
    for _seg, g in d[m].groupby("segment"):
        g = g.sort_values("dt")
        v = g["value"].to_numpy(float)
        if len(v) < 3:
            continue
        prev, nxt = np.roll(v, 1), np.roll(v, -1)
        dev_p = np.abs(v - prev) / np.maximum(np.abs(prev), 1)
        dev_n = np.abs(v - nxt) / np.maximum(np.abs(nxt), 1)
        agree = np.abs(nxt - prev) / np.maximum(np.abs(prev), 1) < rel / 2
        bad = (dev_p > rel) & (dev_n > rel) & agree
        bad[0] = bad[-1] = False
        d.loc[g.index[bad], "value"] = np.nan
    return d

def parse_push_file(name, data: bytes) -> pd.DataFrame:
    """PUSH 원천: 기존/신규/Total 3섹션 × (수신동의 누적·신규추가·이탈) 전부 파싱.
    순증감은 저장하지 않고 화면에서 신규추가-이탈로 파생 (마스킹 일관성 유지)."""
    try:
        rows = read_grid(name, data)
    except Exception:
        return pd.DataFrame()
    if not rows or len(rows) < 8: return pd.DataFrame()
    if not any("Date" in str(c) for c in rows[0]): return pd.DataFrame()

    date_row = rows[1]

    # 1) 유효한 날짜 컬럼을 모두 추출하고 월(month)을 파싱
    date_cols = []
    for ci in range(2, len(date_row)):
        d = _cell(date_row[ci])
        m = re.match(r"^(\d{1,2})/\d{1,2}$", d)
        if m:
            date_cols.append((ci, d, int(m.group(1))))

    if not date_cols:
        return pd.DataFrame()

    # 2) 뒤에서부터 역순으로 읽으면서 해(year)가 바뀌는 지점 계산
    # 맨 마지막 데이터를 0으로 두고, (1월 <- 12월)로 넘어갈 때마다 연도를 -1
    rel_years = {}
    current_rel_year = 0
    last_month = date_cols[-1][2]

    for ci, d, month in reversed(date_cols):
        # 역순으로 읽을 때 월이 1에서 12로 커지면 (실제로는 12월 -> 1월) 해가 바뀐 것
        if last_month < 6 and month > 6:
            current_rel_year -= 1
        elif last_month > 6 and month < 6:
            current_rel_year += 1

        rel_years[ci] = current_rel_year
        last_month = month

    # 3) 섹션(A열) × 행유형(B열) → (지표, 세그먼트) 매핑으로 대상 행 수집
    SECTION_SEG = {"기존": "기존", "신규": "신규", "Total": "*TOTAL"}
    ROW_METRIC = {"수신동의": "앱푸시_동의자수", "신규추가(+)": "앱푸시_신규추가",
                  "기존이탈(-)": "앱푸시_이탈"}
    targets = []  # (metric, segment, row)
    cur_seg = None
    for row in rows:
        c0 = _cell(row[0] if len(row) > 0 else "")
        c1 = _cell(row[1] if len(row) > 1 else "")
        if c0 in SECTION_SEG: cur_seg = SECTION_SEG[c0]
        if cur_seg and c1 in ROW_METRIC:
            targets.append((ROW_METRIC[c1], cur_seg, row))
            # 신규 섹션 신규추가는 기존 지표명으로도 저장 (구버전 누적 데이터·동의율 로직과 연속성)
            if cur_seg == "신규" and c1 == "신규추가(+)":
                targets.append(("앱푸시수신동의", "*TOTAL", row))

    if not targets: return pd.DataFrame()

    records = []
    for metric, seg, trow in targets:
        for ci, d, month in date_cols:
            val = _num(trow[ci] if ci < len(trow) else None)
            if pd.isna(val): continue
            records.append({
                "gran": "일", "metric": metric, "segment": seg,
                "year": rel_years[ci], "label": d, "sortkey": 0, "close": "final", "value": val
            })
    return pd.DataFrame(records)

def looks_like_push_name(name: str) -> bool:
    """파일명이 앱푸시 원천으로 보이는가 (한글명 포함)"""
    up = name.upper()
    return "PUSH" in up or any(h in name for h in ("앱푸시", "푸시", "수신동의"))

def route_push(n, b):
    """엑셀 1건 → (push_df 또는 None, 일반_df 또는 None).
    이름 힌트가 있으면 PUSH 우선 시도, 없으면 일반 파싱 후 빈 결과면 내용 기반으로 PUSH 재시도.
    → 파일명이 'PUSH'가 아니어도(예: 앱푸시수신동의현황.xlsx) 인식된다."""
    is_xlsx = n.lower().endswith((".xlsx", ".xls"))
    if is_xlsx and looks_like_push_name(n):
        pf = parse_push_file(n, b)
        if not pf.empty: return pf, None
    d = parse_file(n, b)
    if not d.empty: return None, d
    if is_xlsx:  # 이름 힌트 없이도 헤더가 PUSH 형식이면 인식 (내용 기반 폴백)
        pf = parse_push_file(n, b)
        if not pf.empty: return pf, None
    return None, None

@st.cache_data(show_spinner=False)
def combine_files(file_tuples) -> pd.DataFrame:
    """업로드 파일들 → 통합 long DF. 동일 키는 마지막 파일 우선"""
    frames = []
    push_frames = []
    for n, b in file_tuples:
        pf, df = route_push(n, b)
        if pf is not None: push_frames.append(pf)
        elif df is not None: frames.append(df)

    inferred_years = set()
    for f in frames:
        if "year" in f.columns:
            inferred_years.update(f["year"].dropna().unique())
            
    default_year = int(max(inferred_years)) if inferred_years else today_kst().year

    # PUSH 파일엔 연도가 없어 상대연도(rel)만 있다 — 함께 올린 일자 데이터와 (연도,일자)
    # 겹침이 가장 큰 기준연도를 고른다 (예: 2025 가입자수만 곁들여 올려도 2024/2025 오배정 방지)
    if push_frames and frames:
        daily_keys = set()
        for f in frames:
            d = f[f["gran"] == "일"][["year", "label"]].dropna()
            daily_keys.update(map(tuple, d.values.tolist()))
        if daily_keys:
            def _overlap(base):
                return sum(1 for pf in push_frames
                           for ry, lb in pf[["year", "label"]].values.tolist()
                           if (base + ry, lb) in daily_keys)
            cands = sorted({default_year, default_year + 1, today_kst().year})
            best = max(cands, key=_overlap)
            if _overlap(best) > _overlap(default_year):
                default_year = int(best)

    for pf in push_frames:
        pf["year"] = default_year + pf["year"]
        def calc_sortkey(row):
            m = re.match(r"(\d+)/(\d+)", str(row["label"]))
            if m: return int(row["year"]) * 10000 + int(m.group(1))*100 + int(m.group(2))
            return 0
        pf["sortkey"] = pf.apply(calc_sortkey, axis=1)
        frames.append(pf)

    if not frames: return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    df = df.drop_duplicates(subset=KEY_COLS, keep="last")
    return df

# ══════════════════════════════════════════════════════
# 데이터 누적 저장소 — 업로드할 때마다 병합·저장
# ══════════════════════════════════════════════════════
DATA_STORE = "wr_data_store.csv"
KEY_COLS = ["gran", "metric", "segment", "year", "label", "close"]
STORE_COLS = KEY_COLS + ["sortkey", "value"]

def load_store() -> pd.DataFrame:
    if os.path.exists(DATA_STORE):
        try:
            d = pd.read_csv(DATA_STORE, encoding="utf-8-sig")
            if set(STORE_COLS) <= set(d.columns):
                # 과거 버전이 요일 헤더 행을 '채널' 세그먼트(전부 NaN)로 오파싱해 저장한 쓰레기 제거
                d = d[~((d["segment"] == "채널") & d["value"].isna())]
                return d[STORE_COLS]
        except Exception:
            pass
    return pd.DataFrame()

def save_store(df: pd.DataFrame):
    df[STORE_COLS].to_csv(DATA_STORE, index=False, encoding="utf-8-sig")

def merge_store(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """기존 누적 + 신규 업로드 병합 — 같은 (단위·지표·채널·기간) 키는 신규 우선"""
    if old is None or old.empty: return new
    if new is None or new.empty: return old
    return (pd.concat([old[STORE_COLS], new[STORE_COLS]], ignore_index=True)
            .drop_duplicates(subset=KEY_COLS, keep="last"))

def _period_set(d, cols):
    if d is None or d.empty: return set()
    return set(map(tuple, d[cols].drop_duplicates().values.tolist()))

def upload_diff(stored, df_new):
    """업로드 데이터가 기존 누적 대비 추가/갱신하는 기간 수 (저장 전 미리보기용)"""
    cols = ["gran", "year", "label"]
    nc, oc = _period_set(df_new, cols), _period_set(stored, cols)
    return len(nc - oc), len(nc & oc)

# ── 백업/업로드 고도화 헬퍼 ────────────────────────────────
@st.cache_data(show_spinner=False)
def classify_uploads(file_tuples):
    """업로드된 각 파일이 무엇으로 인식됐는지 (파일명, 인식결과, 행수) 목록.
    미인식 파일을 눈에 보이게 해 조용한 누락을 방지한다."""
    out = []
    for n, b in file_tuples:
        is_backup = False
        if n.lower().endswith(".csv"):
            head = b[:400].decode("utf-8", "ignore").splitlines()
            if head and all(k in head[0] for k in ("gran", "metric", "sortkey")):
                is_backup = True
        pf, d = route_push(n, b)  # combine_files와 동일한 라우팅 (한글명 PUSH 포함)
        if pf is not None:
            nseries = pf.groupby(["metric", "segment"]).ngroups
            out.append((n, f"✅ 앱푸시 원천 · {nseries}시리즈", len(pf)))
        elif d is None:
            hint = ("❌ PUSH 인식 실패 (헤더 확인)" if looks_like_push_name(n)
                    else "❌ 미인식 (파일명·형식 확인)")
            out.append((n, hint, 0))
        elif is_backup:
            out.append((n, "♻ 누적 백업 복원", len(d)))
        else:
            grans = "".join(g for g in ("일", "주", "월") if g in set(d["gran"]))
            nm = d["metric"].nunique()
            kind = (f"{grans} · {nm}지표" if nm > 1
                    else f"{grans} · {d['metric'].iloc[0]}")
            out.append((n, f"✅ {kind}", len(d)))
    return out

def make_backup_zip(df, texts) -> bytes:
    """누적 데이터 CSV + 보고란·메모 JSON + manifest를 한 ZIP으로 묶는다 (통합 백업)."""
    import zipfile
    buf = io.BytesIO()
    ts = datetime.datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")
    yrs = (f"{int(df['year'].min())}–{int(df['year'].max())}년"
           if not df.empty else "-")
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("wr_data_store.csv",
                    df[STORE_COLS].to_csv(index=False).encode("utf-8-sig"))
        zf.writestr("wr_insights.json",
                    json.dumps(texts, ensure_ascii=False, indent=2).encode("utf-8"))
        zf.writestr("manifest.txt",
                    (f"LF Mall 주간보고 통합 백업\n생성: {ts}\n"
                     f"누적 데이터: {len(df):,}행 · {df['metric'].nunique() if not df.empty else 0}지표 · {yrs}\n"
                     f"보고란·메모: {len(texts)}개 항목\n"
                     "복원: 사이드바 '백업 복원'에 이 ZIP을 그대로 올리세요.\n").encode("utf-8"))
    return buf.getvalue()

def _apply_memos(memos: dict):
    """메모 dict를 세션·파일에 병합 (업로드 값 우선)"""
    merged = {**st.session_state.wr_texts, **memos}
    st.session_state.wr_texts = merged
    all_d = load_insights(); all_d.update(merged); save_insights(all_d)

def restore_from_upload(upload):
    """백업 파일(zip/csv/json) 자동 인식 복원 → (요약 메시지, 데이터스토어_갱신여부).
    형식 불일치는 ValueError로 던진다 (호출측에서 표시)."""
    name = upload.name.lower()
    raw = upload.getvalue()
    data_restored, msgs = False, []

    def _try_csv(b):
        nonlocal data_restored
        d = pd.read_csv(io.BytesIO(b), encoding="utf-8-sig")
        if set(STORE_COLS) <= set(d.columns):
            save_store(d[STORE_COLS]); data_restored = True
            msgs.append(f"누적 데이터 {len(d):,}행")
            return True
        return False

    def _try_json(b):
        memos = json.loads(b.decode("utf-8"))
        if isinstance(memos, dict):
            _apply_memos(memos); msgs.append(f"메모 {len(memos)}개")
            return True
        return False

    if name.endswith(".zip"):
        import zipfile
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            for info in zf.infolist():
                if info.is_dir(): continue
                base = os.path.basename(info.filename).lower()
                if "__macosx" in info.filename.lower(): continue
                b = zf.read(info)
                if base.endswith(".csv"): _try_csv(b)
                elif base.endswith(".json"): _try_json(b)
    elif name.endswith(".csv"):
        if not _try_csv(raw):
            raise ValueError("누적 데이터 CSV 형식이 아닙니다 (gran·metric·sortkey 등 필수 컬럼 누락)")
    elif name.endswith(".json"):
        if not _try_json(raw):
            raise ValueError("메모 JSON 형식이 아닙니다 (최상위가 객체여야 함)")
    else:
        raise ValueError("지원 형식: .zip / .csv / .json")

    if not msgs:
        raise ValueError("복원할 내용을 찾지 못했습니다 (ZIP 안에 백업 파일 없음)")
    return " · ".join(msgs), data_restored

def restore_widget(key, label="백업 복원 (ZIP / CSV / JSON)"):
    """백업 복원 업로더 + 처리. 빈 상태(메인)·사이드바 공용. 성공 시 rerun."""
    up = st.file_uploader(label, type=["zip", "csv", "json"], key=key)
    if up is not None and st.session_state.get("wr_restored") != up.name:
        try:
            summary, data_changed = restore_from_upload(up)
            st.session_state["wr_restored"] = up.name
            if data_changed:
                st.session_state.pop("wr_saved_sig", None)
                st.cache_data.clear()
            st.success(f"복원 완료 ✓ — {summary}"); st.rerun()
        except Exception as e:
            st.error(f"복원 실패: {e}")

def source_upload_widget(key):
    """빈 상태 메인용 원천 업로더 — 인식되면 즉시 저장 후 rerun (첫 업로드는 누적할 게 없음)."""
    mfiles = st.file_uploader("엑셀 / CSV / ZIP (복수 선택)",
                              type=["xlsx", "xls", "csv", "zip"],
                              accept_multiple_files=True, key=key)
    if mfiles:
        exp = expand_uploads(mfiles)
        dnew = combine_files(tuple(exp)) if exp else pd.DataFrame()
        if not dnew.empty:
            save_store(dnew)
            st.session_state.pop("wr_saved_sig", None)
            st.cache_data.clear()
            st.success(f"{len(dnew):,}행 인식 — 저장됨 ✓"); st.rerun()
        else:
            st.error("인식된 데이터가 없습니다. 파일명·형식을 확인하세요.")


# ══════════════════════════════════════════════════════
# 조회 헬퍼 — mtd(당월/당주 일마감) vs final 선택
# ══════════════════════════════════════════════════════
def pick(df, gran, metric, seg, year, label, prefer="final"):
    sub = df[(df["gran"] == gran) & (df["metric"] == metric) &
             (df["segment"] == seg) & (df["year"] == year) & (df["label"] == label)]
    if sub.empty: return np.nan
    order = ["final", "mtd"] if prefer == "final" else ["mtd", "final"]
    for c in order:
        s = sub[sub["close"] == c]["value"].dropna()
        if len(s): return s.iloc[-1]
    return np.nan

def series_by_label(df, gran, metric, seg, year, prefer="final"):
    """한 연도의 기간라벨 → 값 Series (sortkey 순)"""
    sub = df[(df["gran"] == gran) & (df["metric"] == metric) &
             (df["segment"] == seg) & (df["year"] == year)].copy()
    if sub.empty: return pd.Series(dtype=float)
    pref = {"final": 0, "mtd": 1} if prefer == "final" else {"mtd": 0, "final": 1}
    sub["_p"] = sub["close"].map(pref)
    sub = sub.sort_values(["sortkey", "_p"]).drop_duplicates("label", keep="first")
    return sub.set_index("label")["value"]

def labels_sorted(df, gran, years=None):
    sub = df[df["gran"] == gran]
    if years is not None: sub = sub[sub["year"].isin(years)]
    return (sub[["label", "sortkey"]].assign(k=lambda d: d["sortkey"] % 10000)
            .drop_duplicates("label").sort_values("k")["label"].tolist())

def latest_period(df, gran):
    """가장 최근 (year, label)"""
    sub = df[(df["gran"] == gran) & df["value"].notna()]
    if sub.empty: return None, None
    row = sub.loc[sub["sortkey"].idxmax()]
    return int(row["year"]), row["label"]

def prev_label(df, gran, year, label):
    """직전 기간 (연도 경계 포함)"""
    sub = (df[(df["gran"] == gran) & df["value"].notna()]
           [["year", "label", "sortkey"]].drop_duplicates().sort_values("sortkey"))
    keys = sub[["year", "label"]].apply(tuple, axis=1).tolist()
    try: i = keys.index((year, label))
    except ValueError: return None, None
    return keys[i - 1] if i > 0 else (None, None)

def week_ref(df, ref_year, ref_week):
    """기준 주차 (year, label): 사이드바 선택 우선, 없으면 데이터상 최신 주차"""
    return (ref_year, ref_week) if ref_week else latest_period(df, "주")

# ══════════════════════════════════════════════════════
# 보고란 영속화
# ══════════════════════════════════════════════════════
def load_insights():
    if os.path.exists(INSIGHT_FILE):
        with open(INSIGHT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_insights(d):
    with open(INSIGHT_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def report_text_block(key, title, default="", regen=None, ai_fn=None):
    """편집 가능한 보고란 텍스트 박스 (JSON 저장).
    regen 텍스트를 주면 '자동 생성'(템플릿) 버튼, ai_fn을 주면 'AI 생성'(Claude) 버튼이 뜬다.
    좁은 컬럼 안에서도 안 깨지도록 버튼은 박스 위/아래에 배치한다."""
    store = st.session_state.wr_texts
    if not store.get(key): store[key] = default
    ekey = f"__wr_edit_{key}__"
    if ekey not in st.session_state: st.session_state[ekey] = False

    # 제목 + 액션 버튼 (제목 한 줄, 버튼은 그 아래 정상 너비 컬럼)
    st.markdown(f"**{title}**")

    if st.session_state[ekey]:
        if HAS_QUILL:
            # Word 수준 리치 에디터 (글자 크기·색·굵게·기울임·밑줄·목록·정렬 등)
            toolbar = [
                [{"size": ["small", False, "large", "huge"]}],
                ["bold", "italic", "underline", "strike"],
                [{"color": []}, {"background": []}],
                [{"list": "ordered"}, {"list": "bullet"}],
                [{"align": []}], ["clean"],
            ]
            # value(초기값)는 최초 진입 또는 store[key]가 바뀐 경우에만 주입한다.
            # 매 rerun마다 value를 다시 넣으면 타이핑 중 저장값으로 되돌아가
            # '계속 리프레시'되는 버그가 발생한다. (st_quill은 입력마다 rerun 유발)
            qkey = f"wr_quill_{key}"
            skey = f"{qkey}__seed"
            if st.session_state.get(skey) != store[key]:
                st.session_state[skey] = store[key]
                new = st_quill(value=store[key], html=True, toolbar=toolbar, key=qkey)
            else:
                new = st_quill(html=True, toolbar=toolbar, key=qkey)
        else:
            new = st.text_area("", store[key], key=f"wr_ta_{key}", height=180,
                               label_visibility="collapsed")
        if st.button("저장", key=f"wr_save_{key}", type="primary",
                     width="stretch"):
            store[key] = new if new is not None else store[key]
            all_d = load_insights(); all_d[key] = store[key]; save_insights(all_d)
            st.session_state[ekey] = False; st.rerun()
    else:
        st.markdown(f"<div class='report-box'>{store[key] or '내용을 입력하세요.'}</div>",
                    unsafe_allow_html=True)
                    
    # AI 참고 메모: 데이터에 안 나오는 배경(프로모션·이벤트·이슈)을 적으면
    # AI 생성 시 [배경 메모]로 분리 주입돼 원인·맥락 해석에 활용된다.
    memo_val = ""
    if ai_fn is not None:
        mkey = f"{key}__memo"
        if mkey not in store: store[mkey] = ""
        with st.expander("🧠 AI 참고 메모 (프로모션·이벤트·운영 이슈 등 배경)",
                         expanded=bool(store[mkey])):
            memo_val = st.text_area(
                "데이터에 안 나오는 배경을 적으면 AI가 원인·맥락 해석에 활용합니다. "
                "(수치는 데이터에서만 인용)",
                store[mkey], key=f"wr_memo_{key}", height=120)
            if st.button("메모 저장", key=f"wr_memosave_{key}",
                         width="stretch"):
                store[mkey] = memo_val
                all_d = load_insights(); all_d[mkey] = memo_val; save_insights(all_d)
                st.rerun()
                    
    n = 2 + (1 if regen is not None else 0) + (1 if ai_fn is not None else 0)
    bcols = st.columns(n)
    bi = 0
    edit_on = st.session_state[ekey]
    if bcols[bi].button("편집" if not edit_on else "보기",
                        key=f"wr_edit_{key}", width="stretch"):
        st.session_state[ekey] = not edit_on; st.rerun()
    bi += 1
    if regen is not None:
        if bcols[bi].button("자동 생성", key=f"wr_regen_{key}", width="stretch",
                            help="기준 주차 실적으로 템플릿 문구를 채웁니다 (기존 내용 대체)"):
            store[key] = regen
            all_d = load_insights(); all_d[key] = regen; save_insights(all_d)
            st.session_state[ekey] = False; st.rerun()
        bi += 1
    if ai_fn is not None:
        if bcols[bi].button("AI 생성", key=f"wr_ai_{key}", width="stretch",
                            help="Claude가 데이터(+참고 메모)를 보고 인사이트 문구를 작성합니다 (기존 내용 대체)"):
            store[mkey] = memo_val
            all_d = load_insights(); all_d[mkey] = memo_val; save_insights(all_d)
            with st.spinner("AI가 인사이트를 작성 중…"):
                text, err = ai_fn(memo_val)
            if err:
                st.error(err)
            else:
                store[key] = text
                all_d = load_insights(); all_d[key] = text; save_insights(all_d)
                st.session_state[ekey] = False; st.rerun()
    return store[key]

# ══════════════════════════════════════════════════════
# YoY 요약표 / 추이표 빌더
# ══════════════════════════════════════════════════════
def month_label(n): return f"{n}월"

def week_disp(year, label):
    """주차 표시: 2026년 06월 2주차"""
    return f"{year}년 {label}" if label else "-"

def wow_summary_table(df, wy, wlabel, metrics):
    """전주비 요약표 (실적 요약과 동일 구조): 전주·기준주·전주비 + 전년동주·전년비"""
    py, plb = prev_label(df, "주", wy, wlabel)
    cols = [week_disp(py, plb), week_disp(wy, wlabel), "전주비",
            week_disp(wy - 1, wlabel), "전년비"]
    rows = []
    for met in metrics:
        cur = pick(df, "주", met, "*TOTAL", wy, wlabel, "mtd")
        prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
        yoy = pick(df, "주", met, "*TOTAL", wy - 1, wlabel, "final")
        rows.append({
            "구분": met,
            cols[0]: fmt_value(met, prv), cols[1]: fmt_value(met, cur),
            cols[2]: fmt_delta(met, cur, prv) or "–",
            cols[3]: fmt_value(met, yoy), cols[4]: fmt_delta(met, cur, yoy) or "–",
        })
    return pd.DataFrame(rows).set_index("구분")

def yoy_summary_table(df, ref_year, ref_month, metrics):
    """참고본 '실적 요약' 표: 전월·당월 × (전년, 당년, 전년비)"""
    rows = []
    pm_y, pm_m = (ref_year, ref_month - 1) if ref_month > 1 else (ref_year - 1, 12)
    cols = [f"{pm_y-1}년 {pm_m}월", f"{pm_y}년 {pm_m}월", "전년비(전월)",
            f"{ref_year-1}년 {ref_month}월", f"{ref_year}년 {ref_month}월", "전년비(당월)"]
    for met in metrics:
        pm_prev = pick(df, "월", met, "*TOTAL", pm_y - 1, month_label(pm_m), "final")
        pm_cur  = pick(df, "월", met, "*TOTAL", pm_y,     month_label(pm_m), "final")
        cm_prev = pick(df, "월", met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
        cm_cur  = pick(df, "월", met, "*TOTAL", ref_year,     month_label(ref_month), "mtd")
        rows.append({
            "구분": met,
            cols[0]: fmt_value(met, pm_prev), cols[1]: fmt_value(met, pm_cur),
            cols[2]: fmt_delta(met, pm_cur, pm_prev) or "–",
            cols[3]: fmt_value(met, cm_prev), cols[4]: fmt_value(met, cm_cur),
            cols[5]: fmt_delta(met, cm_cur, cm_prev) or "–",
        })
    return pd.DataFrame(rows).set_index("구분"), (pm_y, pm_m)

def trend_table(df, gran, metrics, years, seg="*TOTAL"):
    """추이표: 행=지표, 열=(연도, 기간)"""
    out, columns = {}, []
    for y in years:
        for lb in labels_sorted(df, gran, [y]):
            sub = df[(df["gran"] == gran) & (df["year"] == y) & (df["label"] == lb) &
                     df["value"].notna()]
            if sub.empty: continue
            columns.append((y, lb))
    if not columns:
        return pd.DataFrame()
    for met in metrics:
        vals = []
        for y, lb in columns:
            vals.append(pick(df, gran, met, seg, y, lb, "final"))
        out[met] = vals
    tbl = pd.DataFrame(out, index=pd.MultiIndex.from_tuples(columns, names=["연도", "기간"])).T
    return tbl

def style_trend(tbl, metrics):
    # 최신 pandas는 float 컬럼에 문자열 대입을 금지하므로 object로 변환 후 포맷
    disp = tbl.astype(object).copy()
    for met in disp.index:
        disp.loc[met] = [fmt_value(met, v) for v in tbl.loc[met]]
    return disp

# ══════════════════════════════════════════════════════
# YoY 라인차트 (Plotly)
# ══════════════════════════════════════════════════════
def base_layout(h=300, ysuffix="", title=""):
    return dict(
        paper_bgcolor="rgba(248,249,252,0)", plot_bgcolor="rgba(248,249,252,0)",
        font=dict(color="#475569", size=11), margin=dict(l=10, r=10, t=40, b=10),
        height=h, showlegend=True,
        legend=dict(orientation="h", y=1.12, bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#64748b", size=10)),
        title=dict(text=title, font=dict(color="#94a3b8", size=13)),
        xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=10)),
        yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=10), ticksuffix=ysuffix),
    )

def yoy_chart(df, gran, metric, years, seg="*TOTAL", h=300):
    unit, div = METRIC_UNIT.get(metric, ("", 1))
    if metric in PCT_METRICS: div, unit = 0.01, "%"
    if gran == "월":
        x_all = [month_label(i) for i in range(1, 13)]
    else:
        x_all = labels_sorted(df, gran, years)
    fig = go.Figure()
    for i, y in enumerate(sorted(years)):
        # 연도마다 없는 주차(5주차 등)는 건너뛰고 선을 잇는다
        s = series_by_label(df, gran, metric, seg, y, prefer="final").reindex(x_all).dropna()
        fig.add_trace(go.Scatter(
            x=s.index.tolist(), y=(s / div).tolist(), mode="lines+markers", name=str(y),
            line=dict(color=clr(YEAR_PAL[i % len(YEAR_PAL)]), width=2),
            marker=dict(size=5),
        ))
    gname = "월별" if gran == "월" else "주차별"
    ly = base_layout(h, ysuffix=unit if unit == "%" else "",
                     title=f"{metric} {gname} 추이 ({unit})")
    ly["xaxis"]["categoryorder"] = "array"
    ly["xaxis"]["categoryarray"] = x_all
    if gran == "주": ly["xaxis"]["tickangle"] = -45; ly["xaxis"]["nticks"] = 20
    fig.update_layout(**ly)
    return fig

# ══════════════════════════════════════════════════════
# 엑셀 워크북 내보내기 (월/주/첫구매_요약 + 차트)
# ══════════════════════════════════════════════════════
def build_workbook(df, texts, ref_year, ref_month, chart_years):
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_font = Font(bold=True, size=11)
    head_fill = PatternFill("solid", fgColor="EEF3FA")
    title_font = Font(bold=True, size=13)

    # ── 데이터 시트 (월/주): 지표 블록 스택
    extra_metrics = [m for m in df["metric"].unique() if m not in METRICS7]
    for gran, sheet in (("월", "월"), ("주", "주")):
        ws = wb.create_sheet(sheet)
        sub_g = df[df["gran"] == gran]
        if sub_g.empty: continue
        cols = (sub_g[["year", "label", "sortkey"]].drop_duplicates()
                .sort_values("sortkey")[["year", "label"]].apply(tuple, axis=1).tolist())
        r = 1
        for met in METRICS7 + sorted(extra_metrics):
            sub_m = sub_g[sub_g["metric"] == met]
            if sub_m.empty: continue
            ws.cell(r, 1, met).font = title_font
            for j, (y, lb) in enumerate(cols):
                ws.cell(r, 2 + j, y).font = head_font
                ws.cell(r + 1, 2 + j, lb).font = head_font
                ws.cell(r + 1, 2 + j).fill = head_fill
            segs = ["*TOTAL"] + [s for s in CHANNELS if s in set(sub_m["segment"])]
            for i, seg in enumerate(segs):
                ws.cell(r + 2 + i, 1, seg).font = head_font
                for j, (y, lb) in enumerate(cols):
                    v = pick(df, gran, met, seg, y, lb, "final")
                    c = ws.cell(r + 2 + i, 2 + j)
                    if not (isinstance(v, float) and np.isnan(v)):
                        c.value = v
                        c.number_format = "0.00%" if met in PCT_METRICS else "#,##0"
            r += 2 + len(segs) + 2

    # ── 요약 시트
    ws = wb.create_sheet("첫구매_요약", 0)
    wb.remove(wb["Sheet"])
    r = 1
    ws.cell(r, 1, f"첫구매 주간보고 — {ref_year}년 {ref_month}월 기준").font = Font(bold=True, size=15)
    r += 2

    # 실적 요약 YoY 표
    ws.cell(r, 1, "실적 요약 (일평균)").font = title_font; r += 1
    tbl, _ = yoy_summary_table(df, ref_year, ref_month, METRICS7)
    ws.cell(r, 1, "구분").font = head_font
    for j, cname in enumerate(tbl.columns):
        c = ws.cell(r, 2 + j, cname); c.font = head_font; c.fill = head_fill
    red_font = Font(color="DC2626")
    green_font = Font(color="16A34A")
    for i, met in enumerate(tbl.index):
        ws.cell(r + 1 + i, 1, met).font = head_font
        for j, cname in enumerate(tbl.columns):
            c = ws.cell(r + 1 + i, 2 + j, tbl.loc[met, cname])
            s = str(c.value)
            if "전년비" in str(cname):
                if s.startswith("△"): c.font = red_font
                elif s.startswith("+"): c.font = green_font
    r += len(tbl) + 3

    # 보고란
    ws.cell(r, 1, "전주 주요 지표 현황").font = title_font
    ws.cell(r, 8, "금주 집행 내용 요약").font = title_font
    c1 = ws.cell(r + 1, 1, texts.get("wr_metrics_summary", ""))
    c2 = ws.cell(r + 1, 8, texts.get("wr_exec_summary", ""))
    for c in (c1, c2): c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 8, end_column=6)
    ws.merge_cells(start_row=r + 1, start_column=8, end_row=r + 8, end_column=13)
    r += 11

    # 차트 데이터 블록 + 라인차트 (월별/주차별 × 거래액·고객수·객단가)
    chart_metrics = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]
    for gran, gname in (("월", "월별"), ("주", "주차별")):
        x_all = ([month_label(i) for i in range(1, 13)] if gran == "월"
                 else labels_sorted(df, gran, chart_years))
        if not x_all: continue
        anchor_row = r
        for k, met in enumerate(chart_metrics):
            ws.cell(r, 1, f"{met} {gname} (차트 데이터)").font = head_font
            for j, lb in enumerate(x_all):
                ws.cell(r, 2 + j, lb).fill = head_fill
            nrow = 0
            for y in sorted(chart_years):
                s = series_by_label(df, gran, met, "*TOTAL", y).reindex(x_all)
                if s.dropna().empty: continue
                nrow += 1
                ws.cell(r + nrow, 1, y)
                for j, v in enumerate(s.tolist()):
                    if not (isinstance(v, float) and np.isnan(v)):
                        c = ws.cell(r + nrow, 2 + j, v)
                        c.number_format = "#,##0"
            if nrow:
                ch = LineChart()
                ch.title = f"{met} {gname} 추이"
                ch.height, ch.width = 7.5, 13
                data = Reference(ws, min_col=1, min_row=r + 1,
                                 max_col=1 + len(x_all), max_row=r + nrow)
                cats = Reference(ws, min_col=2, min_row=r, max_col=1 + len(x_all))
                ch.add_data(data, titles_from_data=True, from_rows=True)
                ch.set_categories(cats)
                ws.add_chart(ch, f"{get_column_letter(2 + len(x_all) + 1 + (k % 3) * 8)}{anchor_row}")
            r += nrow + 2
        r += 14

    # 채널별 실적 (당월 YoY)
    ws.cell(r, 1, f"채널별 실적 — {ref_year}년 {ref_month}월 (전년비)").font = title_font; r += 1
    for met in chart_metrics:
        ws.cell(r, 1, met).font = head_font
        heads = [f"{ref_year-1}년 {ref_month}월", f"{ref_year}년 {ref_month}월", "전년비"]
        for j, hd in enumerate(heads):
            c = ws.cell(r, 2 + j, hd); c.font = head_font; c.fill = head_fill
        segs = ["*TOTAL"] + CHANNELS
        for i, seg in enumerate(segs):
            pv = pick(df, "월", met, seg, ref_year - 1, month_label(ref_month), "mtd")
            cv = pick(df, "월", met, seg, ref_year, month_label(ref_month), "mtd")
            ws.cell(r + 1 + i, 1, seg)
            ws.cell(r + 1 + i, 2, None if np.isnan(pv) else pv).number_format = "#,##0"
            ws.cell(r + 1 + i, 3, None if np.isnan(cv) else cv).number_format = "#,##0"
            dcell = ws.cell(r + 1 + i, 4, fmt_delta(met, cv, pv) or "–")
            ds = str(dcell.value)
            if ds.startswith("△"): dcell.font = red_font
            elif ds.startswith("+"): dcell.font = green_font
        r += len(segs) + 3

    ws.column_dimensions["A"].width = 16
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════
# 자동 보고 초안
# ══════════════════════════════════════════════════════
def _delta_html(d):
    """증감 문자열 → 색상 span (역신장 △ 빨강 / 신장 + 초록)"""
    if not d: return ""
    if d.startswith("△"): return f'<span style="color:#dc2626;font-weight:700">{d}</span>'
    return f'<span style="color:#16a34a;font-weight:700">{d}</span>'

def auto_draft(df, ref_year, ref_month, ref_week=None):
    """기준 주차(없으면 기준 월) 실적으로 보고 문구 자동 생성 (HTML, 역신장 빨강)"""
    lines = []
    for met in METRICS7:
        if ref_week:
            cur = pick(df, "주", met, "*TOTAL", ref_year, ref_week, "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            py, plb = prev_label(df, "주", ref_year, ref_week)
            prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
            yoy = pick(df, "주", met, "*TOTAL", ref_year - 1, ref_week, "final")
            parts = [f" - {met} — {fmt_value(met, cur)}"]
            d_w = fmt_delta(met, cur, prv)
            d_y = fmt_delta(met, cur, yoy)
            if d_w: parts.append(f"전주비 {_delta_html(d_w)}")
            if d_y: parts.append(f"전년비 {_delta_html(d_y)}")
            lines.append(", ".join(parts))
        else:
            cur = pick(df, "월", met, "*TOTAL", ref_year, month_label(ref_month), "mtd")
            prv = pick(df, "월", met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            d = fmt_delta(met, cur, prv)
            tail = f", 전년비 {_delta_html(d)}" if d else ""
            lines.append(f" - {met} — {fmt_value(met, cur)}" + tail)
    return "<br>".join(lines)

# ══════════════════════════════════════════════════════
# AI 인사이트 생성 (Claude API) — 모델 교체 가능
# ══════════════════════════════════════════════════════
AI_MODELS = {
    "Claude Sonnet 5 (균형·기본)": "claude-sonnet-5",
    "Claude Opus 4.8 (최고 품질)": "claude-opus-4-8",
    "Claude Haiku 4.5 (빠름·저렴)": "claude-haiku-4-5",
}
DEFAULT_AI_MODEL = "claude-sonnet-5"

def _anthropic_key():
    """Streamlit secrets 또는 환경변수에서 API 키 조회"""
    try:
        if "ANTHROPIC_API_KEY" in st.secrets:
            return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY")

def _ai_metric_facts(df, ref_year, ref_month, ref_week=None):
    """모델에 넘길 지표·증감 요약 텍스트"""
    rows = []
    for met in METRICS7:
        if ref_week:
            cur = pick(df, "주", met, "*TOTAL", ref_year, ref_week, "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            py, plb = prev_label(df, "주", ref_year, ref_week)
            prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
            yoy = pick(df, "주", met, "*TOTAL", ref_year - 1, ref_week, "final")
            wm = re.match(r"(\d{1,2})월 (\d)주차", ref_week)
            mom = np.nan
            if wm:
                mo, wk = int(wm.group(1)), int(wm.group(2))
                my, mm = (ref_year, mo - 1) if mo > 1 else (ref_year - 1, 12)
                mom = pick(df, "주", met, "*TOTAL", my, f"{mm:02d}월 {wk}주차", "final")
            rows.append(f"- {met}: {fmt_value(met, cur)} "
                        f"(전주비 {fmt_delta(met, cur, prv) or '–'}, "
                        f"전월비 {fmt_delta(met, cur, mom) or '–'}, "
                        f"전년비 {fmt_delta(met, cur, yoy) or '–'})")
        else:
            cur = pick(df, "월", met, "*TOTAL", ref_year, month_label(ref_month), "mtd")
            prv = pick(df, "월", met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            rows.append(f"- {met}: {fmt_value(met, cur)} (전년비 {fmt_delta(met, cur, prv) or '–'})")
    return "\n".join(rows)

def ai_generate_insight(df, ref_year, ref_month, ref_week, model,
                        focus="전주 주요 지표 현황", memo=""):
    """Claude API로 보고 인사이트 생성 → HTML(역신장 빨강) 반환. (텍스트, 에러) 튜플.
    memo: 사용자가 적은 정성 배경(프로모션·이벤트 등). 수치와 분리해 [배경 메모]로 주입."""
    key = _anthropic_key()
    if not key:
        return None, ("ANTHROPIC_API_KEY가 설정되지 않았습니다. "
                      "Streamlit Cloud → Settings → Secrets에 ANTHROPIC_API_KEY를 추가하세요.")
    try:
        import anthropic
    except ImportError:
        return None, "anthropic 패키지가 설치되지 않았습니다. requirements.txt 반영 후 재배포하세요."

    period = f"{ref_year}년 {ref_week}" if ref_week else f"{ref_year}년 {ref_month}월"
    facts = _ai_metric_facts(df, ref_year, ref_month, ref_week)
    system = (
        "당신은 LF몰 CRM 첫구매 보고서를 작성하는 데이터 분석가입니다. "
        "한국어 실무 보고 문구를 작성할 때, 긴 문장형 불릿(• ...은 ...으로 ...)을 피하고 "
        "반드시 '- '(메인 요약)와 'ㄴ '(세부 수치/해석)를 사용하는 계층형 불릿 구조로 출력하세요.\n\n"
        "예시:\n"
        "- 첫구매 거래액 및 고객수 동반 감소\n"
        "ㄴ 거래액 88.9백만원 (전주비 △1.5% · 전년비 △24.9%)\n"
        "ㄴ 고객수 639명 (전주비 △3.7% · 전년비 △34.1%)\n"
        "ㄴ 프로모션 종료 여파로 유입 대비 전환 효율이 저조한 것으로 보임\n\n"
        "[확정 수치]에 있는 숫자만 인용하고 수치를 지어내지 마세요. "
        "[배경 메모]는 원인 추정/해석에만 활용하며(단정짓지 말고 '~로 보임' 등 사용), 비어있으면 수치 팩트만 기재하세요. "
        "출력은 HTML로만 (<br> 로 줄바꿈). "
        "증감 수치 중 역신장(감소)은 <span style=\"color:#dc2626;font-weight:700\">…</span>, "
        "신장(증가)은 <span style=\"color:#16a34a;font-weight:700\">…</span>로 감싸세요. "
        "서론·맺음말 없이 바로 계층형 불릿만 출력하세요."
    )
    memo_block = (memo or "").strip() or "(없음)"
    user = (f"[기준: {period}]\n\n"
            f"[확정 수치]\n{facts}\n\n"
            f"[배경 메모]\n{memo_block}\n\n"
            f"위 자료로 '{focus}' 문구를 작성하세요.")
    try:
        client = anthropic.Anthropic(api_key=key)
        # Sonnet 5는 thinking 미지정 시 adaptive가 기본 켜져 max_tokens를 나눠 쓰므로 여유 있게
        resp = client.messages.create(
            model=model, max_tokens=4000, system=system,
            messages=[{"role": "user", "content": user}],
        )
        text = "".join(b.text for b in resp.content if b.type == "text").strip()
        return (text or None), (None if text else "빈 응답이 반환됐습니다.")
    except anthropic.AuthenticationError:
        return None, "API 키 인증에 실패했습니다. 키를 확인하세요."
    except anthropic.RateLimitError:
        return None, "요청이 많아 일시적으로 제한됐습니다. 잠시 후 다시 시도하세요."
    except Exception as e:
        return None, f"생성 중 오류: {e}"

# ══════════════════════════════════════════════════════
# 06. 앱푸시 동의 현황 페이지
# ══════════════════════════════════════════════════════
def render_push_page(df, ref_year, chart_years):
    st.markdown("## 앱푸시 동의 현황")
    gran_opt = st.radio("차트 보기 기준", ["일자별", "주차별", "월별"], horizontal=True, key="push_gran_opt")
    gran_map = {"일자별": "일", "주차별": "주", "월별": "월"}
    sel_gran = gran_map[gran_opt]

    cyrs = sorted(chart_years)  # multiselect 클릭 순서에 의존하지 않도록 정렬
    prev_year = cyrs[-2] if len(cyrs) >= 2 else None

    res_df = pd.DataFrame()  # 데이터 없으면 빈 채로 남음 — 하단 '상세 데이터' 가드용
    df_gran = df[df["gran"] == sel_gran]

    if df_gran.empty:
        st.info(f"{sel_gran} 단위 데이터가 없습니다.")
    else:
        st.subheader(f"{ref_year}년 {gran_opt} 신규가입 및 앱푸시 수신동의율 추이")
        dates = labels_sorted(df, sel_gran, [ref_year])

        if not dates:
            st.info(f"{ref_year}년 {gran_opt} 데이터가 없습니다.")
        else:
            rows = []
            # PUSH 데이터 집계용 (일자별 데이터를 주/월로 변환)
            df_push_daily = df[(df["metric"] == "앱푸시수신동의") & (df["gran"] == "일") & (df["segment"] == "*TOTAL") & (df["close"] == "final")].copy()
            if not df_push_daily.empty:
                df_push_daily["month"] = df_push_daily["label"].apply(lambda x: int(str(x).split('/')[0]) if '/' in str(x) else 0)
                df_push_daily["day"] = df_push_daily["label"].apply(lambda x: int(str(x).split('/')[1]) if '/' in str(x) else 0)
                # 대량 이관/재동의 스파이크 마스킹 (연도별 중앙값 규칙)
                df_push_daily = mask_push_spikes(df_push_daily)

                df_push_daily["주"] = df_push_daily.apply(lambda r: f"{r['month']:02d}월 {(int(r['day'])-1)//7 + 1}주차" if r['month']>0 else "", axis=1)
                df_push_daily["월"] = df_push_daily.apply(lambda r: f"{r['month']}월" if r['month']>0 else "", axis=1)
                df_push_daily["일"] = df_push_daily["label"]

                # 집계
                push_agg = df_push_daily.groupby(["year", sel_gran])["value"].mean().reset_index()
            else:
                push_agg = pd.DataFrame()

            # 월 필터링 로직 (업로드한 해당 월까지만)
            if not df_push_daily.empty:
                last_month = int(df_push_daily["month"].max())
                if sel_gran == "일":
                    dates = [d for d in dates if (int(d.split('/')[0]) if '/' in d else 0) <= last_month]
                elif sel_gran == "주":
                    dates = [d for d in dates if (int(d.split('월')[0]) if '월' in d else 0) <= last_month]
                elif sel_gran == "월":
                    dates = [d for d in dates if (int(d.replace('월', '')) if '월' in d else 0) <= last_month]

            for d in dates:
                join_val = pick(df, sel_gran, "가입자수", "*TOTAL", ref_year, d, "final")

                push_val = float('nan')
                if not push_agg.empty:
                    v = push_agg[(push_agg["year"] == ref_year) & (push_agg[sel_gran] == d)]["value"]
                    if not v.empty:
                        push_val = v.values[0]
                        if push_val == 0: push_val = float('nan')

                rate = float('nan')
                if pd.notna(push_val) and pd.notna(join_val) and join_val > 0:
                    rate = push_val / join_val
                    if rate > 1.0:
                        rate = float('nan')

                prev_rate = float('nan')
                if prev_year:
                    prev_join_val = pick(df, sel_gran, "가입자수", "*TOTAL", prev_year, d, "final")
                    prev_push_val = float('nan')

                    if not push_agg.empty:
                        v = push_agg[(push_agg["year"] == prev_year) & (push_agg[sel_gran] == d)]["value"]
                        if not v.empty:
                            prev_push_val = v.values[0]
                            if prev_push_val == 0: prev_push_val = float('nan')

                    if pd.notna(prev_push_val) and pd.notna(prev_join_val) and prev_join_val > 0:
                        prev_rate = prev_push_val / prev_join_val
                        if prev_rate > 1.0:
                            prev_rate = float('nan')

                rows.append({
                    "날짜": d,
                    "가입자수": join_val,
                    "앱푸시수신동의": push_val,
                    "동의율": rate,
                    "전년동의율": prev_rate
                })

            res_df = pd.DataFrame(rows)

            if not res_df.empty:
                fig = go.Figure()
                fig.add_trace(go.Bar(x=res_df["날짜"], y=res_df["가입자수"], name="가입자수", marker_color=clr("slate"), yaxis="y1"))
                fig.add_trace(go.Bar(x=res_df["날짜"], y=res_df["앱푸시수신동의"], name="앱푸시수신동의", marker_color=clr("blue"), yaxis="y1"))

                fig.add_trace(go.Scatter(x=res_df["날짜"], y=res_df["동의율"]*100, name=f"{ref_year} 동의율(%)", mode="lines+markers", line=dict(color=clr("red"), width=2), connectgaps=True, yaxis="y2"))

                if prev_year:
                    fig.add_trace(go.Scatter(x=res_df["날짜"], y=res_df["전년동의율"]*100, name=f"{prev_year} 동의율(%)", mode="lines", line=dict(color=clr("slate"), width=2, dash="dot"), connectgaps=True, yaxis="y2"))

                max_rate = max(res_df["동의율"].dropna()*100, default=10)
                if prev_year:
                    max_prev = max(res_df["전년동의율"].dropna()*100, default=10)
                    max_rate = max(max_rate, max_prev)

                if pd.isna(max_rate) or max_rate == 0:
                    max_rate = 10

                fig.update_layout(
                    barmode="group",
                    yaxis=dict(title="명", side="left", showgrid=False),
                    yaxis2=dict(title="%", side="right", overlaying="y", range=[0, max_rate * 1.2], showgrid=False),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    margin=dict(l=0, r=0, t=30, b=0),
                    xaxis=dict(type="category")
                )
                st.plotly_chart(fig, width="stretch")

    # ── 확장 뷰: 동의자 잔고·이탈·순증·요일 패턴 (앱푸시_* 지표)
    ext = df[(df["gran"] == "일") & df["metric"].isin(
        ["앱푸시_동의자수", "앱푸시_신규추가", "앱푸시_이탈"]) & (df["close"] == "final")].copy()
    if ext.empty:
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.info("동의자 잔고·이탈·순증 뷰는 PUSH 원천 엑셀을 다시 업로드하면 표시됩니다. "
                "(기존/신규/Total 섹션 전체를 새로 인식해 누적에 저장합니다)")
    else:
        ext[["_m", "_d"]] = ext["label"].str.extract(r"^(\d{1,2})/(\d{1,2})$")
        ext = ext.dropna(subset=["_m"]).copy()
        ext["dt"] = pd.to_datetime(dict(year=ext["year"].astype(int),
                                        month=ext["_m"].astype(int),
                                        day=ext["_d"].astype(int)), errors="coerce")
        ext = ext.dropna(subset=["dt"])
        ext = mask_push_spikes(ext)       # 흐름 지표 스파이크 (이관·재동의)
        ext = mask_level_glitches(ext)    # 잔고 하루짜리 글리치 (0 기록 등)

        bal = ext[ext["metric"] == "앱푸시_동의자수"]
        add = ext[ext["metric"] == "앱푸시_신규추가"]
        out = ext[ext["metric"] == "앱푸시_이탈"]

        # KPI — 최신 잔고(Total) + 당월 흐름
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        bt = bal[bal["segment"] == "*TOTAL"].sort_values("dt")
        if not bt.empty:
            last_dt = bt["dt"].iloc[-1]
            cur_bal = bt["value"].iloc[-1]
            pm_end = bt[bt["dt"] < last_dt.replace(day=1)]  # 전월말 잔고
            mom = (f"{cur_bal - pm_end['value'].iloc[-1]:+,.0f}명 (전월말比)"
                   if not pm_end.empty else None)
            in_mo = lambda s: s[(s["segment"] == "*TOTAL") &
                                (s["dt"].dt.year == last_dt.year) &
                                (s["dt"].dt.month == last_dt.month)]["value"].sum()
            mo_add, mo_out = in_mo(add), in_mo(out)
            k1, k2, k3, k4 = st.columns(4)
            k1.metric(f"총 동의자 ({last_dt:%m/%d} 기준)", f"{cur_bal:,.0f}명", mom)
            k2.metric(f"{last_dt.month}월 신규추가", f"{mo_add:,.0f}명")
            k3.metric(f"{last_dt.month}월 이탈", f"{mo_out:,.0f}명")
            k4.metric(f"{last_dt.month}월 순증", f"{mo_add - mo_out:+,.0f}명")

        # 동의자 잔고 추이 — 기존/신규/Total (선택 단위 기말값)
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("동의자 잔고 추이 — 기존/신규/Total")
        rule = {"일": "D", "주": "W-SUN", "월": "MS"}[sel_gran]
        fig_b = go.Figure()
        seg_pal = {"*TOTAL": "slate", "기존": "blue", "신규": "teal"}
        for seg in ["*TOTAL", "기존", "신규"]:
            s = (bal[bal["segment"] == seg].set_index("dt").sort_index()["value"]
                 .resample(rule).last().dropna())
            if s.empty: continue
            fig_b.add_trace(go.Scatter(x=s.index, y=s.values, mode="lines",
                                       name="Total" if seg == "*TOTAL" else seg,
                                       line=dict(color=clr(seg_pal[seg]), width=2)))
        fig_b.update_layout(**base_layout(300, title="수신동의 누적 잔고 (명)"))
        st.plotly_chart(fig_b, width="stretch")

        # 순증감 분해 — 신규추가(+) vs 이탈(−) + 순증 라인 (Total)
        st.subheader("동의 순증감 분해 — 신규추가 vs 이탈 (Total)")
        a = (add[add["segment"] == "*TOTAL"].set_index("dt").sort_index()["value"]
             .resample(rule).sum(min_count=1))
        o = (out[out["segment"] == "*TOTAL"].set_index("dt").sort_index()["value"]
             .resample(rule).sum(min_count=1))
        flow = pd.DataFrame({"신규추가": a, "이탈": o}).dropna(how="all")
        if not flow.empty:
            fig_f = go.Figure()
            fig_f.add_trace(go.Bar(x=flow.index, y=flow["신규추가"], name="신규추가(+)",
                                   marker_color=clr("green")))
            fig_f.add_trace(go.Bar(x=flow.index, y=-flow["이탈"], name="이탈(−)",
                                   marker_color=clr("red")))
            net = flow["신규추가"].fillna(0) - flow["이탈"].fillna(0)
            fig_f.add_trace(go.Scatter(x=flow.index, y=net, name="순증",
                                       mode="lines+markers",
                                       line=dict(color=clr("slate"), width=2),
                                       marker=dict(size=4)))
            lyf = base_layout(320, title=f"{gran_opt} 신규추가·이탈·순증 (명)")
            lyf["barmode"] = "relative"
            fig_f.update_layout(**lyf)
            st.plotly_chart(fig_f, width="stretch")
            st.caption("※ 대량 이관·재동의 스파이크(연도별 중앙값 10배 초과)는 왜곡 방지를 위해 제외")

        # 요일 패턴 — 신규추가/이탈 요일별 일평균 (기준 연도)
        st.subheader(f"요일별 평균 — 신규추가·이탈 ({ref_year}년)")
        dows = ["월", "화", "수", "목", "금", "토", "일"]
        fig_d = go.Figure()
        for nm, src, cname in [("신규추가", add, "green"), ("이탈", out, "red")]:
            s = src[(src["segment"] == "*TOTAL") & (src["dt"].dt.year == ref_year)]
            if s.empty: continue
            m = s.groupby(s["dt"].dt.dayofweek)["value"].mean().reindex(range(7))
            fig_d.add_trace(go.Bar(x=dows, y=m.values, name=nm, marker_color=clr(cname)))
        fig_d.update_layout(**base_layout(280, title="요일별 일평균 (명)"))
        st.plotly_chart(fig_d, width="stretch")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("월별 전년비 증감 추이표 (YoY)")

    push_metrics = ["가입자수", "앱푸시수신동의", "동의율"]
    sub = df[(df["gran"] == "일") & (df["year"].isin(cyrs)) & (df["metric"].isin(["가입자수", "앱푸시수신동의"])) & (df["segment"] == "*TOTAL") & (df["close"] == "final")].copy()

    if sub.empty:
        st.info("일자별 가입자수·앱푸시수신동의 데이터가 없어 월별 YoY 표를 표시할 수 없습니다.")
    else:
        sub["month"] = sub["label"].apply(lambda x: int(str(x).split('/')[0]) if '/' in str(x) else 0)
        sub["day"] = sub["label"].apply(lambda x: int(str(x).split('/')[1]) if '/' in str(x) else 0)

        # 푸시동의 스파이크 이상치 마스킹 (가입자수 행은 규칙상 제외됨)
        sub = mask_push_spikes(sub)

        # 업로드한 최신 월 및 일(MTD) 계산
        max_month = 12
        max_day = 31
        if cyrs:
            last_year = cyrs[-1]
            for m in range(12, 0, -1):
                r = sub[(sub["year"] == last_year) & (sub["month"] == m) & sub["value"].notna()]
                if not r.empty:
                    max_month = m
                    max_day = r["day"].max()
                    break

        # 당월(max_month)에 대해서는 전년도 데이터도 MTD(max_day)까지만 필터링하여 누적 비교
        if max_month > 0:
            sub = sub[~((sub["month"] == max_month) & (sub["day"] > max_day))]

        grp = sub.groupby(["year", "month", "metric"])["value"].sum().unstack("metric").reset_index()

        if "가입자수" in grp.columns and "앱푸시수신동의" in grp.columns:
            grp["동의율"] = grp["앱푸시수신동의"] / grp["가입자수"]
            grp.loc[grp["동의율"] > 1.0, "동의율"] = float('nan')
        else:
            grp["동의율"] = float('nan')

        out_tbl = {}
        cols = []
        for m in range(1, max_month + 1):
            if m == max_month:
                cols.append(f"{m}월 (MTD)")
            else:
                cols.append(f"{m}월")

        for met in push_metrics:
            for y in cyrs:
                row_name = f"{met} ({y})"
                out_tbl[row_name] = []
                for m in range(1, max_month + 1):
                    r = grp[(grp["year"] == y) & (grp["month"] == m)]
                    out_tbl[row_name].append(r[met].values[0] if (not r.empty and met in r.columns) else float('nan'))

            # 최근 2개 연도 간 YoY
            if len(cyrs) >= 2:
                yoy_name = f"{met} (YoY)"
                out_tbl[yoy_name] = []
                prev_y, cur_y = cyrs[-2], cyrs[-1]
                for i, m in enumerate(range(1, max_month + 1)):
                    prev_val = out_tbl[f"{met} ({prev_y})"][i]
                    cur_val = out_tbl[f"{met} ({cur_y})"][i]
                    if pd.isna(cur_val) and pd.isna(prev_val):
                        out_tbl[yoy_name].append("-")
                    else:
                        out_tbl[yoy_name].append(fmt_delta(met, cur_val, prev_val) or "-")

        tbl_df = pd.DataFrame(out_tbl, index=cols).T

        disp = tbl_df.astype(object).copy()
        for row_name in disp.index:
            if "(YoY)" in row_name:
                continue
            else:
                met = row_name.split(" ")[0]
                disp.loc[row_name] = [fmt_value(met, v) if pd.notna(v) else "-" for v in tbl_df.loc[row_name]]

        def style_yoy_rows(tbl):
            def _color(v):
                s = str(v)
                if s.startswith("△"): return "color:#dc2626;font-weight:600"
                if s.startswith("+"): return "color:#16a34a;font-weight:600"
                return ""
            def _highlight(df_):
                styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                for idx in df_.index:
                    if "(YoY)" in str(idx):
                        for col in df_.columns:
                            styles.loc[idx, col] = _color(df_.loc[idx, col])
                return styles
            return tbl.style.apply(_highlight, axis=None)

        st.dataframe(style_yoy_rows(disp), width="stretch")

    if not res_df.empty:
        st.markdown("### 상세 데이터")
        disp_df = res_df.copy()
        disp_df["앱푸시수신동의"] = disp_df["앱푸시수신동의"].apply(lambda x: f"{int(x):,}" if not pd.isna(x) else "–")
        disp_df["가입자수"] = disp_df["가입자수"].apply(lambda x: f"{int(x):,}" if not pd.isna(x) else "–")
        disp_df["동의율"] = disp_df["동의율"].apply(lambda x: f"{x*100:.1f}%" if not pd.isna(x) else "–")
        st.dataframe(disp_df.set_index("날짜"), width="stretch")

# ══════════════════════════════════════════════════════
# 페이지 PDF 저장 (브라우저 인쇄 → PDF, 차트 포함)
# ══════════════════════════════════════════════════════
def print_button(label="이 페이지 PDF 저장 / 인쇄"):
    # st.components.v1.html은 제거 예정(2026-06-01 이후) — 후속 API인 st.iframe 사용
    st.iframe(
        f"""<button onclick="window.parent.print()"
        style="float:right;background:#2E68B0;color:#fff;border:0;border-radius:6px;
        padding:7px 14px;font-size:13px;font-weight:600;cursor:pointer;
        font-family:'Pretendard',-apple-system,sans-serif">{label}</button>
        <div style="clear:both"></div>""", height=44)
    st.caption("버튼이 동작하지 않으면 Ctrl+P(Mac ⌘+P) → 대상을 'PDF로 저장'으로 인쇄하세요.")

# ══════════════════════════════════════════════════════
# 메인 앱
# ══════════════════════════════════════════════════════
def main():
    if "wr_texts" not in st.session_state:
        st.session_state.wr_texts = load_insights()

    with st.sidebar:
        st.markdown("## 📋 주간보고 통합")
        files = st.file_uploader(
            "원천 엑셀/CSV/ZIP 업로드 (복수 선택)",
            type=["xlsx", "xls", "csv", "zip"], accept_multiple_files=True, key="wr_up",
            help="주간 폴더를 zip으로 묶어 통째로 올려도 됩니다. "
                 "전체관점 마스터(일/주/월) + 지표별 파일(가입율·가입자수·당일가입 첫구매율·비회원 트래픽)을 자동 인식합니다.")
        st.markdown("---")
        PAGES = ["01. 주간보고 요약", "02. 월별 추이", "03. 주차별 추이",
                 "04. 채널별 실적", "05. 통합 데이터·다운로드", "06. 앱푸시 동의 현황", "07. 첫구매 고객 세그먼트 성과"]
        page = st.radio("페이지", PAGES, key="wr_page")

    stored = load_store()
    expanded = expand_uploads(files) if files else []
    df_new = combine_files(tuple(expanded)) if expanded else pd.DataFrame()

    has_any = not stored.empty or not df_new.empty
    if files and df_new.empty and stored.empty:
        st.error("업로드한 파일에서 데이터를 읽지 못했습니다. 파일명 형식을 확인해주세요.")
        st.stop()
    if not has_any:
        st.markdown("## 📋 주간보고 통합 — 시작하기")
        st.caption("누적 데이터가 없습니다. 원천 파일을 올리거나, 예전 백업(ZIP)을 복원하세요.")
        cU, cR = st.columns(2)
        with cU:
            st.markdown("#### ① 원천 파일 업로드")
            source_upload_widget("wr_up_main")
        with cR:
            st.markdown("#### ② 또는 백업 복원")
            restore_widget("wr_restore_empty", label="백업 ZIP / CSV / JSON 올리기")
            st.caption("이전에 받은 `주간보고백업_*.zip`을 그대로 올리면 데이터·메모가 통째로 복원됩니다.")
        st.markdown("""
---
**인식되는 원천 파일**
- **마스터**: `전체관점 - 일자별/주별/월별 실적 (기본)`
- **지표별**: `월_가입율(일평균)`, `주_가입자수(일평균)`, `일_비회원 트래픽(일평균)`, `월_당일가입 첫구매율 (일평균)` …
- **앱푸시**: 파일명에 `PUSH`/`앱푸시`/`수신동의` 포함 또는 헤더가 앱푸시 형식이면 자동 인식
- 주간 폴더를 **zip으로 묶어 통째로** 올려도 됩니다.
""")
        st.stop()

    # 누적 저장소와 병합 — 미리보기(저장 전까지 영구 반영 안 함)
    df = merge_store(stored, df_new)

    has_new = not df_new.empty
    sig = tuple(sorted((n, len(b)) for n, b in expanded))
    with st.sidebar:
        # 업로드한 파일이 각각 무엇으로 인식됐는지 (미인식 파일 가시화)
        if expanded:
            cls = classify_uploads(tuple(expanded))
            n_ok = sum(1 for _, k, _ in cls if k.startswith(("✅", "♻")))
            n_bad = len(cls) - n_ok
            head = f"📄 업로드 인식 {n_ok}/{len(cls)}" + (f" · ⚠ 미인식 {n_bad}" if n_bad else "")
            with st.expander(head, expanded=bool(n_bad)):
                for nm, kind, nrows in cls:
                    base = os.path.basename(nm)
                    st.markdown(f"- `{base}` → {kind}"
                                + (f" ({nrows:,}행)" if nrows else ""))
                if n_bad:
                    st.caption("미인식 파일은 파일명(전체관점/월_가입율 등)·형식을 확인하세요. "
                               "인식된 파일만 저장에 반영됩니다.")
        if has_new:
            added, updated = upload_diff(stored, df_new)
            saved = st.session_state.get("wr_saved_sig") == sig
            if saved:
                st.success("저장됨 ✓ (누적 반영 완료)")
            else:
                st.warning(f"새 데이터 감지 — 추가 {added}기간 · 갱신(겹침) {updated}기간\n\n"
                           "**저장** 눌러야 누적에 반영됩니다.")
                if st.button("💾 저장 (누적 반영)", key="wr_commit",
                             type="primary", width="stretch"):
                    if not df_new.empty: save_store(df)
                    st.session_state["wr_saved_sig"] = sig
                    st.rerun()

    if df.empty:
        st.warning("첫구매(전체관점/지표별) 데이터가 없습니다. 원천 파일을 업로드해주세요.")
        st.stop()

    # ── 인식 결과 + 필터
    years_all = sorted(df["year"].dropna().unique().astype(int))
    ly, llabel = latest_period(df, "월")
    ref_year_default = ly or years_all[-1]
    with st.sidebar:
        st.markdown("---")
        src = f"인식된 파일 {len(expanded)}개" if expanded else "누적 데이터 사용 중"
        st.caption(f"{src} · 지표 {df['metric'].nunique()}종 · "
                   f"{years_all[0]}–{years_all[-1]}년")
        st.markdown("**기준 기간**")
        ref_year = st.selectbox("기준 연도", years_all[::-1],
                                index=years_all[::-1].index(ref_year_default), key="wr_refy")
        months_avail = sorted({int(re.match(r"(\d+)월", l).group(1))
                               for l in df[(df["gran"] == "월") & (df["year"] == ref_year)]["label"]})
        ref_month = st.selectbox("기준 월", months_avail[::-1], key="wr_refm")
        weeks_avail = (df[(df["gran"] == "주") & (df["year"] == ref_year) & df["value"].notna()]
                       [["label", "sortkey"]].drop_duplicates()
                       .sort_values("sortkey")["label"].tolist())
        if weeks_avail:
            # 최신 주차가 진행 중(일마감 데이터이거나 오늘이 속한 주차)이면 직전 주를 기본값으로
            latest_w = weeks_avail[-1]
            today = today_kst()
            cur_week_lbl = f"{today.month:02d}월 {(today.day - 1) // 7 + 1}주차"
            is_partial = not df[(df["gran"] == "주") & (df["year"] == ref_year) &
                                (df["label"] == latest_w) & (df["close"] == "mtd")].empty
            in_progress = is_partial or (ref_year == today.year and latest_w == cur_week_lbl)
            default_week = (weeks_avail[-2] if in_progress and len(weeks_avail) >= 2
                            else latest_w)
            ref_week = st.selectbox("기준 주차", weeks_avail[::-1],
                                    index=weeks_avail[::-1].index(default_week),
                                    key="wr_refw",
                                    help="주간보고 대상 주차. 최신 주차가 진행 중이면 직전 완료 주차가 기본값입니다.")
        else:
            ref_week = None
        st.markdown("**차트 연도**")
        default_yrs = years_all[-2:] if len(years_all) >= 2 else years_all
        chart_years = st.multiselect("비교 연도", years_all, default=default_yrs, key="wr_cyrs")
        st.markdown("**채널**")
        ch_sel = st.multiselect("채널 선택", CHANNELS, default=CHANNELS, key="wr_ch")
        if not chart_years: chart_years = default_yrs

        st.markdown("---")
        st.markdown("**AI 인사이트 모델**")
        ai_label = st.selectbox("모델 선택", list(AI_MODELS.keys()), key="wr_ai_model_label")
        st.session_state["wr_ai_model"] = AI_MODELS[ai_label]
        st.caption("✅ API 키 설정됨" if _anthropic_key()
                   else "⚠ ANTHROPIC_API_KEY 미설정 — Secrets에 추가하세요")

        st.markdown("---")
        st.markdown("**백업 · 복원**")
        saved_rows = len(load_store())
        pend = " · 저장 시 반영" if (has_new and not st.session_state.get("wr_saved_sig") == sig) else ""
        st.caption(f"누적 {saved_rows:,}행 저장됨 / 현재 보기 {len(df):,}행{pend} · "
                   f"메모 {len(st.session_state.wr_texts)}개")

        # 통합 백업 — 누적 데이터 + 메모를 한 ZIP으로 (하나만 받아도 전부 보존)
        st.download_button(
            "⬇ 통합 백업 (ZIP · 데이터+메모)",
            make_backup_zip(df, st.session_state.wr_texts),
            f"주간보고백업_{today_kst():%Y%m%d}.zip", "application/zip",
            width="stretch", type="primary",
            help="누적 데이터 CSV와 보고란·메모 JSON을 한 파일로 백업합니다. "
                 "재배포로 초기화돼도 이 ZIP을 '백업 복원'에 올리면 통째로 되살아납니다.")

        # 통합 복원 — zip/csv/json 자동 인식
        restore_widget("wr_restore")

        with st.expander("개별 백업 (데이터만 / 메모만)"):
            st.download_button("⬇ 누적 데이터 (CSV)",
                               df[STORE_COLS].to_csv(index=False).encode("utf-8-sig"),
                               f"wr_data_store_{today_kst():%Y%m%d}.csv", "text/csv",
                               width="stretch")
            st.download_button("⬇ 보고란·메모 (JSON)",
                               json.dumps(st.session_state.wr_texts, ensure_ascii=False,
                                          indent=2).encode("utf-8"),
                               f"wr_insights_{today_kst():%Y%m%d}.json", "application/json",
                               width="stretch")

        # 초기화 — 2단계 확인 (실수 방지)
        if st.session_state.get("wr_confirm_clear"):
            st.warning("정말 초기화할까요? 누적 데이터가 모두 삭제되며 되돌릴 수 없습니다. "
                       "(메모는 유지 · 초기화 전 통합 백업 권장)")
            cc1, cc2 = st.columns(2)
            if cc1.button("삭제 확인", key="wr_clear_yes", type="primary", width="stretch"):
                if os.path.exists(DATA_STORE): os.remove(DATA_STORE)
                st.cache_data.clear()
                st.session_state["wr_confirm_clear"] = False
                st.session_state.pop("wr_saved_sig", None)
                st.rerun()
            if cc2.button("취소", key="wr_clear_no", width="stretch"):
                st.session_state["wr_confirm_clear"] = False; st.rerun()
        else:
            if st.button("🗑 누적 데이터 초기화", key="wr_clear_store", width="stretch"):
                st.session_state["wr_confirm_clear"] = True; st.rerun()

    texts = st.session_state.wr_texts
    print_button()

    # ════════════ 01. 주간보고 요약 ════════════
    if page == "01. 주간보고 요약":
        st.markdown(f"## 첫구매 주간보고 — {ref_year}년 {ref_month}월")
        wy, wlabel = week_ref(df, ref_year, ref_week)
        if wlabel:
            st.caption(f"기준 주차: {week_disp(wy, wlabel)}")

        # KPI 카드 (기준 주차 — 전주비·전월비·전년비 모두 주차 기준)
        if wlabel:
            cols = st.columns(4)
            kpi_metrics = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가", "가입자수"]
            py, plb = prev_label(df, "주", wy, wlabel)
            # 전월 동일 주차 (예: 06월 1주차 → 05월 1주차)
            mom_y = mom_lbl = None
            wm = re.match(r"(\d{1,2})월 (\d)주차", wlabel)
            if wm:
                mo, wk = int(wm.group(1)), int(wm.group(2))
                mom_y, mom_m = (wy, mo - 1) if mo > 1 else (wy - 1, 12)
                mom_lbl = f"{mom_m:02d}월 {wk}주차"
            for col, met in zip(cols, kpi_metrics):
                cur = pick(df, "주", met, "*TOTAL", wy, wlabel, "mtd")
                deltas = []
                prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
                deltas.append((fmt_delta(met, cur, prv), "전주비"))
                if mom_lbl:
                    mom = pick(df, "주", met, "*TOTAL", mom_y, mom_lbl, "final")
                    deltas.append((fmt_delta(met, cur, mom), "전월비"))
                yoy = pick(df, "주", met, "*TOTAL", wy - 1, wlabel, "final")
                deltas.append((fmt_delta(met, cur, yoy), "전년비"))
                pills = ""
                for d, lab in deltas:
                    if d:
                        neg = d.startswith("△")
                        cls = "down" if neg else "up"
                        prefix = "" if neg else "↑ "
                        pills += f'<div class="kpi-delta {cls}">{prefix}{d} ({lab})</div>'
                    else:
                        pills += f'<div class="kpi-delta na">– ({lab})</div>'
                col.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">{met} ({week_disp(wy, wlabel)})</div>'
                    f'<div class="kpi-value">{fmt_value(met, cur)}</div>{pills}</div>',
                    unsafe_allow_html=True)
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 실적 요약 — 전주비 (주차 기준)
        if wlabel:
            st.subheader("실적 요약 (일평균 · 전주비)")
            st.caption(f"기준 주차: {week_disp(wy, wlabel)} — 전주·전년 동주 대비")
            st.dataframe(style_delta_cols(wow_summary_table(df, wy, wlabel, METRICS7)),
                         width="stretch")
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 실적 요약 YoY 표
        st.subheader("실적 요약 (일평균 · 전년비)")
        tbl, (pm_y, pm_m) = yoy_summary_table(df, ref_year, ref_month, METRICS7)
        st.caption(f"전월({pm_m}월)은 월마감, 당월({ref_month}월)은 일마감(MTD) 기준 동일기간 비교")
        st.dataframe(style_delta_cols(tbl), width="stretch")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 다각도 뷰 — 전환 퍼널(단계 카드) + 채널 기여. 비교 기준 전환 가능
        cmp_mode = st.radio("비교 기준", ["주간 — 전년 동주", "월누적(MTD) — 전년 동월"],
                            horizontal=True, key="wr_multi_cmp")
        weekly_mode = cmp_mode.startswith("주간")
        if weekly_mode and not wlabel:
            st.info("주차 데이터가 없습니다. 월누적(MTD) 비교를 선택하세요.")
        else:
            if weekly_mode:
                period_lbl, base_lbl = week_disp(wy, wlabel), "전년 동주"
                x_prv, x_cur = f"{wy-1}년", f"{wy}년"
                def get_cur(met, seg="*TOTAL"):
                    return pick(df, "주", met, seg, wy, wlabel, "mtd")
                def get_prv(met, seg="*TOTAL"):
                    return pick(df, "주", met, seg, wy - 1, wlabel, "final")
            else:
                period_lbl = f"{ref_year}년 {ref_month}월 누적(MTD)"
                base_lbl = "전년 동월 MTD"
                x_prv, x_cur = f"{ref_year-1}년 {ref_month}월", f"{ref_year}년 {ref_month}월"
                def get_cur(met, seg="*TOTAL"):
                    return pick(df, "월", met, seg, ref_year, month_label(ref_month), "mtd")
                def get_prv(met, seg="*TOTAL"):
                    # 전년 동월도 동일기간(MTD) 잘린 값 우선 — 실적 요약 표와 동일 기준
                    return pick(df, "월", met, seg, ref_year - 1, month_label(ref_month), "mtd")

            st.subheader("전환 퍼널 — 트래픽→가입→첫구매")
            st.caption(f"{period_lbl} vs {base_lbl} (일평균) · "
                       "첫구매 고객에는 과거 가입자도 포함되어 단계 비율은 참고용")
            stages = ["비회원트래픽", "가입자수", "첫구매 고객수"]
            cur_v = [get_cur(m) for m in stages]
            pry_v = [get_prv(m) for m in stages]
            if any(np.isnan(v) for v in cur_v):
                st.info("해당 주차 퍼널 데이터가 부족합니다.")
            else:
                # 트래픽이 가입·첫구매의 수십~수백 배라 면적형 퍼널은 왜곡됨 →
                # 단계 카드 + 전환율 pill 로 표현 (전환율을 1급 정보로)
                def _stage_rate(a, b):
                    return a / b if (not np.isnan(a) and not np.isnan(b) and b > 0) else np.nan
                cells = []
                for i, (mname, cv, pv) in enumerate(zip(stages, cur_v, pry_v)):
                    d = fmt_delta(mname, cv, pv)
                    pill = (f'<div class="kpi-delta {"down" if d.startswith("△") else "up"}">'
                            f'{d} (전년동주)</div>' if d else
                            '<div class="kpi-delta na">– (전년동주)</div>')
                    cells.append(
                        f'<div class="kpi-card" style="flex:1.2;min-width:0">'
                        f'<div class="kpi-label">{mname}</div>'
                        f'<div class="kpi-value">{cv:,.0f}명</div>{pill}</div>')
                    if i < 2:
                        cr = _stage_rate(cur_v[i + 1], cv)
                        pr = _stage_rate(pry_v[i + 1], pv)
                        rate_lbl = "가입율" if i == 0 else "가입 대비 첫구매"
                        cur_s = "–" if np.isnan(cr) else f"{cr*100:.2f}%"
                        sub = ""
                        if not (np.isnan(cr) or np.isnan(pr)):
                            pp = (cr - pr) * 100
                            colr = "#dc2626" if pp < 0 else "#16a34a"
                            sign = "△" if pp < 0 else "+"
                            sub = (f'<div style="font-size:11px;color:#64748b">전년 {pr*100:.2f}% '
                                   f'<span style="color:{colr};font-weight:700">'
                                   f'{sign}{abs(pp):.2f}%p</span></div>')
                        cells.append(
                            f'<div style="flex:0.9;display:flex;flex-direction:column;'
                            f'justify-content:center;text-align:center;min-width:0">'
                            f'<div style="color:#94a3b8;font-size:16px;line-height:1">→</div>'
                            f'<div style="font-size:12px;color:#64748b;margin-top:2px">{rate_lbl}</div>'
                            f'<div style="font-size:17px;font-weight:700;color:#1e293b">{cur_s}</div>'
                            f'{sub}</div>')
                st.markdown('<div style="display:flex;gap:10px;align-items:stretch">'
                            + "".join(cells) + '</div>', unsafe_allow_html=True)

            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
            st.subheader("첫구매 거래액 YoY — 채널 기여")
            st.caption(f"{period_lbl} {base_lbl} 대비 증감 분해 (일평균·만원) · "
                       "세로축은 변동 구간만 확대 표시")
            base_t = get_prv("첫구매 거래액")
            cur_t = get_cur("첫구매 거래액")
            if np.isnan(base_t) or np.isnan(cur_t):
                st.info("해당 주차 채널 데이터가 없습니다.")
            else:
                MW = 1e4  # 만원
                labels, deltas = [], []
                for chn in CHANNELS:
                    pv = get_prv("첫구매 거래액", chn)
                    cv = get_cur("첫구매 거래액", chn)
                    if np.isnan(pv) and np.isnan(cv): continue
                    d = (0 if np.isnan(cv) else cv) - (0 if np.isnan(pv) else pv)
                    labels.append(chn); deltas.append(d / MW)
                resid = (cur_t - base_t) / MW - sum(deltas)
                if abs(resid) > 5:  # 5만원 이상 잔차만 '기타'로 표시
                    labels.append("기타"); deltas.append(resid)
                fig_w = go.Figure(go.Waterfall(
                    x=[x_prv] + labels + [x_cur],
                    measure=["absolute"] + ["relative"] * len(labels) + ["total"],
                    y=[base_t / MW] + deltas + [0],
                    text=[f"{base_t/MW:,.0f}"] + [f"{d:+,.0f}" for d in deltas]
                         + [f"{cur_t/MW:,.0f}"],
                    textposition="outside", cliponaxis=False,
                    increasing=dict(marker=dict(color=clr("green"))),
                    decreasing=dict(marker=dict(color=clr("red"))),
                    totals=dict(marker=dict(color=clr("slate"))),
                    connector=dict(line=dict(color="#e2e8f0")),
                    hovertemplate="%{x}<br>변동 %{delta:+,.0f}만원 · 누계 %{final:,.0f}만원<extra></extra>",
                ))
                # 총액 막대가 델타를 압도하지 않도록 세로축을 변동 구간 주변으로 제한
                run, acc = [base_t / MW], base_t / MW
                for d in deltas:
                    acc += d; run.append(acc)
                run.append(cur_t / MW)
                lo, hi = min(run), max(run)
                span = max(hi - lo, hi * 0.01, 1.0)
                lyw = base_layout(360, title="주간 일평균 거래액 (만원) — 변동 구간 확대")
                lyw["showlegend"] = False
                lyw["yaxis"]["range"] = [lo - span * 0.45, hi + span * 0.55]
                fig_w.update_layout(**lyw)
                st.plotly_chart(fig_w, width="stretch")
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 보고란
        draft = auto_draft(df, ref_year, ref_month, ref_week=wlabel)
        ai_model = st.session_state.get("wr_ai_model", DEFAULT_AI_MODEL)
        cL, cR = st.columns(2)
        with cL:
            report_text_block("wr_metrics_summary", "전주 주요 지표 현황",
                              default=draft, regen=draft,
                              ai_fn=lambda memo: ai_generate_insight(df, ref_year, ref_month,
                                                                wlabel, ai_model, memo=memo))
        with cR:
            report_text_block("wr_exec_summary", "금주 집행 내용 요약")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 핵심 차트 (주차별 YoY 3종)
        st.subheader("주차별 추이 — 전년 비교")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]):
            with col:
                st.plotly_chart(yoy_chart(df, "주", met, chart_years, h=280),
                                width="stretch")

    # ════════════ 02. 월별 추이 ════════════
    elif page == "02. 월별 추이":
        st.markdown("## 월별 추이")
        st.subheader("월별 추이 차트 — 전년 비교")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]):
            with col:
                st.plotly_chart(yoy_chart(df, "월", met, chart_years, h=280),
                                width="stretch")
        c4, c5, c6 = st.columns(3)
        for col, met in zip((c4, c5, c6), ["비회원트래픽", "가입자수", "가입율"]):
            with col:
                st.plotly_chart(yoy_chart(df, "월", met, chart_years, h=280),
                                width="stretch")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("월별 추이표 (일평균)")
        tbl = trend_table(df, "월", METRICS7, chart_years)
        st.dataframe(style_trend(tbl, METRICS7), width="stretch")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        ai_model = st.session_state.get("wr_ai_model", DEFAULT_AI_MODEL)
        report_text_block(
            f"wr_month_memo_{ref_year}_{ref_month}",
            f"{ref_year}년 {ref_month}월 액션·이슈사항",
            ai_fn=lambda memo: ai_generate_insight(df, ref_year, ref_month, None, ai_model,
                                              focus=f"{ref_year}년 {ref_month}월 액션·이슈 및 인사이트",
                                              memo=memo))

    # ════════════ 03. 주차별 추이 ════════════
    elif page == "03. 주차별 추이":
        st.markdown("## 주차별 추이")
        st.subheader("주차별 추이 차트 — 전년 비교")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]):
            with col:
                st.plotly_chart(yoy_chart(df, "주", met, chart_years, h=280),
                                width="stretch")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"주차별 추이표 — {ref_year}년")
        tbl = trend_table(df, "주", METRICS7, [ref_year])
        if not tbl.empty:
            recent = tbl.columns[-16:]
            st.dataframe(style_trend(tbl[recent], METRICS7), width="stretch")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("전주비(WoW)·전년비(YoY) 증감")
        wy, wlabel = week_ref(df, ref_year, ref_week)
        if wlabel:
            st.caption(f"기준 주차: {week_disp(wy, wlabel)}")
            st.dataframe(style_delta_cols(wow_summary_table(df, wy, wlabel, METRICS7)),
                         width="stretch")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        ai_model = st.session_state.get("wr_ai_model", DEFAULT_AI_MODEL)
        report_text_block(
            f"wr_week_memo_{wy}_{wlabel}",
            f"{wy}년 {wlabel} 액션·이슈사항" if wlabel else "주차별 액션·이슈사항",
            ai_fn=lambda memo: ai_generate_insight(df, ref_year, ref_month, wlabel, ai_model,
                                              focus=f"{wlabel} 주차 액션·이슈 및 인사이트",
                                              memo=memo))

    # ════════════ 04. 채널별 실적 ════════════
    elif page == "04. 채널별 실적":
        st.markdown("## 채널별 실적")
        avail = [m for m in METRICS7 if (df["metric"] == m).any()]
        met = st.selectbox("지표 선택", avail, key="wr_chmet")

        st.subheader(f"{met} — {ref_year}년 {ref_month}월 채널별 전년비")
        rows = []
        for seg in ["*TOTAL"] + [c for c in CHANNELS if c in ch_sel]:
            pv = pick(df, "월", met, seg, ref_year - 1, month_label(ref_month), "mtd")
            cv = pick(df, "월", met, seg, ref_year, month_label(ref_month), "mtd")
            rows.append({"채널": seg,
                         f"{ref_year-1}년 {ref_month}월": fmt_value(met, pv),
                         f"{ref_year}년 {ref_month}월": fmt_value(met, cv),
                         "전년비": fmt_delta(met, cv, pv) or "–"})
        st.dataframe(style_delta_cols(pd.DataFrame(rows).set_index("채널")),
                     width="stretch")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"{met} — 채널별 월 추이 ({ref_year}년)")
        unit, div = METRIC_UNIT.get(met, ("", 1))
        if met in PCT_METRICS: div, unit = 0.01, "%"
        fig = go.Figure()
        x = [month_label(i) for i in range(1, 13)]
        for seg in [c for c in CHANNELS if c in ch_sel]:
            s = series_by_label(df, "월", met, seg, ref_year).reindex(x).dropna()
            if s.empty: continue
            fig.add_trace(go.Scatter(
                x=s.index.tolist(), y=(s / div).tolist(), mode="lines+markers", name=seg,
                line=dict(color=clr(CHANNEL_PAL.get(seg, "blue")), width=1.8),
                marker=dict(size=4)))
        ly = base_layout(340, ysuffix=unit if unit == "%" else "",
                         title=f"{met} 채널별 ({unit})")
        fig.update_layout(**ly)
        st.plotly_chart(fig, width="stretch")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"{met} — 채널 × 월 표 ({ref_year}년)")
        rows = []
        for seg in ["*TOTAL"] + [c for c in CHANNELS if c in ch_sel]:
            s = series_by_label(df, "월", met, seg, ref_year)
            row = {"채널": seg}
            for lb in [month_label(i) for i in range(1, 13)]:
                if lb in s.index and not np.isnan(s[lb]):
                    row[lb] = fmt_value(met, s[lb])
            rows.append(row)
        st.dataframe(pd.DataFrame(rows).set_index("채널"), width="stretch")

    # ════════════ 05. 통합 데이터·다운로드 ════════════
    elif page == "05. 통합 데이터·다운로드":
        st.markdown("## 통합 데이터 · 다운로드")
        st.caption("업로드한 모든 파일을 합친 통합 long 데이터입니다.")
        st.dataframe(df.sort_values(["gran", "metric", "segment", "sortkey"]).head(2000),
                     width="stretch", height=420)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("통합 워크북 다운로드")
        st.caption("`첫구매_요약`(요약표·보고란·YoY 차트·채널표) + `월`·`주`(통합 데이터) 3개 시트")
        if st.button("📥 엑셀 워크북 생성", key="wr_build"):
            with st.spinner("워크북 생성 중…"):
                xls = build_workbook(df, st.session_state.wr_texts,
                                     ref_year, ref_month, chart_years)
            st.download_button(
                "다운로드 — 첫구매_주간보고.xlsx", xls,
                file_name=f"첫구매_주간보고_{ref_year}{ref_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("통합 long 데이터 CSV", csv,
                           f"통합데이터_{today_kst():%Y%m%d}.csv", "text/csv")

    # ════════════ 06. 앱푸시 동의 현황 ════════════
    elif page == "06. 앱푸시 동의 현황":
        render_push_page(df, ref_year, chart_years)

    # ════════════ 07. 첫구매 고객 세그먼트 성과 ════════════
    elif page == "07. 첫구매 고객 세그먼트 성과":
        st.markdown("## 첫구매 고객 세그먼트 성과")
        
        # 세그먼트 선택 필터 추가
        all_segs = ["1_신규", "1_당월신규", "2_기가입신규", "3_기존"]
        
        # 데이터프레임에 존재하는 세그먼트만 추출
        available_segs = []
        for s in all_segs:
            if not df[df["segment"] == s].empty:
                available_segs.append(s)
                
        if not available_segs:
            available_segs = all_segs
            
        sel_seg = st.selectbox("세그먼트 선택", available_segs)
        segs = [sel_seg]
        
        def wow_segment_table(wy, wlabel):
            df_gran = df[df["gran"] == "주"]
            # 해당 세그먼트에 존재하는 모든 지표 추출 (순서 유지)
            metrics = df_gran[df_gran["segment"] == sel_seg]["metric"].dropna().unique().tolist()
            if not metrics:
                metrics = ["거래액", "고객수", "객단가", "CR"]
                
            py, plb = prev_label(df, "주", wy, wlabel)
            cols = [week_disp(py, plb), week_disp(wy, wlabel), "전주비",
                    week_disp(wy - 1, wlabel), "전년비"]
            rows = []
            
            def _get_val(yr, lb, c, seg, met):
                v = pick(df, "주", met, seg, yr, lb, c)
                if pd.isna(v) and "객단가" in met:
                    # Fallback
                    rev_m = next((m for m in metrics if "거래액" in m and "비중" not in m), None)
                    cus_m = next((m for m in metrics if "고객수" in m and "비중" not in m), None)
                    if rev_m and cus_m:
                        rev = pick(df, "주", rev_m, seg, yr, lb, c)
                        cus = pick(df, "주", cus_m, seg, yr, lb, c)
                        if pd.notna(rev) and pd.notna(cus) and cus > 0:
                            v = rev / cus
                return v

            for seg in segs:
                if not df_gran[df_gran["segment"] == seg].empty:
                    for met in metrics:
                        cur = _get_val(wy, wlabel, "mtd", seg, met)
                        prv = _get_val(py, plb, "final", seg, met) if plb else np.nan
                        yoy = _get_val(wy - 1, wlabel, "final", seg, met)
                        rows.append({
                            "지표": met,
                            cols[0]: fmt_value(met, prv), cols[1]: fmt_value(met, cur),
                            cols[2]: fmt_delta(met, cur, prv) or "-",
                            cols[3]: fmt_value(met, yoy), cols[4]: fmt_delta(met, cur, yoy) or "-",
                        })
            if rows:
                return pd.DataFrame(rows).set_index("지표")
            return pd.DataFrame(columns=["지표"] + cols).set_index("지표")
            
        def yoy_segment_table(ry, rm):
            df_gran = df[df["gran"] == "월"]
            metrics = df_gran[df_gran["segment"] == sel_seg]["metric"].dropna().unique().tolist()
            if not metrics:
                metrics = ["거래액", "고객수", "객단가", "CR"]
                
            pm_y, pm_m = (ry, rm - 1) if rm > 1 else (ry - 1, 12)
            cols = [f"{pm_y-1}년 {pm_m}월", f"{pm_y}년 {pm_m}월", "전년비 (전월)",
                    f"{ry-1}년 {rm}월", f"{ry}년 {rm}월", "전년비 (당월)"]
            rows = []
            
            def _get_val(yr, m, c, seg, met):
                lb = month_label(m)
                v = pick(df, "월", met, seg, yr, lb, c)
                if pd.isna(v) and "객단가" in met:
                    rev_m = next((m2 for m2 in metrics if "거래액" in m2 and "비중" not in m2), None)
                    cus_m = next((m2 for m2 in metrics if "고객수" in m2 and "비중" not in m2), None)
                    if rev_m and cus_m:
                        rev = pick(df, "월", rev_m, seg, yr, lb, c)
                        cus = pick(df, "월", cus_m, seg, yr, lb, c)
                        if pd.notna(rev) and pd.notna(cus) and cus > 0:
                            v = rev / cus
                return v
                
            for seg in segs:
                if not df_gran[df_gran["segment"] == seg].empty:
                    for met in metrics:
                        pm_prev = _get_val(pm_y - 1, pm_m, "final", seg, met)
                        pm_cur  = _get_val(pm_y, pm_m, "final", seg, met)
                        cm_prev = _get_val(ry - 1, rm, "mtd", seg, met)
                        cm_cur  = _get_val(ry, rm, "mtd", seg, met)
                        rows.append({
                            "지표": met,
                            cols[0]: fmt_value(met, pm_prev), cols[1]: fmt_value(met, pm_cur),
                            cols[2]: fmt_delta(met, pm_cur, pm_prev) or "-",
                            cols[3]: fmt_value(met, cm_prev), cols[4]: fmt_value(met, cm_cur),
                            cols[5]: fmt_delta(met, cm_cur, cm_prev) or "-",
                        })
            if rows:
                return pd.DataFrame(rows).set_index("지표"), (pm_y, pm_m)
            return pd.DataFrame(columns=["지표"] + cols).set_index("지표"), (pm_y, pm_m)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        wy, wlabel = week_ref(df, ref_year, ref_week)
        
        if wlabel:
            st.subheader(f"실적 요약 (전주비)")
            st.caption(f"기준 주차: {week_disp(wy, wlabel)} — 전주·전년 동주 대비")
            tbl1 = wow_segment_table(wy, wlabel)
            if not tbl1.empty:
                st.dataframe(style_delta_cols(tbl1), width="stretch")
            else:
                st.info("해당 주차 데이터가 없습니다.")
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
            
        st.subheader(f"실적 요약 (전년비)")
        tbl2, (pm_y, pm_m) = yoy_segment_table(ref_year, ref_month)
        st.caption(f"전월({pm_m}월) 마감 및 당월({ref_month}월·MTD) 기준 동일기간 비교")
        if not tbl2.empty:
            st.dataframe(style_delta_cols(tbl2), width="stretch")
        else:
            st.info("해당 월 데이터가 없습니다.")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        with st.expander("📊 지표 산출식 및 용어 설명 (클릭하여 펼치기)", expanded=False):
            st.markdown("""
            **산출식**
            * **거래액비중**: 첫구매 거래액 ÷ 전체 거래액
            * **고객비중**: 첫구매 고객수 ÷ 전체 고객수
            * **첫구매 객단가**: 첫구매 거래액 ÷ 첫구매 고객수
            * **유입율**: DAU ÷ 유효회원수
            * **CR (전환율)**: 첫구매 고객수 ÷ DAU
            
            **용어 설명**
            * **유효회원수**: 서비스에 정상적으로 가입되어 활동 가능한 전체 회원 수입니다.
            * **DAU (Daily Active Users)**: 하루 동안 서비스에 1회 이상 방문하여 활동한 사용자 수입니다.
            """)

if st.runtime.exists():
    main()
