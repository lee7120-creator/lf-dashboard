"""주간보고 — 첫구매 통합 실적 대시보드
여러 원천 엑셀(전체관점 마스터 + 지표별 파일)을 업로드하면 하나의 통합 뷰로 종합하고,
전년(YoY)·전주(WoW) 증감현황과 보고란을 갖춘 주간보고 화면을 만든다.
통합 결과는 (월/주/첫구매_요약 + 차트) 엑셀 워크북으로 다운로드할 수 있다.
"""

import datetime
import gzip, io, itertools, json, os, re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# table_export는 같은 폴더의 모듈이다. 테스트 하네스가 앱만 임시 폴더로 복사해 돌리는
# 경우가 있어 경로를 직접 얹는다 (없으면 엑셀 버튼만 빠지고 앱은 정상 동작).
import sys as _sys
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in _sys.path:
    _sys.path.insert(0, _HERE)
try:
    from table_export import xlsx_bytes
except Exception:                                         # noqa: BLE001
    xlsx_bytes = None

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

/* 표 첫 행(헤더) — 기본값이 옅은 회색 12px이라 잘 안 보인다. 표는 캔버스
   (glide-data-grid)로 그려져 보통 CSS가 안 닿으므로, 그 라이브러리가 읽는
   `--gdg-*` 변수로 바꾼다. 글자만 진하게·크게 하고 배경은 살짝 눌러 첫 행이
   본문과 구분되게 한다. 폰트 패밀리는 건드리지 않는다(Material 아이콘 보호). */
[data-testid="stDataFrame"], [data-testid="stDataFrameResizable"],
[data-testid="stDataEditor"], .stDataFrame {
  --gdg-text-header: #1e293b;
  --gdg-text-header-selected: #0f172a;
  --gdg-header-font-style: 600 13px;
  --gdg-bg-header: #eef2f7;
  --gdg-bg-header-hovered: #e2e8f0;
  --gdg-bg-header-has-focus: #e2e8f0;
  --gdg-border-color: #dbe3ec;
}
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
# **오래된 해부터** 이 순서로 색이 붙는다(`sorted(years)`) — 올해가 파랑, 전년이 회색.
# 색을 반대로 매기면 최신 흐름이 회색으로 눌려 한눈에 안 들어온다.
YEAR_PAL = ["slate", "blue", "red", "green", "purple", "amber", "teal"]
# 항목(조직·카테고리)별 선 색 — 연도가 아니라 항목을 색으로 가를 때 쓴다.
# 회색(slate)은 '전년'이 이미 쓰는 자리라 맨 뒤로 뺀다.
ORGCAT_PAL = ["blue", "amber", "green", "purple", "red", "teal", "orange", "slate"]

# ══════════════════════════════════════════════════════
# 지표 정의
# ══════════════════════════════════════════════════════
METRICS7 = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가",
            "비회원트래픽", "가입자수", "가입율", "당일가입CR"]
PCT_METRICS = {"가입율", "당일가입CR", "진입률", "CR", "거래액비중", "고객비중", "동의율", "유입율",
               "상품CR"}
# 추이 차트 6종 — 「01 요약」 하단과 「03 월별 추이」가 같은 목록을 본다. 결과(거래액·고객수·
# 객단가)만 보면 '왜 빠졌는지'가 안 갈려서 앞단(트래픽·가입자수·가입율)을 같이 세운다.
TREND_CHARTS = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가",
                "비회원트래픽", "가입자수", "가입율"]
# 마스터 파일 지표 → 보고서 지표 매핑
MASTER_MAP = {"일평균거래액": "첫구매 거래액", "일평균고객수": "첫구매 고객수",
              "일평균객단가": "첫구매 객단가"}
# 지표별 파일명 → 보고서 지표 매핑 (공백 제거 후 매칭)
METRIC_FILE_MAP = {
    "가입율": "가입율", "가입률": "가입율",
    "가입자수": "가입자수",
    "당일가입첫구매율": "당일가입CR", "당일가입CR": "당일가입CR",
    "비회원트래픽": "비회원트래픽",
    # 아직 안 올라오는 원천 — 「02. 첫구매 퍼널별 상세 실적」 하단이 이걸 기다린다.
    # 긴 이름부터 봐야 '신규앱설치'가 '앱설치'에 먼저 걸리지 않는다(둘 다 같은 지표지만
    # 순서가 곧 우선순위라 규칙을 눈에 보이게 둔다).
    "신규앱설치": "앱설치", "앱다운로드": "앱설치", "앱설치": "앱설치",
}

METRIC_UNIT = {
    "첫구매 거래액": ("백만원", 1e6), "첫구매 고객수": ("명", 1),
    "첫구매 객단가": ("원", 1), "비회원트래픽": ("명", 1),
    "가입자수": ("명", 1), "가입율": ("%", 1), "당일가입CR": ("%", 1),
    "앱설치": ("명", 1), "앱_전체설치": ("명", 1), "앱_재설치": ("명", 1),
    "앱_스토어방문": ("명", 1), "앱_삭제": ("명", 1),
    "앱_Push활성기기": ("명", 1), "앱_순증설치": ("명", 1),
    "앱푸시수신동의": ("명", 1), "앱푸시_동의자수": ("명", 1),
    "앱푸시동의_앱무관": ("명", 1), "앱푸시대상_앱무관": ("명", 1),
    "앱푸시_신규추가": ("명", 1), "앱푸시_이탈": ("명", 1),
    "앱푸시_유효회원": ("명", 1), "앱푸시_수신동의전체": ("명", 1),
    "상품UV": ("명", 1), "상품CR": ("%", 1), "거래액비중": ("%", 1), "고객비중": ("%", 1),
}

# 숫자·영문으로 끝나는 이름(e-영업1·SPACE-R)도 읽는 소리로 받침을 판단한다
_JOSA_TAIL = {**{d: t for d, t in zip("0123456789", [1, 1, 0, 1, 0, 0, 1, 1, 1, 0])},
              **{c: 1 for c in "LMNR"}, **{c.lower(): 1 for c in "LMNR"}}


def josa(word, pair="은는"):
    """받침 유무로 조사를 고른다 — '객단가이 가장'·'조직는' 같은 문장을 막는다."""
    w = str(word).strip()
    if not w:
        return pair[1]
    ch = w[-1]
    if "가" <= ch <= "힣":
        return pair[0] if (ord(ch) - 0xAC00) % 28 else pair[1]
    return pair[0] if _JOSA_TAIL.get(ch) else pair[1]


def esc(v):
    """업로드 파일에서 온 문자열을 unsafe_allow_html에 넣기 전에 이스케이프.
    조직·카테고리 이름에 &·< 가 섞이면 태그로 오해석된다."""
    return (str("" if v is None else v)
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;").replace("'", "&#39;"))


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
    """텍스트 export 디코딩. UTF-16을 **먼저 재 보면 안 된다** — 길이가 짝수인 UTF-8
    파일은 UTF-16으로도 예외 없이 풀려서, 줄바꿈이 사라진 한 줄짜리 쓰레기가 된다
    (헤더를 못 찾아 '미인식'으로만 보인다). BOM이나 NUL 바이트가 있을 때만 UTF-16이다."""
    head = data[:4096]
    if data[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in head:
        for enc in ("utf-16", "utf-16-le", "utf-16-be"):
            try: return data.decode(enc)
            except (UnicodeDecodeError, UnicodeError): pass
    for enc in ("utf-8-sig", "cp949", "utf-8"):
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

# 금액 칸에 붙어 오는 표기. **명시한 것만** 떼어 낸다 — 숫자 아닌 글자를 싹 지우면
# '3건'·'2025년' 같은 라벨이 조용히 숫자로 둔갑한다.
_CUR_MARKS = ("₩", "$", "€", "£", "¥", "KRW", "USD", "원")
_NUM_EMPTY = ("", "-", "–", "—", "nan", "None")


def _num(v):
    """셀 값 → float. 콤마·통화기호·회계식 음수 `(1,234)`·`%`를 읽는다(%는 비율로).

    수기 export라 같은 금액이 `153,000` · `₩153,000` · `153,000원` · `(63,818)`로 섞여 온다.
    못 읽으면 NaN이 되고 합계에서 그냥 빠져서, 증상이 '금액이 좀 적네'로만 보인다."""
    if v is None: return np.nan
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(",", "")
    if s.lower() in _NUM_EMPTY: return np.nan
    neg = s.startswith("(") and s.endswith(")")      # 회계식 음수
    if neg: s = s[1:-1].strip()
    for m in _CUR_MARKS:
        if s.startswith(m): s = s[len(m):].strip()
        if s.endswith(m): s = s[:-len(m)].strip()
    if s.lower() in _NUM_EMPTY: return np.nan
    pct = s.endswith("%")
    if pct: s = s[:-1].strip()
    try: f = float(s)
    except ValueError: return np.nan
    if neg: f = -f
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

# ══════════════════════════════════════════════════════
# 조직 × 카테고리 실적 (MICRO 대시보드 export)
# ══════════════════════════════════════════════════════
# 마스터(전체관점)와 달리 세그먼트 축이 **둘**이다 — 구분06=조직(BPU), 구분07=카테고리.
# 기존 store는 `segment` 한 칸뿐이라 여기에 밀어 넣으면 두 축이 뭉개진다. 그래서
# 별도 store(`wr_orgcat_store.csv`)에 org/cat 두 칸으로 쌓는다.
#
# LFMS 포함여부(Y/N)는 헤더 2행에 실려 오는 **다른 모집단**이다. 키에 넣지 않으면
# 같은 기간·같은 조직 값이 서로를 조용히 덮어쓴다 — 화면에서 필터로 고르게 둔다.
ORGCAT_STORE = "wr_orgcat_store.csv"
# 실제 뎁스는 **영업조직 > 카테고리 > 브랜드 > 상품** 네 단계고, export의 구분06~09가
# 차례로 대응한다. 브랜드·상품은 아직 값이 안 들어와 전부 '-'로 오지만, 컬럼은 미리
# 잡아 둔다 — 값이 들어오는 날 파서·화면이 알아서 그 단계를 열게.
# 컬럼명 `prod`는 DataFrame.prod(곱) 메서드와 겹쳐 `d.prod`가 조용히 메서드를 준다 — item으로.
ORGCAT_LEVELS = [("org", "조직", "구분06"), ("cat", "카테고리", "구분07"),
                 ("brand", "브랜드", "구분08"), ("item", "상품", "구분09")]
ORGCAT_LV = [c for c, _, _ in ORGCAT_LEVELS]
# **채널은 레벨이 아니라 «축»이다.** 새 export는 구분08에 채널(직접·광고·EP…)을 싣고 오는데,
# 이걸 브랜드 레벨로 받으면 조직>카테고리>«직접» 같은 뎁스가 생겨 파고들기가 뭉개진다.
# LFMS처럼 **키에 넣고 화면에서 고르는 축**으로 둔다 — `*TOTAL`이 '전 채널'이라,
# 채널 칸이 없던 옛 백업도 그대로 전 채널 값으로 읽힌다.
ORGCAT_CH_ALL = "*TOTAL"
ORGCAT_KEY = (["gran", "metric"] + ORGCAT_LV
              + ["ch", "lfms", "year", "label", "close"])
ORGCAT_COLS = ORGCAT_KEY + ["sortkey", "value"]
# 마스터 파일과 같은 지표는 같은 이름으로 — fmt_value·PCT_METRICS를 그대로 태운다
ORGCAT_MAP = {"일평균거래액": "첫구매 거래액", "일평균고객수": "첫구매 고객수",
              "일평균객단가": "첫구매 객단가"}
ORGCAT_TOTAL = "*TOTAL"
# 그 레벨이 아예 없음(원본 '-')은 빈 문자열로 둔다. `*TOTAL`(그 레벨의 합계)과 다르다.
ORGCAT_NONE = ""


def orgcat_fill(d):
    """저장소·백업에서 읽은 프레임을 현재 스키마로 맞춘다.

    레벨이 2단(org·cat)뿐이던 시절의 CSV도 그대로 복원돼야 한다 — 없는 레벨은 빈 칸으로
    채운다. CSV는 빈 칸을 NaN으로 돌려주므로 **문자열로 되돌리는 게 필수**다.
    안 그러면 키 비교가 전부 어긋나 병합이 조용히 중복을 만든다.
    """
    d = d.copy()
    for col in ORGCAT_LV:
        d[col] = (d[col] if col in d.columns else ORGCAT_NONE)
        d[col] = d[col].fillna(ORGCAT_NONE).astype(str).replace(
            {"nan": ORGCAT_NONE, "-": ORGCAT_NONE, "–": ORGCAT_NONE})
    # **채널 칸이 없던 백업은 «전 채널»이다.** 빈 칸으로 두면 ⑤가 `*TOTAL`만 보는 순간
    # 옛 데이터가 통째로 안 잡혀 화면이 빈다 — 증상이 '데이터가 없다'로만 보인다.
    d["ch"] = (d["ch"] if "ch" in d.columns else ORGCAT_CH_ALL)
    d["ch"] = d["ch"].fillna(ORGCAT_CH_ALL).astype(str).replace(
        {"nan": ORGCAT_CH_ALL, "": ORGCAT_CH_ALL, "-": ORGCAT_CH_ALL,
         "–": ORGCAT_CH_ALL})
    return d[[c for c in ORGCAT_COLS if c in d.columns]]


def is_orgcat_grid(rows):
    """구분06·구분07 헤더가 있으면 조직×카테고리 export."""
    for r in rows[:8]:
        cells = [_cell(c) for c in r]
        if "구분06" in cells and "구분07" in cells:
            return True
    return False


def parse_orgcat_grid(rows):
    """조직×카테고리(구분06~09) 그리드 → long DataFrame.

    헤더 구성이 단위마다 다르다. 월·주는 `연도 / LFMS / 기간 / (구분+마감)` 4행이고,
    일은 마감 행이 없어 `연도 / LFMS / (구분+기간)` 3행이다. 그래서 행 번호를 박지 않고
    내용으로 찾는다 — 구분06이 있는 행, 연도가 있는 행, Y/N만 있는 행, 기간 라벨 행.
    """
    hdr, lv_at = None, {}
    for ri, r in enumerate(rows[:8]):
        cells = [_cell(c) for c in r]
        if "구분06" in cells and "구분07" in cells:
            hdr = ri
            lv_at = {col: cells.index(h) for col, _lb, h in ORGCAT_LEVELS if h in cells}
            break
    if hdr is None:
        return pd.DataFrame()
    ncols = max(len(r) for r in rows)

    def cell(ri, ci):
        return _cell(rows[ri][ci]) if ri is not None and ci < len(rows[ri]) else ""

    year_row = lfms_row = period_row = close_row = None
    for ri in range(min(hdr + 1, len(rows))):
        cells = [_cell(c) for c in rows[ri]]
        nz = [c for c in cells if c]
        nY = sum(1 for c in cells if YEAR_RE.match(c))
        nP = sum(1 for c in cells if PERIOD_RE.match(c))
        nC = sum(1 for c in cells if "마감" in c)
        if year_row is None and nY >= 1 and nP == 0: year_row = ri
        if lfms_row is None and nz and all(c in ("Y", "N") for c in nz): lfms_row = ri
        if period_row is None and nP >= 2: period_row = ri
        if close_row is None and nC >= 2: close_row = ri
    if period_row is None:
        return pd.DataFrame()
    data_start = max(r for r in (hdr, year_row, period_row, close_row)
                     if r is not None) + 1

    plabels = [_cell(c) for c in rows[period_row] if PERIOD_RE.match(_cell(c))]
    if any("주차" in p for p in plabels):  gran = "주"
    elif any("/" in p for p in plabels):   gran = "일"
    else:                                   gran = "월"

    # 연도·LFMS는 병합셀이라 왼쪽 값을 오른쪽으로 이어받는다 (한 파일에 2개년이 온다)
    col_year, col_lfms, cur_y, cur_l = {}, {}, None, ""
    for ci in range(ncols):
        m = YEAR_RE.match(cell(year_row, ci)) if year_row is not None else None
        if m: cur_y = int(m.group(1))
        v = cell(lfms_row, ci) if lfms_row is not None else ""
        if v in ("Y", "N"): cur_l = v
        col_year[ci], col_lfms[ci] = cur_y, cur_l

    # 기간 라벨이 빈 '일마감'(MTD) 칼럼은 직전 라벨을 이어받는다 (마스터 파서와 동일)
    data_cols, col_label, last_lbl = [], {}, None
    for ci in range(ncols):
        lbl = cell(period_row, ci)
        if PERIOD_RE.match(lbl) and col_year[ci]:
            last_lbl = lbl; col_label[ci] = lbl; data_cols.append(ci)
        elif (close_row is not None and "마감" in cell(close_row, ci)
              and last_lbl and col_year[ci]):
            col_label[ci] = last_lbl; data_cols.append(ci)

    lv_cols = [c for c in ORGCAT_LV if c in lv_at]
    # **구분08이 브랜드가 아니라 «채널»로 오는 export가 있다.** 값만 보고 가른다 —
    # 파일명·시트명은 둘 다 「전체관점 …」이라 근거가 안 된다. 그대로 브랜드로 받으면
    # 조직>카테고리>«직접» 뎁스가 생겨 파고들기가 뭉개지고, 화면은 멀쩡히 떠서
    # 눈으로는 안 잡힌다.
    ch_col = None
    if "brand" in lv_at:
        _seen = set()
        for _ri in range(data_start, len(rows)):
            _v = cell(_ri, lv_at["brand"])
            if _v and _v not in ("-", "–", ORGCAT_TOTAL):
                _seen.add(_v)
        if _seen and _seen <= set(CHANNELS):
            ch_col = lv_at["brand"]
            lv_cols = [c for c in lv_cols if c not in ("brand", "item")]

    records, metric = [], None
    cur = {c: ORGCAT_NONE for c in ORGCAT_LV}
    cur_ch = ORGCAT_CH_ALL
    for ri in range(data_start, len(rows)):
        m0 = cell(ri, 0)
        if m0 and m0 not in ("-", "–"):
            metric = ORGCAT_MAP.get(m0, m0)
        # 상위 레벨은 병합셀이라 아래로 이어받는다. 어떤 레벨에 값이 새로 찍히면
        # 그 아래 레벨은 초기화해야 한다 — 안 그러면 앞 조직의 카테고리가 따라붙는다.
        for i, col in enumerate(lv_cols):
            v = cell(ri, lv_at[col])
            if v in ("-", "–"):
                cur[col] = ORGCAT_NONE
                for deeper in lv_cols[i + 1:]:
                    cur[deeper] = ORGCAT_NONE
            elif v:
                cur[col] = v
                for deeper in lv_cols[i + 1:]:
                    cur[deeper] = ORGCAT_NONE
        # 채널 칸도 병합셀이라 아래로 이어받는다. 상위(카테고리)가 새로 찍히면 초기화 —
        # 안 그러면 앞 카테고리의 마지막 채널이 따라붙어 값이 통째로 어긋난다.
        if ch_col is not None:
            _cv = cell(ri, ch_col)
            if cell(ri, lv_at["cat"]):        # 카테고리가 새로 찍힌 줄
                cur_ch = ORGCAT_CH_ALL
            if _cv in ("-", "–"):
                cur_ch = ORGCAT_CH_ALL
            elif _cv:
                cur_ch = _cv
        # 카테고리가 없는 행(원본 '-')은 조직 합계와 값이 겹치는 자리표시라 버린다
        if not metric or not cur["org"] or not cur["cat"]:
            continue
        for ci in data_cols:
            pp = period_parts(gran, col_year[ci], col_label[ci])
            if pp is None: continue
            label, sortkey = pp
            close = cell(close_row, ci) if close_row is not None else ""
            rec = {"gran": gran, "metric": metric}
            rec.update({c: cur[c] for c in ORGCAT_LV})
            rec.update({
                "ch": cur_ch,
                "lfms": col_lfms[ci] or "N", "year": col_year[ci],
                "label": label, "sortkey": sortkey,
                "close": "mtd" if "일마감" in close and gran != "일" else "final",
                "value": _num(rows[ri][ci] if ci < len(rows[ri]) else None)})
            records.append(rec)
    return pd.DataFrame(records, columns=ORGCAT_COLS) if records else pd.DataFrame()


@st.cache_data(show_spinner=False, max_entries=4)
def parse_orgcat_file(name, data: bytes) -> pd.DataFrame:
    """업로드 1건 → 조직×카테고리 long DF. 이 형식이 아니면 빈 DF.

    라우팅·인식목록·누적 병합 세 군데서 같은 파일을 물어보므로 캐시해 둔다
    (일별 파일이 250칼럼짜리라 매번 다시 읽으면 업로드가 눈에 띄게 느려진다)."""
    # 백업 CSV 재업로드는 그대로 복원
    if name.lower().endswith(".csv"):
        head = data[:400].decode("utf-8", "ignore").splitlines()
        if head and all(k in head[0] for k in ("org", "cat", "lfms")):
            try:
                d = pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
                if _is_orgcat_csv(d.columns):
                    return orgcat_fill(d)
            except Exception:
                pass
        return pd.DataFrame()
    if not name.lower().endswith((".xlsx", ".xls")):
        return pd.DataFrame()
    try:
        rows = read_grid(name, data)
    except Exception:
        return pd.DataFrame()
    if not rows or not is_orgcat_grid(rows):
        return pd.DataFrame()
    return parse_orgcat_grid(rows)


# ══════════════════════════════════════════════════════
# 브랜드·상품 상세 (결제 원장 export)
# ══════════════════════════════════════════════════════
# MICRO export는 조직×카테고리까지고 구분08·09(브랜드·상품)가 늘 비어 온다. 브랜드·상품은
# **결제 원장**으로 따로 온다 — 한 줄이 (결제일자 × 조직 × 유입채널 × 카테고리 × 브랜드 × 상품)
# 하나다.
#
# 두 원천을 한 store에 섞으면 안 된다:
#  · **카테고리 어휘가 다르다** — MICRO 구분07은 골프·남성·슈즈·잡화, 원장은 가방·지갑·
#    패션소품·남성의류·향수… 훨씬 잘다. 같은 칸에 넣으면 축이 조용히 뭉개진다.
#  · 원장엔 **유입채널(AF대분류명)** 축이 하나 더 있고 MICRO엔 없다.
#  · 원장 값은 그날의 **합계**, MICRO 값은 **일평균**이다.
# 그래서 따로 쌓고, 화면에서 「데이터 소스」로 갈라 본다.
DETAIL_STORE = "wr_detail_store.csv.gz"
DETAIL_LV = ["org", "cat", "brand", "item"]
DETAIL_KEY = ["date"] + DETAIL_LV + ["ch"]
DETAIL_COLS = ["date", "org", "ch", "cat", "brand", "item", "rev", "cust"]
DETAIL_GRANS = ["월", "주", "일"]
DETAIL_ALL = "전체"
# 원장엔 상품UV·상품CR이 없다. 대신 **거래액 = 고객수 × 객단가**가 정의상 정확히 성립한다.
DETAIL_FACTORS = [("첫구매 고객수", "고객수"), ("첫구매 객단가", "객단가")]
DETAIL_METS = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]
DETAIL_NA = "(미상)"
# 헤더 표기가 흔들려도 찾도록 **부분 일치**로 건다 (조각이 전부 들어 있어야 매칭).
# 한 칼럼이 두 자리를 겸하지 않게 이미 잡힌 열은 건너뛴다 — 순서가 곧 우선순위다.
DETAIL_HDR = [("date", ("결제", "일자")), ("org", ("BPU",)), ("ch", ("AF",)),
              ("cat", ("카테고리",)), ("brand", ("브랜드",)), ("item", ("상품명",)),
              ("rev", ("거래액",)), ("cust", ("고객수",))]
_DATE8_RE = re.compile(r"^(20\d{2})\D?(\d{1,2})\D?(\d{1,2})$")


def _detail_header(cells):
    """헤더 행이면 {저장칼럼: 열 인덱스}, 아니면 None."""
    at = {}
    for key, frags in DETAIL_HDR:
        for ci, c in enumerate(cells):
            if ci in at.values():
                continue
            if all(f in c for f in frags):
                at[key] = ci
                break
    return at if len(at) == len(DETAIL_HDR) else None


def _detail_rows(name, data: bytes):
    """탭만 보는 read_grid의 보완 — 원장은 콤마 CSV로도 온다."""
    if name.lower().endswith((".csv", ".txt", ".tsv")):
        text = _decode_text(data)
        lines = text.splitlines()
        if not lines:
            return []
        if "\t" in lines[0]:
            return [ln.split("\t") for ln in lines]
        import csv as _csv
        return list(_csv.reader(io.StringIO(text)))
    return read_grid(name, data)


def is_detail_grid(rows):
    for r in rows[:8]:
        if _detail_header([_cell(c) for c in r]):
            return True
    return False


# 엑셀 날짜 일련번호(1900 날짜체계, 1900년 윤년 버그 때문에 기준이 1899-12-30)의 실용 범위.
# YYYYMMDD(2000_0101 이상)와 겹치지 않아 서로 헷갈릴 일이 없다.
_XL_EPOCH = datetime.date(1899, 12, 30)
_XL_MIN, _XL_MAX = 20000, 60000          # 1954-10 ~ 2064-03
# 맨 아래 요약행은 원래 날짜가 없다 — 이것까지 경고에 세면 매 업로드마다 ⚠가 떠서
# 정작 진짜 문제(날짜 형식을 못 읽음)를 무시하게 된다.
_SUMMARY_ROW = ("합계", "총계", "소계", "계", "total", "sum", "전체")


def _detail_date(v):
    """20250101 · 2025-01-01 · 엑셀 날짜셀 · 엑셀 일련번호 → date. 못 읽으면 None.

    **일련번호(45658 같은 숫자)를 빠뜨리면 안 된다.** xlsx의 날짜 칸이 날짜서식이 아니라
    숫자로 오는 export가 있는데, 그러면 `openpyxl`이 숫자를 그대로 돌려줘서 원장 전 행이
    통째로 버려진다. 증상이 '원장이 안 올라가요'로만 보여 원인이 안 드러난다."""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    m = _DATE8_RE.match(s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    try:                                  # 엑셀 일련번호
        n = int(float(s))
    except (TypeError, ValueError):
        return None
    if _XL_MIN <= n <= _XL_MAX:
        return _XL_EPOCH + datetime.timedelta(days=n)
    return None


def parse_detail_grid(rows):
    """결제 원장 그리드 → (date, org, ch, cat, brand, item, rev, cust) long DF."""
    at = None
    hdr = 0
    for ri, r in enumerate(rows[:8]):
        a = _detail_header([_cell(c) for c in r])
        if a:
            hdr, at = ri, a
            break
    if at is None:
        return pd.DataFrame(columns=DETAIL_COLS)
    recs, dropped = [], []
    for r in rows[hdr + 1:]:
        def g(k):
            ci = at[k]
            return _cell(r[ci]) if ci < len(r) else ""

        def raw(k):
            ci = at[k]
            return r[ci] if ci < len(r) else None

        d = _detail_date(raw("date"))
        if d is None:
            # 합계·머리말이면 정상이지만, 못 읽는 날짜 형식이면 매출이 통째로 사라진다.
            # 금액이 붙어 있는 줄만 센다 — 그런 줄이 버려졌으면 알려야 한다.
            _lab = _cell(raw("date"))[:20]
            if (_num(raw("rev")) == _num(raw("rev"))
                    and _lab.strip().lower() not in _SUMMARY_ROW):
                dropped.append(_lab)
            continue
        # 빈 칸은 자기 노드로 남긴다 — 지워 버리면 그 매출이 조용히 사라지고,
        # 상위 합계에 흡수시키면 하위 합이 상위와 안 맞는다.
        recs.append({"date": d.isoformat(),
                     "org": g("org") or DETAIL_NA, "ch": g("ch") or DETAIL_NA,
                     "cat": g("cat") or DETAIL_NA, "brand": g("brand") or DETAIL_NA,
                     "item": g("item") or DETAIL_NA,
                     "rev": _num(raw("rev")), "cust": _num(raw("cust"))})
    if not recs:
        out = pd.DataFrame(columns=DETAIL_COLS)
        out.attrs["date_dropped"] = dropped
        return out
    d = pd.DataFrame(recs, columns=DETAIL_COLS)
    # 같은 키가 여러 줄로 오면(상품코드 색상 분리 등) 합친다 — 상품코드는 안 쌓는다.
    # 화면 뎁스가 상품'명'까지라 코드는 축이 아니고, 원장 크기만 15%쯤 키운다.
    out = (d.groupby(DETAIL_KEY, as_index=False, sort=False)[["rev", "cust"]]
           .sum()[DETAIL_COLS])
    out.attrs["date_dropped"] = dropped
    return out


def detail_fill(d):
    """저장소·백업에서 읽은 원장을 현재 스키마·dtype으로 맞춘다.

    CSV 라운드트립은 숫자를 문자열로, 빈 칸을 NaN으로 돌려준다 — 되돌리지 않으면
    키 비교가 어긋나 병합이 조용히 중복을 만들고 합계가 문자열 이어붙이기가 된다."""
    d = d.copy()
    for c in ["date"] + DETAIL_LV + ["ch"]:
        d[c] = (d[c] if c in d.columns else "")
        d[c] = d[c].fillna("").astype(str).replace({"nan": ""})
    for c in ("rev", "cust"):
        d[c] = pd.to_numeric(d.get(c), errors="coerce")
    return d[DETAIL_COLS]


@st.cache_data(show_spinner=False, max_entries=4)
def parse_detail_file(name, data: bytes) -> pd.DataFrame:
    """업로드 1건 → 결제 원장 long DF. 이 형식이 아니면 빈 DF."""
    if name.lower().endswith(".csv") and not name.lower().endswith(".csv.gz"):
        head = data[:400].decode("utf-8", "ignore").splitlines()
        if head and all(k in head[0] for k in ("date", "brand", "cust")):
            try:
                d = pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
                if _is_detail_csv(d.columns):
                    return detail_fill(d)
            except Exception:
                pass
    if not name.lower().endswith((".xlsx", ".xls", ".csv", ".txt", ".tsv")):
        return pd.DataFrame(columns=DETAIL_COLS)
    try:
        rows = _detail_rows(name, data)
    except Exception:
        return pd.DataFrame(columns=DETAIL_COLS)
    if not rows or not is_detail_grid(rows):
        return pd.DataFrame(columns=DETAIL_COLS)
    return parse_detail_grid(rows)


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
# 잔고(수준) 지표 — 하루짜리 급변 글리치 규칙 대상
PUSH_LEVEL_METRICS = {"앱푸시_동의자수", "앱푸시_유효회원", "앱푸시_수신동의전체"}

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
    m = d["metric"].isin(PUSH_LEVEL_METRICS)
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
    if not rows or len(rows) < 6: return pd.DataFrame()

    # 1) 날짜 헤더 행 탐색 — 파일마다 'Date' 라벨 유무·행 위치가 달라(0행 또는 1행)
    #    M/D 형태 셀이 가장 많은 행을 날짜 행으로 본다 (앞 6행 중)
    def _date_cols_of(row):
        out = []
        for ci in range(2, len(row)):
            m = re.match(r"^(\d{1,2})/\d{1,2}$", _cell(row[ci]))
            if m: out.append((ci, _cell(row[ci]), int(m.group(1))))
        return out
    date_cols, date_ri = [], None
    for ri in range(min(6, len(rows))):
        cand = _date_cols_of(rows[ri])
        if len(cand) > len(date_cols):
            date_cols, date_ri = cand, ri
    if len(date_cols) < 3:
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
    # 행 라벨은 공백 유무가 파일마다 달라(수신동의/수신 동의) 공백 제거 후 매칭
    _rk = lambda s: _cell(s).replace(" ", "")
    labels = {_rk(r[1]) for r in rows if len(r) > 1}
    # 3단 표(전체 유효 회원/수신 동의/타겟팅 가능)인지 판별.
    # 이 표에서 '수신 동의'는 앱 수신동의 전체(수백만)이고, 실제 발송 가능 모수는
    # '타겟팅 가능'이다. 구 2단 표의 '수신동의'는 이 '타겟팅 가능'과 동일한 수치이므로
    # (실측 대조 확인) 두 레이아웃 모두 앱푸시_동의자수로 이어붙여 시계열 연속성을 유지한다.
    three_tier = "타겟팅가능" in labels
    ROW_METRIC = {"신규추가(+)": "앱푸시_신규추가", "기존이탈(-)": "앱푸시_이탈",
                  "전체유효회원": "앱푸시_유효회원", "타겟팅가능": "앱푸시_동의자수"}
    ROW_METRIC["수신동의"] = "앱푸시_수신동의전체" if three_tier else "앱푸시_동의자수"
    targets = []  # (metric, segment, row)
    cur_seg = None
    for row in rows:
        c0 = _cell(row[0] if len(row) > 0 else "")
        c1 = _rk(row[1] if len(row) > 1 else "")
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

# ══════════════════════════════════════════════════════
# 앱 보유와 무관한 푸시 수신동의 (Push Total / Y / N)
# ══════════════════════════════════════════════════════
# 기존 PUSH 원천은 **앱을 가진 회원**의 수신동의라, '앱을 안 깐 신규회원까지 포함하면
# 동의율이 얼마인가'에 답하지 못한다. 이 원천이 그 자리를 메운다.
#
#   (빈칸) 1/1  1/2  1/3 …      ← 1행: 날짜, 연도가 없다
#   (빈칸) (수) (목) (금) …     ← 2행: 요일
#   Push Total  1451 1626 …     ← 그날 신규회원 전체
#   Push Y Cnt   660  732 …     ← 그중 수신동의
#   Push N Cnt   791  894 …     ← 미동의 (= Total - Y, 실파일 616일 전부 정확히 성립)
#
# **연도는 기존 PUSH와 같은 길로 푼다** — 상대연도로 내보내고 `combine_files`가 마스터
# 일자와 겹침이 가장 큰 기준연도를 고른다. 여기서 따로 추론하면 규칙이 둘로 갈린다.
# `Push N Cnt`는 안 쌓는다 — Total-Y로 언제든 나오고, 축이 아니라 파생이다.
PUSHALL_ROWS = {"pushtotal": "앱푸시대상_앱무관", "pushycnt": "앱푸시동의_앱무관"}


def parse_pushall_file(name, data: bytes) -> pd.DataFrame:
    """앱 보유 무관 푸시 수신동의 원천 → long DF (연도는 상대값)."""
    try:
        rows = read_grid(name, data)
    except Exception:                                     # noqa: BLE001
        return pd.DataFrame()
    if not rows or len(rows) < 3:
        return pd.DataFrame()

    _k = lambda v: re.sub(r"[^a-z]", "", _cell(v).lower())
    have = {_k(r[0]) for r in rows if r}
    # 세 줄이 다 있어야 이 원천이다 — 한 줄만 보고 잡으면 기존 PUSH 표를 가로챈다
    if not ({"pushtotal", "pushycnt", "pushncnt"} <= have):
        return pd.DataFrame()

    # 날짜 행 — M/D 셀이 가장 많은 행 (앞 4행 중)
    def _dcols(row):
        return [(ci, _cell(row[ci])) for ci in range(1, len(row))
                if re.match(r"^\d{1,2}/\d{1,2}$", _cell(row[ci]))]
    date_cols, best = [], None
    for ri in range(min(4, len(rows))):
        cand = _dcols(rows[ri])
        if len(cand) > len(date_cols):
            date_cols, best = cand, ri
    if len(date_cols) < 3:
        return pd.DataFrame()

    # 뒤에서 앞으로 읽으며 월이 커지면 한 해 전 (기존 PUSH 파서와 같은 규칙)
    rel, cur_rel = {}, 0
    last_m = int(date_cols[-1][1].split("/")[0])
    for ci, d in reversed(date_cols):
        mo = int(d.split("/")[0])
        if last_m < 6 and mo > 6:
            cur_rel -= 1
        elif last_m > 6 and mo < 6:
            cur_rel += 1
        rel[ci] = cur_rel
        last_m = mo

    recs = []
    for row in rows:
        met = PUSHALL_ROWS.get(_k(row[0] if row else ""))
        if not met:
            continue
        for ci, d in date_cols:
            v = _num(row[ci] if ci < len(row) else None)
            if pd.isna(v):
                continue
            recs.append({"gran": "일", "metric": met, "segment": "*TOTAL",
                         "year": rel[ci], "label": d, "sortkey": 0,
                         "close": "final", "value": v})
    return pd.DataFrame(recs)


def looks_like_push_name(name: str) -> bool:
    """파일명이 앱푸시 원천으로 보이는가 (한글명 포함)"""
    up = name.upper()
    return "PUSH" in up or any(h in name for h in ("앱푸시", "푸시", "수신동의"))

# ══════════════════════════════════════════════════════
# 앱설치 원천 (LFmall 앱 대시보드 export — Daily · Weekly · Monthly)
# ══════════════════════════════════════════════════════
# 「02. 첫구매 퍼널별 상세 실적」 하단 앱 블록이 기다리던 원천이다. 같은 대시보드에서
# 세 벌이 따로 떨어지는데, **일별 하나만 쌓는다.** 나머지 둘은 인식만 하고 버린다.
#
#   · Daily   `9월 5일,710,394,316,1080,,547529`  ← 연도가 없다
#   · Weekly  `~ 26.09.05,5337,...`               ← 주가 **일~토**다
#   · Monthly `2026. 08,19509,0,19509`            ← 칼럼이 다르다(전체/Android/iOS)
#
# **주간 파일을 같이 쌓으면 안 된다.** 이 export의 주는 일~토인데 보고서의 주는 월~일이라
# 창이 하루씩 어긋난다. 그런데 목요일로 라벨을 매기면 **같은 라벨**(`08월 4주차`)이 나와서,
# 저장하면 일별에서 만든 값을 조용히 덮어쓴다. 실파일로 대조해 보면 주간 파일의
# `~26.08.29`(8/23~8/29 합 22,788)와 보고서의 08월 4주차(8/24~8/30 합 22,382)가 다른 창이다.
#
# **월간 파일도 안 쌓는다.** 커버리지가 일별보다 좁고(6개월 vs 189일), 실파일에서
# **최신 달이 깨져 온다** — 2026-08이 파일 19,509(Android 0)인데 일별 합은 92,696이다.
# 그 값이 일별에서 만든 월 값을 덮어쓰면 그 달이 통째로 틀어진다.
#
# 값은 전부 **일평균**으로 낸다 — 같은 화면의 가입자수·앱푸시 수신동의가 일평균이라
# 합계로 섞으면 비율이 통째로 틀어진다. 주·월로 묶는 규칙은 원장과 같은
# `detail_periods()`를 쓴다(주차 = 목요일이 속한 달의 몇 번째 주).
# **`앱설치`는 「신규 설치」다.** 이 화면이 묻는 건 '신규 가입자가 앱까지 오나'라,
# 재설치까지 섞인 「전체 설치」로는 답이 안 나온다. 전체·재설치는 상세로 내려 둔다.
APPINSTALL_MAP = [("신규설치", "앱설치"), ("전체설치", "앱_전체설치"),
                  ("재설치", "앱_재설치"), ("스토어방문", "앱_스토어방문"),
                  ("삭제", "앱_삭제"), ("push활성기기", "앱_Push활성기기")]
APPINSTALL_METS = [m for _h, m in APPINSTALL_MAP]
# 이 둘이 한 행에 같이 있어야 일별·주간 export로 본다 — 다른 원천과 겹치지 않는 조합이다
APPINSTALL_NEED = {"전체설치", "신규설치"}
# 월간 export는 칼럼이 통째로 다르다 (전체 설치 (전체)/(Android)/(iOS))
APPINSTALL_MON_NEED = {"전체설치(전체)", "전체설치(android)"}

# 수집이 끊긴 날 판정 — `Push 활성 기기`는 그날의 **잔고**라 하루 만에 반토막 났다가
# 다음 날 돌아올 수 없다. 실파일에서 8/31·9/2~9/5가 112만 → 54.7만으로 떨어지는데
# (월간 파일의 8월 Android 0과 같은 증상), 한쪽 플랫폼이 빠진 export다.
# **고치지 않고 표시만 한다** — 어느 날이 그런지 말해 주면 사람이 판단할 수 있고,
# 지어내는 것보다 낫다. 기준은 **적재 전체의 중앙값**이다. 6개월 동안 111만~112만으로
# 1%밖에 안 움직이는 잔고라 안정적이고, 이상한 날이 며칠 더 쌓여도 흔들리지 않는다.
APPINSTALL_LEVEL_DROP = 0.7
APPINSTALL_FLAG = "앱_기기수이상일"

_AI_DAY_RE = re.compile(r"^(\d{1,2})\s*월\s*(\d{1,2})\s*일$")        # 9월 5일
_AI_WEEK_RE = re.compile(r"^~?\s*(\d{2})\.\s*(\d{1,2})\.\s*(\d{1,2})\.?$")   # ~ 26.09.05
_AI_MON_RE = re.compile(r"^(20\d{2})\s*\.\s*(\d{1,2})\.?$")          # 2026. 08


def _ai_key(v):
    """헤더 비교용 정규화 — `Push 활성 기기` → `push활성기기`."""
    return _cell(v).replace(" ", "").replace("_", "").lower()


def appinstall_kind(rows):
    """이 그리드가 앱설치 export면 '일'·'주'·'월', 아니면 None.

    단위는 헤더가 아니라 **날짜 칸 모양**이 정한다 — 일별과 주간은 헤더가 똑같다.
    """
    hdr = None
    for ri, r in enumerate(rows[:8]):
        keys = {_ai_key(c) for c in r}
        if APPINSTALL_MON_NEED <= keys:
            return "월", ri
        if APPINSTALL_NEED <= keys:
            hdr = ri
            break
    if hdr is None:
        return None, None
    for r in rows[hdr + 1:]:
        d = _cell(r[0] if len(r) else "")
        if not d:
            continue
        if _AI_DAY_RE.match(d):
            return "일", hdr
        if _AI_WEEK_RE.match(d):
            return "주", hdr
        if _AI_MON_RE.match(d):
            return "월", hdr
        # 헤더는 앱설치인데 날짜 칸 모양을 모른다. None으로 흘리면 '미인식'으로만 보여
        # 원인이 안 드러나니, 그 사실을 들고 나가 인식 목록에서 말하게 한다.
        return "?", hdr
    return None, None


def is_appinstall_grid(rows):
    return appinstall_kind(rows)[0] is not None


def _ai_days(labels, today=None):
    """`9월 5일` 목록 → 날짜 목록. 연도가 없어서 **오늘을 기준으로 거꾸로 채운다.**

    export가 늘 '최근 N일'이라 가장 최근 행이 오늘 이전의 그 월·일이다. 거기서부터
    거슬러 올라가며 달이 커지면(=한 해 넘어감) 연도를 하나 뺀다. 읽은 범위는 인식 목록에
    찍어 두니, 옛 파일을 올려 연도가 어긋나면 눈으로 바로 잡힌다.
    """
    md = []
    for s in labels:
        m = _AI_DAY_RE.match(str(s).strip())
        md.append((int(m.group(1)), int(m.group(2))) if m else None)
    real = [x for x in md if x]
    if not real:
        return [None] * len(md)
    # 파일이 내림차순인지 오름차순인지 — 이웃끼리 비교해서 많은 쪽으로 정한다.
    # (연말을 넘는 한 쌍은 반대로 세지만 소수라 결론이 안 바뀐다)
    desc = sum(1 for a, b in zip(real, real[1:]) if b < a) >= len(real) / 2
    order = list(range(len(md))) if desc else list(range(len(md)))[::-1]
    today = today or today_kst()
    out, year, prev = [None] * len(md), None, None
    for i in order:
        if md[i] is None:
            continue
        mo, dd = md[i]
        if year is None:                       # 가장 최근 행 — 오늘 이전이 되게 연도를 고른다
            year = today.year if (mo, dd) <= (today.month, today.day) else today.year - 1
        elif mo > prev:                        # 거슬러 올라가다 달이 커졌다 = 해를 넘었다
            year -= 1
        prev = mo
        try:
            out[i] = datetime.date(year, mo, dd)
        except ValueError:                     # 2/29 — 그 해에 없는 날
            out[i] = None
    return out


def parse_appinstall_grid(rows):
    """앱설치 그리드 → 일·주·월 long DF (값은 일평균). 일별 파일만 값을 낸다."""
    kind, hdr = appinstall_kind(rows)
    empty = pd.DataFrame(columns=STORE_COLS)
    if kind is None:
        return empty
    empty.attrs["appinstall_kind"] = kind
    if kind != "일":
        # 주간·월간은 인식만 한다. 왜 안 쌓는지는 위 주석과 인식 목록에 적혀 있다.
        # 날짜 모양을 모르는 파일('?')은 어떤 값이 들어 있었는지 같이 들고 나간다.
        if kind == "?":
            empty.attrs["date_sample"] = list(dict.fromkeys(
                _cell(r[0]) for r in rows[hdr + 1:] if len(r) and _cell(r[0])))[:3]
        return empty
    hkeys = [_ai_key(c) for c in rows[hdr]]
    col = {}
    for h, met in APPINSTALL_MAP:        # 앞 항목이 우선 — 한 칸이 두 지표를 겸하지 않게
        if h in hkeys:
            col.setdefault(hkeys.index(h), met)
    body = [r for r in rows[hdr + 1:] if any(_cell(c) for c in r)]
    if not body or not col:
        return empty
    days = _ai_days([_cell(r[0]) if len(r) else "" for r in body])
    recs, dropped = [], 0
    for r, d in zip(body, days):
        if d is None:
            dropped += 1
            continue
        rec = {"date": d.isoformat()}
        for ci, met in col.items():
            rec[met] = _num(r[ci]) if ci < len(r) else np.nan
        recs.append(rec)
    if not recs:
        empty.attrs["date_dropped"] = dropped
        return empty
    base = pd.DataFrame(recs).drop_duplicates(subset=["date"], keep="last")
    mets = [m for m in col.values() if base[m].notna().any()]
    # 수집이 끊긴 날 표시 — 값은 그대로 두고 '며칠이 그랬는지'만 따로 쌓는다
    drop_days = []
    if "앱_Push활성기기" in base:
        _lv = pd.to_numeric(base["앱_Push활성기기"], errors="coerce")
        _med = _lv.median()
        if _med == _med and _med > 0:
            _bad = _lv < _med * APPINSTALL_LEVEL_DROP
            base[APPINSTALL_FLAG] = _bad.astype(float)
            drop_days = sorted(base.loc[_bad.fillna(False), "date"])
    frames = []
    for gran in ("일", "주", "월"):
        per = detail_periods(base["date"], gran)
        if per.empty or not mets:
            continue
        m = base.merge(per[["date", "year", "label", "sortkey", "close"]], on="date")
        # 빈 칸은 NaN으로 두고 평균에서 빠지게 한다 — 0으로 채우면 없는 날을 '0건'으로
        # 읽어 일평균이 내려간다 (실파일에서 8/31 이후 삭제 칸이 통째로 비어 온다).
        _keys = ["year", "label", "sortkey", "close"]
        g = m.groupby(_keys, as_index=False)[mets].mean()
        _vals = list(mets)
        if APPINSTALL_FLAG in m:
            # 이건 평균이 아니라 **개수**다 — 그 기간에 며칠이 그랬는지를 세야 한다
            g = g.merge(m.groupby(_keys, as_index=False)[[APPINSTALL_FLAG]].sum(), on=_keys)
            _vals.append(APPINSTALL_FLAG)
        long = g.melt(id_vars=_keys, value_vars=_vals,
                      var_name="metric", value_name="value")
        # 이상 없는 기간까지 0으로 쌓지는 않는다 (없으면 없는 것)
        long = long[(long["metric"] != APPINSTALL_FLAG) | (long["value"] > 0)]
        long["gran"], long["segment"] = gran, "*TOTAL"
        frames.append(long.dropna(subset=["value"]))
    out = (pd.concat(frames, ignore_index=True)[STORE_COLS] if frames else empty.copy())
    out.attrs["appinstall_kind"] = "일"
    out.attrs["date_dropped"] = dropped
    out.attrs["date_range"] = (base["date"].min(), base["date"].max())
    out.attrs["level_drop"] = drop_days
    return out


@st.cache_data(show_spinner=False, max_entries=4)
def parse_appinstall_file(name, data: bytes) -> pd.DataFrame:
    """라우팅·인식목록·누적 병합 세 군데서 같은 파일을 물어보므로 캐시해 둔다."""
    try:
        rows = _detail_rows(name, data)      # 원장처럼 콤마 CSV로 온다
    except Exception:                        # noqa: BLE001
        return pd.DataFrame(columns=STORE_COLS)
    if not rows:
        return pd.DataFrame(columns=STORE_COLS)
    return parse_appinstall_grid(rows)


def route_push(n, b):
    """엑셀 1건 → (push_df 또는 None, 일반_df 또는 None).
    이름 힌트가 있으면 PUSH 우선 시도, 없으면 일반 파싱 후 빈 결과면 내용 기반으로 PUSH 재시도.
    → 파일명이 'PUSH'가 아니어도(예: 앱푸시수신동의현황.xlsx) 인식된다.
    조직×카테고리 export·브랜드 상품 원장은 각각 combine_orgcat·combine_detail이
    따로 받으므로 여기선 둘 다 None."""
    if not parse_orgcat_file(n, b).empty:
        return None, None
    if not parse_detail_file(n, b).empty:
        return None, None
    ai = parse_appinstall_file(n, b)      # 앱설치는 마스터와 같은 store에 쌓인다
    if not ai.empty:
        return None, ai
    if ai.attrs.get("appinstall_kind"):   # 앱설치인데 안 쌓는 단위 — 마스터로 흘리지 않는다
        return None, None
    # 파일명이 PUSH라 기존 파서가 먼저 집어 가려 한다 — 세 줄(Total/Y/N)이 다 있는
    # 이 원천만 여기서 먼저 가른다. 상대연도라 push 프레임으로 내보낸다.
    pa = parse_pushall_file(n, b)
    if not pa.empty:
        return pa, None
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

@st.cache_data(show_spinner=False, max_entries=2)
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
    """누적 마스터. **파일 서명으로 캐시한다** — 안 하면 리런마다 CSV를 다시 판다.

    저장하면 mtime이 바뀌어 캐시가 저절로 갈린다. 원장(`load_detail_store`)과 같은 규칙.
    """
    if not os.path.exists(DATA_STORE):
        return pd.DataFrame()
    stt = os.stat(DATA_STORE)
    return _read_store(DATA_STORE, stt.st_mtime_ns, stt.st_size)


@st.cache_data(show_spinner=False, max_entries=1)
def _read_store(path, mtime, size):
    try:
        d = pd.read_csv(path, encoding="utf-8-sig")
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

def _is_orgcat_csv(cols):
    """조직×카테고리 저장소 CSV인가 — 브랜드·상품 레벨이 없던 옛 백업도 인정한다."""
    need = {"gran", "metric", "org", "cat", "lfms", "year", "label", "close", "value"}
    return need <= set(cols)


def load_orgcat_store() -> pd.DataFrame:
    """누적 조직×카테고리. 실파일이 70만 행이라 **파일 서명으로 캐시한다**.

    캐시가 없으면 CSV 파싱 0.4초 + `orgcat_fill`의 문자열 변환 0.26초를 **어느 페이지를
    보든 매 리런** 낸다 — 조직×카테고리를 안 쓰는 페이지까지 같이 느려진다.
    """
    if not os.path.exists(ORGCAT_STORE):
        return pd.DataFrame(columns=ORGCAT_COLS)
    stt = os.stat(ORGCAT_STORE)
    return _read_orgcat_store(ORGCAT_STORE, stt.st_mtime_ns, stt.st_size)


@st.cache_data(show_spinner=False, max_entries=1)
def _read_orgcat_store(path, mtime, size):
    try:
        d = pd.read_csv(path, encoding="utf-8-sig")
        if _is_orgcat_csv(d.columns):
            return orgcat_fill(d)
    except Exception:
        pass
    return pd.DataFrame(columns=ORGCAT_COLS)


def save_orgcat_store(df: pd.DataFrame):
    orgcat_fill(df)[ORGCAT_COLS].to_csv(ORGCAT_STORE, index=False, encoding="utf-8-sig")


def merge_orgcat(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """조직×카테고리 누적 병합 — 같은 (단위·지표·조직·카테고리·LFMS·기간) 키는 신규 우선"""
    if old is None or old.empty: return new if new is not None else pd.DataFrame(columns=ORGCAT_COLS)
    if new is None or new.empty: return old
    return (pd.concat([old[ORGCAT_COLS], new[ORGCAT_COLS]], ignore_index=True)
            .drop_duplicates(subset=ORGCAT_KEY, keep="last"))


@st.cache_data(show_spinner=False, max_entries=2)
def combine_orgcat(file_tuples) -> pd.DataFrame:
    """업로드 파일들 중 조직×카테고리 형식만 모아 하나로 — 같은 키는 마지막 파일 우선"""
    frames = []
    for n, b in file_tuples:
        d = parse_orgcat_file(n, b)
        if not d.empty:
            frames.append(d)
    if not frames:
        return pd.DataFrame(columns=ORGCAT_COLS)
    return (pd.concat(frames, ignore_index=True)
            .drop_duplicates(subset=ORGCAT_KEY, keep="last"))



def _is_detail_csv(cols):
    """결제 원장 저장소 CSV인가. 조직×카테고리 CSV와는 date·ch·rev·cust로 갈린다."""
    return set(DETAIL_COLS) <= set(cols)


def load_detail_store() -> pd.DataFrame:
    """누적 원장. 행이 수십만을 넘길 수 있어 파일 서명으로 캐시한다 (매 리런 재파싱 방지)."""
    if not os.path.exists(DETAIL_STORE):
        return pd.DataFrame(columns=DETAIL_COLS)
    stt = os.stat(DETAIL_STORE)
    return _read_detail_store(DETAIL_STORE, stt.st_mtime_ns, stt.st_size)


@st.cache_data(show_spinner=False, max_entries=1)
def _read_detail_store(path, mtime, size):
    try:
        d = pd.read_csv(path, encoding="utf-8-sig")   # .gz는 pandas가 알아서 푼다
        if _is_detail_csv(d.columns):
            return detail_fill(d)
    except Exception:
        pass
    return pd.DataFrame(columns=DETAIL_COLS)


def save_detail_store(df: pd.DataFrame):
    # 확장자가 .gz라 pandas가 알아서 압축한다 — 원장은 행이 많아 그대로 쌓으면 수백 MB다
    detail_fill(df).to_csv(DETAIL_STORE, index=False, encoding="utf-8-sig")


def merge_detail(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """원장 누적 병합 — 같은 (일자·조직·채널·카테고리·브랜드·상품) 키는 신규 우선"""
    if old is None or old.empty:
        return new if new is not None else pd.DataFrame(columns=DETAIL_COLS)
    if new is None or new.empty:
        return old
    return (pd.concat([old[DETAIL_COLS], new[DETAIL_COLS]], ignore_index=True)
            .drop_duplicates(subset=DETAIL_KEY, keep="last"))


@st.cache_data(show_spinner=False, max_entries=2)
def combine_detail(file_tuples) -> pd.DataFrame:
    """업로드 파일들 중 결제 원장 형식만 모아 하나로 — 같은 키는 마지막 파일 우선"""
    frames = [d for d in (parse_detail_file(n, b) for n, b in file_tuples) if not d.empty]
    if not frames:
        return pd.DataFrame(columns=DETAIL_COLS)
    return (pd.concat(frames, ignore_index=True)
            .drop_duplicates(subset=DETAIL_KEY, keep="last"))


def _period_set(d, cols):
    if d is None or d.empty: return set()
    return set(map(tuple, d[cols].drop_duplicates().values.tolist()))

def upload_diff(stored, df_new):
    """업로드 데이터가 기존 누적 대비 추가/갱신하는 기간 수 (저장 전 미리보기용)"""
    cols = ["gran", "year", "label"]
    nc, oc = _period_set(df_new, cols), _period_set(stored, cols)
    return len(nc - oc), len(nc & oc)

# ── 백업/업로드 고도화 헬퍼 ────────────────────────────────
@st.cache_data(show_spinner=False, max_entries=2)
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
        oc = parse_orgcat_file(n, b)
        if not oc.empty:
            _g = "".join(g for g in ("일", "주", "월") if g in set(oc["gran"]))
            _lf = "/".join(sorted(set(oc["lfms"].astype(str))))
            _lv = " > ".join(lb for c, lb, _h in ORGCAT_LEVELS
                             if (oc[c].astype(str) != ORGCAT_NONE).any())
            out.append((n, f"✅ 조직×카테고리 · {_g} · LFMS {_lf} · {_lv}", len(oc)))
            continue
        dt = parse_detail_file(n, b)
        _drop = dt.attrs.get("date_dropped") or []
        if not dt.empty:
            _d0, _d1 = str(dt["date"].min())[:10], str(dt["date"].max())[:10]
            _warn = (f" · ⚠ 날짜를 못 읽어 뺀 줄 {len(_drop):,}건"
                     f"({', '.join(dict.fromkeys(_drop))[:40]})" if _drop else "")
            out.append((n, f"✅ 브랜드·상품 원장 · {_d0}~{_d1} · "
                           f"브랜드 {dt['brand'].nunique():,} · 상품 {dt['item'].nunique():,}"
                           + _warn, len(dt)))
            continue
        if _drop:
            # 헤더는 원장인데 한 줄도 못 읽었다 — 날짜 형식이 원인이다.
            out.append((n, f"❌ 브랜드·상품 원장인데 날짜를 하나도 못 읽었어요 "
                           f"({len(_drop):,}줄, 예: {', '.join(dict.fromkeys(_drop))[:30]})", 0))
            continue
        ai = parse_appinstall_file(n, b)
        _aikind = ai.attrs.get("appinstall_kind")
        _aidrop = int(ai.attrs.get("date_dropped") or 0)
        if not ai.empty:
            _d0, _d1 = ai.attrs.get("date_range", ("?", "?"))
            # 지표마다 결측이 달라(삭제 칸이 며칠 비어 온다) 행 수를 나누면 하루씩 어긋난다
            _nd = ai[ai["gran"] == "일"]["label"].nunique()
            _warn = f" · ⚠ 날짜를 못 읽어 뺀 줄 {_aidrop:,}건" if _aidrop else ""
            _bad = ai.attrs.get("level_drop") or []
            if _bad:
                _warn += (f" · ⚠ 기기 수가 절반 이하인 날 {len(_bad)}일"
                          f"({', '.join(str(x)[5:] for x in _bad[:4])}"
                          + (" 외" if len(_bad) > 4 else "") + ") — 한쪽 플랫폼 누락 의심")
            out.append((n, f"✅ 앱설치 원천(일별) · {_d0}~{_d1} · {_nd:,}일 · "
                           f"{ai['metric'].nunique() - (1 if _bad else 0)}지표" + _warn, len(ai)))
            continue
        if _aikind in ("주", "월"):
            # 인식은 했지만 안 쌓는다. 왜인지 말해 주지 않으면 '올렸는데 왜 안 보이지'가 된다.
            _why = ("주 경계가 일~토라 보고서의 월~일과 어긋나요"
                    if _aikind == "주" else "커버리지가 좁고 최신 달이 깨져 와요")
            out.append((n, f"➖ 앱설치({_aikind}간) — 저장 안 함 ({_why}). "
                           "일별 파일이 주·월을 만들어요", 0))
            continue
        if _aikind:
            # 헤더는 앱설치인데 값이 안 나왔다 — 날짜 칸이 원인이다.
            _ex = ", ".join(ai.attrs.get("date_sample") or [])
            out.append((n, "❌ 앱설치 원천인데 날짜 칸을 못 읽었어요"
                           + (f" (예: {_ex[:40]})" if _ex else f" ({_aidrop:,}줄)"), 0))
            continue
        pf, d = route_push(n, b)  # combine_files와 동일한 라우팅 (한글명 PUSH 포함)
        if pf is not None:
            # 앱 보유 무관 원천은 기존 PUSH와 답하는 질문이 달라 따로 찍는다 —
            # 같은 문구로 두면 어느 쪽을 올렸는지 확인할 방법이 없다
            if set(pf["metric"]) <= set(PUSHALL_ROWS.values()):
                _d = pf["label"].nunique()
                out.append((n, f"✅ 앱푸시 수신동의(앱 보유 무관) · {_d:,}일", len(pf)))
            else:
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

def make_backup_zip(df, texts, odf=None, ddf=None) -> bytes:
    """누적 데이터 CSV + 조직×카테고리 CSV + 브랜드·상품 원장 CSV + 메모 JSON + manifest."""
    import zipfile
    buf = io.BytesIO()
    ts = datetime.datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")
    yrs = (f"{int(df['year'].min())}–{int(df['year'].max())}년"
           if not df.empty else "-")
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("wr_data_store.csv",
                    df[STORE_COLS].to_csv(index=False).encode("utf-8-sig"))
        if odf is not None and not odf.empty:
            zf.writestr("wr_orgcat_store.csv",
                        odf[ORGCAT_COLS].to_csv(index=False).encode("utf-8-sig"))
        if ddf is not None and not ddf.empty:
            # ZIP이 알아서 압축하니 안에는 평문 CSV로 넣는다 (복원 쪽이 확장자로 가른다)
            zf.writestr("wr_detail_store.csv",
                        ddf[DETAIL_COLS].to_csv(index=False).encode("utf-8-sig"))
        zf.writestr("wr_insights.json",
                    json.dumps(texts, ensure_ascii=False, indent=2).encode("utf-8"))
        zf.writestr("manifest.txt",
                    (f"LF Mall 주간보고 통합 백업\n생성: {ts}\n"
                     f"누적 데이터: {len(df):,}행 · {df['metric'].nunique() if not df.empty else 0}지표 · {yrs}\n"
                     f"조직×카테고리: {0 if odf is None else len(odf):,}행\n"
                     f"브랜드·상품 원장: {0 if ddf is None else len(ddf):,}행\n"
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
        # gzip으로 받은 원장(.csv.gz)도 그대로 올릴 수 있게. 버퍼로 주면 pandas가
        # 확장자를 못 보고 compression을 추론하지 못해서 매직바이트로 직접 가른다.
        if b[:2] == b"\x1f\x8b":
            import gzip as _gz
            b = _gz.decompress(b)
        d = pd.read_csv(io.BytesIO(b), encoding="utf-8-sig")
        # 원장이 먼저 — date·ch·rev·cust가 있어 다른 둘과 헷갈릴 일은 없다
        if _is_detail_csv(d.columns):
            save_detail_store(detail_fill(d)); data_restored = True
            msgs.append(f"브랜드·상품 원장 {len(d):,}행")
            return True
        # 조직×카테고리 백업이 그다음 — segment가 없어 STORE_COLS와 헷갈릴 일은 없다
        if _is_orgcat_csv(d.columns) and "segment" not in d.columns:
            save_orgcat_store(orgcat_fill(d)); data_restored = True
            msgs.append(f"조직×카테고리 {len(d):,}행")
            return True
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
                if base.endswith((".csv", ".csv.gz")): _try_csv(b)
                elif base.endswith(".json"): _try_json(b)
    elif name.endswith((".csv", ".csv.gz")):
        if not _try_csv(raw):
            raise ValueError("누적 데이터 CSV 형식이 아니에요. gran·metric·sortkey 같은 필수 컬럼이 없어요.")
    elif name.endswith(".json"):
        if not _try_json(raw):
            raise ValueError("메모 JSON 형식이 아니에요. 최상위가 객체여야 해요.")
    else:
        raise ValueError("지원 형식: .zip / .csv / .csv.gz / .json")

    if not msgs:
        raise ValueError("복원할 내용을 찾지 못했어요. ZIP 안에 백업 파일이 없어요.")
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

def lazy_download(key, label, filename, mime, sig=None, **kw):
    """누를 때 만들어서 받는 다운로드 버튼.

    `st.download_button`은 data를 **미리** 받아야 해서, 큰 프레임의 직렬화 결과를 그대로
    넘기면 백업을 안 받는 리런에서도 매번 CSV를 다시 찍고 그 바이트를 브라우저로 보낸다.
    실측 120만 행 원장에서 통합 ZIP 5.6초 + 개별 CSV 3.4초(122MB)라, 페이지를 바꾸거나
    필터를 만질 때마다 9초씩 그냥 버리고 있었다.

    그래서 두 단계로 나눈다 — 평범한 버튼을 먼저 보여 주고, 누르면 그때 만들어 진짜
    다운로드 버튼으로 바꾼다. 만든 바이트는 `sig`가 같은 동안만 세션에 들고 있다가
    데이터가 바뀌면 버린다(옛 백업을 받게 두면 안 된다).

    반환값은 `make()`를 받아 실제 버튼을 그리는 함수, 아직 안 눌렀으면 None."""
    flag = f"_lazydl_{key}"
    cur = st.session_state.get(flag)
    if not (isinstance(cur, tuple) and cur[0] == sig):
        st.session_state.pop(flag, None)
        if st.button(label, key=f"{flag}_go", width="stretch", **kw):
            st.session_state[flag] = (sig, None)
            st.rerun()
        return None

    def _render(make):
        _s, data = st.session_state[flag]
        if data is None:
            try:
                with st.spinner("백업 파일을 만드는 중이에요…"):
                    data = make()
            except Exception as e:                            # noqa: BLE001
                st.session_state.pop(flag, None)
                st.error(f"백업 파일을 만들지 못했어요 — {e}")
                return
            st.session_state[flag] = (_s, data)
        st.download_button(f"{label} · 받기 ({len(data) / 1e6:,.1f}MB)", data, filename, mime,
                           key=f"{flag}_dl", width="stretch", **kw)
        if st.button("✕ 닫기", key=f"{flag}_x", width="stretch",
                     help="만들어 둔 백업 파일을 메모리에서 비워요."):
            st.session_state.pop(flag, None); st.rerun()

    return _render


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
            st.error("인식된 데이터가 없어요. 파일명과 형식을 확인해 주세요.")


# ── 표 렌더 + 엑셀 내려받기 ────────────────────────────────────────────
# Streamlit 기본 툴바의 'Download as CSV'는 서식이 다 날아간다. 발송성과 대시보드와
# 같은 검은 헤더 스타일 xlsx를 같은 자리에서 받을 수 있게 감싼다.
# 바이트는 콜러블로 넘겨 '누를 때만' 만든다.
_WTBL_SEQ = itertools.count(1)


def wtable(data, *args, dl=True, dl_name=None, dl_data=None, **kw):
    """st.dataframe 과 같게 쓰되, 아래에 엑셀 다운로드 버튼을 붙인다.

    `dl_data`를 주면 **내려받기만** 그 객체로 만든다. Styler 배경색은 엑셀까지
    따라가므로(`table_export`), 선택 하이라이트 같은 화면 전용 서식을 파일에
    싣지 않으려면 색을 안 입힌 쪽을 여기로 넘긴다.
    """
    ev = st.dataframe(data, *args, **kw)
    if dl and xlsx_bytes is not None:
        _i = next(_WTBL_SEQ)
        _nm = dl_name or "표"
        _fn = re.sub(r"[^\w가-힣.\- ]", "", str(_nm)).strip() or "표"
        # set_index로 축을 세운 표가 많아 인덱스도 같이 내보낸다(기본 RangeIndex면 제외)
        _src = data if dl_data is None else dl_data
        _df = getattr(_src, "data", _src)
        _idx = (not kw.get("hide_index", False)) and not isinstance(
            getattr(_df, "index", None), pd.RangeIndex)
        try:
            st.download_button(
                "⬇️ 엑셀", key=f"_wxl{_i}",
                data=lambda d=_src, n=_nm, x=_idx: xlsx_bytes(d, sheet_name=n, title=n, index=x),
                file_name=f"{_fn}_{today_kst():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="화면에 보이는 서식 그대로 받아요. 숫자는 엑셀 숫자로 들어가서 정렬·합계가 돼요.")
        except Exception:                                 # noqa: BLE001
            pass                                          # 다운로드가 안 되더라도 표는 보여야 한다
    return ev


# ══════════════════════════════════════════════════════
# 조회 헬퍼 — mtd(당월/당주 일마감) vs final 선택
# ══════════════════════════════════════════════════════
# ── 기준 기간은 **한 칸일 수도, 여러 칸일 수도** 있다 ──────────────────
# 일자별은 하루가 너무 성겨서(퍼널이 통째로 비는 날이 흔하다) 조회 기간으로 고른다.
# 주·월은 지금대로 한 칸이다. 아래 두 헬퍼가 그 차이를 한 군데서 흡수한다 — 소비처마다
# `isinstance(list)`를 흩뿌리면 한 곳을 빠뜨렸을 때 그 화면만 조용히 옛 동작으로 남는다.
def as_labels(label):
    """기준 기간을 **라벨 튜플**로. 한 칸이면 길이 1."""
    if isinstance(label, (list, tuple, set, pd.Index, np.ndarray)):
        return tuple(str(x) for x in label)
    return (str(label),)


def is_range(label):
    """기준 기간이 여러 칸인가 — 값 규칙이 갈리는 자리에서 본다."""
    return len(as_labels(label)) > 1


def lbl_disp(label):
    """화면에 찍을 기준 기간 — 여러 칸이면 `8/23~9/12`."""
    ls = as_labels(label)
    return ls[0] if len(ls) == 1 else f"{ls[0]}~{ls[-1]}"


# ── `pick` 조회 인덱스 ────────────────────────────────────────────────
# 한 페이지 렌더에 `pick`이 200번 가까이 불리는데, 예전엔 그때마다 프레임 전체에
# 마스크를 다섯 번 씌우고 `label`을 통째로 문자열로 바꿨다(실측 184회 0.35초).
# **값 규칙은 그대로 두고 조회만** 딕셔너리로 바꾼다(184회 0.0002초, 구축 0.03초).
#
# 메모는 한 칸이고 **프레임 객체 자체를 들고 `is`로 비교한다.** `id()`만 키로 쓰면
# 옛 프레임이 해제된 자리에 새 프레임이 앉을 때 옛 인덱스를 그대로 주고, 업로드가
# 행 수를 안 바꾸는 갱신이면 길이로도 못 가른다. 참조를 들고 있으면 그 자리가
# 재사용될 수 없어서 애초에 생기지 않는 문제다.
_PICK_MEMO = {"df": None, "map": None}
_PICK_NEED = ("gran", "metric", "segment", "year", "label", "close", "value")


def _pick_map(df):
    """`(gran, metric, segment, year, label, close)` → 그 칸의 **마지막 유효값**.

    뒤 행이 앞 행을 덮는 게 곧 예전 `sub[...]["value"].dropna().iloc[-1]`이다.
    """
    if _PICK_MEMO["df"] is not df:
        m = {}
        if all(c in getattr(df, "columns", ()) for c in _PICK_NEED):
            d = df.dropna(subset=["value"])
            for k, v in zip(zip(d["gran"], d["metric"], d["segment"], d["year"],
                                d["label"].astype(str), d["close"]), d["value"]):
                m[k] = v
        _PICK_MEMO["df"], _PICK_MEMO["map"] = df, m
    return _PICK_MEMO["map"]


def pick(df, gran, metric, seg, year, label, prefer="final"):
    """한 칸의 값. `label`이 여러 개면 그 기간의 **일평균**(칸별 값의 평균)이다.

    이 앱의 값은 전부 일평균이라, 3주를 골라도 하루를 골라도 **단위가 같아야** 나란히
    읽힌다. 합으로 묶으면 3주가 하루의 21배로 찍혀 전년비도 카드도 통째로 뒤집힌다.
    """
    labs = as_labels(label)
    m = _pick_map(df)
    order = ("final", "mtd") if prefer == "final" else ("mtd", "final")
    # 칸마다 close 우선순위로 **하나씩만** 고른다 — 같은 라벨이 final·mtd 둘 다 있을 때
    # 둘 다 세면 그 날만 두 번 반영된다.
    vals = []
    for lb in labs:
        for c in order:
            v = m.get((gran, metric, seg, year, lb, c))
            if v is not None:
                vals.append(v)
                break
    if not vals:
        return np.nan
    return vals[0] if len(labs) == 1 else float(np.mean(vals))

def series_by_label(df, gran, metric, seg, year, prefer="final"):
    """한 연도의 기간라벨 → 값 Series (sortkey 순)"""
    sub = df[(df["gran"] == gran) & (df["metric"] == metric) &
             (df["segment"] == seg) & (df["year"] == year)].copy()
    if sub.empty: return pd.Series(dtype=float)
    pref = {"final": 0, "mtd": 1} if prefer == "final" else {"mtd": 0, "final": 1}
    sub["_p"] = sub["close"].map(pref)
    sub = sub.sort_values(["sortkey", "_p"]).drop_duplicates("label", keep="first")
    return sub.set_index("label")["value"]

# ── 보고서가 읽는 '한 칸의 값' 규칙 ────────────────────────────────────
# **파일 값을 그대로 읽지 않는 칸이 있다.** 가입율은 `가입자수 ÷ 비회원트래픽`으로
# 계산한다 — 파일에도 「가입율」이 오지만 그건 **일별 비율의 평균**이라 트래픽이 적은
# 날이 과대 대표돼 값이 다르다. 이 앱의 값은 전부 일평균이라 두 일평균을 나누면 곧
# 기간 합계끼리의 비율이 된다(「6. 유입 퍼널」의 단계 비율과 같은 이유).
#
# 예전엔 「01 KPI·02 퍼널」만 계산값을 쓰고 「01 하단·03·04 추이」는 파일값을 써서
# **같은 가입율이 페이지마다 다른 숫자**로 떴다(합성 데이터에서 1.00% vs 50.00%).
# 규칙은 여기 한 곳이고, 낱개 조회(`funnel_val`)와 기간 조회(`report_series`)가
# 같은 표를 본다 — 한쪽만 고치면 또 갈린다.
# 비율 칸을 두 카운트로 만들 때의 (분자, 분모)
FUNNEL_DERIVED = {"가입율": ("가입자수", "비회원트래픽"),
                  "첫구매 객단가": ("첫구매 거래액", "첫구매 고객수")}
# 파일 값을 **먼저** 쓰는 칸 — 객단가는 「01」 KPI 카드와 같은 숫자여야 한다.
# 가입율은 반대로 계산이 먼저다(그래야 퍼널 앞 두 칸이 딱 맞는다).
FUNNEL_FILE_FIRST = {"첫구매 객단가"}


def funnel_val(g, met, seg="*TOTAL", file_first=True):
    """한 칸의 값. `g(metric, segment)`가 원값을 돌려주는 함수다.

    비율 칸은 되도록 **두 카운트에서 만든다** — 채널로 쪼갤 때 합계와 같은 규칙으로
    나와야 해서다. 당일가입CR만은 분자(당일가입 첫구매 고객수)가 데이터에 없어 늘 파일
    값이고, 채널별 파일이 없으면 그 칸은 빈다.

    **여러 날을 묶어 볼 땐 `file_first=False`다.** 객단가를 파일 값으로 읽으면 날마다의
    객단가를 그냥 평균 내게 되는데, 그러면 거래가 적은 날이 많은 날과 같은 무게를 갖는다
    (가입율을 계산으로 되돌린 것과 같은 이유). 두 카운트에서 만들면 거래액 합 ÷ 고객수 합이
    되어 무게가 맞는다. 한 칸일 땐 둘이 같은 값이라 예전 동작 그대로다.
    """
    if file_first and met in FUNNEL_FILE_FIRST:
        v = g(met, seg)
        if not pd.isna(v):
            return v
    if met in FUNNEL_DERIVED:
        num, den = FUNNEL_DERIVED[met]
        a, b = g(num, seg), g(den, seg)
        if not pd.isna(a) and not pd.isna(b) and b > 0:
            return a / b
    return g(met, seg)


def report_val(df, gran, metric, seg, year, label, prefer="final"):
    """한 칸의 값 — **보고서 규칙**으로 읽는다(`funnel_val`을 기간 조회에 얹은 껍데기).

    `pick`을 그대로 쓰면 파일의 가입율(일별 비율의 평균)이 나와 ①·② 카드와 갈린다.

    기준 기간이 여러 칸이면 파일 우선 규칙을 끈다 — 위 `funnel_val` 주석 참고.
    """
    # **`seg`를 `funnel_val`에도 넘겨야 한다.** 람다에 기본인자로 묶어 둬도 `funnel_val`이
    # `g(met, seg)`로 부르면서 자기 기본값 `*TOTAL`로 덮어써, 채널을 뭘 넣든 전체 값이
    # 나왔다. 증상은 표의 **모든 채널 줄이 전체와 같은 숫자**로 뜨는 것 — 화면은 멀쩡히
    # 떠서 눈으로 세어 보지 않으면 안 드러난다.
    return funnel_val(lambda m, s=seg: pick(df, gran, m, s, year, label, prefer),
                      metric, seg, file_first=not is_range(label))


def report_series(df, gran, metric, seg, year, prefer="final"):
    """`series_by_label`의 **보고서 규칙 판** — 추이표·차트가 이걸 쓴다.

    `funnel_val`과 칸별로 결과가 같아야 한다. 낱개로 훑으면 기간 수만큼 프레임을
    필터하게 되므로(주차 50개 × 지표 6개 × 2개년) 시리즈 단위로 한 번에 만든다.
    """
    file_s = series_by_label(df, gran, metric, seg, year, prefer)
    if metric not in FUNNEL_DERIVED:
        return file_s
    num, den = FUNNEL_DERIVED[metric]
    a = series_by_label(df, gran, num, seg, year, prefer)
    b = series_by_label(df, gran, den, seg, year, prefer)
    idx = file_s.index.union(a.index).union(b.index)
    a, b, f = a.reindex(idx), b.reindex(idx), file_s.reindex(idx)
    der = a / b.where(b > 0)                              # 분모 0·결측은 NaN
    # 파일 우선 칸은 파일이 빈 자리만 계산으로 메우고, 나머지는 그 반대다.
    return f.fillna(der) if metric in FUNNEL_FILE_FIRST else der.fillna(f)


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

def week_like(df, year, month, wk):
    """(year, 'MM월 N주차') 라벨 찾기 — 같은 주차 번호가 없으면 그 달의 '마지막 주차'로 대체.

    주차는 달마다 4~5개로 들쭉날쭉해서(전체 342주 중 31주가 5주차) 07월 5주차의
    전월인 06월엔 5주차가 아예 없다. 그대로 두면 전월비가 '–'로 비어버리므로
    가장 가까운 대응 주차(= 그 달 마지막 주)로 떨어뜨린다.
    반환: (label, exact) — exact=False면 대체된 것이라 화면에 기준을 같이 표기해야 한다.
          label이 None이면 그 달 데이터 자체가 없음.
    """
    labs = set(df[(df["gran"] == "주") & (df["year"] == year)]["label"].dropna().unique())
    exact = f"{month:02d}월 {wk}주차"
    if exact in labs:
        return exact, True

    def _wknum(s):
        """라벨에서 주차 번호. 'N주차'가 없으면 -1 (후보에서 제외).

        파서는 주간 라벨을 'MM월 N주차'로 강제하지만, 백업 CSV 복원은 컬럼만 보고
        내용은 검증하지 않는다. 형식이 어긋난 행이 하나 섞이면 여기서 None.group()으로
        터져 앱 전체가 초기 렌더부터 죽으므로 방어한다.
        """
        m = re.search(r"(\d)주차", s)
        return int(m.group(1)) if m else -1

    cand = [l for l in labs if l.startswith(f"{month:02d}월") and _wknum(l) > 0]
    if not cand:
        return None, False
    return max(cand, key=_wknum), False

def growth_pace_note(s_cur, s_prev=None):
    """잔고 시계열의 증가속도 분석 문구(HTML). s_cur/s_prev: dt 인덱스 Series(당해/전년).
    연초 대비 증감·일평균 속도, 전년 동기 속도 대비, 최근 4주 가속/감속, 연말 단순추정."""
    s_cur = s_cur.dropna().sort_index()
    if len(s_cur) < 8:
        return None
    as_of = s_cur.index[-1]
    span = max((as_of - s_cur.index[0]).days, 1)
    ytd = float(s_cur.iloc[-1] - s_cur.iloc[0])
    pace = ytd / span

    def _c(v, txt):
        return f'<span style="color:{"#dc2626" if v < 0 else "#16a34a"};font-weight:700">{txt}</span>'
    bits = [f"연초 대비 {_c(ytd, f'{ytd:+,.0f}명')} · 일평균 {_c(pace, f'{pace:+,.1f}명/일')}"]

    # 전년 동기(같은 연중 위치)까지의 속도와 비교
    if s_prev is not None:
        sp = s_prev.dropna().sort_index()
        if len(sp) >= 8:
            py = int(sp.index[0].year)
            try:
                cutoff = as_of.replace(year=py)
            except ValueError:          # 2/29 → 평년
                cutoff = as_of.replace(year=py, day=28)
            spc = sp[sp.index <= cutoff]
            if len(spc) >= 8:
                p_pace = (float(spc.iloc[-1] - spc.iloc[0])
                          / max((spc.index[-1] - spc.index[0]).days, 1))
                diff = pace - p_pace
                bits.append(f"전년 동기 {p_pace:+,.1f}명/일 대비 "
                            f"{_c(diff, f'{diff:+,.1f}명/일')} {'빠름' if diff > 0 else '느림'}")

    # 최근 4주 속도 → 추세 판정. 감소 추세에서는 '가속/감속'이 뒤집혀 읽히므로
    # 진행 방향(pace 부호)에 맞춰 문구를 고른다.
    recent = s_cur[s_cur.index >= as_of - pd.Timedelta(days=28)]
    if len(recent) >= 5:
        r_pace = (float(recent.iloc[-1] - recent.iloc[0])
                  / max((recent.index[-1] - recent.index[0]).days, 1))
        tol = max(abs(pace) * 0.05, 0.5)      # 이 폭 안이면 속도 변화 없음으로 본다
        if abs(r_pace - pace) <= tol:
            tag = "속도 유지"
        elif pace >= 0:
            tag = "감소 전환" if r_pace < 0 else ("증가 가속" if r_pace > pace else "증가 둔화")
        else:
            tag = "증가 전환" if r_pace > 0 else ("감소 둔화" if r_pace > pace else "감소 심화")
        bits.append(f"최근 4주 {_c(r_pace, f'{r_pace:+,.1f}명/일')} → <b>{tag}</b>")

    # 연말 단순 선형 추정 (현 속도 유지 가정)
    rest = (datetime.date(as_of.year, 12, 31) - as_of.date()).days
    if rest > 0:
        bits.append(f"현 속도 유지 시 연말 ≈ {float(s_cur.iloc[-1]) + pace * rest:,.0f}명")
    return " · ".join(bits)

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
        st.markdown(f"<div class='report-box'>{store[key] or '내용을 입력해 주세요.'}</div>",
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
                "데이터에 안 나오는 배경을 적으면 AI가 원인과 맥락을 풀 때 참고해요. "
                "수치는 데이터에서만 가져와요.",
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
                            help="기준 주차 실적으로 템플릿 문구를 채워요. 기존 내용은 지워져요."):
            store[key] = regen
            all_d = load_insights(); all_d[key] = regen; save_insights(all_d)
            st.session_state[ekey] = False; st.rerun()
        bi += 1
    if ai_fn is not None:
        if bcols[bi].button("AI 생성", key=f"wr_ai_{key}", width="stretch",
                            help="Claude가 데이터와 참고 메모를 보고 인사이트 문구를 써요. 기존 내용은 지워져요."):
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

def month_trim(s):
    """표시용 — 달의 앞자리 0을 뗀다 ('08월 2주차' → '8월 2주차').

    **저장 라벨은 그대로 둔다.** 라벨이 곧 조회 키(`pick`·`label` 비교)라 여기서
    바꾸면 조인이 통째로 어긋난다. 화면·엑셀에 찍기 **직전에만** 쓴다.
    `10월`은 0 앞이 숫자라 lookbehind에 걸려 안 바뀐다.
    """
    return re.sub(r"(?<!\d)0(\d월)", r"\1", str(s)) if s else s


def week_disp(year, label):
    """주차 표시: 2026년 6월 2주차"""
    return f"{year}년 {month_trim(label)}" if label else "-"

def week_back(df, wy, wlabel, years):
    """wlabel의 `years`년 전 대응 주차 (label, exact). 5주차 등은 그 달 마지막 주로 대체."""
    wm = re.match(r"(\d{1,2})월 (\d)주차", wlabel or "")
    if not wm:
        return wlabel, True
    return week_like(df, wy - years, int(wm.group(1)), int(wm.group(2)))

def wow_summary_table(df, wy, wlabel, metrics):
    """전주비 요약표 (실적 요약과 동일 구조): 전주·기준주·전주비 + 전년동주·전년비"""
    py, plb = prev_label(df, "주", wy, wlabel)
    ylb, _ = week_back(df, wy, wlabel, 1)
    cols = [week_disp(py, plb), week_disp(wy, wlabel), "전주비",
            week_disp(wy - 1, ylb), "전년비"]
    rows = []
    for met in metrics:
        cur = pick(df, "주", met, "*TOTAL", wy, wlabel, "mtd")
        prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
        yoy = pick(df, "주", met, "*TOTAL", wy - 1, ylb, "final") if ylb else np.nan
        rows.append({
            "구분": met,
            cols[0]: fmt_value(met, prv), cols[1]: fmt_value(met, cur),
            cols[2]: fmt_delta(met, cur, prv) or "–",
            cols[3]: fmt_value(met, yoy), cols[4]: fmt_delta(met, cur, yoy) or "–",
        })
    return pd.DataFrame(rows).set_index("구분")

def yoy2_summary_table(df, wy, wlabel, metrics):
    """전년비 · 전전년비 비교표 — 올해 흐름이 전전년에도 있었는지 확인용.

    같은 주차의 3개 연도 값과 세 가지 증감을 나란히 둔다.
      · 전년비        = 금년 vs 전년   (올해 무슨 일이 있었나)
      · 전전년비      = 금년 vs 전전년 (2년 새 얼마나 움직였나)
      · 전년의 전년비 = 전년 vs 전전년 (작년에도 같은 방향이었나)
    '추세' 칸은 전년의 전년비 → 전년비 순서로 방향을 읽어 '2년 연속'인지 판정한다.
    반환: (표, 전년라벨, 전전년라벨, 전년정확여부, 전전년정확여부)
    """
    l1, e1 = week_back(df, wy, wlabel, 1)
    l2, e2 = week_back(df, wy, wlabel, 2)

    def _neg(d):
        return None if not d or d == "–" else d.startswith("△")

    def _trend(d_now, d_prev):
        a, b = _neg(d_now), _neg(d_prev)
        if a is None or b is None: return "–"
        if a and b:       return "2년 연속 감소"
        if not a and not b: return "2년 연속 증가"
        return "증가 → 감소" if a else "감소 → 증가"

    cols = [week_disp(wy, wlabel), week_disp(wy - 1, l1), week_disp(wy - 2, l2),
            "전년비", "전전년비", "전년의 전년비", "추세"]
    rows = []
    for met in metrics:
        cur = pick(df, "주", met, "*TOTAL", wy, wlabel, "mtd")
        p1 = pick(df, "주", met, "*TOTAL", wy - 1, l1, "final") if l1 else np.nan
        p2 = pick(df, "주", met, "*TOTAL", wy - 2, l2, "final") if l2 else np.nan
        d1 = fmt_delta(met, cur, p1)   # 금년 vs 전년
        d2 = fmt_delta(met, cur, p2)   # 금년 vs 전전년
        dp = fmt_delta(met, p1, p2)    # 전년 vs 전전년
        rows.append({
            "구분": met,
            cols[0]: fmt_value(met, cur), cols[1]: fmt_value(met, p1),
            cols[2]: fmt_value(met, p2),
            cols[3]: d1 or "–", cols[4]: d2 or "–", cols[5]: dp or "–",
            cols[6]: _trend(d1, dp),
        })
    return pd.DataFrame(rows).set_index("구분"), l1, l2, e1, e2

def yoy_summary_table(df, ref_year, ref_month, metrics):
    """참고본 '실적 요약' 표: 전월·당월 × (전년, 당년, 전년비) + 당월의 전월비.

    맨 끝 '전월비(당월)'은 같은 해 당월(MTD) ↔ 전월(월마감) 비교다. 당월은 아직
    진행 중이라 누계로는 못 맞대지만, 이 표의 값이 전부 **일평균**이라 기간 길이가
    달라도 그대로 견줄 수 있다.
    """
    rows = []
    pm_y, pm_m = (ref_year, ref_month - 1) if ref_month > 1 else (ref_year - 1, 12)
    cols = [f"{pm_y-1}년 {pm_m}월", f"{pm_y}년 {pm_m}월", "전년비(전월)",
            f"{ref_year-1}년 {ref_month}월", f"{ref_year}년 {ref_month}월", "전년비(당월)",
            "전월비(당월)"]
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
            cols[6]: fmt_delta(met, cm_cur, pm_cur) or "–",
        })
    return pd.DataFrame(rows).set_index("구분"), (pm_y, pm_m)

def trend_table(df, gran, metrics, years, seg="*TOTAL", delta_year=None,
                delta_side="inline", win=None):
    """추이표: 행=지표, 열=(연도, 기간)

    `delta_year`를 주면 그 해의 각 기간에 전년 같은 기간 대비 증감 칸을 붙인다.
    2025년 1월과 2026년 1월이 열두 칸 떨어져 있어서 그냥은 못 맞대기 때문이다.

    `delta_side`가 자리를 정한다:
      - `"inline"` — 기간 **바로 오른쪽**에 끼운다(03·04가 쓰는 배치).
        한 기간의 값과 증감을 붙여 읽기 좋다.
      - `"right"` — 값을 다 놓고 **증감을 오른쪽에 몰아**둔다.
        증감만 가로로 훑어 '어느 달부터 꺾였나'를 보기 좋다.

    비교 대상은 **화면에 그린 연도와 무관하게** `delta_year - 1`을 직접 조회한다.
    전년을 안 그리고 있어도 증감은 나와야 한다.

    `win`을 주면 **그 기간만** 낸다(화면 위 「조회 기간」). 안 주면 그 해 전체다 —
    기간을 좁혀 놓고 추이만 연중으로 남으면 위아래가 다른 기간을 말하게 된다.
    """
    _win = None if win is None else set(as_labels(win))
    spec, tail = [], []             # (연도, 표시 라벨, 조회 라벨, 비교 연도 or None)
    for y in years:
        for lb in labels_sorted(df, gran, [y]):
            if _win is not None and str(lb) not in _win:
                continue
            sub = df[(df["gran"] == gran) & (df["year"] == y) & (df["label"] == lb) &
                     df["value"].notna()]
            if sub.empty: continue
            spec.append((y, lb, lb, None))
            if delta_year is not None and y == delta_year:
                (tail if delta_side == "right" else spec).append(
                    (y, f"{lb} 증감", lb, y - 1))
    spec += tail
    if not spec:
        return pd.DataFrame()
    # 값은 **보고서 규칙**으로 읽는다(`report_series`) — 파일의 가입율은 일별 비율의
    # 평균이라 ①·② 카드와 다른 숫자가 된다. 연도마다 한 번만 만들어 재사용한다.
    _ser = {}

    def _v(met, y, lb):
        s = _ser.get((met, y))
        if s is None:
            s = _ser[(met, y)] = report_series(df, gran, met, seg, y, "final")
        v = s.get(lb, np.nan)
        return np.nan if v is None else v

    out = {}
    for met in metrics:
        vals = []
        for y, _show, lb, base in spec:
            v = _v(met, y, lb)
            if base is None:
                vals.append(v)
            else:
                # 이 칸만 미리 문자열이다 — style_trend가 fmt_value를 다시 안 먹인다
                vals.append(fmt_delta(met, v, _v(met, base, lb)) or "–")
        out[met] = vals
    # 표시용으로만 앞자리 0을 뗀다 — 조회는 위에서 원래 라벨로 이미 끝났다
    columns = [(y, month_trim(show)) for y, show, _lb, _b in spec]
    tbl = pd.DataFrame(out, index=pd.MultiIndex.from_tuples(columns, names=["연도", "기간"])).T
    return tbl

def _is_delta_col(c):
    """추이표에서 증감 칸인지 — MultiIndex면 마지막 레벨(기간 라벨)을 본다."""
    return "증감" in str(c[-1] if isinstance(c, tuple) else c)

def style_trend(tbl, metrics):
    # 최신 pandas는 float 컬럼에 문자열 대입을 금지하므로 object로 변환 후 포맷
    disp = tbl.astype(object).copy()
    # 증감 칸은 이미 문자열이라 fmt_value를 먹이면 '△5.2%'가 뭉개진다
    dmask = [_is_delta_col(c) for c in tbl.columns]
    for met in disp.index:
        disp.loc[met] = [v if d else fmt_value(met, v)
                         for v, d in zip(tbl.loc[met], dmask)]
    return style_delta_cols(disp) if any(dmask) else disp

# ══════════════════════════════════════════════════════
# YoY 라인차트 (Plotly)
# ══════════════════════════════════════════════════════
def base_layout(h=300, ysuffix="", title=""):
    # 제목은 최상단(container 기준), 범례는 그 아래 별도 줄(plot 상단 바로 위)로 분리해
    # 제목·범례가 같은 높이에서 겹치지 않게 한다. 상단 여백(t)을 넉넉히 확보.
    return dict(
        paper_bgcolor="rgba(248,249,252,0)", plot_bgcolor="rgba(248,249,252,0)",
        font=dict(color="#475569", size=11), margin=dict(l=10, r=10, t=64, b=10),
        height=h, showlegend=True,
        legend=dict(orientation="h", yref="paper", yanchor="bottom", y=1.0,
                    xanchor="left", x=0, bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#64748b", size=10)),
        title=dict(text=title, font=dict(color="#94a3b8", size=13),
                   x=0, xanchor="left", yref="container", y=0.98, yanchor="top"),
        xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=10)),
        yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=10), ticksuffix=ysuffix),
    )

def yoy_chart(df, gran, metric, years, seg="*TOTAL", h=300, win=None):
    """`win`을 주면 그 기간만 그린다 — 화면 위 「조회 기간」을 따라가는 자리."""
    unit, div = METRIC_UNIT.get(metric, ("", 1))
    if metric in PCT_METRICS: div, unit = 0.01, "%"
    if gran == "월":
        x_all = [month_label(i) for i in range(1, 13)]
    else:
        x_all = labels_sorted(df, gran, years)
    if win is not None:
        _w = set(as_labels(win))
        x_all = [x for x in x_all if str(x) in _w]
    fig = go.Figure()
    for i, y in enumerate(sorted(years)):
        # 연도마다 없는 주차(5주차 등)는 건너뛰고 선을 잇는다
        s = report_series(df, gran, metric, seg, y, "final").reindex(x_all).dropna()
        fig.add_trace(go.Scatter(
            x=[month_trim(v) for v in s.index], y=(s / div).tolist(),
            mode="lines+markers", name=str(y),
            line=dict(color=clr(YEAR_PAL[i % len(YEAR_PAL)]), width=2),
            marker=dict(size=5),
        ))
    gname = {"일": "일자별", "주": "주차별", "월": "월별"}.get(gran, gran)
    ly = base_layout(h, ysuffix=unit if unit == "%" else "",
                     title=f"{metric} {gname} 추이 ({unit})")
    ly["xaxis"]["categoryorder"] = "array"
    ly["xaxis"]["categoryarray"] = [month_trim(v) for v in x_all]
    # 일별은 라벨이 한 해 365개라 그대로 두면 축이 새까매진다 — 눈금 수를 묶는다.
    if gran in ("주", "일"):
        ly["xaxis"]["tickangle"] = -45
        ly["xaxis"]["nticks"] = 20 if gran == "주" else 14
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
                ws.cell(r + 1, 2 + j, month_trim(lb)).font = head_font
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
            if "전년비" in str(cname) or "전월비" in str(cname):
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
                ws.cell(r, 2 + j, month_trim(lb)).fill = head_fill
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
            wm = re.match(r"(\d{1,2})월 (\d)주차", ref_week)
            mom = np.nan
            ylb = ref_week
            if wm:
                # 5주차처럼 대응 주차가 없으면 그 달 마지막 주차로 대체 (KPI 카드와 동일 규칙)
                mo, wk = int(wm.group(1)), int(wm.group(2))
                my, mm = (ref_year, mo - 1) if mo > 1 else (ref_year - 1, 12)
                mlb, _ = week_like(df, my, mm, wk)
                if mlb:
                    mom = pick(df, "주", met, "*TOTAL", my, mlb, "final")
                ylb, _ = week_like(df, ref_year - 1, mo, wk)
            yoy = pick(df, "주", met, "*TOTAL", ref_year - 1, ylb, "final") if ylb else np.nan
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
        return None, ("ANTHROPIC_API_KEY가 없어요. "
                      "Streamlit Cloud → Settings → Secrets에 추가해 주세요.")
    try:
        import anthropic
    except ImportError:
        return None, "anthropic 패키지가 없어요. requirements.txt에 넣고 다시 배포해 주세요."

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
        return (text or None), (None if text else "응답이 비어 있어요.")
    except anthropic.AuthenticationError:
        return None, "API 키 인증에 실패했어요. 키를 확인해 주세요."
    except anthropic.RateLimitError:
        return None, "요청이 많아 잠시 제한됐어요. 조금 뒤에 다시 시도해 주세요."
    except Exception as e:
        return None, f"생성 중 오류: {e}"

# ══════════════════════════════════════════════════════
# 04. 앱푸시 동의 현황 페이지
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
        st.info(f"{sel_gran} 단위 데이터가 없어요.")
    else:
        st.subheader(f"{ref_year}년 {gran_opt} 신규가입 및 앱푸시 수신동의율 추이")
        dates = labels_sorted(df, sel_gran, [ref_year])

        if not dates:
            st.info(f"{ref_year}년 {gran_opt} 데이터가 없어요.")
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
        st.info("동의자 잔고·이탈·순증 뷰는 PUSH 원천 엑셀을 다시 올리면 보여요. "
                "기존·신규·Total 섹션을 새로 인식해 누적에 저장해요.")
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

        # 타겟팅 가능 모수 — 연중 추이 (전년 비교). 세그먼트별로 각각.
        # 잔고(앱푸시_동의자수)는 실측 대조 결과 원천의 '타겟팅 가능' 행과 동일한 모수다.
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("타겟팅 가능 모수 — 연중 추이 (전년 비교)")
        st.caption("실제로 발송할 수 있는 수신동의 모수예요. 연도별 라인을 같은 연중 위치(월·일)에 "
                   "겹쳐 작년과 비교해요.")
        byr = bal.copy()
        # 각 연도를 같은 X축(연중 위치)에 겹치려면 공통 연도로 정규화 (2000=윤년이라 2/29 안전)
        byr["mdt"] = byr["dt"].apply(lambda d: d.replace(year=2000))
        yrs_b = sorted(byr["year"].dropna().unique().astype(int))
        # 연도가 쌓일수록 선이 4~5개가 되어 최근 흐름이 안 읽힌다. 기본은 **최근 2개년**
        # (제목 그대로 전년 비교)이고, 과거는 필요할 때 켜서 본다.
        _ydef = yrs_b[-2:]
        guard_multi("push_yoy_years", yrs_b)
        _ykw = {} if "push_yoy_years" in st.session_state else {"default": _ydef}
        sel_yrs = sorted(st.multiselect(
            "비교 연도", yrs_b, key="push_yoy_years", format_func=lambda y: f"{y}년",
            help="겹쳐 볼 연도를 골라요. 색은 연도에 고정이라 연도를 빼도 남은 선의 색은 "
                 "안 바뀌어요.", **_ykw))
        # 색은 **연도에 고정**한다(전체 목록 기준). 선택한 것만으로 색을 매기면 한 해를
        # 뺄 때마다 남은 선의 색이 바뀌어 눈이 흔들린다.
        _ycolor = {y: YEAR_PAL[i % len(YEAR_PAL)] for i, y in enumerate(yrs_b)}
        SEG_LABEL = {"*TOTAL": "Total", "기존": "기존", "신규": "신규"}
        if not sel_yrs:
            st.info("비교할 연도를 하나 이상 골라 주세요.")
        for seg in (["*TOTAL", "기존", "신규"] if sel_yrs else []):
            sb = byr[byr["segment"] == seg]
            if sb.empty: continue
            figy = go.Figure()
            for yr in sel_yrs:
                s = (sb[sb["year"] == yr].sort_values("mdt")
                     .set_index("mdt")["value"].dropna())
                if s.empty: continue
                figy.add_trace(go.Scatter(
                    x=s.index, y=s.values, mode="lines", name=str(yr),
                    line=dict(color=clr(_ycolor[yr]), width=2),
                    hovertemplate="%{x|%m/%d}<br>" + str(yr) + " %{y:,.0f}명<extra></extra>"))
            lyy = base_layout(260, title=f"{SEG_LABEL[seg]} — 타겟팅 가능 (명)")
            lyy["xaxis"]["tickformat"] = "%m월"
            lyy["xaxis"]["dtick"] = "M1"
            figy.update_layout(**lyy)
            st.plotly_chart(figy, width="stretch")

            # 증가속도 코멘트는 **화면에 그린 연도**를 따라간다 — 안 보이는 해와 비교한
            # 문장이 붙으면 차트와 말이 어긋난다. (실제 날짜 인덱스로 계산 — mdt 아님)
            _ser = lambda y: (sb[sb["year"] == y].sort_values("dt")
                              .set_index("dt")["value"]) if y in yrs_b else None
            note = growth_pace_note(_ser(sel_yrs[-1]),
                                    _ser(sel_yrs[-2]) if len(sel_yrs) >= 2 else None)
            if note:
                st.markdown(
                    "<div style='font-size:12px;color:#64748b;margin:-6px 0 16px 2px'>"
                    f"📈 <b>{SEG_LABEL[seg]} 증감 속도</b> — {note}</div>",
                    unsafe_allow_html=True)

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
        st.info("일자별 가입자수·앱푸시 수신동의 데이터가 없어서 월별 YoY 표를 보여줄 수 없어요.")
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

        # min_count=1 — 값이 하나도 없는 달을 0으로 만들지 않는다.
        # (기본 sum()은 전부 NaN이면 0을 돌려줘서, 아직 안 올라온 당월 가입자수가
        #  '0명 · YoY △100%'로 찍히고 동의율은 288/0=inf → 마스킹돼 '-'로 빠졌다.)
        grp = (sub.groupby(["year", "month", "metric"])["value"].sum(min_count=1)
               .unstack("metric").reset_index())

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

        wtable(style_yoy_rows(disp), width="stretch")

    if not res_df.empty:
        st.markdown("### 상세 데이터")
        disp_df = res_df.copy()
        disp_df["앱푸시수신동의"] = disp_df["앱푸시수신동의"].apply(lambda x: f"{int(x):,}" if not pd.isna(x) else "–")
        disp_df["가입자수"] = disp_df["가입자수"].apply(lambda x: f"{int(x):,}" if not pd.isna(x) else "–")
        disp_df["동의율"] = disp_df["동의율"].apply(lambda x: f"{x*100:.1f}%" if not pd.isna(x) else "–")
        wtable(disp_df.set_index("날짜"), width="stretch", dl_name="상세 데이터")

# ══════════════════════════════════════════════════════
# 페이지 PDF 저장 (브라우저 인쇄 → PDF, 차트 포함)
# ══════════════════════════════════════════════════════
def guard_select(key, opts, default=None):
    """옵션 목록이 바뀌면 세션에 남은 옛 선택값이 목록 밖이 된다 — 조용한 리셋·예외 방지.

    `default`를 주면 **처음 한 번만** 그 값을 초기값으로 심는다(옵션에 있을 때만).
    위젯을 만들기 전에 세션에 넣는 방식이라, `key`가 붙은 위젯에 `index=`를 같이
    넘길 때 나는 경고를 피하고 사용자가 고른 값도 덮지 않는다.
    """
    if key in st.session_state and st.session_state[key] not in opts:
        st.session_state.pop(key, None)
    if key not in st.session_state and default is not None and default in opts:
        st.session_state[key] = default


def guard_range(key, opts, default):
    """select_slider(범위)용 가드 — `guard_select`의 두 칸짜리 판.

    옵션이 바뀌면(파일을 새로 올려 날짜가 늘거나 줄면) 세션에 남은 옛 경계가 목록 밖이
    되어 위젯이 예외로 죽는다. **손 안 댄 쪽은 새 범위로 따라가고, 직접 좁힌 쪽은 새
    경계 안으로 clamp**한다 — 기간 필터의 세션 처리와 같은 규칙이다.

    돌려주는 값을 `value=`로 넘긴다. 세션값이 성하면 그대로라 사용자가 고른 걸 안 덮는다.
    """
    cur = st.session_state.get(key)
    if isinstance(cur, (list, tuple)) and len(cur) == 2 and all(v in opts for v in cur):
        return tuple(cur)
    st.session_state.pop(key, None)
    return tuple(default)


def guard_multi(key, opts):
    """multiselect용 가드 — 사라진 값이 세션에 남으면 위젯이 예외로 죽는다."""
    cur = st.session_state.get(key)
    if isinstance(cur, (list, tuple)):
        keep = [v for v in cur if v in opts]
        if len(keep) != len(cur):
            st.session_state[key] = keep


def follow_options(key, opts):
    """«전체 선택»으로 시작하는 다중선택이 **데이터를 따라가게** 한다.

    `key`가 붙은 위젯은 `default=`가 최초 1회만 먹는다. 그래서 전체로 시작한 필터는
    새 항목이 생겨도 옛 목록을 계속 써서 **그 항목이 화면에서 통째로 사라진다**
    (기간 필터에서 겪은 것과 같은 자리 — 증상이 'F5 누르면 보인다'로만 나타난다).

    직전 옵션 전체를 같이 기억해 두고 **세션값이 그것과 같으면**(=손을 안 댄 필터면)
    새 전체 목록으로 갱신한다. 직접 좁혀 둔 선택은 경계 안으로만 자르고 안 건드린다.
    """
    seen = f"_{key}__opts"
    cur, prev = st.session_state.get(key), st.session_state.get(seen)
    if not isinstance(cur, (list, tuple)):
        st.session_state[key] = list(opts)
    elif prev is not None and set(cur) == set(prev):
        st.session_state[key] = list(opts)
    else:
        st.session_state[key] = [v for v in cur if v in opts]
    st.session_state[seen] = list(opts)


# 조직·카테고리 목록엔 실적이 사실상 없는 행이 섞여 온다 — 미매칭·기타·라움워치·리빙사업부·
# PROJECT-C·e-Corner·SPACE-R. 실파일 전 기간을 다 더해도 수천 원이라, 살아 있는 조직(15~33%)과
# 수만 배 차이다. 표·차트·히트맵 자리만 차지하면서 아무것도 말해 주지 않는다.
# 이름으로 지우면 새 조직이 생길 때마다 또 손대야 하니 **비중으로** 거른다.
ORGCAT_MIN_SHARE = 0.001        # 전 기간 거래액이 부모의 0.1% 미만이면 숨김


def orgcat_depth(sub):
    """이 데이터가 실제로 가진 뎁스 — 값이 하나라도 있는 레벨까지.

    브랜드·상품 칸은 아직 '-'로만 와서 지금은 2단이다. 값이 들어오는 날 자동으로
    3·4단이 열리도록 **데이터를 보고** 정한다(코드를 안 고쳐도 되게).
    """
    n = 0
    for col, _lb, _h in ORGCAT_LEVELS:
        if col in sub.columns and (sub[col].astype(str) != ORGCAT_NONE).any():
            n += 1
        else:
            break
    return max(n, 1)



class OrgcatView:
    """`sub` 한 벌을 **한 번만** 훑어 만든 조회 인덱스.

    화면은 노드마다 `opick`을 부르는데, 그때마다 프레임 전체를 다시 훑으면 브랜드 300개 ×
    지표 3종 × 2개년 = 1,800번 풀스캔이 된다. 실측(원장 120만 행, 브랜드 단계)에서 조회에만
    3.9초, 마스크(`orgcat_node`)에 1.6초가 더 들었다. 값 조회는 딕셔너리 하나로 끝내고,
    자식 목록과 전 기간 거래액 합계도 같은 한 번의 순회에서 같이 만든다.

    노드 판정 규칙은 `orgcat_node`와 같다 — 앞쪽 레벨은 실제 값, 그 뒤는 전부 `*TOTAL`
    (그 레벨의 합계) 또는 빈 칸(레벨 자체가 없음). 중간이 비었는데 더 깊은 레벨에 값이
    있는 행은 성한 노드가 아니라 버린다.
    """

    __slots__ = ("_final", "_mtd", "_kids", "_rev", "depth")

    # 카테고리를 좁혀 본 상태인가. 파일 값을 그대로 쓰는 원본은 늘 False다
    # (`_CatFilteredView`가 True로 덮는다) — 화면 문구가 이걸로 갈린다.
    narrowed = False

    def __init__(self, sub):
        self._final, self._mtd, self._kids, self._rev = {}, {}, {}, {}
        self.depth = 0
        lv = [c for c in ORGCAT_LV if c in getattr(sub, "columns", ())]
        if sub is None or not len(sub) or not lv:
            return
        cols = [sub[c].astype(str).to_numpy() for c in lv]
        real = np.stack([~np.isin(v, (ORGCAT_TOTAL, ORGCAT_NONE)) for v in cols], axis=1)
        # 앞에서부터 연속으로 값이 있는 개수 = 그 행의 뎁스. 그 뒤에도 값이 있으면
        # (중간이 빈 행) 노드로 안 친다 — 합계 자리와 겹쳐 값이 틀어진다.
        lead = np.cumprod(real, axis=1).sum(axis=1)
        ok = lead == real.sum(axis=1)
        self.depth = int(lead.max()) if len(lead) else 0
        met = sub["metric"].astype(str).to_numpy()
        yr = pd.to_numeric(sub["year"], errors="coerce").to_numpy()
        lab = sub["label"].astype(str).to_numpy()
        cl = (sub["close"].astype(str).to_numpy() if "close" in sub.columns
              else np.full(len(sub), "final"))
        val = pd.to_numeric(sub["value"], errors="coerce").to_numpy(dtype=float)
        for i in range(len(sub)):
            if not ok[i]:
                continue
            node = tuple(c[i] for c in cols[:lead[i]])
            if node:                                  # 부모 → 자식 목록 (첫 등장 순서 유지)
                self._kids.setdefault(node[:-1], {})[node[-1]] = None
            v = val[i]
            if v != v:
                continue
            if met[i] == "첫구매 거래액":              # 전 기간 합 — 실적 없는 항목 숨김 판정용
                self._rev[node] = self._rev.get(node, 0.0) + v
            y = yr[i]
            k = (met[i], (int(y) if y == y else None), lab[i], node)
            (self._final if cl[i] == "final" else self._mtd)[k] = v   # 같은 키는 뒤 행 우선

    def get(self, path, metric, year, label, prefer="final"):
        """값 하나. `label`이 여러 개면 그 기간의 **일평균** — `pick`과 같은 규칙이다.

        ⑤가 ①과 다른 기간을 보면 두 원천 비교(합계 대사)가 통째로 헛돈다.
        """
        y = pd.to_numeric(pd.Series([year]), errors="coerce").iloc[0]
        y = int(y) if y == y else None
        a, b = ((self._final, self._mtd) if prefer == "final"
                else (self._mtd, self._final))
        node, met = tuple(path), str(metric)
        got = []
        for lb in as_labels(label):
            v = a.get((met, y, lb, node), b.get((met, y, lb, node)))
            if v is not None and v == v:
                got.append(v)
        return float(np.mean(got)) if got else np.nan

    def children(self, path):
        return list(self._kids.get(tuple(path), ()))

    def live(self, path=()):
        """실적이 사실상 없는 자식을 걸러 낸 목록 → (보일 것, 숨긴 것).

        판단은 **적재된 전 기간**의 거래액 합으로 한다 — 고른 기간으로 재면 달마다 목록이
        바뀌어 파고들기 선택이 튀고, 한두 달 쉰 항목이 통째로 사라진다."""
        path = tuple(path)
        kids = self.children(path)
        if not kids:
            return [], []
        parent = self._rev.get(path, 0.0)
        base = abs(parent) if (np.isfinite(parent) and parent) else 1.0
        live, hidden = [], []
        for k in kids:
            kid = self._rev.get(path + (k,), 0.0)
            (live if abs(kid) / base >= ORGCAT_MIN_SHARE else hidden).append(k)
        if not live:                  # 전부 잘리면 거르지 않는다 (빈 화면보다 낫다)
            return kids, []
        return live, hidden


# 같은 프레임으로 연달아 물어보는 경우(낱개 헬퍼를 루프에서 부르는 코드)를 위한 1칸 메모.
# `is`로만 맞춰 보므로 다른 데이터를 잘못 돌려줄 일이 없다 — 프레임을 붙들고 있어서
# id가 재활용될 수도 없다. 붙들리는 건 한 벌뿐이라 메모리도 걱정 없다.
_OCVIEW_LAST = [None, None]


def orgcat_view(sub):
    """`sub`의 조회 인덱스. 화면은 이걸 **한 번 만들어 재사용**할 것."""
    if _OCVIEW_LAST[0] is sub:
        return _OCVIEW_LAST[1]
    v = OrgcatView(sub)
    _OCVIEW_LAST[0], _OCVIEW_LAST[1] = sub, v
    return v



# 상위 값을 «고른 카테고리»에서 다시 만들 때 쓰는 축. 화면 지표 다섯을 이 셋으로 덮는다.
_CATFIX_BASE = ("첫구매 거래액", "첫구매 고객수", "상품UV")


class _CatFilteredView:
    """카테고리 몇 개만 남기고 본 조회 뷰 — 상위 값을 **다시 만든다**.

    "전년에 슈즈에서 잘 팔리던 브랜드가 올해 빠졌다" 같은 일이 생기면 그 카테고리를
    빼고 나머지를 봐야 한다. 그런데 **뺀 뒤의 조직 합계·전체 합계는 파일에 없다** —
    파일 값을 그대로 두면 뺀 카테고리가 그 안에 그대로 들어 있어 표가 스스로 모순된다.

    그래서 상위 값은 고른 카테고리에서 다시 만든다. 규칙은 카테고리 합산 표
    (`_funnel_cat_rollup`)와 **같은 것**이다 — 거래액·고객수·상품UV는 더하고,
    객단가 = 거래액합 ÷ 고객수합, 상품CR = 고객수합 ÷ 상품UV합
    (`거래액 = 상품UV × 상품CR × 객단가`와 `거래액 = 고객수 × 객단가`에서 나오는
    항등식이다). 고객수·상품UV는 유니크 값이라 합이 조금 부풀려지는 것까지 같다 —
    지어내지 않고 화면에 밝힌다.

    카테고리 «아래»(브랜드·상품)와 카테고리 자신은 손대지 않고 그대로 넘긴다.
    """

    narrowed = True

    def __init__(self, view, cats):
        self._v = view
        self._cs = set(cats)
        self._memo = {}
        self.depth = getattr(view, "depth", 0)

    def children(self, path):
        kids = self._v.children(path)
        return [k for k in kids if k in self._cs] if len(tuple(path)) == 1 else kids

    def live(self, path=()):
        path = tuple(path)
        live, hidden = self._v.live(path)
        if len(path) == 1:
            return [k for k in live if k in self._cs], hidden
        if not path:
            # 고른 카테고리를 하나도 안 가진 조직은 뺀다 — 빈 줄만 남는다
            return [o for o in live
                    if any(c in self._cs for c in self._v.live((o,))[0])], hidden
        return live, hidden

    def _cats(self, path):
        """이 노드 아래에서 값을 모을 (조직, 카테고리) 목록."""
        path = tuple(path)
        if len(path) == 1:
            return [path + (c,) for c in self._v.live(path)[0] if c in self._cs]
        return [(o, c) for o in self._v.live(())[0]
                for c in self._v.live((o,))[0] if c in self._cs]

    def get(self, path, metric, year, label, prefer="final"):
        path = tuple(path)
        if len(path) >= 2:                                # 카테고리 이하는 그대로
            return self._v.get(path, metric, year, label, prefer)
        key = (path, str(metric), year, tuple(as_labels(label)), prefer)
        if key in self._memo:
            return self._memo[key]
        acc = {k: np.nan for k in _CATFIX_BASE}
        for node in self._cats(path):
            for k in _CATFIX_BASE:
                v = self._v.get(node, k, year, label, prefer)
                if not pd.isna(v):
                    acc[k] = v if pd.isna(acc[k]) else acc[k] + v
        rev, cust, uv = (acc[k] for k in _CATFIX_BASE)
        if metric == "첫구매 객단가":
            out = (rev / cust if not (pd.isna(rev) or pd.isna(cust) or not cust)
                   else np.nan)
        elif metric == "상품CR":
            out = (cust / uv if not (pd.isna(cust) or pd.isna(uv) or not uv)
                   else np.nan)
        else:
            out = acc.get({"첫구매 고객수": "첫구매 고객수",
                           "상품UV": "상품UV"}.get(str(metric), "첫구매 거래액"), np.nan)
        self._memo[key] = out
        return out


def cat_filtered_view(view, cats, all_cats):
    """고른 카테고리가 전체와 같으면 **원본을 그대로** 돌려준다.

    손을 안 댄 화면이 예전과 한 글자도 달라지지 않게 하려는 것이다 — 전체일 땐
    상위 값이 파일 값 그대로여야 하고(다시 만들면 고객수가 살짝 부풀려진다),
    괜한 재계산 비용도 없다.
    """
    if not cats or set(cats) >= set(all_cats):
        return view
    return _CatFilteredView(view, cats)

def orgcat_node(sub, path, depth=None):
    """path(선택한 값들)가 가리키는 **합계 행**만 남긴 마스크.

    선택한 레벨은 그 값으로, 그 아래 레벨은 `*TOTAL`(그 레벨의 합계) 또는 빈 칸
    (레벨 자체가 없음)이어야 한다. 이 두 가지를 구분 못 하면 카테고리 합계 대신
    브랜드 한 줄이 잡혀 값이 통째로 틀어진다.
    """
    m = pd.Series(True, index=sub.index)
    for i, (col, _lb, _h) in enumerate(ORGCAT_LEVELS):
        if col not in sub.columns:
            continue
        v = sub[col].astype(str)
        m &= (v == path[i]) if i < len(path) else v.isin([ORGCAT_TOTAL, ORGCAT_NONE])
    return m


# 아래 셋은 **낱개 조회용 얇은 껍데기**다 — 구현은 OrgcatView 하나뿐이라 조용히 갈릴 일이
# 없다. 화면은 뷰를 직접 한 번 만들어 재사용할 것(노드마다 부르면 매번 다시 인덱싱한다).
def orgcat_children(sub, path):
    """path 바로 아래 단계의 항목들 (합계·빈 칸 제외). 더 들어갈 데가 없으면 빈 리스트."""
    return orgcat_view(sub).children(path)


def orgcat_live(sub, path=()):
    """실적이 사실상 없는 자식을 걸러 낸 목록 → (보일 것, 숨긴 것)."""
    return orgcat_view(sub).live(path)


def opick(sub, path, metric, year, label, prefer="final"):
    """path가 가리키는 노드의 값 하나 — 마스터의 pick()과 같은 규칙(final 우선)."""
    return orgcat_view(sub).get(path, metric, year, label, prefer)


# 상위와 합이 맞는 지표는 **거래액 하나뿐**이다. 실파일 검증에서 조직 합이 전체와
# 오차 0.00%로 일치했다. 고객수(+2.4%)·상품UV(+33%)는 유니크 값이라 같은 사람이 여러
# 조직에 잡혀 합이 상위를 넘는다 — 기여도 분해를 하면 안 된다.
ORGCAT_ADDITIVE = {"첫구매 거래액"}
# 거래액 = 상품UV × 상품CR × 객단가 (실파일에서 모든 레벨에 오차 0.000%로 성립)
ORGCAT_FACTORS = [("상품UV", "유입"), ("상품CR", "전환"), ("첫구매 객단가", "객단가")]


def _logmean(a, b):
    if a <= 0 or b <= 0: return np.nan
    if abs(a - b) < 1e-12: return float(a)
    return (a - b) / np.log(a / b)


def factor_split(prev, cur):
    """거래액 증감을 요인들로 쪼갠다 → (요인별 기여액, 총증감). 못 쪼개면 None.

    LMDI(로그평균 디비지아)라 다 더하면 실제 증감과 **정확히** 같고, 어느 요인을 먼저
    대입하느냐로 답이 달라지지 않는다(순차 대입법의 순서 의존을 피한다).

    **요인 개수는 원천마다 다르다.** MICRO는 (상품UV, 상품CR, 객단가) 셋이고, 결제 원장은
    상품UV·CR이 없어 (고객수, 객단가) 둘이다. 곱이 거래액이기만 하면 몇 개든 성립한다.
    0·음수·결측이 하나라도 있으면 로그가 정의되지 않으니 숫자를 지어내지 않는다.
    """
    prev, cur = list(prev), list(cur)
    if not prev or len(prev) != len(cur):
        return None
    if any(v is None or not np.isfinite(v) or v <= 0 for v in prev + cur):
        return None
    v0 = float(np.prod(prev))
    v1 = float(np.prod(cur))
    L = _logmean(v1, v0)
    if not np.isfinite(L): return None
    return tuple(L * np.log(cur[i] / prev[i]) for i in range(len(prev))), v1 - v0


def _won_m(v, digits=1):
    """백만원 표기 — 증감액은 부호를 살려 △ 대신 −를 쓴다(표의 전년비와 구분)."""
    if v is None or not np.isfinite(v): return "–"
    return f"{v / 1e6:+,.{digits}f}백만"


# ── 결제 원장 → 화면이 그대로 읽는 조직×카테고리 큐브 ──────────────────
@st.cache_data(show_spinner=False, max_entries=8)
def detail_periods(dates, gran):
    """고유 일자 → (year, label, sortkey, nd, close) 표. 기간 규칙은 여기 한 곳뿐이다.

    원장 전체에 `pd.to_datetime`을 거는 대신 **고유 일자만** 계산한다(2년치라 700개 남짓).
    120만 행에 그대로 걸면 그것만 0.9초라 리런마다 그 값을 다시 낸다.

    나눌 일수(`nd`)는 **적재 범위와 겹치는 달력 일수**다 — 매출이 0인 날은 원장에 줄 자체가
    없어서, 줄이 있는 날만 세면 일평균이 부풀려진다. 주차는 **목요일이 속한 달의 몇 번째
    주**(ISO와 같은 규칙) — 월요일로 정하면 한 주가 두 달에 걸릴 때 어느 달인지 흔들린다.
    """
    u = pd.Index(pd.Series(dates, dtype="object").dropna().astype(str).unique())
    dt = pd.to_datetime(pd.Series(u.values), errors="coerce")
    keep = dt.notna()
    u, dt = u[keep.values], dt[keep].reset_index(drop=True)
    if not len(u):
        return pd.DataFrame(columns=["date", "year", "label", "sortkey", "nd", "close"])
    per = dt.dt.to_period({"월": "M", "주": "W-SUN"}.get(gran, "D"))
    ps, pe = per.dt.start_time.dt.normalize(), per.dt.end_time.dt.normalize()
    if gran == "월":
        yy, mm = ps.dt.year, ps.dt.month
        lab, sk = mm.astype(str) + "월", yy * 10000 + mm * 100
    elif gran == "주":
        thu = ps + pd.Timedelta(days=3)
        yy, mm = thu.dt.year, thu.dt.month
        wk = (thu.dt.day - 1) // 7 + 1          # 그 달 첫 목요일이 든 주가 1주차
        lab = mm.astype(str).str.zfill(2) + "월 " + wk.astype(str) + "주차"
        sk = yy * 10000 + mm * 100 + wk
    else:
        yy, mm, dd = ps.dt.year, ps.dt.month, ps.dt.day
        lab = mm.astype(str) + "/" + dd.astype(str)
        sk = yy * 10000 + mm * 100 + dd
    lo, hi = dt.min().normalize(), dt.max().normalize()
    cs, ce = ps.clip(lower=lo), pe.clip(upper=hi)
    return pd.DataFrame({
        "date": u.values, "year": yy.astype(int).values, "label": lab.values,
        "sortkey": sk.astype(int).values,
        "nd": ((ce - cs).dt.days + 1).clip(lower=1).values,
        "close": np.where((ps >= lo) & (pe <= hi), "final", "mtd")})



# ── 일별에서 주·월을 만든다 ──────────────────────────────────────────
# 같은 실적인데 일·주·월 **세 파일**을 따로 받아 올려야 했다. 실파일 대조 결과 일별을
# 그 기간 평균으로 묶으면 월별 파일과 **원 단위까지 같다**(2025년 12개월, 거래액·고객수·
# 상품UV 전부 오차 0.0%). 고객수·상품UV가 기간 단위로 중복제거된 값이 아니라는 뜻이라
# 파생이 성립한다 — 유니크였다면 일별 평균이 월값보다 클 수밖에 없다.
#
# **비율은 평균 내지 말고 다시 만든다.** 파일의 객단가·상품CR은 '날마다의 비율을 평균'한
# 값이라 합÷합과 다르다(실측 0.2%·1.0%). 앱은 다른 자리에서도 비율을 합÷합으로 읽으므로
# (가입율 규칙과 같은 자리) 여기서도 두 카운트에서 만든다.
ORGCAT_DERIVE_MEAN = ["첫구매 거래액", "첫구매 고객수", "상품UV"]
# 비율 = (분자, 분모) — 둘 다 위 평균값에서 만든다
ORGCAT_DERIVE_RATIO = {"첫구매 객단가": ("첫구매 거래액", "첫구매 고객수"),
                       "상품CR": ("첫구매 고객수", "상품UV")}
# 비중 = 그 노드 ÷ 같은 기간의 *TOTAL 노드
ORGCAT_DERIVE_SHARE = {"거래액비중": "첫구매 거래액", "고객비중": "첫구매 고객수"}


@st.cache_data(show_spinner=False, max_entries=1)
def orgcat_derive_periods(df):
    """일별 행에서 주·월 행을 만들어 **빈자리만** 채운다.

    **파일이 준 값이 늘 이긴다** — 주·월 파일을 올렸으면 그게 그대로 쓰이고, 파생은 그
    파일에 없는 (기간·노드·지표)만 메운다. 그래야 지금까지 세 파일을 올리던 방식이
    그대로 살고, 일별만 올려도 화면이 빈 채로 남지 않는다.

    만든 기간은 `df.attrs["orgcat_derived"]`에 `{(gran, year, label)}`로 남긴다 —
    화면이 '이건 일별에서 만든 값'이라고 밝힐 수 있게. `attrs`는 merge·concat에서
    사라지니 **받은 자리에서 바로 읽을 것**(`af_rejected`와 같은 규칙).
    """
    if df is None or df.empty or not (df["gran"] == "일").any():
        return df
    day = df[df["gran"] == "일"]
    # 일 라벨 'M/D' + 연도 → 실제 날짜. 못 읽는 라벨은 그냥 버린다(지어내지 않는다).
    _d = (day["year"].astype(str) + "-"
          + day["label"].astype(str).str.replace("/", "-", regex=False))
    dt = pd.to_datetime(_d, errors="coerce")
    day = day[dt.notna()].copy()
    if day.empty:
        return df
    day["_date"] = dt[dt.notna()].dt.strftime("%Y-%m-%d").values

    keys = ["metric"] + ORGCAT_LV + ["ch", "lfms"]
    made, marks = [], set()
    for tg in ("주", "월"):
        # 캐시 함수라 해시 가능한 값으로 넘긴다 — Arrow 문자열 배열은 못 해시한다
        per = detail_periods(list(pd.unique(day["_date"].astype(object))), tg)
        if per.empty:
            continue
        pm = per.set_index("date")
        d = day.join(pm[["year", "label", "sortkey", "close"]], on="_date",
                     rsuffix="_p")
        d = d[d["label_p"].notna()]
        if d.empty:
            continue
        # **접미어 붙은 쪽을 쓸 것.** `day`에 이미 year·label·sortkey·close가 있어서
        # 조인이 새 값에 `_p`를 붙인다. `sortkey`를 그냥 쓰면 **일별 정렬키**로 묶여
        # 하루씩 따로 그룹이 되고, 월 값이 그 달 1일 값 그대로 나온다(실측 최대 59% 오차).
        g = (d[d["metric"].isin(ORGCAT_DERIVE_MEAN)]
             .groupby(keys + ["year_p", "label_p", "sortkey_p", "close_p"],
                      as_index=False, sort=False)["value"].mean())
        if g.empty:
            continue
        g = g.rename(columns={"year_p": "year", "label_p": "label",
                              "sortkey_p": "sortkey", "close_p": "close"})
        g["gran"] = tg
        # 비율·비중은 평균이 아니라 **다시 만든다**
        idx = keys[1:] + ["year", "label", "sortkey", "close"]
        wide = g.pivot_table(index=idx, columns="metric", values="value",
                             aggfunc="first")
        extra = []
        for met, (num, den) in ORGCAT_DERIVE_RATIO.items():
            if num in wide.columns and den in wide.columns:
                v = wide[num] / wide[den].where(wide[den] > 0)
                extra.append(v.rename("value").reset_index().assign(metric=met))
        # 비중 — 같은 기간·같은 축의 *TOTAL 노드로 나눈다
        tot_key = ["ch", "lfms", "year", "label", "sortkey", "close"]
        tot = wide.reset_index()
        tot = tot[(tot["org"] == ORGCAT_TOTAL) & (tot["cat"] == ORGCAT_TOTAL)]
        for met, base in ORGCAT_DERIVE_SHARE.items():
            if base not in wide.columns or tot.empty:
                continue
            t = tot[tot_key + [base]].rename(columns={base: "_tot"})
            j = wide.reset_index().merge(t, on=tot_key, how="left")
            v = j[base] / j["_tot"].where(j["_tot"] > 0)
            extra.append(j[idx].assign(value=v.values, metric=met))
        parts = [g.reindex(columns=ORGCAT_COLS)]
        for e in extra:
            e = e.assign(gran=tg)
            parts.append(e.reindex(columns=ORGCAT_COLS))
        out = pd.concat(parts, ignore_index=True)
        out = out[out["value"].notna()]
        if out.empty:
            continue
        made.append(out)
        marks |= {(tg, int(y), str(lb))
                  for y, lb in out[["year", "label"]].drop_duplicates().values}
    if not made:
        return df
    add = pd.concat(made, ignore_index=True)
    # **파일이 있는 기간엔 아예 안 만든다 — 칸 단위가 아니라 «기간» 단위로 비킨다.**
    # 키(`ORGCAT_KEY`)에 `close`가 들어 있어서, 칸 단위로 겹칠 때만 버리면 파일의
    # `final`과 파생의 `mtd`가 **둘 다 남는다**. 그런데 화면은 올해를 `prefer="mtd"`로
    # 읽으므로 파생 쪽을 먼저 집어, 사용자가 올린 마감값이 조용히 밀려난다.
    # 한 기간 안에서 두 원천이 섞이는 것도 막아 준다.
    have = set(map(tuple, df[df["gran"].isin(("주", "월"))]
                   [["gran", "year", "label"]].astype(str).drop_duplicates().values))
    key3 = list(zip(add["gran"].astype(str), add["year"].astype(str),
                    add["label"].astype(str)))
    add = add[[k not in have for k in key3]]
    if add.empty:
        return df
    kept = pd.concat([df[ORGCAT_COLS], add[ORGCAT_COLS]], ignore_index=True)
    kept.attrs["orgcat_derived"] = {
        m for m in marks if (m[0], str(m[1]), m[2]) not in have}
    return kept


def detail_stamp(ddf, gran):
    """원장 + 기간 칼럼(year·label·sortkey) + 그 기간의 일수(nd)·완결여부(close).

    값은 화면에 **일평균**으로 낸다. 합계로 내면 진행 중인 달(3일치)이 전년 같은 달(30일치)과
    맞붙어 △90% 가짜 급락이 뜬다. 나눌 일수는 **적재된 날짜 범위와 겹치는 달력 일수**다 —
    매출이 0인 날은 원장에 줄 자체가 없어서, 줄이 있는 날만 세면 일평균이 부풀려진다.

    주차는 **목요일이 속한 달의 몇 번째 주**다(ISO와 같은 규칙). 월요일로 달을 정하면
    한 주가 두 달에 걸릴 때 어느 달에 넣을지가 흔들린다.
    """
    per = detail_periods(ddf["date"] if len(ddf) else pd.Series([], dtype="object"), gran)
    if per.empty:
        d = ddf.iloc[0:0].copy()
        for c, tp in (("year", "int64"), ("label", "object"), ("sortkey", "int64"),
                      ("nd", "int64"), ("close", "object")):
            d[c] = pd.Series(dtype=tp)
        return d
    # 일자 → 기간표 행번호를 Categorical 코드로 한 번에 얻는다. dict.map을 칼럼마다
    # 돌리면 120만 행 × 5회라 그것만 0.8초다 (코드 색인은 배열 인덱싱이라 사실상 공짜).
    # `detail_fill`이 date를 이미 문자열로 맞춰 두므로 astype(str)은 120만 행짜리
    # 통째 복사만 하고 끝난다 — object dtype이면 그대로 쓴다.
    _dt = ddf["date"] if ddf["date"].dtype == object else ddf["date"].astype(str)
    code = pd.Categorical(_dt, categories=pd.Index(per["date"])).codes
    ok = code >= 0
    d = (ddf if ok.all() else ddf[ok]).reset_index(drop=True)
    cc = code if ok.all() else code[ok]
    for c in ("year", "label", "sortkey", "nd", "close"):
        d[c] = per[c].to_numpy()[cc]
    return d


@st.cache_data(show_spinner=False, max_entries=16)
def detail_level(sdf, gran, axis, restrict, depth):
    """원장 → 특정 뎁스 한 겹의 ORGCAT_COLS 행들. `restrict`로 그 위 경로에 가둔다.

    한 겹마다 원장 전체(수십~수백만 행)를 groupby하므로 **리런을 넘겨 캐시한다**.
    돌려주는 건 수천 행짜리라 캐시에 얹어도 가볍다. 캐시 키에 들어가는 `sdf` 해시는
    120만 행에서도 0.04초라, 다시 group by 하는 것보다 훨씬 싸다."""
    d = sdf if axis == DETAIL_ALL else sdf[sdf["ch"] == axis]
    for i, v in enumerate(restrict):
        d = d[d[DETAIL_LV[i]] == v]
    if d.empty:
        return pd.DataFrame(columns=ORGCAT_COLS)
    keys = ["year", "label", "sortkey", "nd", "close"] + DETAIL_LV[:depth]
    a = d.groupby(keys, as_index=False, sort=False)[["rev", "cust"]].sum()
    for col in DETAIL_LV[depth:]:
        a[col] = ORGCAT_TOTAL           # 그 아래 단계의 '합계' 행이라는 표시
    cu = a["cust"].where(a["cust"] > 0)  # 0·음수 고객수는 객단가를 지어내지 않는다
    vals = {"첫구매 거래액": a["rev"] / a["nd"], "첫구매 고객수": a["cust"] / a["nd"],
            "첫구매 객단가": a["rev"] / cu}
    base = a[["year", "label", "sortkey", "close"] + DETAIL_LV]
    out = []
    for met, v in vals.items():
        m = base.copy()
        m["metric"], m["value"] = met, np.asarray(v, dtype=float)
        out.append(m)
    r = pd.concat(out, ignore_index=True)
    r["gran"], r["lfms"] = gran, axis
    # 원장은 자기 유입채널(AF대분류)을 이미 `axis`로 걸러 `lfms` 자리에 실어 보낸다.
    # MICRO의 채널 축(`ch`)과는 다른 것이라 여기선 «전 채널»로 둔다 — 두 원천이 같은
    # 칼럼을 다른 뜻으로 쓰면 필터가 조용히 엉뚱한 행을 집는다.
    r["ch"] = ORGCAT_CH_ALL
    return r[ORGCAT_COLS]


def detail_provider(sdf, gran, axis):
    """`cube(path)` — 그 경로 주변만 담은 조직×카테고리 프레임을 돌려주는 함수.

    상품까지 전부 펼치면 (기간 × 채널 × 상품)이 수백만 행이라 화면이 멈춘다. 화면이 실제로
    읽는 건 ①경로 각 단계의 형제 ②지금 노드의 자식 ③손자(④ 히트맵)뿐이라, 그만큼만
    만들면 결과는 같고 크기는 수천 행이다. 파고드는 동안 겹치는 겹은 memo로 재사용한다.
    """
    memo = {}

    def cube(path):
        P = tuple(path)
        need = []
        for depth in range(0, min(len(P) + 2, len(DETAIL_LV)) + 1):
            key = (P[:max(0, depth - 1)], depth)
            if key not in memo:
                memo[key] = detail_level(sdf, gran, axis, key[0], depth)
            need.append(memo[key])
        if not need:
            return pd.DataFrame(columns=ORGCAT_COLS)
        return (pd.concat(need, ignore_index=True)
                .drop_duplicates(subset=ORGCAT_KEY, keep="last"))

    return cube


# ══════════════════════════════════════════════════════
# 02. 첫구매 퍼널별 상세 실적
# ══════════════════════════════════════════════════════
# 「01 요약」에 얹혀 있던 전환 퍼널·채널 기여 분해를 빼내어, 퍼널 한 줄을 끝까지 따라가며
# 진단하는 화면으로 다시 짰다. 읽는 순서가 곧 화면 순서다 — ① 합계 퍼널로 전체 흐름을
# 보고 → ② 그게 추세인지 연중 흐름으로 확인하고 → ③ 어느 채널이 어느 단계에서 빠지는지
# 한 표로 훑고 → ④ 문제 단계의 채널 기여를 분해하고 → ⑤ 첫구매 3지표는 조직 > 카테고리로
# 한 단계 더 내려간다 → ⑥ 앱 수신동의.
#
# **비율 두 칸의 출처가 다르다.** 가입율은 두 카운트에서 계산하고(그래야 채널로 쪼갤 때
# 합계와 같은 규칙이 된다), 당일가입CR은 파일 값을 그대로 쓴다 — 분자인 '당일가입 첫구매
# 고객수'가 데이터에 없어 역산할 수 없다. 그래서 `가입자수 × 당일가입CR ≠ 첫구매 고객수`고,
# 마지막 칸은 과거 가입자까지 포함한 **전체 첫구매**다. 곱셈이 안 닫히는 걸 화면에 밝힌다.
FUNNEL_STEPS = ["비회원트래픽", "가입율", "가입자수", "당일가입CR",
                "첫구매 고객수", "첫구매 객단가", "첫구매 거래액"]
FUNNEL_TAIL = ["첫구매 고객수", "첫구매 객단가", "첫구매 거래액"]
# ④에서 고를 수 있는 지표 — 거래액이 왜 빠졌는지는 `거래액 = 상품UV × 상품CR × 객단가`
# 세 요인을 다 봐야 갈린다(원본에서 오차 0.000%로 성립하는 항등식이다).
FUNNEL_OC_METS = FUNNEL_TAIL + ["상품UV", "상품CR"]
# 조직을 가로질러 모을 수 있는 지표. 그냥 더해도 되는 건 거래액 하나뿐이고
# 나머지는 **다시 만든다** — 객단가 = 거래액 합 ÷ 고객수 합,
# 상품CR = 고객수 합 ÷ 상품UV 합(`거래액 = 상품UV × 상품CR × 객단가`와
# `거래액 = 고객수 × 객단가`에서 나오는 항등식이다).
FUNNEL_ROLLUP_METS = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가",
                      "상품UV", "상품CR"]
# 채널 합 ≈ 전체가 성립하는 가산 지표 — 기여도 분해는 여기서만 성립한다
FUNNEL_ADDITIVE = ["비회원트래픽", "가입자수", "첫구매 고객수", "첫구매 거래액"]
# 값 규칙(FUNNEL_DERIVED·FUNNEL_FILE_FIRST·funnel_val)은 조회 계층 옆으로 옮겼다 —
# 추이표·차트도 같은 규칙을 타야 페이지마다 숫자가 갈리지 않는다.
# 화면 전체가 **이 하나**를 본다 — ①의 카드부터 ⑥의 앱 추이까지. 예전엔 위의
# 「비교 기준」과 ②·⑤의 「기간 단위」가 따로 놀아 단위 하나 바꾸려고 세 군데를 눌러야 했다.
# 「03」 차트에 기본으로 올릴 채널 — 여덟 줄을 다 켜면 엉켜서 안 읽힌다. 평소 맞대 보는
# 넷만 켜 두고 나머지는 필요할 때 켠다(데이터에 없는 건 자동으로 빠진다).
CHART_CH_DEFAULT = ["직접", "광고", "EP", "제휴"]
# 첫구매 고객 세그먼트 — 「05」가 채널별과 **같은 함수**를 쓴다(축만 다르다).
# 색은 목록 순서로 굳힌다 — 고르고 빼도 남은 선의 색이 안 바뀌어야 맞댈 수 있다.
SEGMENTS = ["1_신규", "1_당월신규", "2_기가입신규", "3_기존"]
SEG_PAL = {s: ORGCAT_PAL[i % len(ORGCAT_PAL)] for i, s in enumerate(SEGMENTS)}
# 일자별의 기본 조회 기간 — 하루는 퍼널이 비는 날이 흔하고, 3주면 주차별과 겹쳐 읽힌다.
DAY_RANGE_DEFAULT = 21
FUNNEL_GRAN_LABEL = {"일": "일자별", "주": "주차별", "월": "월별"}
FUNNEL_GRAN_ORDER = ["일", "주", "월"]
# 페이지 이름은 **여기 한곳**에서만 적는다. 번호는 `PAGE_ORDER`로 목록을 만들 때만
# 붙이고, 분기는 번호를 뗀 이름으로 가린다.
#
# 예전엔 목록과 분기가 번호까지 박은 문자열을 **따로** 들고 있어서, 페이지를 옮길 때마다
# 두 곳을 같이 고쳐야 했다. 한쪽을 빠뜨리면 그 페이지는 **예외 없이 빈 화면**이 된다 —
# `page.startswith("09.")` 예외가 실제로 그렇게 끊겼고, 증상이 안 드러난다.
# 이제 목록과 분기가 같은 상수를 보므로 번호를 바꿔도 어긋날 수가 없다
# (`smoke_weekly_report.py`가 페이지마다 본문 `##` 제목이 나오는지 본다).
PAGE_SUMMARY = "주간보고 요약"
PAGE_FUNNEL = "첫구매 퍼널별 상세 실적"
PAGE_ORGCAT_CH = "조직·카테고리·채널별 첫구매 상세"
PAGE_CHANNEL = "채널별 실적"
PAGE_PUSH = "앱푸시 동의 현황"
PAGE_SEGMENT = "첫구매 고객 세그먼트 성과"
PAGE_ORGCAT = "조직·카테고리별 실적"
PAGE_DOWNLOAD = "통합 데이터·다운로드"

# 화면에 거는 순서. 「월별 추이」·「주차별 추이」는 「02」가 기간 단위를 갖게 되면서
# 같은 걸 두 벌 보여 주게 돼 접었고, 받아 가는 화면인 「통합 데이터」는 맨 뒤다.
PAGE_ORDER = [PAGE_SUMMARY, PAGE_FUNNEL, PAGE_ORGCAT_CH, PAGE_CHANNEL,
              PAGE_PUSH, PAGE_SEGMENT, PAGE_ORGCAT, PAGE_DOWNLOAD]
PAGES = [f"{i:02d}. {nm}" for i, nm in enumerate(PAGE_ORDER, 1)]


def page_name(label):
    """화면이 고른 페이지 라벨에서 **번호를 뗀 이름**. 분기는 전부 이걸로 가린다."""
    return label.split(". ", 1)[-1] if label else ""


def gran_options(df, metrics):
    """지표가 **실제로 있는** 기간 단위만. 없는 단위를 올리면 눌렀을 때 빈 화면이 된다.

    앱푸시 수신동의는 원천이 일자 헤더 표라 **늘** 일별로 쌓인다 — `gran=="일"`만 보면
    정작 볼 지표가 없는데도 일자별이 켜진다. 그래서 지표를 지정해 받는다.
    """
    m = df["metric"].isin(metrics) & df["value"].notna()
    return [g for g in FUNNEL_GRAN_ORDER if (m & (df["gran"] == g)).any()]


def gran_ref(df, gran, ref_year, ref_month, wy, wlabel, metrics, daykey, box=None):
    """고른 단위의 **기준 기간 한 벌**. 못 잡으면 None(화면엔 이유를 남긴다).

    두 화면(「02」·「03」)이 같은 규칙을 봐야 한다 — 따로 짜면 같은 단위인데 기준 기간이
    갈려서 숫자가 어긋난다. 비교 기준도 단위에서 따라 나온다:
    일=전년 같은 날 · 주=전년 동주 · 월=전년 동월 MTD.

    사이드바 「기준 기간」엔 연·월·주차뿐이라 **일자만 화면에서 고른다**(`daykey`가 그
    위젯 키다 — 화면마다 달라야 서로의 선택을 안 덮는다).
    """
    if gran == "주":
        if not wlabel:
            st.info("주차 데이터가 없어요. 다른 기간 단위를 골라 주세요.")
            return None
        return dict(clabel=wlabel, win=None, cy=wy, py=wy - 1,
                    period_lbl=week_disp(wy, wlabel), base_lbl="전년 동주",
                    base_tag="전년동주", prv_close="final",
                    x_prv=f"{wy - 1}년", x_cur=f"{wy}년")
    if gran == "월":
        return dict(clabel=month_label(ref_month), win=None,
                    cy=ref_year, py=ref_year - 1,
                    period_lbl=f"{ref_year}년 {ref_month}월 누적(MTD)",
                    base_lbl="전년 동월 MTD", base_tag="전년동월",
                    # 전년 동월도 동일기간(MTD)으로 잘린 값 우선 — 「01」 요약 표와 같은 기준
                    prv_close="mtd",
                    x_prv=f"{ref_year - 1}년 {ref_month}월",
                    x_cur=f"{ref_year}년 {ref_month}월")
    days = (df[(df["gran"] == "일") & (df["year"] == ref_year)
               & df["metric"].isin(metrics) & df["value"].notna()]
            [["label", "sortkey"]].drop_duplicates()
            .sort_values("sortkey")["label"].tolist())
    if not days:
        st.info(f"{ref_year}년은 일자별 데이터가 없어요. 다른 기간 단위를 골라 주세요.")
        return None
    # **하루가 아니라 기간으로 고른다.** 하루치는 퍼널이 통째로 비는 날이 흔해서
    # (①이 '데이터가 부족해요'로 끝난다) 기준일 하나로는 화면이 안 선다. 기본은
    # **최근 3주**(`DAY_RANGE_DEFAULT`) — 주차별과 겹쳐 보기 좋은 길이다.
    lo, hi = guard_range(daykey, days,
                         default=(days[max(0, len(days) - DAY_RANGE_DEFAULT)], days[-1]))
    _box = box if box is not None else st
    lo, hi = _box.select_slider("조회 기간", options=days, value=(lo, hi), key=daykey,
                                help="사이드바엔 일자 선택이 없어서 여기서 골라요. "
                                     f"기본은 최근 {DAY_RANGE_DEFAULT}일이에요. "
                                     "값은 고른 기간의 일평균이에요.")
    i, j = days.index(lo), days.index(hi)
    clabel = days[min(i, j):max(i, j) + 1]
    disp = lbl_disp(clabel)
    # **고른 기간이 곧 아래 추이의 창이다.** 위에서 3주를 골라 놓고 아래 차트만 연중이면
    # 같은 화면이 두 기간을 말하게 된다. 주·월은 범위 선택이 없어 `win=None`(연중)이다.
    return dict(clabel=clabel, win=list(clabel), cy=ref_year, py=ref_year - 1,
                period_lbl=f"{ref_year}년 {disp}",
                base_lbl="전년 같은 기간" if len(clabel) > 1 else "전년 같은 날",
                base_tag="전년동일", prv_close="final",
                x_prv=f"{ref_year - 1}년 {disp}", x_cur=f"{ref_year}년 {disp}")


# ② 추이에 기본으로 올릴 지표 — 퍼널을 앞에서 뒤로 훑는 순서 그대로다. 객단가는 빼 뒀다
# (거래액·고객수가 이미 있어 셋을 다 켜면 차트가 여덟 장이 된다). 필요하면 골라서 켠다.
FUNNEL_TREND_DEFAULT = ["비회원트래픽", "가입율", "가입자수", "당일가입CR",
                        "첫구매 고객수", "첫구매 거래액"]
# 표에 보일 기간 수 — 일은 한 달치, 주차는 넉 달치, 월은 한 해치
FUNNEL_TREND_KEEP = {"일": 31, "주": 16, "월": 12}
# 하단 앱 블록 — 앱설치는 아직 원천이 안 올라와서, 없으면 '–'로 비우고 왜인지 밝힌다
APP_STEPS = ["가입자수", "앱설치", "앱푸시수신동의"]
# 원천이 일자 헤더 표라 `gran='일'`로만 쌓이는 지표 — 주·월은 일평균으로 묶어야 한다
PUSH_DAILY_ONLY = {"앱푸시수신동의", "앱푸시동의_앱무관", "앱푸시대상_앱무관"}
# 앱설치 원천이 같이 주는 나머지 칸 — 카드 세 장 아래 접이식으로만 보여 준다.
# 스토어 방문 → 설치 → 삭제까지가 한 흐름이라 같이 봐야 '설치가 준 건지 삭제가 는 건지'가 갈린다.
APP_EXTRA = ["앱_전체설치", "앱_재설치", "앱_삭제", "앱_순증설치",
             "앱_스토어방문", "앱_Push활성기기"]


def push_period_avg(pdf, year, label, metric="앱푸시수신동의"):
    """일별로만 쌓이는 지표의 그 기간 **일평균**과 근사 여부.

    이 지표는 원천이 일자 헤더 표라 **일별로만** 쌓인다(`gran='일'`, `label='M/D'`).
    주·월 칸을 그냥 집으면 늘 비므로 일별 값을 묶어 준다. 화면의 다른 값이 전부
    일평균이라 합이 아니라 평균으로 맞춘다.

    달은 라벨에서 그대로 읽어 정확하지만, **주는 원천에 주차가 없다**. `(일-1)//7+1`로
    나누는데 마스터의 주차와 경계가 어긋날 수 있어, 근사라는 사실을 같이 돌려준다.

    돌려주는 건 `(값, 근사여부, 며칠치)`. 며칠치는 화면이 '이게 몇 날 평균인지'를 밝히는
    데 쓴다 — 안 밝히면 3일치와 30일치가 같은 얼굴로 나란히 선다.
    """
    m = re.match(r"(\d{1,2})\s*월(?:\s*(\d)\s*주차)?", str(label or ""))
    if not m:
        return np.nan, False, 0
    d = pdf[(pdf["metric"] == metric) & (pdf["gran"] == "일")
            & (pdf["segment"] == "*TOTAL") & (pdf["year"] == year)]
    d = d[d["label"].astype(str).str.contains("/", na=False)]
    if d.empty:
        return np.nan, False, 0
    d = mask_push_spikes(d.copy())
    parts = d["label"].astype(str).str.split("/")
    mo = pd.to_numeric(parts.str[0], errors="coerce")
    dy = pd.to_numeric(parts.str[1], errors="coerce")
    sel = mo == int(m.group(1))
    approx = False
    if m.group(2):
        sel &= ((dy - 1) // 7 + 1) == int(m.group(2))
        approx = True
    v = pd.to_numeric(d.loc[sel.fillna(False), "value"], errors="coerce").dropna()
    return (float(v.mean()) if len(v) else np.nan), approx, len(v)


def _fn_pill(met, cur, prv, tag):
    """카드 아래 전년비 pill 한 조각 — 없으면 '–'로 자리를 지킨다."""
    d = fmt_delta(met, cur, prv)
    if not d:
        return f'<div class="kpi-delta na">– ({tag})</div>'
    return f'<div class="kpi-delta {"down" if d.startswith("△") else "up"}">{d} ({tag})</div>'


def _fn_rate_cell(label, cur, prv, tag):
    """단계 사이 화살표 칸 — 비율은 %p로 읽는 게 맞아서 따로 그린다."""
    cur_s = "–" if pd.isna(cur) else f"{cur * 100:.2f}%"
    sub = ""
    if not (pd.isna(cur) or pd.isna(prv)):
        pp = (cur - prv) * 100
        colr = "#dc2626" if pp < 0 else "#16a34a"
        sign = "△" if pp < 0 else "+"
        sub = (f'<div style="font-size:11px;color:#64748b">전년 {prv * 100:.2f}% '
               f'<span style="color:{colr};font-weight:700">{sign}{abs(pp):.2f}%p</span></div>')
    return (f'<div style="flex:0.9;display:flex;flex-direction:column;justify-content:center;'
            f'text-align:center;min-width:0">'
            f'<div style="color:#94a3b8;font-size:16px;line-height:1">→</div>'
            f'<div style="font-size:12px;color:#64748b;margin-top:2px">{esc(label)}</div>'
            f'<div style="font-size:17px;font-weight:700;color:#1e293b">{cur_s}</div>'
            f'{sub}</div>')


def render_channel_page(df, ref_year, ref_month, wy, wlabel, ch_sel):
    """04. 채널별 실적 — 축이 채널인 판."""
    render_axis_page(df, ref_year, ref_month, wy, wlabel,
                     title="채널별 실적", axis="채널", kp="wr_ch",
                     opts=[c for c in CHANNELS if c in ch_sel],
                     pal=CHANNEL_PAL, dft=CHART_CH_DEFAULT)


def render_axis_page(df, ref_year, ref_month, wy, wlabel, *,
                     title, axis, kp, opts, pal, dft, mets=None):
    """축(채널·세그먼트) 하나를 놓고 **지표를 고르지 않고 다 보는** 화면.

    예전엔 두 화면이 따로 짜여 있었다 — 채널별은 지표 셀렉트 하나에 월 고정, 세그먼트는
    주·월 표 두 개가 박혀 있었다. 지표를 갈아 끼우며 보게 되니 항목 사이의 이야기가 안
    이어졌고, 두 화면의 얼굴이 달라 오갈 때마다 다시 찾아야 했다.

    구성은 「02」와 같다 — **기간 단위 하나**(일/주/월), 표는 **왼쪽 실적 · 오른쪽
    전년비**, 차트는 지표마다 한 장씩(선 = 축 항목). **한 구현을 둘이 나눠 쓴다** —
    복사해 두면 한쪽만 고쳐져 조용히 갈린다.

    `kp`가 위젯 키 접두어다 — 화면마다 달라야 서로의 선택을 안 덮는다.
    """
    st.markdown(f"## {title}")
    mets = [m for m in (mets or METRICS7) if (df["metric"] == m).any()]
    if not mets:
        st.info("실적 지표가 없어요. 실적 파일을 올려 주세요.")
        return
    gr_avail = gran_options(df, mets)
    if not gr_avail:
        st.info("실적 데이터가 없어요.")
        return
    gsel, dsel = st.columns([1.3, 1.7])
    with gsel:
        guard_select(f"{kp}_gran", gr_avail,
                     default="월" if "월" in gr_avail else gr_avail[-1])
        gran = st.radio("기간 단위", gr_avail, horizontal=True, key=f"{kp}_gran",
                        format_func=lambda g: FUNNEL_GRAN_LABEL[g],
                        help="표와 차트가 같이 이 단위를 따라가요.")
    ref = gran_ref(df, gran, ref_year, ref_month, wy, wlabel, mets,
                   f"{kp}_day", box=dsel)
    if ref is None:
        return
    clabel, cy, py = ref["clabel"], ref["cy"], ref["py"]
    segs = ["*TOTAL"] + list(opts)

    # ── 표 — 왼쪽 실적 · 오른쪽 전년비 ──
    st.subheader(f"{title} — {ref['period_lbl']}")
    st.caption(f"{ref['base_lbl']} 대비예요. 왼쪽에 실적, 오른쪽에 전년비를 모아 뒀어요 — "
               f"전년비만 가로로 훑으면 어느 {axis}이 빠지는지 한눈에 보여요. "
               "비율 지표(가입율·당일가입CR)는 %p 차이예요.")
    rows = []
    for seg in segs:
        row = {axis: "전체" if seg == "*TOTAL" else seg}
        for met in mets:
            cv = report_val(df, gran, met, seg, cy, clabel, "mtd")
            pv = report_val(df, gran, met, seg, py, clabel, ref["prv_close"])
            row[met] = fmt_value(met, cv)
            row[f"{met} 전년비"] = fmt_delta(met, cv, pv) or "–"
        rows.append(row)
    tbl = pd.DataFrame(rows).set_index(axis)
    # 실적을 다 놓고 전년비를 오른쪽에 몬다 (「02」 ②와 같은 규칙)
    tbl = tbl[[m for m in mets] + [f"{m} 전년비" for m in mets]]
    wtable(style_delta_cols(tbl), width="stretch",
           dl_name=f"{title} ({ref['period_lbl']})")

    # ── 차트 — 지표마다 한 장, 선 = 축 항목 ──
    # **차트에서 항목을 넣고 뺀다.** 사이드바 선택만 따라가면 표까지 같이 좁혀지는데,
    # 표는 다 놓고 보면서 차트만 두셋으로 줄여 맞대고 싶은 경우가 대부분이다.
    # **연도도 켜고 끈다** — 전년 선이 있어야 '이번이 낮은 건지 원래 그런 건지'가 갈린다.
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader(f"{axis}별 {FUNNEL_GRAN_LABEL[gran]} 추이")
    # **값이 실제로 있는 항목만** 올린다 — 눌러도 아무 선이 안 생기는 선택지는 '왜 안
    # 그려지지'만 남긴다(기간 단위를 데이터 있는 것만 올리는 것과 같은 규칙).
    _live_ch = set(df[(df["gran"] == gran) & df["metric"].isin(mets)
                      & df["value"].notna()]["segment"].astype(str))
    _chan_opts = [c for c in opts if c in _live_ch]
    if not _chan_opts:
        st.info(f"이 기간 단위엔 {axis}별 값이 없어요. 다른 단위를 골라 주세요.")
        return
    _cc, _yc = st.columns([2.4, 1])
    with _cc:
        guard_multi(f"{kp}_chart_ch", _chan_opts)
        if f"{kp}_chart_ch" not in st.session_state:
            # 전부 켜면 선이 여덟 줄이라 엉킨다. 평소 맞대 보는 것만 기본으로 두고
            # 나머지는 켜서 본다(데이터에 없는 건 자동으로 빠진다).
            _pre = [c for c in dft if c in _chan_opts]
            st.session_state[f"{kp}_chart_ch"] = _pre or _chan_opts
        chart_ch = st.multiselect(f"차트에 올릴 {axis}", _chan_opts,
                                  key=f"{kp}_chart_ch",
                                  help="표는 그대로 두고 차트만 좁혀 봐요. "
                                       f"{axis} 색은 어느 걸 골라도 그대로예요.")
    with _yc:
        st.caption("연도")
        _y1 = st.checkbox(f"{cy}년", value=True, key=f"{kp}_chart_cy")
        _has_p = not df[(df["gran"] == gran) & (df["year"] == py)].empty
        _y2 = (st.checkbox(f"{py}년", value=False, key=f"{kp}_chart_py",
                           help="전년은 얇은 점선이에요.") if _has_p else False)
    chart_yrs = [y for y, on in ((py, _y2), (cy, _y1)) if on]
    if not chart_ch or not chart_yrs:
        st.info(f"차트에 올릴 {axis if not chart_ch else '연도'}를 골라 주세요.")
        return
    st.caption(f"지표를 고르지 않고 다 그려요. **색은 {axis}, 선 모양은 연도**예요 — "
               f"{cy}년은 실선, 전년은 얇은 점선이에요. "
               + ("전년 데이터가 없어서 올해만 그려요." if not _has_p else ""))

    # x축은 **올해 라벨**로 세운다. 다만 고른 채널·연도 어디에도 값이 없는 칸은 뺀다 —
    # 그 라벨은 *모든 지표·모든 채널*이 쓴 것이라 지금 보는 것과 무관한 칸이 섞이고,
    # 그 자리에서 모든 선이 나란히 끊겨 데이터가 빠진 것처럼 보인다.
    _x_all = labels_sorted(df, gran, [cy])
    if ref.get("win") is not None:
        _w = set(as_labels(ref["win"]))
        _x_all = [x for x in _x_all if str(x) in _w]
    if not _x_all:
        st.info(f"{cy}년 «{FUNNEL_GRAN_LABEL[gran]}» 값이 없어요.")
        return
    _ser = {(met, seg, yr): report_series(df, gran, met, seg, yr, "final").reindex(_x_all)
            for met in mets for seg in chart_ch for yr in chart_yrs}
    _x = [lb for lb in _x_all if any(not pd.isna(v.get(lb, np.nan)) for v in _ser.values())]
    if not _x:
        st.info("고른 채널·연도에 값이 없어요.")
        return
    for i in range(0, len(mets), 3):
        for _col, met in zip(st.columns(3), mets[i:i + 3]):
            with _col:
                unit, div = METRIC_UNIT.get(met, ("", 1))
                if met in PCT_METRICS:
                    div, unit = 0.01, "%"
                fig = go.Figure()
                for seg in chart_ch:
                    for yr in chart_yrs:
                        s = _ser[(met, seg, yr)].reindex(_x)
                        if s.dropna().empty:
                            continue
                        _cur = yr == cy
                        # 한쪽 해에만 없는 칸은 남겨서 선을 끊는다 — 이으면 없는 데이터를
                        # 지어낸다(`connectgaps=False`).
                        fig.add_trace(go.Scatter(
                            x=[month_trim(v) for v in s.index],
                            y=[None if pd.isna(v) else v / div for v in s],
                            mode="lines+markers",
                            name=seg if _cur else f"{seg} ({yr})",
                            connectgaps=False,
                            legendgroup=seg,
                            line=dict(color=clr(pal.get(seg, "blue")),
                                      width=1.8 if _cur else 1.1,
                                      dash=None if _cur else "dot"),
                            marker=dict(size=4 if _cur else 3)))
                ly = base_layout(300, ysuffix=unit if unit == "%" else "",
                                 title=f"{met} ({unit})")
                ly["xaxis"]["categoryorder"] = "array"
                ly["xaxis"]["categoryarray"] = [month_trim(v) for v in _x]
                # 일별은 라벨이 한 해 365개라 그대로 두면 축이 새까매진다
                if gran in ("주", "일"):
                    ly["xaxis"]["tickangle"] = -45
                    ly["xaxis"]["nticks"] = 20 if gran == "주" else 14
                fig.update_layout(**ly)
                st.plotly_chart(fig, width="stretch")

    # ── 채널 × 기간 표 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader(f"{axis} × 기간 — {cy}년")
    guard_select(f"{kp}_xmet", mets)
    xmet = st.selectbox("표로 볼 지표", mets, key=f"{kp}_xmet",
                        help="위 차트는 다 그리고, 이 표만 한 지표를 자세히 봐요.")
    # 차트에서 좁힌 항목은 **표에 안 옮긴다** — 표는 다 놓고 보면서 차트만 줄이는 게
    # 이 화면의 쓰임새다. 그래서 차트가 값 없는 칸을 뺀 `_x`가 아니라 `_x_all`을 쓴다.
    # **자르지 않는다 — 올해 전체다.** 최근 N개만 보여 주면 '어느 기간부터 꺾였나'를
    # 보려고 매번 엑셀을 받아야 한다(②·⑤와 같은 규칙).
    keep_x = _x_all
    xrows, xdlt = {}, {}
    for seg in segs:
        nm = "전체" if seg == "*TOTAL" else seg
        cs = report_series(df, gran, xmet, seg, cy, "final")
        ps = report_series(df, gran, xmet, seg, py, "final")
        xrows[nm] = {}
        xdlt[nm] = {}
        for lb in keep_x:
            v, pv = cs.get(lb, np.nan), ps.get(lb, np.nan)
            xrows[nm][month_trim(lb)] = fmt_value(xmet, v) if not pd.isna(v) else "–"
            xdlt[nm][f"{month_trim(lb)} 전년비"] = fmt_delta(xmet, v, pv) or "–"
    # 실적을 다 놓고 전년비를 오른쪽에 몬다 — 위 표·②·⑤와 같은 얼굴이다
    xtbl = pd.concat([pd.DataFrame(xrows).T, pd.DataFrame(xdlt).T], axis=1)
    xtbl.index.name = axis
    wtable(style_delta_cols(xtbl), width="stretch",
           dl_name=f"{axis}×기간 {xmet} ({cy}년)")
    st.caption(f"왼쪽은 {cy}년 실적, 오른쪽은 **{py}년 같은 기간 대비 전년비**예요. "
               f"{cy}년 **{len(keep_x)}개 기간 전체**를 담았어요 — 옆으로 밀어서 보세요.")


def render_funnel_page(df, odf, ref_year, ref_month, wy, wlabel):
    """02. 첫구매 퍼널별 상세 실적 — 퍼널 한 줄을 끝까지 따라가며 진단하는 화면."""
    st.markdown("## 첫구매 퍼널별 상세 실적")
    # ── 기간 단위 — 이 화면의 유일한 단위 스위치 ──
    # 퍼널 지표가 **실제로 있는** 단위만 올린다. `gran=="일"`만 보면 앱푸시 수신동의가
    # 늘 일별로 쌓이는 탓에 퍼널 데이터가 없는데도 일자별이 켜진다.
    gr_avail = gran_options(df, FUNNEL_STEPS)
    if not gr_avail:
        st.info("퍼널 지표 데이터가 없어요. 실적 파일을 올려 주세요.")
        return
    gsel, dsel = st.columns([1.3, 1.7])
    with gsel:
        guard_select("wr_fn_gran", gr_avail,
                     default="주" if "주" in gr_avail else gr_avail[-1])
        gran = st.radio("기간 단위", gr_avail, horizontal=True, key="wr_fn_gran",
                        format_func=lambda g: FUNNEL_GRAN_LABEL[g],
                        help="이 화면 전체가 이 단위를 따라가요 — 카드·추이·채널·조직·앱까지.")
    ref = gran_ref(df, gran, ref_year, ref_month, wy, wlabel,
                   FUNNEL_STEPS, "wr_fn_day", box=dsel)
    if ref is None:
        return
    clabel, cy, py = ref["clabel"], ref["cy"], ref["py"]
    period_lbl, base_lbl, base_tag = ref["period_lbl"], ref["base_lbl"], ref["base_tag"]
    prv_close, x_prv, x_cur = ref["prv_close"], ref["x_prv"], ref["x_cur"]
    win = ref.get("win")          # 「조회 기간」 — 아래 추이 전부가 이 창을 본다

    def gcur(met, seg="*TOTAL"):
        return pick(df, gran, met, seg, cy, clabel, "mtd")

    def gprv(met, seg="*TOTAL"):
        return pick(df, gran, met, seg, py, clabel, prv_close)

    # 기간을 묶어 볼 땐 파일 우선을 끈다 — 객단가를 날마다 평균 내면 거래가 적은 날이
    # 많은 날과 같은 무게를 갖는다(`funnel_val` 주석 참고). 한 칸일 땐 값이 같다.
    _ff = not is_range(clabel)

    def vcur(met, seg="*TOTAL"):
        return funnel_val(gcur, met, seg, file_first=_ff)

    def vprv(met, seg="*TOTAL"):
        return funnel_val(gprv, met, seg, file_first=_ff)

    st.caption(f"{period_lbl} vs {base_lbl} · 값은 모두 **일평균**이에요."
               + (f" {len(as_labels(clabel))}일치를 묶은 값이에요."
                  if is_range(clabel) else ""))

    # ── ① 전체 퍼널 ────────────────────────────────────
    st.subheader("① 전체 퍼널")
    trf_c, trf_p = vcur("비회원트래픽"), vprv("비회원트래픽")
    jn_c, jn_p = vcur("가입자수"), vprv("가입자수")
    if pd.isna(trf_c) and pd.isna(jn_c):
        st.info("이 기간은 퍼널 데이터가 부족해요. 사이드바에서 다른 기준 기간을 고르거나 "
                "비회원 트래픽·가입자수 파일을 올려 주세요.")
    else:
        cells = [f'<div class="kpi-card" style="flex:1.1;min-width:0">'
                 f'<div class="kpi-label">비회원트래픽</div>'
                 f'<div class="kpi-value">{fmt_value("비회원트래픽", trf_c)}</div>'
                 f'{_fn_pill("비회원트래픽", trf_c, trf_p, base_tag)}</div>',
                 _fn_rate_cell("가입율", vcur("가입율"), vprv("가입율"), base_tag),
                 f'<div class="kpi-card" style="flex:1.1;min-width:0">'
                 f'<div class="kpi-label">가입자수</div>'
                 f'<div class="kpi-value">{fmt_value("가입자수", jn_c)}</div>'
                 f'{_fn_pill("가입자수", jn_c, jn_p, base_tag)}</div>',
                 _fn_rate_cell("당일가입 첫구매율", vcur("당일가입CR"), vprv("당일가입CR"),
                               base_tag)]
        # 마지막 칸은 고객수를 머리로 두고 객단가·거래액을 아래에 붙인다 — 셋이 한 묶음이라
        # 카드를 셋으로 쪼개면 퍼널의 '끝'이 어디인지 안 읽힌다.
        head_c, head_p = vcur("첫구매 고객수"), vprv("첫구매 고객수")
        subs = ""
        for _m in ("첫구매 객단가", "첫구매 거래액"):
            _c, _p = vcur(_m), vprv(_m)
            _d = fmt_delta(_m, _c, _p)
            _dtxt = ""
            if _d:
                _dc = "#dc2626" if _d.startswith("△") else "#16a34a"
                _dtxt = f' <span style="color:{_dc};font-weight:700">{_d}</span>'
            subs += (f'<div style="font-size:12px;color:#64748b;margin-top:2px">'
                     f'{esc(_m.replace("첫구매 ", ""))} '
                     f'<b style="color:#1e293b">{fmt_value(_m, _c)}</b>{_dtxt}</div>')
        cells.append(f'<div class="kpi-card" style="flex:1.6;min-width:0">'
                     f'<div class="kpi-label">첫구매 고객수</div>'
                     f'<div class="kpi-value">{fmt_value("첫구매 고객수", head_c)}</div>'
                     f'{_fn_pill("첫구매 고객수", head_c, head_p, base_tag)}{subs}</div>')
        st.markdown('<div style="display:flex;gap:10px;align-items:stretch">'
                    + "".join(cells) + '</div>', unsafe_allow_html=True)
        st.caption("**가입율**은 `가입자수 ÷ 비회원트래픽`으로 계산해요. "
                   "**당일가입 첫구매율**은 원천 파일 값을 그대로 써요 — 분자인 당일가입 "
                   "첫구매 고객수가 데이터에 없어서예요. 그래서 `가입자수 × 당일가입 "
                   "첫구매율`이 오른쪽 첫구매 고객수와 안 맞아요. 마지막 칸은 과거 가입자까지 "
                   "포함한 **전체 첫구매**거든요.")

    # ── ② 퍼널 지표 추이 ───────────────────────────────
    # ①은 한 기간의 사진이라 '이번이 낮은 건지 원래 그런 건지'를 못 본다. 바로 여기서
    # 같은 지표의 연중 흐름을 전년과 맞대 본 뒤 ③ 이하로 파고든다.
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("② 퍼널 지표 추이")
    st.caption("①이 한 기간의 사진이라면 여기는 흐름이에요. 이상한 구간을 찾으면 "
               "사이드바에서 그 기간으로 옮겨 다시 보세요.")
    _render_funnel_trend(df, gran, cy, py, win)

    # ── ③ 채널별 퍼널 ──────────────────────────────────
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("③ 채널별 퍼널")
    st.caption("어느 채널이 **어느 단계에서** 빠지는지 한 표로 훑어요.")
    chans = [c for c in CHANNELS
             if not df[(df["gran"] == gran) & (df["segment"] == c)
                       & (df["label"].astype(str).isin(as_labels(clabel)))].empty]
    if not chans:
        st.info("이 기간은 채널별 데이터가 없어요. 전체 값만 있어요.")
    else:
        rows = []
        for seg, name in [("*TOTAL", "전체")] + [(c, c) for c in chans]:
            row = {"채널": name}
            for stp in FUNNEL_STEPS:
                cv, pv = vcur(stp, seg), vprv(stp, seg)
                row[stp] = fmt_value(stp, cv)
                row[f"{stp} 전년비"] = fmt_delta(stp, cv, pv) or "–"
            rows.append(row)
        tbl = pd.DataFrame(rows).set_index("채널")
        wtable(style_delta_cols(tbl), width="stretch",
               dl_name=f"채널별 퍼널 상세 ({period_lbl})")
        st.caption("가로로 스크롤하면 뒤 단계까지 볼 수 있어요. "
                   "**가입율·객단가는 채널별 두 카운트로 계산**하고, "
                   "**당일가입 첫구매율은 채널별 파일 값**이라 원천에 없으면 '–'로 비어요. "
                   "비율 지표는 채널 합이 전체와 다른 게 정상이에요.")

    # ── ④ 채널별 증감 ──────────────────────────────────
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("④ 채널별 증감")
    # 지표는 **일곱 단계 전부** 고를 수 있다. 다만 그리는 그림이 갈린다 —
    # 가산 지표(발송·고객수·거래액)는 부분의 합이 전체라 워터폴이 성립하지만,
    # 비율·평균(가입율·당일가입CR·객단가)은 채널 합 ≠ 전체라 워터폴이 거짓말이 된다.
    # 그래서 비율은 **채널마다 얼마나 움직였는지**를 나란히 세운다.
    def _has_ch(m):
        return any(not pd.isna(vcur(m, c)) or not pd.isna(vprv(m, c)) for c in CHANNELS)

    avail_dec = [m for m in FUNNEL_STEPS if _has_ch(m)]
    if not avail_dec:
        st.info("채널별로 나눠 볼 지표 데이터가 없어요.")
    else:
        # 기본은 **첫구매 거래액** — 퍼널이 결국 설명하려는 결과 지표다.
        # 목록 순서대로 두면 맨 위 비회원트래픽으로 열려 매번 바꿔야 한다.
        guard_select("wr_decomp_met", avail_dec, default="첫구매 거래액")
        dec_met = st.selectbox("지표", avail_dec, key="wr_decomp_met",
                               help="고른 지표가 전년 대비 얼마나 움직였는지 채널별로 봐요.")
        if dec_met not in FUNNEL_ADDITIVE:
            _pct = dec_met in PCT_METRICS
            _suf, _fmt = ("%p", "{:+.2f}") if _pct else ("원", "{:+,.0f}")
            _lab, _dlt, _rows = [], [], []
            for chn in CHANNELS:
                cv, pv = vcur(dec_met, chn), vprv(dec_met, chn)
                if pd.isna(cv) and pd.isna(pv):
                    continue
                d = ((cv - pv) * (100 if _pct else 1)
                     if not (pd.isna(cv) or pd.isna(pv)) else np.nan)
                _lab.append(chn); _dlt.append(d)
                _rows.append({"채널": chn, f"{py}년": fmt_value(dec_met, pv),
                              f"{cy}년": fmt_value(dec_met, cv),
                              "전년비": fmt_delta(dec_met, cv, pv) or "–"})
            if not _lab:
                st.info(f"이 기간은 «{dec_met}» 채널별 데이터가 없어요. 전체 값만 있어요.")
            else:
                st.caption(f"{period_lbl} {base_lbl} 대비 «{dec_met}» 채널별 변화 "
                           f"({_suf}) · **비율·평균 지표라 채널을 더해도 전체가 안 돼요** — "
                           "합이 아니라 '어느 채널이 움직였나'로 읽으세요.")
                _fig_r = go.Figure(go.Bar(
                    x=_lab, y=_dlt, marker_color=[clr("red") if (d == d and d < 0)
                                                  else clr("green") for d in _dlt],
                    text=[("–" if d != d else _fmt.format(d)) for d in _dlt],
                    textposition="outside", cliponaxis=False))
                _ly_r = base_layout(320, title=f"«{dec_met}» 채널별 전년비 변화 ({_suf})")
                _ly_r["showlegend"] = False
                _fig_r.update_layout(**_ly_r)
                st.plotly_chart(_fig_r, width="stretch")
                wtable(style_delta_cols(pd.DataFrame(_rows).set_index("채널")),
                       width="stretch", dl_name=f"채널별 {dec_met} ({period_lbl})")
            base_t = cur_t = np.nan          # 아래 워터폴 경로는 타지 않는다
        else:
            # 거래액만 만원(더 잘게), 나머지 카운트는 명 그대로
            div, unit = (1e4, "만원") if dec_met == "첫구매 거래액" else (1, "명")
            st.caption(f"{period_lbl} {base_lbl} 대비 «{dec_met}» 증감 분해 "
                       f"(일평균·{unit}) · 세로축은 변동 구간만 확대 표시")
            base_t, cur_t = gprv(dec_met), gcur(dec_met)
        if dec_met in FUNNEL_ADDITIVE and (pd.isna(base_t) or pd.isna(cur_t)):
            st.info("이 기간엔 채널 데이터가 없어요.")
        elif dec_met in FUNNEL_ADDITIVE:
            labels, deltas = [], []
            for chn in CHANNELS:
                pv, cv = gprv(dec_met, chn), gcur(dec_met, chn)
                # 한쪽만 있으면 결측(채널분해 부재) — 0 취급하면 가짜 급락이 되므로 제외.
                # 실제 0은 pick이 0.0으로 돌려주므로 유지된다.
                if pd.isna(pv) or pd.isna(cv):
                    continue
                labels.append(chn); deltas.append((cv - pv) / div)
            if not labels:
                st.info(f"이 기간은 «{dec_met}» 채널별 분해 데이터가 없어요. 전체 값만 있어요. "
                        "사이드바에서 다른 기준 주차를 고르거나, 위 비교 기준을 "
                        "**월누적(MTD)** 으로 바꿔 보세요.")
            else:
                # 잔차(채널합↔TOTAL 차이·미분류)를 '기타'로 표시해 총계 막대가 항상
                # 라벨(=당년 TOTAL)과 정확히 맞게 한다. 0.5단위 미만(반올림 0)만 생략.
                resid = (cur_t - base_t) / div - sum(deltas)
                if abs(resid) >= 0.5:
                    labels.append("기타"); deltas.append(resid)
                fig_w = go.Figure(go.Waterfall(
                    x=[x_prv] + labels + [x_cur],
                    measure=["absolute"] + ["relative"] * len(labels) + ["total"],
                    y=[base_t / div] + deltas + [0],
                    text=[f"{base_t/div:,.0f}"] + [f"{d:+,.0f}" for d in deltas]
                         + [f"{cur_t/div:,.0f}"],
                    textposition="outside", cliponaxis=False,
                    increasing=dict(marker=dict(color=clr("green"))),
                    decreasing=dict(marker=dict(color=clr("red"))),
                    totals=dict(marker=dict(color=clr("slate"))),
                    connector=dict(line=dict(color="#e2e8f0")),
                    hovertemplate=("%{x}<br>변동 %{delta:+,.0f}" + unit
                                   + " · 누계 %{final:,.0f}" + unit + "<extra></extra>"),
                ))
                # 총액 막대가 델타를 압도하지 않도록 세로축을 변동 구간 주변으로 제한
                run, acc = [base_t / div], base_t / div
                for d in deltas:
                    acc += d; run.append(acc)
                run.append(cur_t / div)
                lo, hi = min(run), max(run)
                span = max(hi - lo, abs(hi) * 0.01, 1.0)
                lyw = base_layout(360, title=f"일평균 {dec_met} ({unit}) — 변동 구간 확대")
                lyw["showlegend"] = False
                lyw["yaxis"]["range"] = [lo - span * 0.45, hi + span * 0.55]
                fig_w.update_layout(**lyw)
                st.plotly_chart(fig_w, width="stretch")

            # 산출식 (접이식 첨부) — 워터폴을 그린 경우에만
            with st.expander("📐 산출식 · 계산 방법", expanded=False):
                st.markdown(f"""
**비교 기준**: {period_lbl} vs {base_lbl} · 값은 모두 **일평균**, 단위 **{unit}**
{"(거래액은 가독성을 위해 원 → 만원)" if dec_met == "첫구매 거래액" else ""}

**막대 구성** (왼쪽 → 오른쪽 누적):

- **시작(전년 전체)** = 전년 동기 «{dec_met}» 전체(*TOTAL) 값
- **채널별 증감** = `당년 채널값 − 전년 채널값`
  ㄴ 전년·당년 **양쪽에 값이 있는 채널만** 포함 (한쪽만 있으면 결측으로 보고 제외 — 0으로 두면 가짜 급락이 되므로)
- **기타** = `당년 전체 − 전년 전체 − Σ(채널별 증감)`
  ㄴ 채널 합과 전체의 차이·채널 미분류분을 담는 잔차 (0.5{unit} 미만이면 생략)
- **끝(당년 전체)** = 당년 «{dec_met}» 전체(*TOTAL) 값
  ㄴ 항상 `시작 + Σ(채널별 증감) + 기타` 와 정확히 일치

**대상 지표**: 채널 합이 전체와 일치하는 **가산 지표만** 제공
(거래액·고객수·가입자수·비회원트래픽). 객단가·가입율·당일가입CR 등 **비율 지표는
채널 합 ≠ 전체** 라 분해가 성립하지 않아 뺐어요.
""")

    # ── ⑤ 조직·카테고리별 첫구매 ────────────────────────
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("⑤ 조직·카테고리별 첫구매")
    _render_funnel_orgcat(df, odf, gran, cy, py, clabel, period_lbl, base_lbl,
                          prv_close, win)

    # ── ⑥ 신규회원 앱 설치·수신동의 ─────────────────────
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("⑥ 신규회원 앱 설치·수신동의")
    _render_funnel_app(df, gran, cy, py, clabel, base_tag, prv_close,
                       period_lbl, base_lbl)

    # ── 액션·이슈사항 ───────────────────────────────────
    # 「03 월별 추이」·「04 주차별 추이」를 접으면서 그 메모를 여기로 옮겨 왔다.
    # **키는 그대로 둔다** — 이미 써 둔 메모가 열쇠를 잃으면 파일엔 남는데 화면에서
    # 영영 안 보인다. 일자별엔 메모를 두지 않는다(하루짜리 액션은 쌓을 자리가 아니다).
    if gran in ("주", "월"):
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        _aim = st.session_state.get("wr_ai_model", DEFAULT_AI_MODEL)
        if gran == "주":
            report_text_block(
                f"wr_week_memo_{cy}_{clabel}", f"{cy}년 {month_trim(clabel)} 액션·이슈사항",
                ai_fn=lambda memo: ai_generate_insight(
                    df, ref_year, ref_month, clabel, _aim,
                    focus=f"{clabel} 주차 액션·이슈 및 인사이트", memo=memo))
        else:
            report_text_block(
                f"wr_month_memo_{ref_year}_{ref_month}",
                f"{ref_year}년 {ref_month}월 액션·이슈사항",
                ai_fn=lambda memo: ai_generate_insight(
                    df, ref_year, ref_month, None, _aim,
                    focus=f"{ref_year}년 {ref_month}월 액션·이슈 및 인사이트", memo=memo))


def _render_funnel_trend(df, gran, cy, py, win=None):
    """② 퍼널 지표 추이 — ①의 한 장면이 흐름 위 어디쯤인지.

    ①은 한 기간의 **사진**이라 '이번이 낮은 건지 원래 그런 건지'를 못 본다. 같은 퍼널
    지표를 연중으로 펼쳐 전년과 맞대면 그게 갈린다. 여기서 이상한 구간을 찾으면
    사이드바에서 그 기간으로 옮겨 가 ①~④를 다시 보면 된다.

    **표는 03·04와 같은 얼굴**이다 — 행=지표, 열=기간, 각 기간 **바로 오른쪽**에 전년 같은
    기간 대비 증감(`trend_table(delta_year=)`). 두 해가 열두 칸 떨어져 있으면 눈으로는
    못 맞댄다. 비율 지표(가입율·당일가입CR)는 %p 차이다.

    **자를 땐 기간을 센다.** 증감 칸까지 섞어 `columns[-N:]`으로 집으면 보이는 기간이
    반토막 난다(04에서 실제로 8주로 줄었다). 기간을 N개 고른 뒤 짝이 되는 증감 칸을
    딸려 보낸다.

    **기간 단위는 화면 위 「기간 단위」 하나를 따라간다.** 예전엔 여기에 자체 라디오가
    있어서, 단위를 바꾸려면 위와 여기를 따로 눌러야 했다(⑤까지 세 군데였다).
    """
    tg = gran
    # **지표는 안 고른다 — 있는 걸 다 그린다.** 고르게 해 두면 매번 같은 걸 다시 켜야 하고,
    # 퍼널은 앞단(트래픽·가입)과 뒷단(고객수·거래액)을 같이 봐야 어디서 빠졌는지가 보인다.
    sel = [m for m in FUNNEL_STEPS
           if ((df["gran"] == tg) & (df["metric"] == m)).any()]
    if not sel:
        st.info(f"«{FUNNEL_GRAN_LABEL.get(tg, tg)}» 단위에 퍼널 지표가 없어요.")
        return

    # ── 차트 — 03·04와 같이 한 줄에 세 장 ──
    _yrs = [y for y in (py, cy) if ((df["gran"] == tg) & (df["year"] == y)).any()]
    for i in range(0, len(sel), 3):
        for _col, _met in zip(st.columns(3), sel[i:i + 3]):
            with _col:
                st.plotly_chart(yoy_chart(df, tg, _met, _yrs, h=280, win=win),
                                width="stretch")

    # ── 표 — 기간마다 오른쪽에 전년 대비 증감 ──
    # 값을 왼쪽에 모으고 **증감은 오른쪽에 몰아**둔다 — 증감만 가로로 훑어야
    # '어느 기간부터 꺾였나'가 보인다. 03·04는 기간마다 끼우는 배치 그대로다.
    tbl = trend_table(df, tg, sel, [cy], delta_year=cy, delta_side="right", win=win)
    if tbl.empty:
        st.caption(f"{cy}년 «{FUNNEL_GRAN_LABEL.get(tg, tg)}» 값이 아직 없어요.")
        return
    # **자르지 않는다 — 올해 전체를 그대로 놓는다.** 최근 N개만 보여 주면 '어느 기간부터
    # 꺾였나'를 보려고 매번 엑셀을 받아야 한다. 표는 가로로 스크롤되니 폭은 문제가 아니다.
    # (자를 땐 기간을 세야 한다는 옛 규칙은 자르는 표에만 남는다 — 03의 「채널 × 기간」.)
    _unit = {"일": "날", "주": "주차", "월": "월"}.get(tg, tg)
    _n = len([c for c in tbl.columns if not _is_delta_col(c)])
    st.caption(f"왼쪽은 {cy}년 실적, 오른쪽은 **전년 같은 {_unit} 대비 증감**을 모아 뒀어요. "
               f"비율 지표(가입율·당일가입 첫구매율)는 %p 차이예요. 전년에 그 {_unit}가 "
               f"없으면 '–'로 둬요. 값은 전체(채널 합산) 기준이에요. "
               f"{cy}년 **{_n}개 {_unit} 전체**를 담았어요 — 옆으로 밀어서 보세요.")
    wtable(style_trend(tbl, sel), width="stretch",
           dl_name=f"퍼널 지표 {_unit}별 추이 ({cy}년)")


# ⑤의 값은 MICRO export, ①의 값은 마스터 export — **원천이 두 벌이다.** 같은 기간·같은
# 지표인데도 어긋날 수 있는데, 지금까지 화면엔 그냥 나란히 떠서 '어느 쪽이 맞나'로 끝났다.
# 숫자를 억지로 맞추는 게 아니라 **얼마나 다른지 밝히는** 게 여기서 할 일이다.
ORGCAT_GAP_WARN = 0.01                                    # 1% 넘게 벌어지면 띄운다
# 마스터에도 있는 지표만 대사할 수 있다 — 상품UV·상품CR은 MICRO에만 있다.
ORGCAT_GAP_METS = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]


def _lfms_like_master(df, base, gran, cy, clabel, lfmss):
    """마스터(①)와 **같은 모집단**인 LFMS 쪽. 고를 근거가 없으면 None.

    LFMS 포함 여부는 모집단이 다른 축이라 Y와 N의 전체 값이 다르다. 그런데 기본값이
    `sorted()[0]`(= 'N')이라, 마스터가 Y쪽 모집단이면 ①과 ④가 **같은 기간·같은 지표인데도
    수십 %씩 어긋난 채로** 나란히 떴다. 합성 재현에서 25% 차이가 났다.

    거래액으로 먼저 맞춘다 — 실파일에서 조직 합이 전체와 오차 0.00%로 맞는 유일한
    가산 지표라 대사 기준으로 가장 믿을 만하다. 없으면 고객수로 떨어진다.
    """
    for met in ("첫구매 거래액", "첫구매 고객수"):
        ref = pick(df, gran, met, "*TOTAL", cy, clabel, "mtd")
        if pd.isna(ref) or not ref:
            continue
        best, bgap = None, None
        for lf in lfmss:
            v = orgcat_view(base[base["lfms"] == lf]).get((), met, cy, clabel, "mtd")
            if pd.isna(v):
                continue
            g = abs(v - ref) / abs(ref)
            if bgap is None or g < bgap:
                best, bgap = lf, g
        if best is not None:
            return best
    return None


def _orgcat_master_gap(df, view, gran, met, cy, clabel):
    """⑤ 파일 합계와 ① 마스터 값의 차이 → `(마스터, 파일, 상대차)`. 비교 불가면 None."""
    if met not in ORGCAT_GAP_METS:
        return None
    m = pick(df, gran, met, "*TOTAL", cy, clabel, "mtd")
    o = view.get((), met, cy, clabel, "mtd")
    if pd.isna(m) or pd.isna(o) or not m:
        return None
    return float(m), float(o), (float(o) - float(m)) / abs(float(m))


def _orgcat_gap_note(gap, met):
    """합계 대사를 **블록 맨 아래에 접어서** 낸다.

    차이는 접이식 **라벨**에 박는다 — 펼치지 않아도 몇 %인지 보이고, 본문(왜 다를 수
    있는지)은 필요할 때만 편다. 예전엔 표 위에 큰 경고 상자로 있어서, 숫자를 보러 온
    화면인데 설명이 먼저 자리를 먹었다.
    맞물릴 땐(`ORGCAT_GAP_WARN` 미만) 캡션 한 줄로만 둔다 — 매번 펼칠 거리가 아니다.
    """
    if not gap:
        return
    mv, ov, rel = gap
    line = (f"① 전체 퍼널 `{fmt_value(met, mv)}` vs ⑤ 파일 합계 `{fmt_value(met, ov)}`"
            f" · 차이 **{rel * 100:+.1f}%**")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    if abs(rel) < ORGCAT_GAP_WARN:
        st.caption("합계 대사 — " + line + " · 두 원천이 맞물려요.")
        return
    with st.expander(f"합계 대사 — ①과 차이 {rel * 100:+.1f}%", expanded=False):
        st.markdown(
            line + "\n\n두 값은 **원천이 달라요** — ①은 마스터 export, ⑤는 MICRO "
            "조직×카테고리 export예요. 벌어지는 이유는 대개 셋 중 하나예요.\n"
            "- **LFMS 포함 여부**가 다른 모집단을 가리켜요 (위에서 바꿔 볼 수 있어요)\n"
            f"- 두 export의 **커버리지**가 달라요 (「{esc(PAGE_ORGCAT)}」에서 "
            "어느 기간이 있는지 볼 수 있어요)\n"
            "- 한쪽만 **최신 마감분**이 안 올라왔어요\n\n"
            "조직·카테고리 **구성비와 증감**은 ⑤ 안에서 일관되니 그대로 읽어도 돼요. "
            "총량은 ①을 기준으로 보세요.")


def _render_funnel_orgcat(df, odf, gran, cy, py, clabel, period_lbl, base_lbl,
                          prv_close, win=None):
    """⑤ 조직 > 카테고리 — MICRO 조직×카테고리 export를 퍼널과 같은 기간으로 자른다.

    이 원천엔 **채널 축이 없다.** 그래서 ③의 채널 상세와 교차하지 않고 나란히 놓는다 —
    '어느 채널에서 빠졌나'와 '어느 조직에서 빠졌나'는 서로 독립된 두 질문이다.
    """
    if odf is None or odf.empty:
        st.info("조직×카테고리 데이터가 없어요. MICRO 대시보드의 구분06×구분07 export를 "
                "올리면 첫구매 실적을 조직 > 카테고리로 파고들 수 있어요.")
        return
    # **채널 축은 여기서 «전 채널»만 본다.** 새 export가 채널로도 쪼개 오는데, 여기서
    # 섞으면 같은 조직이 채널 수만큼 중복돼 합계가 통째로 부풀어 오른다. 채널로 갈라
    # 보는 건 「조직·카테고리·채널별 첫구매 상세」 전용 화면이 맡는다.
    base = odf[(odf["gran"] == gran) & (odf["ch"] == ORGCAT_CH_ALL)]
    if base.empty:
        st.info(f"조직×카테고리 데이터에 «{gran}» 단위가 없어요. "
                "그 단위의 export를 올리면 여기서 같이 볼 수 있어요.")
        return
    f1, f2 = st.columns([1, 1.4])
    # LFMS 포함여부는 모집단이 다른 축이라 섞으면 안 된다. 고른 단위 안에서만 뽑는다 —
    # 단위마다 받아 온 export가 달라 전역 목록을 쓰면 데이터가 있는데도 빈 화면이 된다.
    lfmss = sorted(base["lfms"].dropna().astype(str).unique())
    # 기본값은 **① 퍼널과 총계가 가장 가까운 쪽**으로 맞춘다. 세션에 이미 고른 값이 있으면
    # 재지 않는다 — 인덱스를 LFMS마다 새로 만드는 값이라 매 리런에 물면 안 된다.
    _lf_auto = (_lfms_like_master(df, base, gran, cy, clabel, lfmss)
                if (len(lfmss) > 1 and "wr_fn_lfms" not in st.session_state) else None)
    with f1:
        if len(lfmss) > 1:
            guard_select("wr_fn_lfms", lfmss, default=_lf_auto)
            lf = st.radio("LFMS 포함", lfmss, horizontal=True, key="wr_fn_lfms",
                          help="모집단이 다른 축이에요. 처음 열 때는 위 ① 전체 퍼널과 "
                               "총계가 가장 가까운 쪽으로 맞춰 둬요.")
        else:
            lf = lfmss[0] if lfmss else "N"
            st.caption(f"LFMS 포함: **{esc(lf)}**")
    base = base[base["lfms"] == lf]
    mets = [m for m in FUNNEL_OC_METS if (base["metric"] == m).any()]
    if base.empty or not mets:
        st.info("고른 조건에 첫구매 지표가 없어요.")
        return
    with f2:
        guard_select("wr_fn_ocmet", mets)
        met = st.selectbox("지표", mets, key="wr_fn_ocmet")
    if base[base["label"].astype(str).isin(as_labels(clabel))
            & (base["year"] == cy)].empty:
        st.info(f"조직×카테고리 데이터에 **{cy}년 {lbl_disp(clabel)}**이 없어요. "
                "커버리지가 마스터와 달라서 그럴 수 있어요 — "
                f"「{PAGE_ORGCAT}」에서 어느 기간이 있는지 볼 수 있어요.")
        return

    view = orgcat_view(base)
    # ── 합계 대사 — ①과 ⑤는 원천이 다르다 ──
    # 같은 기간·같은 지표를 두 파일에서 읽으니 어긋날 수 있다. 조용히 나란히 두면
    # '어느 쪽이 맞나'로 끝나므로 얼마나 다른지와 왜 다를 수 있는지를 밝힌다.
    # **자리는 블록 «맨 아래», 접은 채로** 둔다 — 표를 보러 온 화면인데 설명 상자가
    # 위를 막고 있으면 정작 숫자가 안 보인다. 차이 자체는 접이식 라벨에 박아 둬서
    # 펼치지 않아도 눈에 들어온다(그게 신호고, 본문은 이유 설명이라 접어도 된다).
    _gap = _orgcat_master_gap(df, view, gran, met, cy, clabel)
    _orgs_all = view.live(())[0]
    if not _orgs_all:
        st.info("조직 항목이 없어요.")
        return

    # ── 볼 카테고리 — **⑤ 블록 전체가 이 하나를 본다** ─────────────────
    # 전년에 잘 팔리던 브랜드가 올해 빠지면 그 카테고리 하나가 전년비를 통째로 끌어내려
    # 나머지가 안 보인다. 그럴 땐 빼고 봐야 한다. 표만 거르고 합계·차트·요인 분해를
    # 그대로 두면 같은 화면이 두 모집단을 말하게 되므로 **아래 전부**에 같이 건다.
    _cats_all = []
    for _o in _orgs_all:
        for _c in view.live((_o,))[0]:
            if _c not in _cats_all:
                _cats_all.append(_c)
    if len(_cats_all) > 1:
        follow_options("wr_fn_cats", _cats_all)
        _cat_sel = st.multiselect(
            "볼 카테고리", _cats_all, key="wr_fn_cats",
            help="뺀 카테고리는 아래 표·차트·요인 분해에서 다 빠져요. 조직 합계와 "
                 "전체 합계도 남은 카테고리에서 다시 만들어요.")
        if not _cat_sel:
            st.info("볼 카테고리를 하나도 안 골랐어요. 위에서 골라 주세요.")
            return
    else:
        _cat_sel = _cats_all
    view = cat_filtered_view(view, _cat_sel, _cats_all)
    orgs = view.live(())[0]
    if not orgs:
        st.info("고른 카테고리를 가진 조직이 없어요.")
        return
    if view.narrowed:
        st.caption(f"카테고리 {len(_cats_all)}개 중 **{len(_cat_sel)}개**만 보고 있어요. "
                   "조직 합계·전체 합계는 파일 값이 아니라 **남은 카테고리에서 다시 "
                   "만든 값**이에요 — 고객수·상품UV는 같은 사람이 여러 조직에 잡혀 "
                   "합이 조금 부풀려져요.")
    # 상위와 합이 맞는 지표는 거래액 하나뿐이다 — 나머지는 유니크 값이라 비중을 붙이면
    # 거짓말이 된다(같은 사람이 여러 조직에 잡혀 합이 전체를 넘는다).
    additive = met in ORGCAT_ADDITIVE
    org_tbl = _funnel_level_table(view, [], orgs, "조직", met, cy, py, clabel, prv_close,
                                  view.get((), met, cy, clabel, "mtd") if additive else None)
    st.caption(f"{period_lbl} vs {base_lbl} · {esc(met)}"
               + ("" if additive else f" · «{esc(met)}»는 하위 합이 상위와 안 맞는 지표라 "
                                      "비중 칸을 뺐어요"))
    # 셀렉트박스 대신 **행을 눌러** 내려간다 — 조직을 고르고 다시 표를 보는 왕복이 없어진다.
    # 하이라이트는 **그리기 전에** 세션에서 읽는다 — 클릭이 리런을 일으키니 이 시점의
    # 세션값이 이미 방금 누른 행이다(반환값으로 칠하면 한 박자 늦는다).
    _pre = _picked_row(st.session_state.get("wr_fn_orgsel"), len(org_tbl))
    _org_sty = style_delta_cols(org_tbl)
    ev = wtable(_hl_row(_org_sty, _pre), width="stretch", key="wr_fn_orgsel",
                on_select="rerun", selection_mode="single-cell", hide_index=True,
                dl_name=f"조직별 {met} ({period_lbl})", dl_data=_org_sty)
    st.caption("맨 윗줄은 **합계**예요"
               + ("(남은 카테고리에서 다시 만든 값이에요)."
                  if view.narrowed else
                  "(파일이 준 전체 값이라 아래 합과 꼭 같진 않아요).")
               + " 조직 한 줄에서 **아무 칸이나 누르면** 그 줄에 색이 들어오고 아래에 그 "
                 "조직의 카테고리가 열려요. 합계 줄을 누르면 다시 닫혀요.")

    picked = _picked_child(ev, orgs)

    path, node_lbl = [], "전체"
    if picked:
        path, node_lbl = [picked], picked
        _kids = view.live((picked,))[0]
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.markdown(f"###### {esc(picked)} — 카테고리별")
        if not _kids:
            st.info(f"«{esc(picked)}» 아래에 볼 카테고리가 없어요.")
        else:
            _sub_tot = view.get((picked,), met, cy, clabel, "mtd")
            # 조직마다 위젯 키를 갈라 둔다 — 같은 키를 쓰면 조직을 바꿔도 옛 행 번호가
            # 남아 엉뚱한 카테고리가 열린다
            _ckey = f"wr_fn_catsel_{picked}"
            _cat_sty = style_delta_cols(_funnel_level_table(
                view, [picked], _kids, "카테고리", met, cy, py, clabel, prv_close,
                _sub_tot if additive else None))
            _pre2 = _picked_row(st.session_state.get(_ckey), len(_kids) + 1)
            ev2 = wtable(_hl_row(_cat_sty, _pre2), width="stretch",
                key=_ckey, on_select="rerun",
                selection_mode="single-cell", hide_index=True,
                dl_name=f"{picked} 카테고리별 {met} ({period_lbl})", dl_data=_cat_sty)
            st.caption(f"맨 윗줄은 **{esc(picked)} 합계**예요. 카테고리 한 줄에서 "
                       "**아무 칸이나 누르면** 그 줄에 색이 들어오고 아래 요인 분해가 "
                       "그 카테고리로 바뀌어요.")
            _c2 = _picked_child(ev2, _kids)
            if _c2 is not None:
                path, node_lbl = [picked, _c2], f"{picked} › {_c2}"

    # **표 바로 뒤에 차트를 둔다.** 표는 한 기간의 사진이라 '이번이 낮은 건지 원래 그런
    # 건지'를 못 본다. 예전엔 요인 분해 뒤로 밀려 있어서 표와 같이 못 봤다 — 위 채널별과
    # 같은 얼굴(표 → 차트)로 맞춘다.
    _funnel_orgcat_trend(odf, base, path, node_lbl, met, cy, py, gran, win,
                         cats=_cat_sel, cats_all=_cats_all)

    # 위에서 고른 게 있을 때만 펼친다 — 고르기 전엔 조직 표에 집중하게 둔다.
    # `expanded`는 리런마다 다시 먹으므로 행을 누르면 그 자리에서 열린다.
    _drilled = bool(path)
    with st.expander(f"어디에서 빠졌나 — {node_lbl}", expanded=_drilled):
        if not _drilled:
            st.caption("위 표에서 조직 한 줄을 누르면 그 조직 기준으로 바뀌어요. "
                       "지금은 **전체** 기준이에요.")
        _funnel_factor_block(view, path, node_lbl, cy, py, clabel, prv_close,
                             period_lbl, base_lbl)
    with st.expander("카테고리별 (조직 합산)", expanded=False):
        st.caption("조직을 안 고르고 **카테고리만 가로질러** 봐요 — 같은 카테고리를 "
                   "여러 조직이 나눠 갖고 있어서, 조직을 하나씩 들어가면 안 보이는 "
                   "구멍이 여기서 드러나요.")
        _funnel_cat_rollup(view, orgs, met, cy, py, clabel, period_lbl, base_lbl,
                           prv_close)
    if view.narrowed:
        st.caption("카테고리를 좁혀 놓아서 ① 전체 퍼널과의 「합계 대사」는 건너뛰었어요 "
                   "— 모집단이 달라 맞대도 알 수 있는 게 없어요.")
    else:
        _orgcat_gap_note(_gap, met)


# 표 맨 위에 두는 기준 행의 이름. 합계가 없으면 개별 값이 큰지 작은지 가늠이 안 된다.
TOTAL_ROW = "합계"


def _funnel_level_table(view, path, kids, lv_lbl, met, cy, py, clabel, prv_close, tot=None):
    """한 레벨의 표 — 행=자식, 열=전년·올해·전년비(+비중). `tot`을 주면 비중을 붙인다.

    **맨 위에 상위 합계 행을 둔다.** 기준점이 없으면 '골프 4명'이 큰지 작은지 가늠이
    안 된다. 값은 **파일이 준 상위 값** 그대로다 — 자식을 더하지 않는다. 고객수·상품UV는
    유니크라 자식 합이 상위를 넘고(실파일 +2.4%·+33%), 객단가·상품CR은 애초에 더할 수
    없다. 그래서 `*TOTAL`은 늘 파일 값이라는 규칙을 여기서도 지킨다.

    비중은 **하위 합이 상위와 맞는 지표에만** 붙인다(거래액). 고객수는 유니크 값이라
    같은 사람이 여러 곳에 잡혀 합이 상위를 넘는다 — 붙이면 거짓말이 된다.
    """
    rows = []
    # 합계 행 — 지금 보고 있는 노드(자식들의 상위) 값
    _tc = view.get(tuple(path), met, cy, clabel, "mtd")
    _tp = view.get(tuple(path), met, py, clabel, prv_close)
    _trow = {lv_lbl: TOTAL_ROW, f"{py}년": fmt_value(met, _tp),
             f"{cy}년": fmt_value(met, _tc),
             "전년비": fmt_delta(met, _tc, _tp) or "–"}
    if tot is not None:
        _trow["비중"] = "100.0%"
    rows.append(_trow)
    for k in kids:
        kc = view.get(tuple(list(path) + [k]), met, cy, clabel, "mtd")
        kp = view.get(tuple(list(path) + [k]), met, py, clabel, prv_close)
        row = {lv_lbl: k, f"{py}년": fmt_value(met, kp), f"{cy}년": fmt_value(met, kc),
               "전년비": fmt_delta(met, kc, kp) or "–"}
        if tot is not None:
            row["비중"] = ("–" if (pd.isna(kc) or pd.isna(tot) or not tot)
                          else f"{kc / tot * 100:.1f}%")
        rows.append(row)
    return pd.DataFrame(rows)


# 셀 선택은 **누른 칸에만** 옅은 테두리를 그린다 — 칼럼이 많으면 어느 줄을 눌렀는지
# 눈에 안 들어온다. 그래서 그 행 전체에 배경을 깐다. 파고들기 상태를 화면에서
# 되짚는 유일한 표시이기도 하다(셀렉트박스가 없어졌으니까).
ROW_PICK_BG = "#E8F0FE"


def _hl_row(sty, i, color=ROW_PICK_BG):
    """Styler의 i번째 **행 전체**에 배경을 입힌다. i가 None이면 그대로 둔다.

    위치로 칠한다 — 인덱스 종류에 안 기댄다. 배경은 엑셀까지 따라가므로
    이 Styler는 화면에만 쓰고 내려받기는 `dl_data`로 원본을 넘길 것.
    """
    if i is None:
        return sty

    def _paint(d):
        out = pd.DataFrame("", index=d.index, columns=d.columns)
        if 0 <= i < len(d):
            out.iloc[i, :] = f"background-color:{color}"
        return out

    try:
        return sty.apply(_paint, axis=None)
    except Exception:                                     # noqa: BLE001
        return sty                                        # 색이 없어도 표는 보여야 한다


def _picked_child(ev, kids):
    """표에서 고른 **자식**. 맨 위 합계 행(0번)을 누르면 고르지 않은 것으로 본다.

    `_funnel_level_table`이 0번에 합계 행을 끼우므로 행 번호가 한 칸씩 밀린다.
    그대로 `kids[i]`로 읽으면 **한 칸씩 어긋난 항목이 열린다** — 화면은 멀쩡히 뜨고
    값만 틀려서 눈으로는 안 잡힌다. 합계 줄을 누르는 건 '전체로 되돌리기'로 읽는다.
    """
    i = _picked_row(ev, len(kids) + 1)
    return None if i in (None, 0) else kids[i - 1]


def _picked_row(ev, n):
    """표에서 **행을 눌러** 고른 행 번호. 없으면 None.

    Streamlit의 행 선택(`single-row`)은 화면에 **체크박스 열**로 나온다 — 행을 눌러도
    안 잡히고 체크박스를 정확히 찍어야 한다. 그래서 셀 선택(`single-cell`)을 쓴다.
    아무 칸이나 누르면 `(행 번호, 칼럼명)`이 오니 그게 곧 행 클릭이다.
    **인덱스 칸은 안 돌려주므로** 이름은 일반 칼럼이어야 한다(`hide_index=True`).
    옛 세션에 남은 `rows` 선택도 같이 받아 준다.
    """
    sel = None
    for _get in (lambda: ev["selection"], lambda: getattr(ev, "selection", None)):
        try:
            sel = _get()
        except Exception:                                 # noqa: BLE001
            sel = None
        if sel:
            break
    if not sel:
        return None

    def _f(k):
        try:
            v = sel[k]
        except Exception:                                 # noqa: BLE001
            v = getattr(sel, k, None)
        return v or []

    cells, rows = _f("cells"), _f("rows")
    idx = cells[0][0] if cells else (rows[0] if rows else None)
    return idx if idx is not None and 0 <= idx < n else None


def _funnel_factor_block(view, path, node_lbl, cy, py, clabel, prv_close,
                         period_lbl, base_lbl):
    """어디에서 빠졌는지 — `거래액 = 상품UV × 상품CR × 객단가`를 요인별로 쪼갠다.

    원본에서 이 항등식이 오차 0.000%로 성립해서(모든 조직·카테고리) '왜'의 축이 된다.
    LMDI(로그평균 디비지아)라 세 기여액의 합이 실제 증감과 원 단위까지 같고, 어느 요인을
    먼저 대입하느냐로 답이 달라지지 않는다.

    **0·음수·결측이 하나라도 있으면 로그가 정의되지 않는다** — 그럴 땐 숫자를 지어내지 말고
    왜 못 쪼갰는지 말한다.

    **고객수는 요인이 아니라 중간 결과다.** 원본에서 `고객수 = 상품UV × 상품CR`이
    정확히 성립해서(위 항등식과 `거래액 = 고객수 × 객단가`를 나누면 나온다), 값은
    보여 주되 **기여액은 안 매긴다** — 그 몫이 유입·전환 두 줄에 이미 들어 있어
    같이 세면 두 번 센다. 안 넣으면 '전환율은 있는데 고객수가 없다'로 읽힌다.
    """
    prev = [view.get(tuple(path), f, py, clabel, prv_close) for f, _ in ORGCAT_FACTORS]
    cur = [view.get(tuple(path), f, cy, clabel, "mtd") for f, _ in ORGCAT_FACTORS]
    _cu_p = view.get(tuple(path), "첫구매 고객수", py, clabel, prv_close)
    _cu_c = view.get(tuple(path), "첫구매 고객수", cy, clabel, "mtd")
    rows = []
    for (f, nick), pv, cv in zip(ORGCAT_FACTORS, prev, cur):
        rows.append({"요인": f"{nick} ({f})", f"{py}년": fmt_value(f, pv),
                     f"{cy}년": fmt_value(f, cv), "전년비": fmt_delta(f, cv, pv) or "–"})
        # 유입 × 전환 = 고객수 — 그 자리에 끼워야 곱셈 사슬로 읽힌다
        if f == "상품CR" and not (pd.isna(_cu_p) and pd.isna(_cu_c)):
            rows.append({"요인": "= 고객수 (첫구매 고객수)",
                         f"{py}년": fmt_value("첫구매 고객수", _cu_p),
                         f"{cy}년": fmt_value("첫구매 고객수", _cu_c),
                         "전년비": fmt_delta("첫구매 고객수", _cu_c, _cu_p) or "–"})
    split = factor_split(prev, cur)
    if split is None:
        st.caption(f"{period_lbl} vs {base_lbl} · 세 요인 중 하나라도 값이 없거나 0 이하라 "
                   "기여액은 못 쪼갰어요. 증감률만 보세요.")
    else:
        parts, total = split
        # 기여액은 **요인 세 줄에만** 붙인다. 고객수 줄은 유입 × 전환이라 여기서
        # 숫자를 주면 칼럼을 더했을 때 그 몫이 두 번 세어진다.
        _fr = [r for r in rows if not str(r["요인"]).startswith("=")]
        for r in rows:
            r["기여액"] = "–"
        for r, part in zip(_fr, parts):
            r["기여액"] = f"{part / 1e6:+,.1f}백만원"
        _rv_p = view.get(tuple(path), "첫구매 거래액", py, clabel, prv_close)
        _rv_c = view.get(tuple(path), "첫구매 거래액", cy, clabel, "mtd")
        rows.append({"요인": "합계 (거래액 증감)", f"{py}년": fmt_value("첫구매 거래액", _rv_p),
                     f"{cy}년": fmt_value("첫구매 거래액", _rv_c),
                     "전년비": fmt_delta("첫구매 거래액", _rv_c, _rv_p) or "–",
                     "기여액": f"{total / 1e6:+,.1f}백만원"})
        _worst = min(zip(ORGCAT_FACTORS, parts), key=lambda x: x[1])
        _cap = (f"{period_lbl} vs {base_lbl} · 거래액 증감을 세 요인으로 쪼갰어요 "
                f"(합이 실제 증감과 원 단위까지 같아요). 가장 많이 끌어내린 건 "
                f"**{_worst[0][1]}**({_worst[0][0]}) 이에요.")
        if any(str(r["요인"]).startswith("=") for r in rows):
            _cap += (f" **고객수는 유입 × 전환**이라 기여액을 따로 안 매겨요 — 그 몫 "
                     f"{(parts[0] + parts[1]) / 1e6:+,.1f}백만원은 위 두 줄에 이미 "
                     f"들어 있어요.")
        st.caption(_cap)
    wtable(style_delta_cols(pd.DataFrame(rows).set_index("요인")), width="stretch",
           dl_name=f"요인 분해 {node_lbl} ({period_lbl})")


def _funnel_orgcat_trend(odf, base, path, node_lbl, met, cy, py, gran, win=None,
                         cats=None, cats_all=None):
    """⑤ 하단 — 지금 보고 있는 자리의 **하위 항목별** 연중 흐름.

    표는 한 기간의 사진이라 '이번이 낮은 건지 원래 그런 건지'를 못 본다. 여기서
    연중 흐름을 보고, 이상한 구간을 찾으면 위 표에서 그 기간으로 옮겨 가면 된다.

    **드릴다운을 따라가며 한 단계 아래를 그린다** — 합계면 조직들, 조직을 누르면 그
    조직의 카테고리들, 더 내려갈 데가 없으면 그 노드 자신. 항목을 다 그려야 '어느
    조직이 빠지고 있나'가 한 화면에서 보인다.

    **색은 항목, 선 모양은 연도**다. 올해는 실선, 전년은 얇은 점선. 색을 연도에 쓰면
    항목이 셋만 넘어도 무엇이 무엇인지 못 짚는다. 항목 색은 **전체 자식 목록** 기준으로
    고정한다 — 그린 것만으로 색을 매기면 연도를 끄고 켤 때 색이 바뀐다
    (「05. 앱푸시 동의 현황」의 `_ycolor`와 같은 이유).

    연도는 체크로 넣고 뺀다. 항목이 많으면 두 해가 겹쳐 구분이 안 되므로, 전년을 꺼서
    올해 흐름만 보는 길을 열어 둔다.

    **기간 단위는 화면 위 「기간 단위」 하나를 따라간다** — 예전엔 여기에도 자체 라디오가
    있어서 단위 하나 바꾸는 데 세 군데를 눌러야 했다. 그 단위가 이 원천에 없으면
    왜 못 그리는지 말하고 물러난다(커버리지가 마스터와 다를 수 있다).
    """
    tg = gran
    if not (odf["gran"] == tg).any():
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.caption(f"조직×카테고리 원천엔 «{FUNNEL_GRAN_LABEL.get(tg, tg)}» 단위가 없어서 "
                   "연중 추이는 못 그려요.")
        return
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.markdown(f"###### 연중 추이 — {esc(node_lbl)} · {esc(met)}")
    st.markdown("**차트에 올릴 연도**")
    y1, y2, _y3 = st.columns([1, 1, 4])
    show_cy = y1.checkbox(f"{cy}년", value=True, key="wr_fn_trend_cy")
    show_py = y2.checkbox(f"{py}년", value=True, key="wr_fn_trend_py",
                          help="항목이 많아 겹쳐 보이면 꺼서 올해만 봐요.")
    yrs = [y for y, on in ((cy, show_cy), (py, show_py)) if on]
    if not yrs:
        st.caption("차트에 올릴 연도를 하나도 안 골랐어요. 위에서 연도를 켜 주세요.")
        return

    # 위 표와 같은 축이어야 한다 — 채널을 안 거르면 추이만 채널 수만큼 부풀어 오른다.
    sub = odf[(odf["gran"] == tg) & (odf["ch"] == ORGCAT_CH_ALL)]
    if "lfms" in base.columns and len(base):
        sub = sub[sub["lfms"] == base["lfms"].iloc[0]]
    view = orgcat_view(sub)
    # 위에서 고른 카테고리를 여기도 그대로 건다 — 표만 좁히고 차트를 그대로 두면
    # 같은 화면이 두 모집단을 말하게 된다. `tg`가 달라 뷰를 새로 만들지만 규칙은 같다.
    if cats is not None and cats_all is not None:
        view = cat_filtered_view(view, cats, cats_all)
    # 그릴 대상 — 한 단계 아래 항목 전부. 더 내려갈 데가 없으면 그 노드 자신.
    kids = view.live(tuple(path))[0]
    items = ([(k, tuple(path) + (k,)) for k in kids] if kids
             else [(node_lbl, tuple(path))])
    # **차트에서 항목을 넣고 뺀다** — 조직이 열 곳 넘으면 선이 엉켜 정작 보려던 게 안
    # 보인다. 표는 다 놓고 차트만 두세 곳으로 좁혀 맞대는 게 이 화면의 쓰임새다.
    # 키를 **경로마다 갈라 둔다** — 한 키를 쓰면 조직을 파고들었을 때 옛 조직 이름이
    # 세션에 남아 항목이 통째로 비어 보인다(카테고리 표의 `_ckey`와 같은 이유).
    _all_nm = [nm for nm, _p in items]
    if len(_all_nm) > 1:
        _ikey = "wr_fn_trend_items_" + "/".join(path)
        guard_multi(_ikey, _all_nm)
        if _ikey not in st.session_state:
            st.session_state[_ikey] = _all_nm
        _pick = st.multiselect("차트에 올릴 항목", _all_nm, key=_ikey,
                               help="표는 그대로 두고 차트만 좁혀 봐요. "
                                    "항목 색은 어느 걸 골라도 그대로예요.")
        if not _pick:
            st.caption("차트에 올릴 항목을 하나도 안 골랐어요. 위에서 골라 주세요.")
            return
    else:
        _pick = _all_nm
    # 색은 **전체 목록** 기준으로 미리 굳힌다 — 연도를 끄고 켜도, 항목을 빼도 안 바뀌게.
    # 항목이 하나뿐이면(말단 노드) 색을 **연도**에 쓴다. 그게 앱의 다른 추이 차트와
    # 같은 얼굴이고, 항목이 하나면 색으로 가를 게 연도밖에 없다.
    # 연도 색은 `yoy_chart`와 같은 규칙 — **오래된 해부터** YEAR_PAL 순서라 올해가 파랑이다.
    ycolor = {y: YEAR_PAL[i % len(YEAR_PAL)] for i, y in enumerate(sorted({cy, py}))}
    icolor = {nm: ORGCAT_PAL[i % len(ORGCAT_PAL)] for i, (nm, _p) in enumerate(items)}
    # 색을 굳힌 **뒤에** 좁힌다 — 순서로 색을 매기니 먼저 좁히면 뺄 때마다 색이 밀린다.
    items = [(nm, _p) for nm, _p in items if nm in set(_pick)]
    # 색을 항목에 쓸지 연도에 쓸지는 **있는 항목 수**로 정한다. 그린 것 수로 정하면
    # 필터로 하나만 남겼을 때 색이 연도 색으로 뒤집혀, 맞대던 선을 다시 찾아야 한다.
    # (말단 노드라 애초에 항목이 하나뿐이면 그땐 색으로 가를 게 연도밖에 없다.)
    by_item = len(_all_nm) > 1

    # x축은 **올해 라벨**로 세운다 — 전년에만 있는 기간까지 그리면 축이 늘어져
    # 정작 올해 흐름이 눌린다. 전년은 같은 라벨끼리 맞대진다.
    _lb = (sub[sub["year"] == cy][["label", "sortkey"]].drop_duplicates()
           .sort_values("sortkey")["label"].astype(str).tolist())
    # 화면 위 「조회 기간」을 따라간다 — 위에서 좁혀 놓고 여기만 연중이면 같은 화면이
    # 두 기간을 말하게 된다.
    if win is not None:
        _w = set(as_labels(win))
        _lb = [x for x in _lb if x in _w]
    if not _lb:
        st.caption(f"{cy}년 «{tg}» 데이터가 없어요.")
        return
    unit, div = METRIC_UNIT.get(met, ("", 1))
    if met in PCT_METRICS:
        div, unit = 0.01, "%"

    # 값을 먼저 다 뽑고 **아무 항목·아무 해에도 값이 없는 기간은 축에서 뺀다.** `_lb`는
    # 그 해의 모든 지표·모든 노드가 쓴 라벨이라 지금 보는 것과 무관한 기간이 섞인다.
    # 그대로 두면 **모든 선이 같은 자리에서 나란히 끊겨** 데이터가 빠진 것처럼 보이는데,
    # 실은 그 칸에 애초에 아무것도 없다. 정보가 없는 눈금이라 빼는 게 맞다.
    # (일부 항목·한쪽 해에만 값이 없는 칸은 남긴다 — 그건 진짜 결측이라 끊어 보여 준다.)
    series = {(nm, y): [view.get(_pth, met, y, lb, "mtd") for lb in _lb]
              for nm, _pth in items for y in yrs}
    keep = [j for j in range(len(_lb))
            if any(not pd.isna(v[j]) for v in series.values())]
    dropped = len(_lb) - len(keep)
    _lb = [_lb[j] for j in keep]
    if not _lb:
        st.caption(f"«{esc(node_lbl)}»의 {tg} 단위 값이 없어요.")
        return

    fig = go.Figure()
    drew = 0
    for nm, _pth in items:
        for y in yrs:
            vals = [series[(nm, y)][j] for j in keep]
            if not any(not pd.isna(v) for v in vals):
                continue
            drew += 1
            _cur = (y == cy)
            fig.add_trace(go.Scatter(
                x=[month_trim(v) for v in _lb],
                y=[(v / div if not pd.isna(v) else None) for v in vals],
                mode="lines+markers",
                name=(f"{nm} ({y})" if by_item else f"{y}년"),
                line=dict(color=clr(icolor[nm] if by_item else ycolor[y]),
                          width=2 if _cur else 1.4,
                          dash=None if _cur else ("dot" if by_item else None)),
                opacity=1.0 if _cur else (0.65 if by_item else 1.0),
                marker=dict(size=5 if _cur else 4), connectgaps=False))
    if not drew:
        st.caption(f"«{esc(node_lbl)}»의 {tg} 단위 값이 없어요.")
        return
    # 제목에 **보고 있는 자리**를 박는다 — 이 페이지엔 ②의 퍼널 지표 추이 차트가 여럿
    # 있어서, 지표 이름만 적혀 있으면 어느 블록 것인지 구분이 안 된다.
    ly = base_layout(340 if by_item else 300,
                     ysuffix=unit if unit == "%" else "",
                     title=f"{node_lbl} · {met} {'주차별' if tg == '주' else '월별'} 추이 ({unit})")
    ly["xaxis"]["categoryorder"] = "array"
    ly["xaxis"]["categoryarray"] = [month_trim(v) for v in _lb]
    if tg == "주":
        ly["xaxis"]["tickangle"] = -45
        ly["xaxis"]["nticks"] = 20
    fig.update_layout(**ly)
    st.plotly_chart(fig, width="stretch")

    _cap = "위 표에서 조직·카테고리를 누르면 이 차트도 그 아래 단계로 바뀌어요."
    if by_item:
        _cap += (f" 지금은 **{len(items)}개 항목**을 그렸어요 — 색이 항목, "
                 "**점선이 전년**이에요. 겹쳐 보이면 위에서 전년을 꺼 보세요.")
    if dropped:
        _cap += (f" 어느 항목·어느 해에도 값이 없는 {dropped}개 기간은 축에서 뺐어요 — "
                 "빈 눈금이 끼면 선이 괜히 끊겨 보여요.")
    _cap += " 값이 없는 기간은 **선을 끊어** 둬요 — 이어 버리면 그 사이에 데이터가 있는 것처럼 보여요."
    st.caption(_cap)

    # ── 차트 아래 같은 값을 표로 — ②·03과 같은 얼굴(왼쪽 실적 · 오른쪽 증감) ──
    # 차트는 흐름을, 표는 숫자를 준다. 선만 있으면 '9월 2주차가 정확히 얼마였나'를 못 읽어
    # 결국 엑셀을 받아야 한다. 행=항목, 열=기간이라 ②(행=지표)와 축만 다르고 얼굴은 같다.
    # **맨 위에 합계 행을 둔다** — 기준점이 없으면 'e-영업1 201명'이 큰지 작은지
    # 가늠이 안 된다. 값은 지금 보고 있는 노드(자식들의 상위)의 **파일 값 그대로**지
    # 자식 합이 아니다 — 고객수·상품UV는 유니크라 합이 상위를 넘고 객단가·상품CR은
    # 애초에 더할 수 없다(위 레벨 표·합산 표와 같은 규칙).
    _trows, _drows = {}, {}
    _rows_src = ([(TOTAL_ROW, tuple(path))] if kids else []) + list(items)
    for nm, _pth in _rows_src:
        _cv = {lb: view.get(_pth, met, cy, lb, "mtd") for lb in _lb}
        _pv = {lb: view.get(_pth, met, py, lb, "mtd") for lb in _lb}
        _trows[nm] = {month_trim(lb): fmt_value(met, _cv[lb]) for lb in _lb}
        _drows[nm] = {f"{month_trim(lb)} 증감": (fmt_delta(met, _cv[lb], _pv[lb]) or "–")
                      for lb in _lb}
    if _trows:
        _tt = pd.DataFrame(_trows).T
        _dd = pd.DataFrame(_drows).T
        # 실적을 다 놓고 증감을 오른쪽에 몬다 — 증감만 가로로 훑어야 어디서 꺾였나가 보인다
        _tt = pd.concat([_tt, _dd], axis=1)
        _tt.index.name = "항목"
        wtable(style_delta_cols(_tt), width="stretch",
               dl_name=f"{node_lbl} {met} {tg}별 추이 ({cy}년)")
        _tcap = (f"위 차트와 같은 값이에요. 왼쪽은 {cy}년 실적, 오른쪽은 **{py}년 같은 "
                 f"기간 대비 증감**이에요. {cy}년 **{len(_lb)}개 기간 전체**를 담았어요 — "
                 "옆으로 밀어서 보세요.")
        if kids:
            _tcap += (f" 맨 윗줄 **합계**는 «{esc(node_lbl)}»의 파일 값이라 아래 항목 "
                      "합과 꼭 같진 않아요. 차트엔 안 그려요.")
        st.caption(_tcap)


def _funnel_cat_rollup(view, orgs, met, cy, py, clabel, period_lbl, base_lbl, prv_close):
    """조직을 가로질러 카테고리를 모은 표 — '어느 조직 것이든 이 카테고리가 빠졌나'.

    위 표는 조직을 골라야 카테고리가 보인다. 같은 카테고리를 여러 조직이 나눠 갖고 있어서
    '가방이 빠졌나'를 보려면 조직을 하나씩 다 들어가 봐야 한다. 그래서 조직 선택과 상관없이
    늘 전체를 보여 준다.

    **합칠 수 있는 건 거래액뿐이다.** 고객수는 유니크 값이라 같은 사람이 여러 조직에 잡혀
    합이 전체를 조금 넘고(실파일 +2.4%), 객단가는 애초에 더할 수 없어 `거래액 합 ÷ 고객수
    합`으로 다시 만든다 — 분모가 부풀려진 만큼 살짝 낮게 나온다. 지어내지 않고 **그렇게
    만들었다고 화면에 밝힌다**.

    거래액·고객수 두 축만 모으면 세 지표가 다 나온다 — 지표마다 따로 훑지 않는다.
    """
    if met not in FUNNEL_ROLLUP_METS:
        st.caption(f"«{esc(met)}»는 조직을 가로질러 **모을 수 없는 지표**라 카테고리 합산 "
                   "표를 안 만들었어요.")
        return
    acc = {}
    for _org in orgs:
        for _cat in view.live((_org,))[0]:
            e = acc.setdefault(_cat, {"rev": [np.nan, np.nan],
                                      "cust": [np.nan, np.nan],
                                      "uv": [np.nan, np.nan]})
            for i, (yr, cl) in enumerate(((cy, "mtd"), (py, prv_close))):
                for k, mname in (("rev", "첫구매 거래액"), ("cust", "첫구매 고객수"),
                                 ("uv", "상품UV")):
                    v = view.get((_org, _cat), mname, yr, clabel, cl)
                    if not pd.isna(v):
                        e[k][i] = v if pd.isna(e[k][i]) else e[k][i] + v
    if not acc:
        return

    def _val(e, i):
        if met == "첫구매 객단가":
            r, c = e["rev"][i], e["cust"][i]
            return r / c if not (pd.isna(r) or pd.isna(c) or not c) else np.nan
        if met == "상품CR":
            # 항등식으로 되만든다 — 중복이 분자·분모에 같이 들어가 상당 부분 상쇄된다
            c, u = e["cust"][i], e["uv"][i]
            return c / u if not (pd.isna(c) or pd.isna(u) or not u) else np.nan
        return e[{"첫구매 고객수": "cust", "상품UV": "uv"}.get(met, "rev")][i]

    additive = met in ORGCAT_ADDITIVE
    tot_c = sum(v for v in (_val(e, 0) for e in acc.values()) if not pd.isna(v))
    rows = []
    for _cat, e in acc.items():
        kc, kp = _val(e, 0), _val(e, 1)
        row = {"카테고리": _cat,
               f"{py}년": fmt_value(met, kp), f"{cy}년": fmt_value(met, kc),
               "전년비": fmt_delta(met, kc, kp) or "–",
               "_sort": -kc if not pd.isna(kc) else 1}
        if additive:
            row["비중"] = ("–" if (pd.isna(kc) or not tot_c)
                          else f"{kc / tot_c * 100:.1f}%")
        rows.append(row)
    # 맨 위에 전체 합계 — **파일이 준 최상위 값**이지 카테고리 합이 아니다.
    # 고객수는 유니크라 카테고리 합이 전체를 조금 넘는데(실파일 +2.4%), 그 합을
    # 합계라고 적으면 아래 줄과 안 맞는 숫자를 기준점으로 삼게 된다.
    _gc = view.get((), met, cy, clabel, "mtd")
    _gp = view.get((), met, py, clabel, prv_close)
    _grow = {"카테고리": TOTAL_ROW, f"{py}년": fmt_value(met, _gp),
             f"{cy}년": fmt_value(met, _gc),
             "전년비": fmt_delta(met, _gc, _gp) or "–", "_sort": float("-inf")}
    if additive:
        _grow["비중"] = "100.0%"
    rows.append(_grow)
    tbl = (pd.DataFrame(rows).sort_values("_sort")
           .drop(columns=["_sort"]).set_index("카테고리"))
    _why = {"첫구매 거래액": "조직 합이 전체와 맞는 지표라 그대로 더했어요.",
            "첫구매 고객수": "같은 고객이 여러 조직에 잡혀 **합이 전체를 조금 넘어요** — "
                          "순위·전년비로 읽고 절대값은 위 표를 보세요.",
            "첫구매 객단가": "더할 수 없는 지표라 **거래액 합 ÷ 고객수 합**으로 다시 "
                          "만들었어요. 고객수가 조금 부풀려져 있어 실제보다 살짝 낮아요.",
            "상품UV": "같은 사람이 여러 조직에 잡혀 **합이 전체를 크게 넘어요**(실파일 +33%) — "
                    "절대값 말고 **어느 카테고리가 큰지·전년 대비 어떤지**로 읽으세요.",
            "상품CR": "더할 수 없는 지표라 **고객수 합 ÷ 상품UV 합**으로 다시 만들었어요 "
                    "(`거래액 = 상품UV × 상품CR × 객단가`에서 나오는 항등식이에요). "
                    "중복이 분자·분모에 같이 들어가 상품UV 단독보다는 왜곡이 작아요."}
    st.caption(f"{period_lbl} vs {base_lbl} · {esc(met)} · 맨 윗줄 **합계**는 "
               f"{'고른 카테고리에서 다시 만든 값' if view.narrowed else '파일이 준 전체 값'}"
               f"이라 아래 카테고리 합과 꼭 같진 않아요. {_why.get(met, '')}")
    wtable(style_delta_cols(tbl), width="stretch",
           dl_name=f"카테고리 전체 {met} ({period_lbl})")

def _funnel_app_one(df, gran, met, year, label, close):
    """앱 블록의 값 하나. 앱푸시 수신동의만 원천이 일별이라 그 기간 일평균으로 묶는다."""
    v = pick(df, gran, met, "*TOTAL", year, label, close)
    # 원천이 일자 헤더 표라 주·월 칸이 늘 비는 지표들 — 일별을 그 기간 일평균으로 묶는다
    if met in PUSH_DAILY_ONLY and pd.isna(v):
        v = push_period_avg(df, year, label, metric=met)[0]
    return v


# 최근 몇 개 기간을 나란히 — 카드 한 장으로는 '이번이 낮은 건지 원래 그런 건지'를 못 본다.
# 단위는 화면 위 비교 기준(주간/월누적)을 그대로 따라간다.
APP_TREND_N = {"일": 14, "주": 8, "월": 6}


def _funnel_app_trend(df, gran, cy, clabel):
    """⑤ 최근 기간 추이 — 위 비교 기준과 같은 단위(주/월)로 최근 N개를 나란히 놓는다."""
    src = df[(df["metric"] == "앱설치") & (df["gran"] == gran)
             & (df["segment"] == "*TOTAL") & df["value"].notna()]
    if src.empty:
        return
    per = (src[["year", "label", "sortkey"]].drop_duplicates()
           .sort_values("sortkey").tail(APP_TREND_N.get(gran, 6)))
    _cur = as_labels(clabel)
    _here = ((per["year"] == cy) & (per["label"].astype(str).isin(_cur))).any()
    if not _here:
        _pp = period_parts(gran, cy, _cur[-1])
        if _pp:
            per = pd.concat([per, pd.DataFrame([{"year": cy, "label": _pp[0],
                                                 "sortkey": _pp[1]}])], ignore_index=True)
            per = per.sort_values("sortkey")
    # 「가입자 대비 설치율」은 뺐다 — 설치와 동의 두 갈래의 비율이 한 표에 섞여 있어
    # 어느 쪽 이야기인지 헷갈렸다. 설치는 위의 원값으로 읽고, 비율 칸은 **동의율 둘**만
    # 남겨 분모(앱 보유 무관 ↔ 앱 포함)끼리 맞대게 한다.
    _rate = lambda a, b: ("–" if (pd.isna(a) or pd.isna(b) or not b)
                          else f"{a / b * 100:.2f}%")
    rows = []
    for _, r in per.iterrows():
        y, lb = int(r["year"]), str(r["label"])
        got = {m: _funnel_app_one(df, gran, m, y, lb, "mtd")
               for m in APP_STEPS + ["앱푸시동의_앱무관"]}
        jn, ins = got["가입자수"], got["앱설치"]
        ag, aa = got["앱푸시수신동의"], got["앱푸시동의_앱무관"]
        rows.append({
            "기간": f"{y}년 {month_trim(lb)}"
                    + (" ◀" if (y == cy and lb in _cur) else ""),
            "가입자수": fmt_value("가입자수", jn),
            "앱 신규설치": fmt_value("앱설치", ins),
            "_동의_앱무관": fmt_value("앱푸시동의_앱무관", aa),
            "_동의_앱포함": fmt_value("앱푸시수신동의", ag),
            "_율_앱무관": _rate(aa, jn),
            "_율_앱포함": _rate(ag, jn),
        })
    if not rows:
        return
    tbl = pd.DataFrame(rows).set_index("기간")
    # **앱 보유 무관 원천을 안 올렸으면 예전 화면 그대로 낸다.** 구분할 대상이 하나뿐인데
    # 칼럼 이름에 `(앱포함)`을 달아 두면 있지도 않은 다른 칸을 찾게 되고, `–`만 늘어선
    # 칼럼은 0인지 없는 건지도 안 갈린다. 위 비율 표와 같은 규칙이다.
    has_pa = not ((tbl["_동의_앱무관"] == "–").all() and (tbl["_율_앱무관"] == "–").all())
    if has_pa:
        tbl = tbl.rename(columns={"_동의_앱무관": "앱푸시 동의(앱보유무관)",
                                  "_동의_앱포함": "앱푸시 동의(앱포함)",
                                  "_율_앱무관": "앱푸시(앱보유무관)/신규회원",
                                  "_율_앱포함": "앱푸시(앱포함)/신규회원"})
    else:
        tbl = (tbl.drop(columns=["_동의_앱무관", "_율_앱무관"])
                  .rename(columns={"_동의_앱포함": "앱푸시 수신동의",
                                   "_율_앱포함": "신규회원 수신동의율"}))
    with st.expander(f"최근 {gran} 추이 ({len(rows)}개 기간)", expanded=True):
        wtable(tbl, width="stretch", dl_name=f"앱 설치·수신동의 최근 {gran} 추이")
        _cap = ("`◀` 가 위 카드와 같은 기간이에요. 값은 모두 **일평균**이라 기간 길이가 "
                "달라도 그대로 견줄 수 있어요.")
        if has_pa:
            _cap += (" 동의율 두 칸은 **분모가 달라요** — 두 값의 차이가 곧 "
                     "'앱을 안 깔아서 못 받는 몫'이에요.")
        st.caption(_cap)


def _render_funnel_app(df, gran, cy, py, clabel, base_tag, prv_close,
                       period_lbl="", base_lbl=""):
    """⑤ 신규 가입자가 앱까지 오는지 — 앱설치·앱푸시 수신동의.

    **어느 기간 값인지 여기서 다시 말한다.** 위에서 한참 내려온 자리라 상단 캡션이 안
    보이고, 이 블록만 원천이 달라 커버리지도 다르다(앱설치는 3월부터, 앱푸시는 일별 집계).
    기간을 안 밝히면 빈 칸이 '데이터가 없는 건지 기간이 안 맞는 건지' 안 갈린다.
    """
    def _one(met, year, close):
        return _funnel_app_one(df, gran, met, year, clabel, close)

    if period_lbl:
        st.caption(f"기준: **{period_lbl}** vs {base_lbl} · 값은 모두 **일평균**이에요.")
    # 비율 계산에 쓰는 지표까지 같이 읽는다(카드는 APP_STEPS만 세운다)
    _need = APP_STEPS + ["앱푸시동의_앱무관", "앱푸시대상_앱무관"]
    cur = {m: _one(m, cy, "mtd") for m in _need}
    prv = {m: _one(m, py, prv_close) for m in _need}
    _, approx, n_push = push_period_avg(df, cy, as_labels(clabel)[-1])
    # '원천이 아예 없다'와 '이 기간에만 없다'는 다른 문제다 — 뒤엣것은 커버리지 이야기라
    # 어디까지 들어왔는지 같이 말해 줘야 '업로드가 실패했나'로 안 읽힌다.
    _cv = df[(df["metric"] == "앱설치") & (df["gran"] == gran)
             & (df["segment"] == "*TOTAL") & df["value"].notna()]
    if pd.isna(cur["앱설치"]) and not _cv.empty:
        _last = _cv.sort_values("sortkey").iloc[-1]
        st.warning(f"**{period_lbl or lbl_disp(clabel)}에는 앱 데이터가 없어요.** 앱 원천은 "
                   f"**{int(_last['year'])}년 {_last['label']}**까지 들어와 있어요 — "
                   "마스터(가입자수·거래액)보다 며칠 늦게 끝나요. 아래 추이에서 "
                   "값이 있는 기간을 보세요.")
    if all(pd.isna(cur[m]) for m in APP_STEPS):
        st.info("가입자수·앱설치·앱푸시 수신동의 데이터가 모두 없어요.")
        _funnel_app_trend(df, gran, cy, clabel)
        return
    cells = []
    for m in APP_STEPS:
        cells.append(f'<div class="kpi-card" style="flex:1;min-width:0">'
                     f'<div class="kpi-label">{esc(m)}</div>'
                     f'<div class="kpi-value">{fmt_value(m, cur[m])}</div>'
                     f'{_fn_pill(m, cur[m], prv[m], base_tag)}</div>')
    st.markdown('<div style="display:flex;gap:10px;align-items:stretch">'
                + "".join(cells) + '</div>', unsafe_allow_html=True)
    st.caption("**앱설치는 「신규 설치」예요** — 재설치는 빼고 봐요. 이 블록이 묻는 게 "
               "'신규 가입자가 앱까지 오나'라서예요. 전체 설치·재설치는 아래 상세에 있어요.")

    # 비율 — 분모가 없으면 그 줄만 비운다(0으로 두면 없는 걸 있는 것처럼 읽는다)
    #
    # **수신동의는 두 원천이 서로 다른 질문에 답한다.** 기존 PUSH는 *앱을 가진 회원*의
    # 동의라 '앱을 안 깐 신규회원까지 넣으면 얼마인가'엔 답하지 못한다. 둘을 나란히
    # 놓으면 '앱을 안 깔아서 못 받는 사람'이 얼마나 되는지가 그 차이로 드러난다.
    #
    # **앱 보유 무관 원천을 안 올렸으면 화면은 예전 그대로여야 한다.** 구분할 대상이
    # 하나뿐인데 이름에 `(앱포함)`을 달아 두면, 있지도 않은 다른 지표를 찾게 된다.
    # 그래서 이름·안내를 원천 유무로 갈라 둔다 — 올리면 두 줄, 안 올리면 예전 한 줄.
    has_pa = not (pd.isna(cur.get("앱푸시동의_앱무관"))
                  and pd.isna(prv.get("앱푸시동의_앱무관")))
    RATIOS = ([("앱푸시(앱보유무관)/신규회원", "앱푸시동의_앱무관", "가입자수"),
               ("앱푸시(앱포함)/신규회원", "앱푸시수신동의", "가입자수")] if has_pa
              else [("신규회원 앱 수신동의율", "앱푸시수신동의", "가입자수")]) + [
              ("가입자 대비 앱 신규설치율", "앱설치", "가입자수"),
              ("앱 신규설치 대비 수신동의율", "앱푸시수신동의", "앱설치")]
    rrows = []
    for name, num, den in RATIOS:
        if pd.isna(cur.get(num)) and pd.isna(prv.get(num)):
            continue                    # 원천을 아직 안 올린 줄은 아예 빼 둔다
        rc = (cur[num] / cur[den] if not (pd.isna(cur[num]) or pd.isna(cur[den])
                                          or not cur[den]) else np.nan)
        rp = (prv[num] / prv[den] if not (pd.isna(prv[num]) or pd.isna(prv[den])
                                          or not prv[den]) else np.nan)
        rrows.append({"비율": name, "산식": f"{num} ÷ {den}",
                      f"{py}년": ("–" if pd.isna(rp) else f"{rp * 100:.2f}%"),
                      f"{cy}년": ("–" if pd.isna(rc) else f"{rc * 100:.2f}%"),
                      "전년비": (fmt_delta("동의율", rc, rp) or "–")})
    if rrows:
        wtable(style_delta_cols(pd.DataFrame(rrows).set_index("비율")), width="stretch",
               dl_name="신규회원 앱 수신동의")
    # 원천이 없으면 아무 말도 덧붙이지 않는다 — 안 올린 데이터를 채근하는 문구가
    # 상시로 떠 있으면 그것도 소음이다. 올렸을 때만 두 줄을 어떻게 읽는지 말한다.
    if has_pa:
        # 이 원천은 그날의 신규회원 전체(Push Total)도 같이 준다. 마스터 가입자수와
        # 크게 어긋나면 둘 중 하나가 다른 모수를 세고 있다는 뜻이라 눈으로 잡히게 한다.
        _tot, _jn = cur.get("앱푸시대상_앱무관"), cur.get("가입자수")
        _gap = (abs(_tot - _jn) / _jn if not (pd.isna(_tot) or pd.isna(_jn) or not _jn)
                else np.nan)
        _msg = ("**앱푸시(앱보유무관)**은 앱 설치 여부와 상관없이 신규회원 중 "
                "수신동의한 사람이에요. **앱푸시(앱포함)**은 앱을 가진 회원만 세니 "
                "두 줄의 차이가 곧 **앱을 안 깔아서 못 받는 몫**이에요.")
        if not pd.isna(_gap) and _gap > 0.1:
            _msg += (f" 다만 이 원천이 센 신규회원({fmt_value('가입자수', _tot)})이 "
                     f"마스터 가입자수({fmt_value('가입자수', _jn)})와 "
                     f"{_gap * 100:.0f}% 어긋나요 — 모수 정의가 다를 수 있어요.")
        st.caption(_msg)
    # 원천이 같이 준 나머지 칸 — 접이식으로만 (카드를 아홉 장 세우면 퍼널이 안 읽힌다)
    # 신규설치는 위 카드에도 있지만 여기 같이 둔다 — `전체설치 = 신규 + 재설치`가
    # 한 표에서 닫혀야 구성비가 읽힌다.
    extra = [(m, _one(m, cy, "mtd"), _one(m, py, prv_close))
             for m in ["앱설치"] + APP_EXTRA]
    extra = [(m, c, pv) for m, c, pv in extra if not (pd.isna(c) and pd.isna(pv))]
    if extra:
        # 이 원천은 2026-03부터 쌓여서 전년이 **한 칸도 없다**. 그 상태로 두면
        # 전년·전년비 두 열이 통째로 '–'로 남아 자리만 먹는다 — 빼고 대신 같은 기간
        # 안에서 읽히는 구성비를 넣는다. 전년이 쌓이는 해가 오면 자동으로 되돌아온다.
        _has_prev = any(not pd.isna(pv) for _, _, pv in extra)
        _cvs = {m: c for m, c, _ in extra}
        _base = _cvs.get("앱_전체설치")

        def _share(m, c):
            """전체설치를 분모로 한 비중. 스토어방문만 뒤집어 설치 전환율로 읽는다."""
            if m == "앱_스토어방문":
                return ("–" if pd.isna(_base) or pd.isna(c) or not c
                        else f"설치 전환 {_base / c * 100:.1f}%")
            # 전체설치는 자기 자신이 분모고, Push 활성 기기는 잔고라 비중이 없다
            if m in ("앱_전체설치", "앱_Push활성기기"):
                return "–"
            if pd.isna(_base) or pd.isna(c) or not _base:
                return "–"
            return f"{c / _base * 100:.1f}%"

        _rows = []
        for m, c, pv in extra:
            r = {"지표": "신규설치" if m == "앱설치" else m.replace("앱_", "")}
            if _has_prev:
                r[f"{py}년"] = fmt_value(m, pv)
            r[f"{cy}년"] = fmt_value(m, c)
            if _has_prev:
                r["전년비"] = fmt_delta(m, c, pv) or "–"
            r["전체설치 대비"] = _share(m, c)
            _rows.append(r)
        with st.expander(f"앱설치 상세 ({len(extra)}개 지표)"):
            wtable(style_delta_cols(pd.DataFrame(_rows).set_index("지표")),
                   width="stretch", dl_name="앱설치 상세")
            _cap = ("값은 모두 **일평균**이에요. 전체설치 = 신규설치 + 재설치고, "
                    "Push 활성 기기는 그날의 잔고예요(합이 아니라 수준).")
            if not _has_prev:
                _cap += (" 이 원천은 **올해분부터** 쌓여서 전년 칸이 통째로 비어요 — "
                         "대신 전체설치 대비 비중을 넣었어요. 스토어방문 줄만 "
                         "**설치 전환율**(전체설치 ÷ 스토어방문)이에요.")
            st.caption(_cap)

    _funnel_app_trend(df, gran, cy, clabel)

    # 비율을 어디까지 믿고 읽을지 — 안 적어 두면 '가입한 사람의 몇 %가 앱을 깔았다'로
    # 읽힌다. 원천이 가입↔설치를 개인 단위로 잇지 않아 거기까진 알 수 없다.
    notes = ["비율은 모두 **일평균끼리 나눈 값**이에요. 그날 가입한 사람이 그날 "
             "설치·동의했다는 뜻은 아니에요 — 원천이 가입↔설치를 개인 단위로 "
             "잇지 않아 거기까진 알 수 없어요."]
    # 어느 날까지 들어와 있는지 — 기간은 맞는데 원천이 아직 안 닿았을 수 있다
    _cov = df[(df["metric"] == "앱설치") & (df["gran"] == "일")
              & (df["segment"] == "*TOTAL") & df["value"].notna()]
    if not _cov.empty:
        _cs = _cov.sort_values("sortkey")
        _c0, _c1 = _cs.iloc[0], _cs.iloc[-1]
        notes.append(f"앱설치 원천은 **{int(_c0['year'])}/{_c0['label']} ~ "
                     f"{int(_c1['year'])}/{_c1['label']}** 까지 들어와 있어요 "
                     f"(총 {_cov['label'].nunique():,}일).")
    if n_push:
        notes.append(f"앱푸시 수신동의는 이 기간 **{n_push}일치**를 일평균으로 묶은 값이에요.")
    if pd.isna(cur["앱설치"]) and _cv.empty:
        notes.append("**앱설치** 원천이 아직 안 올라와서 두 줄이 비어 있어요. "
                     "Daily 통계 플랫폼 export(`전체 설치`·`신규 설치` 칸이 있는 파일)를 "
                     "올리면 채워져요.")
    elif pd.isna(prv["앱설치"]):
        notes.append("**앱설치는 전년 데이터가 없어** 전년비가 비어 있어요. "
                     "원천이 올해분부터 쌓이기 시작했어요.")
    if approx:
        notes.append("앱푸시 수신동의는 원천이 **일별로만** 와서 주 단위는 "
                     "`(일-1)//7+1` 규칙으로 묶은 근사값이에요. 마스터 주차와 경계가 "
                     "하루 이틀 어긋날 수 있어요.")
    for n in notes:
        st.caption(n)
    _bad = pick(df, gran, APPINSTALL_FLAG, "*TOTAL", cy, clabel, "mtd")
    if not pd.isna(_bad) and _bad > 0:
        st.warning(f"⚠️ 이 기간에 **앱 기기 수가 절반 이하로 찍힌 날이 {int(_bad)}일** 있어요. "
                   "한쪽 플랫폼(Android)이 빠진 export로 보여요 — 그 날들 때문에 앱설치 "
                   "일평균이 실제보다 낮게 나와요. **값은 원천 그대로 두었어요**. "
                   "원천을 다시 받아 보시고, 같은 증상이면 그 날 수치는 빼고 읽으세요.")



# ══════════════════════════════════════════════════════
# 조직·카테고리·채널별 첫구매 상세
# ══════════════════════════════════════════════════════
# MICRO export가 **구분08에 채널**을 실어 오면서 '어느 조직의 어느 채널에서 빠졌나'를
# 한 화면에서 볼 수 있게 됐다. 「조직·카테고리별 실적」은 채널을 안 가르고(전 채널) 조직을
# 깊이 파고드는 화면이고, 여기는 **조직 × 채널 교차**가 주인공이라 화면을 따로 둔다 —
# 한 화면에 다 넣으면 축이 셋(조직·카테고리·채널)이라 표가 안 읽힌다.
# 페이지 이름 상수(`PAGE_ORGCAT_CH`)는 파일 위쪽 `PAGE_ORDER` 옆에 모아 뒀다.


def _occ_table(views, path, kids, lv_lbl, met, cy, py, clabel, chs):
    """행=자식(맨 위 합계), 열=**왼쪽 실적 · 오른쪽 전년비** — 화면 전체와 같은 얼굴.

    합계 행 값은 **파일이 준 상위 값**이지 자식 합이 아니다. 고객수·상품UV는 유니크라
    자식 합이 상위를 넘고 객단가·상품CR은 애초에 더할 수 없다(「조직·카테고리별 실적」과
    같은 규칙). 「전체」 열도 마찬가지로 파일이 준 전 채널 값이라 고른 채널의 합과
    꼭 같진 않다 — 캡션에 밝힌다.
    """
    cols = [(ORGCAT_CH_ALL, "전체")] + [(c, c) for c in chs]
    vals, dlts = [], []
    for nm, node in [(TOTAL_ROW, tuple(path))] + [(k, tuple(path) + (k,)) for k in kids]:
        v, d = {lv_lbl: nm}, {}
        for key, cn in cols:
            cv = views[key].get(node, met, cy, clabel, "mtd")
            pv = views[key].get(node, met, py, clabel, "final")
            v[cn] = fmt_value(met, cv)
            d[f"{cn} 전년비"] = fmt_delta(met, cv, pv) or "–"
        vals.append(v); dlts.append(d)
    return pd.concat([pd.DataFrame(vals), pd.DataFrame(dlts)], axis=1)


def _occ_trend(views, path, node_lbl, met, cy, py, chs, gran, scope):
    """② 채널별 추이 — 선=채널, 색은 `CHANNEL_PAL`이라 화면을 오가도 같은 색이다.

    차트만 두면 '9월 2주차가 정확히 얼마였나'를 못 읽어 결국 엑셀을 받게 된다 —
    아래에 같은 값을 표로도 낸다(②·⑤·03과 같은 얼굴).
    """
    _lb = (scope[(scope["year"] == cy) & scope["value"].notna()]
           [["label", "sortkey"]].drop_duplicates()
           .sort_values("sortkey")["label"].astype(str).tolist())
    if not _lb:
        st.caption(f"{cy}년 «{FUNNEL_GRAN_LABEL.get(gran, gran)}» 값이 없어요.")
        return
    ser = {c: [views[c].get(path, met, cy, lb, "mtd") for lb in _lb] for c in chs}
    prv = {c: [views[c].get(path, met, py, lb, "final") for lb in _lb] for c in chs}
    # 아무 채널에도 값이 없는 기간은 축에서 뺀다 — 그 자리에서 모든 선이 나란히 끊겨
    # 데이터가 빠진 것처럼 보인다. 일부 채널에만 없는 칸은 남겨 선을 끊는다(진짜 결측).
    keep = [j for j in range(len(_lb))
            if any(not pd.isna(v[j]) for v in ser.values())]
    if not keep:
        st.caption(f"«{esc(node_lbl)}»의 {gran} 단위 값이 없어요.")
        return
    _lb = [_lb[j] for j in keep]
    unit, div = METRIC_UNIT.get(met, ("", 1))
    if met in PCT_METRICS:
        div, unit = 0.01, "%"
    fig = go.Figure()
    for c in chs:
        ys = [ser[c][j] for j in keep]
        if not any(not pd.isna(v) for v in ys):
            continue
        fig.add_trace(go.Scatter(
            x=[month_trim(v) for v in _lb],
            y=[None if pd.isna(v) else v / div for v in ys],
            mode="lines+markers", name=c, connectgaps=False,
            line=dict(color=clr(CHANNEL_PAL.get(c, "blue")), width=1.8),
            marker=dict(size=4)))
    ly = base_layout(340, ysuffix=unit if unit == "%" else "",
                     title=f"{node_lbl} · {met} 채널별 추이 ({unit})")
    ly["xaxis"]["categoryorder"] = "array"
    ly["xaxis"]["categoryarray"] = [month_trim(v) for v in _lb]
    if gran in ("주", "일"):
        ly["xaxis"]["tickangle"] = -45
        ly["xaxis"]["nticks"] = 20 if gran == "주" else 14
    fig.update_layout(**ly)
    st.plotly_chart(fig, width="stretch")

    rows, dl = {}, {}
    for c in chs:
        rows[c] = {month_trim(_lb[i]): fmt_value(met, ser[c][keep[i]])
                   for i in range(len(_lb))}
        dl[c] = {f"{month_trim(_lb[i])} 전년비":
                 (fmt_delta(met, ser[c][keep[i]], prv[c][keep[i]]) or "–")
                 for i in range(len(_lb))}
    tt = pd.concat([pd.DataFrame(rows).T, pd.DataFrame(dl).T], axis=1)
    tt.index.name = "채널"
    wtable(style_delta_cols(tt), width="stretch",
           dl_name=f"{node_lbl} {met} 채널별 추이 ({cy}년)")
    st.caption(f"위 차트와 같은 값이에요. 왼쪽은 {cy}년 실적, 오른쪽은 **{py}년 같은 "
               f"기간 대비 전년비**예요. {cy}년 **{len(_lb)}개 기간 전체**를 담았어요.")


def render_orgcat_channel_page(odf):
    """조직·카테고리·채널별 첫구매 상세 — 행=조직/카테고리, 열=채널."""
    st.markdown("## " + PAGE_ORGCAT_CH)
    if odf is None or odf.empty:
        st.info("조직×카테고리 데이터가 없어요. MICRO 대시보드의 "
                "**구분06 × 구분07 × 구분08(채널)** export를 사이드바에 올려 주세요.")
        return
    if not [c for c in CHANNELS if (odf["ch"] == c).any()]:
        st.info("올려 주신 조직×카테고리 데이터엔 **채널 축이 없어요**. 구분08에 채널"
                "(직접·광고·EP…)이 실린 export를 올리면 여기서 갈라 볼 수 있어요.")
        st.caption(f"채널 축이 없는 파일은 「{esc(PAGE_ORGCAT)}」에서 그대로 보면 돼요.")
        return

    grans = [g for g in ("월", "주", "일") if (odf["gran"] == g).any()]
    f1, f2, f3 = st.columns([1.2, 1, 1.6])
    with f1:
        guard_select("occ_gran", grans)
        gran = st.radio("집계 단위", grans, horizontal=True, key="occ_gran")
    # LFMS는 모집단이 다른 축이라 **고른 단위 안에서** 뽑는다(「조직·카테고리별 실적」과 같은 이유).
    lfmss = sorted(odf[odf["gran"] == gran]["lfms"].dropna().astype(str).unique())
    with f2:
        if len(lfmss) > 1:
            guard_select("occ_lfms", lfmss)
            lf = st.radio("LFMS 포함", lfmss, horizontal=True, key="occ_lfms",
                          help="포함/미포함은 모집단이 달라서 섞어 보면 안 돼요.")
        else:
            lf = lfmss[0] if lfmss else "N"
            st.caption(f"LFMS 포함여부 **{esc(lf)}** 데이터만 있어요")
    scope = odf[(odf["gran"] == gran) & (odf["lfms"] == lf)]
    pref = list(ORGCAT_MAP.values()) + ["상품UV", "상품CR"]
    mets = [m for m in pref if (scope["metric"] == m).any()]
    if not mets:
        st.info("고른 조건에 첫구매 지표가 없어요."); return
    with f3:
        guard_select("occ_met", mets)
        met = st.selectbox("진단 지표", mets, key="occ_met",
                           help="표·차트가 이 지표를 따라가요.")

    # ── 채널 복수 선택 ──
    # 여덟을 다 켜면 표가 옆으로 길어진다. 기본은 「채널별 실적」 차트와 같은 넷
    # (`CHART_CH_DEFAULT`)이라 두 화면을 오갈 때 같은 채널을 보게 된다.
    chs_here = [c for c in CHANNELS if (scope["ch"] == c).any()]
    if not chs_here:
        st.info(f"«{esc(gran)}» 단위엔 채널별 값이 없어요. 다른 단위를 골라 주세요."); return
    guard_multi("occ_chs", chs_here)
    if "occ_chs" not in st.session_state:
        _pre = [c for c in CHART_CH_DEFAULT if c in chs_here]
        st.session_state["occ_chs"] = _pre or chs_here
    chs = st.multiselect("채널 (복수 선택)", chs_here, key="occ_chs",
                         help="고른 채널만 표·차트에 올려요. 「전체」 열은 늘 같이 보여 줘요.")
    if not chs:
        st.info("채널을 하나 이상 골라 주세요."); return

    # ── 기준 기간 — 사이드바가 아니라 여기서 고른다(커버리지가 마스터와 다르다) ──
    labs = (scope[scope["value"].notna()][["year", "label", "sortkey"]]
            .drop_duplicates().sort_values("sortkey"))
    if labs.empty:
        st.info("고른 조건에 값이 없어요."); return
    opts = [(int(r["year"]), str(r["label"])) for _, r in labs.iterrows()][::-1]
    guard_select("occ_per", opts)
    cy, clabel = st.selectbox("기준 기간", opts, key="occ_per",
                              format_func=lambda p: f"{p[0]}년 {month_trim(p[1])}")
    py = cy - 1

    # 채널마다 뷰를 **한 번씩만** 만든다 — 노드마다 프레임을 훑으면 조직 10 × 채널 8이
    # 그대로 풀스캔 횟수가 된다(「조직·카테고리별 실적」에서 겪은 것과 같은 자리).
    views = {c: orgcat_view(scope[scope["ch"] == c]) for c in [ORGCAT_CH_ALL] + chs}
    orgs = views[ORGCAT_CH_ALL].live(())[0]
    if not orgs:
        st.info("조직 항목이 없어요."); return

    st.markdown("### ① 조직 × 채널")
    _pre = _picked_row(st.session_state.get("occ_orgsel"), len(orgs) + 1)
    tbl = _occ_table(views, [], orgs, "조직", met, cy, py, clabel, chs)
    sty = style_delta_cols(tbl)
    ev = wtable(_hl_row(sty, _pre), width="stretch", key="occ_orgsel",
                on_select="rerun", selection_mode="single-cell", hide_index=True,
                dl_name=f"조직×채널 {met} ({cy}년 {month_trim(clabel)})", dl_data=sty)
    st.caption(f"{cy}년 {month_trim(clabel)} · 왼쪽 실적, 오른쪽 전년비예요. 「전체」는 "
               "**파일이 준 전 채널 값**이라 고른 채널의 합과 꼭 같진 않아요. 맨 윗줄은 "
               "합계고요. 조직 한 줄에서 **아무 칸이나 누르면** 아래에 그 조직의 "
               "카테고리가 열려요.")

    picked = _picked_child(ev, orgs)
    path, node_lbl = [], "전체"
    if picked:
        path, node_lbl = [picked], picked
        kids = views[ORGCAT_CH_ALL].live((picked,))[0]
        st.markdown(f"###### {esc(picked)} — 카테고리 × 채널")
        if not kids:
            st.info(f"«{esc(picked)}» 아래에 볼 카테고리가 없어요.")
        else:
            # 조직마다 키를 갈라 둔다 — 같은 키면 조직을 바꿔도 옛 행 번호가 남아
            # 엉뚱한 카테고리가 열린다.
            _ck = f"occ_catsel_{picked}"
            _p2 = _picked_row(st.session_state.get(_ck), len(kids) + 1)
            t2 = _occ_table(views, [picked], kids, "카테고리", met, cy, py, clabel, chs)
            s2 = style_delta_cols(t2)
            ev2 = wtable(_hl_row(s2, _p2), width="stretch", key=_ck, on_select="rerun",
                         selection_mode="single-cell", hide_index=True,
                         dl_name=f"{picked} 카테고리×채널 {met}", dl_data=s2)
            _c2 = _picked_child(ev2, kids)
            if _c2 is not None:
                path, node_lbl = [picked, _c2], f"{picked} › {_c2}"

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.markdown(f"### ② {esc(node_lbl)} — 채널별 추이")
    _occ_trend(views, tuple(path), node_lbl, met, cy, py, chs, gran, scope)


def render_orgcat_page(odf, ddf=None):
    """06. 조직·카테고리별 실적 — 개괄에서 이상한 데를 찾아 그 자리에서 파고드는 화면.

    원천이 둘이다.
      · **MICRO 조직×카테고리**(구분06~09) — 조직 > 카테고리까지 온다. 값은 일평균.
      · **브랜드·상품 원장**(결제 원장) — 조직 > 카테고리 > 브랜드 > 상품 네 단계가 다 있고
        유입채널(AF대분류) 축이 하나 더 있다. 일별 합계를 일평균으로 환산해 쓴다.
    **카테고리 어휘가 서로 달라**(MICRO는 골프·남성·잡화, 원장은 가방·패션소품·향수…)
    한 축에 섞으면 안 된다 — 「데이터 소스」로 갈라 본다.

    사이드바 기준 기간(월·주)에 기대지 않고 **자체 기간 선택**을 쓴다. 이 데이터는 일
    단위까지 오고 커버리지도 마스터와 달라서, 사이드바 값을 그대로 쓰면 빈 화면이 되기
    쉽다. 덕분에 마스터 데이터가 아직 없어도 이 화면만은 열린다.
    """
    st.markdown("## 조직·카테고리별 첫구매 실적")
    has_o = odf is not None and not odf.empty
    has_d = ddf is not None and not ddf.empty
    if not has_o and not has_d:
        st.info("조직·카테고리 데이터가 없어요. MICRO 대시보드의 "
                "**구분06(BPU) × 구분07(카테고리)** 엑셀이나, **브랜드·상품 결제 원장**"
                "(`결제_일자`·`BPU`·`대카테고리명`·`ADMIN브랜드명`·`상품명`·`거래액`)을 "
                "사이드바에 올려 주세요.")
        st.caption("일자별·주별·월별 파일을 따로 올리면 각각 누적돼요. "
                   "원장을 올리면 **브랜드·상품**까지 파고들 수 있어요.")
        return

    SRC_M, SRC_D = "MICRO 조직×카테고리", "브랜드·상품 원장"
    srcs = ([SRC_M] if has_o else []) + ([SRC_D] if has_d else [])
    if len(srcs) > 1:
        guard_select("oc_src", srcs)
        source = st.radio("데이터 소스", srcs, horizontal=True, key="oc_src",
                          help="두 원천은 카테고리 어휘가 서로 달라서 섞어 보면 안 돼요. "
                               "원장 쪽만 브랜드·상품까지 열려요.")
    else:
        source = srcs[0]
    detail = source == SRC_D

    # ── 필터: 집계 단위 · 두 번째 축(LFMS / 유입채널) · 지표 ──
    grans = (DETAIL_GRANS if detail else
             [g for g in ("월", "주", "일") if (odf["gran"] == g).any()])
    if not grans:
        st.info("고른 조건에 데이터가 없어요."); return
    f1, f2, f3 = st.columns([1.2, 1, 1.6])
    with f1:
        guard_select("oc_gran", grans)
        gran = st.radio("집계 단위", grans, horizontal=True, key="oc_gran")

    if detail:
        sdf = detail_stamp(ddf, gran)
        chs = [c for c in CHANNELS if (ddf["ch"] == c).any()]
        chs += sorted(set(ddf["ch"].astype(str)) - set(chs))
        axes = [DETAIL_ALL] + chs
        with f2:
            guard_select("oc_lfms", axes)
            axis = st.selectbox("유입 채널", axes, key="oc_lfms",
                                help="원장의 `AF대분류명`이에요. 고르면 그 채널만 봐요 — "
                                     "채널마다 브랜드 구성이 달라서 섞어 보면 원인이 흐려져요.")
        cube = detail_provider(sdf, gran, axis)
        base = cube(())
        DEPTH = len(ORGCAT_LEVELS)
    else:
        # LFMS 선택지는 **고른 단위 안에서** 뽑는다. 단위마다 받아 온 export가 달라
        # (예: 일별만 LFMS=Y) 전역 목록을 쓰면 '일'로 바꿨을 때 이전 선택 'N'이 남아
        # 데이터가 있는데도 빈 화면이 된다.
        lfmss = sorted(odf[(odf["gran"] == gran) & (odf["ch"] == ORGCAT_CH_ALL)]
                       ["lfms"].dropna().astype(str).unique())
        with f2:
            if len(lfmss) > 1:
                guard_select("oc_lfms", lfmss)
                axis = st.radio("LFMS 포함", lfmss, horizontal=True, key="oc_lfms",
                                help="원본 헤더의 'LFMS 포함여부' 값이에요. 포함/미포함은 "
                                     "모집단이 달라서 섞어 보면 안 돼요.")
            else:
                axis = lfmss[0] if lfmss else "N"
                st.caption(f"LFMS 포함여부 **{axis}** 데이터만 있어요")
        # 채널 축은 「조직·카테고리·채널별 첫구매 상세」가 맡는다 — 여기선 전 채널.
        base = odf[(odf["gran"] == gran) & (odf["lfms"] == axis)
                   & (odf["ch"] == ORGCAT_CH_ALL)]
        DEPTH = orgcat_depth(base)

        def cube(_path):
            return base

    if base.empty:
        st.info("고른 조건에 데이터가 없어요."); return
    pref = (DETAIL_METS if detail else
            list(ORGCAT_MAP.values()) + ["상품UV", "상품CR", "거래액비중", "고객비중"])
    mets = [m for m in pref if (base["metric"] == m).any()]
    mets += [m for m in base["metric"].unique() if m not in mets]
    with f3:
        guard_select("oc_met", mets)
        met = st.selectbox("진단 지표", mets, key="oc_met",
                           help="표·차트·히트맵이 이 지표를 따라가요. 기여도 분해는 "
                                "하위 합이 상위와 일치하는 거래액에서만 나와요.")

    # ── 기준 기간 ──
    labs = (base[base["value"].notna()][["year", "label", "sortkey"]]
            .drop_duplicates().sort_values("sortkey"))
    if labs.empty:
        st.info("고른 조건에 값이 없어요."); return
    opts = [(int(r["year"]), str(r["label"])) for _, r in labs.iterrows()][::-1]
    # 아직 다 안 찬 기간은 close=mtd로 온다 — 값이 다 일평균이라 완결 기간과 견줄 수는
    # 있지만 며칠치 표본이라 크게 흔들린다. 어느 기간이 그런지 라벨에 박아 둔다.
    _fin = {(int(r["year"]), str(r["label"]))
            for _, r in base[(base["close"] == "final") & base["value"].notna()]
            [["year", "label"]].drop_duplicates().iterrows()}
    part = {p for p in opts if p not in _fin}
    guard_select("oc_per", opts)
    cy, clabel = st.selectbox("기준 기간", opts, key="oc_per",
                              format_func=lambda p: f"{p[0]}년 {month_trim(p[1])}"
                              + (" · 부분 기간" if p in part else ""))
    py = cy - 1
    ccol, pcol = f"{cy}년 {month_trim(clabel)}", f"{py}년 {month_trim(clabel)}"
    has_prev = not base[(base["year"] == py) & (base["label"] == clabel)].empty
    if (cy, clabel) in part:
        # MICRO는 '일마감'만 있는 진행 중 기간이고, 원장은 적재 범위 끝(또는 처음)이
        # 달·주 가운데를 자른 경우다 — 둘 다 며칠치라 완결 기간과 무게가 다르다.
        st.info(f"**{ccol}은 "
                + ("아직 진행 중이에요(일마감)." if not detail else
                   "기간이 다 안 찼어요 (적재 범위가 이 기간을 다 덮지 않아요).")
                + "** 값이 모두 일평균이라 완결 기간과 견줄 수는 있지만, 며칠치라 크게 "
                  "흔들려요. 추세 판단은 완결된 기간으로 하세요.")
    if not has_prev:
        st.warning(f"전년({py}년) 같은 기간 데이터가 없어 전년비·기여도는 비어 있어요.")

    unit, div = METRIC_UNIT.get(met, ("", 1))
    is_pct = met in PCT_METRICS
    xdiv = 0.01 if is_pct else div
    additive = met in ORGCAT_ADDITIVE
    LV = ORGCAT_LEVELS[:DEPTH]
    FACTORS = DETAIL_FACTORS if detail else ORGCAT_FACTORS
    # `sub`이 바뀔 때마다 인덱스를 **한 번만** 다시 만든다. 노드마다 프레임을 훑으면
    # 브랜드 300개 × 지표 3종 × 2개년이 그대로 풀스캔 횟수가 된다(실측 5.5초).
    sub, view = base, orgcat_view(base)

    def V(path, year, metric=None):
        return view.get(tuple(path), metric or met, year, clabel)

    # ── 파고들기 — 레벨마다 셀렉트 하나. 위를 '전체'로 되돌리면 아래는 같이 닫힌다 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    show_all = st.checkbox("실적 없는 항목도 보기", key="oc_showall",
                           help="미매칭·기타처럼 전 기간 거래액이 상위의 0.1%에 못 미치는 "
                                "항목이에요. 기본은 숨겨서 표와 차트를 비워 둬요.")
    path, hidden_here, dcols = [], [], st.columns(max(len(LV), 1))
    for d, (col, lb, _h) in enumerate(LV):
        sub = cube(tuple(path))          # 원장 소스는 경로가 깊어질 때마다 그 아래를 만든다
        view = orgcat_view(sub)
        live, hid = view.live(path)
        kids = (view.children(path) if show_all else live)
        if not kids:
            break
        key = f"oc_lv{d}"
        guard_select(key, ["전체"] + kids)
        with dcols[d]:
            pick_v = st.selectbox(f"{d + 1}. {lb}", ["전체"] + kids, key=key,
                                  help=("여기서 고르면 한 단계 더 들어가요."
                                        if d + 1 < len(LV) else "마지막 단계예요."))
        if d == 0:
            hidden_here = hid
        if pick_v == "전체":
            break
        path.append(pick_v)
    sub = cube(tuple(path))
    view = orgcat_view(sub)
    if len(LV) < len(ORGCAT_LEVELS):
        st.caption("브랜드·상품(구분08·09)은 이 원천에 값이 없어 단계가 안 열려요. "
                   "**브랜드·상품 원장**을 올리면 네 단계로 파고들 수 있어요.")
    if hidden_here and not show_all:
        st.caption(f"실적이 거의 없는 {LV[0][1]} {len(hidden_here)}곳은 숨겼어요 — "
                   + ", ".join(esc(str(o)) for o in hidden_here[:6])
                   + (" 외" if len(hidden_here) > 6 else "") + ".")

    crumb = " › ".join(["전체"] + [str(p) for p in path])
    _axis_txt = (f"{axis} 채널" if detail else f"LFMS {axis}")
    st.caption(f"보는 곳: **{esc(crumb)}** · {ccol} · {met} · {esc(_axis_txt)}")

    # 지금 노드와 한 단계 아래
    cur_c, cur_p = V(path, cy), V(path, py)
    kid_live, kid_hidden = view.live(path)
    kid_names = (view.children(path) if show_all else kid_live)
    kid_lbl = LV[len(path)][1] if len(path) < len(LV) else None

    rows = []
    for name in kid_names:
        kp = tuple(path) + (name,)
        vc, vp = V(kp, cy), V(kp, py)
        if not np.isfinite(vc) and not np.isfinite(vp):
            continue
        d_ = (vc - vp) if (np.isfinite(vc) and np.isfinite(vp)) else np.nan
        rows.append({"_n": name, "_p": kp, "_vc": vc, "_vp": vp, "_d": d_,
                     "_share": (d_ / cur_p * 100)
                     if (additive and np.isfinite(d_) and np.isfinite(cur_p) and cur_p)
                     else np.nan})
    kid = pd.DataFrame(rows)
    show = kid

    # ── 진단 요약 ──
    _msg = []
    if np.isfinite(cur_c):
        _delta = f" · 전년비 {fmt_delta(met, cur_c, cur_p) or '–'}" if np.isfinite(cur_p) else ""
        _amt = f" ({_won_m(cur_c - cur_p)})" if additive and np.isfinite(cur_p) else ""
        _msg.append(f"**{esc(crumb)}**의 {ccol} {met}은 **{fmt_value(met, cur_c)}**"
                    f"{_delta}{_amt}이에요.")
    if len(kid) and kid["_d"].notna().any() and kid_lbl:
        _w = kid.reindex(kid["_d"].abs().sort_values(ascending=False).index).iloc[0]
        _dir = "끌어내린" if _w["_d"] < 0 else "끌어올린"
        # 전체 변화 대비 몫. 서로 상쇄되는 구간이면 100%를 넘는 게 정상이지만,
        # 전체가 거의 안 움직인 달엔 300%·800% 같은 무의미한 수가 나와 생략한다.
        _sh = ""
        if (additive and np.isfinite(cur_p) and np.isfinite(cur_c)
                and (cur_c - cur_p) != 0 and np.isfinite(_w["_d"])):
            _r = _w["_d"] / (cur_c - cur_p)
            if abs(_r) <= 3:
                _sh = f", 전체 변화의 **{abs(_r) * 100:.0f}%**"
        _msg.append(f"가장 크게 {_dir} {kid_lbl}{josa(kid_lbl)} **{esc(str(_w['_n']))}**"
                    + (f" ({_won_m(_w['_d'])}{_sh})" if additive else
                       f" ({fmt_delta(met, _w['_vc'], _w['_vp']) or '–'})") + "예요.")
    fp = tuple(V(path, py, m) for m, _ in FACTORS)
    fc = tuple(V(path, cy, m) for m, _ in FACTORS)
    fs = factor_split(fp, fc)
    if fs:
        parts, _tot = fs
        _nm = FACTORS[max(range(len(parts)), key=lambda i: abs(parts[i]))][1]
        _msg.append("거래액 변화를 쪼개면 "
                    + " · ".join(f"{lb} {_won_m(parts[i])}"
                                 for i, (_, lb) in enumerate(FACTORS))
                    + f" — **{_nm}**{josa(_nm, '이가')} 가장 크게 움직였어요.")
    if _msg:
        st.markdown('<div class="vg">📌 ' + "<br>".join(_msg) + '</div>',
                    unsafe_allow_html=True)

    # ── ① 어디가 움직였나 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader(f"① 어디가 움직였나" + (f" — {kid_lbl}별" if kid_lbl else ""))
    if not kid_lbl:
        st.info("더 들어갈 단계가 없어요. 위에서 한 단계 위를 '전체'로 바꾸면 돌아가요.")
    elif not len(kid):
        st.info(f"이 기간에 {kid_lbl} 값이 없어요.")
    else:
        show = kid.reindex(
            (kid["_d"].abs() if additive else kid["_vc"]).sort_values(ascending=False).index)
        trow = []
        for _, r in show.iterrows():
            item = {kid_lbl: r["_n"], pcol: fmt_value(met, r["_vp"]),
                    ccol: fmt_value(met, r["_vc"]),
                    "전년비": fmt_delta(met, r["_vc"], r["_vp"]) or "–"}
            if additive:
                item["증감액"] = _won_m(r["_d"]) if np.isfinite(r["_d"]) else "–"
                item["기여도"] = f"{r['_share']:+.1f}%p" if np.isfinite(r["_share"]) else "–"
            # 같은 줄에서 원인 축까지 보이게 — 요인별 전년비를 붙인다
            for mkey, lb in FACTORS:
                item[f"{lb} 전년비"] = fmt_delta(mkey, V(r["_p"], cy, mkey),
                                               V(r["_p"], py, mkey)) or "–"
            trow.append(item)
        wtable(style_delta_cols(pd.DataFrame(trow).set_index(kid_lbl)),
               width="stretch", dl_name=f"{crumb} {kid_lbl}별 {met}")
        if additive:
            _hid_n = 0 if show_all else len(kid_hidden)
            st.caption(f"**기여도**는 그 {kid_lbl}의 증감액을 **{pcol} {crumb}**로 나눈 값이에요. "
                       + (f"숨긴 {_hid_n}곳까지 더해야 정확히 맞아요(합쳐도 0.1% 미만). "
                          if _hid_n else "다 더하면 위 전년비와 같아요. ")
                       + "정렬은 증감액이 큰 순이에요."
                       + f" 더 들어가려면 위 **{len(path) + 1}. {kid_lbl}**에서 고르세요.")
        elif detail:
            st.caption(f"기여도 분해는 거래액에서만 나와요 — {met}은 더해서 상위가 되는 "
                       "값이 아니에요. 정렬은 값이 큰 순이에요.")
        else:
            st.caption(f"{met}은 유니크 값이라 {kid_lbl} 합이 상위와 달라요 — 기여도 분해는 "
                       "거래액에서만 나와요. 정렬은 값이 큰 순이에요.")

        if additive and np.isfinite(cur_p) and show["_d"].notna().any():
            # 상위 8개만 막대로 세우고 나머지는 '기타'로 합친다. 다 세우면 눈금이 뭉개져
            # 정작 큰 항목이 안 읽힌다.
            TOPN = 8
            wall = show[show["_d"].notna()]
            ws, rest = wall.head(TOPN), wall.iloc[TOPN:]
            resid = (cur_c - cur_p) - ws["_d"].sum()      # 나머지 + 숨긴 항목 + 합계 불일치
            _extra = abs(resid) > 5e4
            _rlab = f"기타 {len(rest) + (0 if show_all else len(kid_hidden))}곳" \
                if (len(rest) or kid_hidden) else "기타·미분류"
            names = [pcol] + [str(x) for x in ws["_n"]] + ([_rlab] if _extra else []) + [ccol]
            meas = (["absolute"] + ["relative"] * len(ws)
                    + (["relative"] if _extra else []) + ["total"])
            vals = [cur_p / 1e6] + [v / 1e6 for v in ws["_d"]] + \
                   ([resid / 1e6] if _extra else []) + [cur_c / 1e6]
            figw = go.Figure(go.Waterfall(
                orientation="v", measure=meas, x=names, y=vals,
                text=[f"{v:+,.1f}" if m == "relative" else f"{v:,.1f}"
                      for v, m in zip(vals, meas)],
                textposition="outside",
                connector=dict(line=dict(color="#cbd5e1", width=1)),
                increasing=dict(marker=dict(color=clr("green"))),
                decreasing=dict(marker=dict(color=clr("red"))),
                totals=dict(marker=dict(color=clr("slate")))))
            lay = base_layout(380, title=f"{crumb} · {kid_lbl}별 증감 분해 (백만원)")
            lay["showlegend"] = False
            figw.update_layout(**lay)
            st.plotly_chart(figw, width="stretch")

    # ── ② 왜 — 요인 분해 ──
    _fnames = "·".join(lb for _, lb in FACTORS)
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader(f"② 왜 그랬나 — {_fnames} (거래액 기준)")
    if not fs:
        st.info("이 대상은 전년 값이 없거나 0이라 요인 분해를 할 수 없어요.")
    else:
        parts, tot_d = fs
        _absum = sum(abs(x) for x in parts)      # 전년과 값이 똑같으면 0이라 나눌 수 없다
        frow = []
        for i, (mkey, lb) in enumerate(FACTORS):
            vp_, vc_ = V(path, py, mkey), V(path, cy, mkey)
            frow.append({"요인": lb, "지표": mkey,
                         pcol: fmt_value(mkey, vp_), ccol: fmt_value(mkey, vc_),
                         "전년비": fmt_delta(mkey, vc_, vp_) or "–",
                         "거래액 기여": _won_m(parts[i]),
                         "설명력": (f"{abs(parts[i]) / _absum * 100:.0f}%"
                                 if _absum else "–")})
        wtable(style_delta_cols(pd.DataFrame(frow).set_index("요인")),
               width="stretch", dl_name=f"{crumb} 요인 분해")
        figf = go.Figure(go.Waterfall(
            orientation="v",
            measure=["absolute"] + ["relative"] * len(parts) + ["total"],
            x=[pcol] + [lb for _, lb in FACTORS] + [ccol],
            y=[(cur_p or 0) / 1e6] + [p / 1e6 for p in parts] + [(cur_c or 0) / 1e6],
            text=[f"{(cur_p or 0)/1e6:,.1f}"] + [f"{p/1e6:+,.1f}" for p in parts]
                 + [f"{(cur_c or 0)/1e6:,.1f}"],
            textposition="outside",
            connector=dict(line=dict(color="#cbd5e1", width=1)),
            increasing=dict(marker=dict(color=clr("green"))),
            decreasing=dict(marker=dict(color=clr("red"))),
            totals=dict(marker=dict(color=clr("slate")))))
        lay = base_layout(360, title=f"{crumb} · 거래액 증감을 요인으로 (백만원)")
        lay["showlegend"] = False
        figf.update_layout(**lay)
        st.plotly_chart(figf, width="stretch")
        if detail:
            st.caption("원장엔 상품UV·상품CR이 없어 **거래액 = 고객수 × 객단가**로 쪼개요. "
                       "정의상 정확히 맞는 식이라 두 기여액을 더하면 실제 증감과 딱 맞고"
                       "(LMDI), 어느 쪽을 먼저 대입하느냐로 답이 달라지지 않아요. "
                       "**고객수**가 크면 산 사람이 줄어든 것, **객단가**면 사긴 샀는데 "
                       "덜 쓴 거예요.")
        else:
            st.caption("**거래액 = 상품UV × 상품CR × 객단가**가 원본에서 정확히 성립해요. "
                       "세 기여액을 더하면 실제 증감과 딱 맞고(LMDI), 어느 요인을 먼저 "
                       "대입하느냐로 답이 달라지지 않아요. "
                       "**유입**이 크면 사람이 덜 들어온 것, **전환**이면 들어와서 안 산 것, "
                       "**객단가**면 사긴 샀는데 덜 쓴 거예요.")

    # ── ③ 추이 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader(f"③ 추이 — {crumb}")
    tsub = sub[(sub["metric"] == met)]
    tsub = tsub[orgcat_node(tsub, tuple(path))]
    if tsub.empty:
        st.info("추이를 그릴 값이 없어요.")
    else:
        order = (tsub[tsub["year"] == cy][["label", "sortkey"]]
                 .drop_duplicates().sort_values("sortkey"))
        xs = [str(x) for x in order["label"]]
        MAXP = {"일": 90, "주": 53, "월": 12}[gran]
        xs = xs[-MAXP:]
        figt = go.Figure()
        for yr, color in ((py, "slate"), (cy, "blue")):
            g = tsub[tsub["year"] == yr]
            mp = {str(r["label"]): r["value"] for _, r in g.iterrows()}
            figt.add_trace(go.Scatter(
                x=xs, y=[(mp[k] / xdiv) if (k in mp and mp[k] == mp[k]) else None for k in xs],
                mode="lines+markers", name=f"{yr}년", connectgaps=False,
                line=dict(color=clr(color), width=2 if yr == cy else 1.5,
                          dash=None if yr == cy else "dot"),
                marker=dict(size=5)))
        figt.update_layout(**base_layout(
            340, ysuffix="%" if is_pct else "",
            title=f"{met} · {cy}년 vs {py}년 ({'%' if is_pct else unit})"))
        st.plotly_chart(figt, width="stretch")

    # ── ④ 한눈에 — 지금 단계 × 그 아래 단계 ──
    gd = len(path) + 1
    if kid_lbl and gd < len(LV) and len(kid):
        gcol, glbl = LV[gd][0], LV[gd][1]
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"④ 한눈에 — {kid_lbl} × {glbl}")
        # 브랜드·상품까지 열리면 교차가 수백 × 수천이 된다. 다 그리면 눈금도 글자도
        # 뭉개져 아무것도 못 읽으니 **위쪽 몇 개씩만** 세운다.
        HMAX = 12
        cells = []
        for _, r in show.head(HMAX).iterrows():
            gl, _gh = view.live(r["_p"])
            for g_ in (view.children(r["_p"]) if show_all else gl):
                gp = tuple(r["_p"]) + (g_,)
                cells.append({"r": r["_n"], "c": g_,
                              "vc": V(gp, cy), "vp": V(gp, py)})
        cdf = pd.DataFrame(cells)
        if cdf.empty or not cdf["vc"].notna().any():
            st.info("이 기간엔 교차 값이 없어요.")
        else:
            if additive and np.isfinite(cur_p) and cur_p:
                cdf["z"] = (cdf["vc"] - cdf["vp"]) / cur_p * 100
                zttl, zsuf = f"기여도(%p) — {crumb} 전년 대비", "%p"
            elif is_pct:
                cdf["z"] = (cdf["vc"] - cdf["vp"]) * 100
                zttl, zsuf = "전년비(%p)", "%p"
            else:
                cdf["z"] = (cdf["vc"] / cdf["vp"] - 1) * 100
                zttl, zsuf = "전년비(%)", "%"
            z = cdf.pivot_table(index="r", columns="c", values="z", aggfunc="last")
            z = z.reindex([n for n in show["_n"] if n in z.index])
            if z.shape[1] > HMAX:      # 열도 같은 이유로 잘라 준다 (움직임이 큰 순)
                keep = z.abs().max(axis=0).sort_values(ascending=False).index[:HMAX]
                z = z[[c for c in z.columns if c in set(keep)]]
            lim = float(np.nanmax(np.abs(z.values))) if np.isfinite(z.values).any() else 1.0
            figh = go.Figure(go.Heatmap(
                z=z.values, x=[str(c) for c in z.columns], y=[str(i) for i in z.index],
                zmid=0, zmin=-lim, zmax=lim,
                colorscale=[[0, clr("red")], [0.5, "#ffffff"], [1, clr("green")]],
                text=[[("" if not np.isfinite(v) else f"{v:+.1f}") for v in row]
                      for row in z.values],
                texttemplate="%{text}", textfont=dict(size=10),
                hovertemplate="%{y} · %{x}<br>" + zttl + ": %{z:+.2f}" + zsuf
                              + "<extra></extra>",
                colorbar=dict(title=dict(text=zsuf, side="right"), thickness=10)))
            lay = base_layout(max(260, 30 * len(z.index) + 120), title=f"{ccol} · {zttl}")
            lay["showlegend"] = False
            lay["yaxis"]["autorange"] = "reversed"
            figh.update_layout(**lay)
            st.plotly_chart(figh, width="stretch")
            st.caption(("빨간 칸이 상위를 끌어내린 곳이에요. " if additive else "")
                       + f"위쪽 {HMAX}개씩만 그려요. "
                       + f"칸을 짚었으면 위 **{gd}. {kid_lbl}** 에서 그 이름을 골라 "
                         "한 단계 더 들어가세요.")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    if detail:
        _dmin, _dmax = str(ddf["date"].min())[:10], str(ddf["date"].max())[:10]
        st.caption("**브랜드·상품 결제 원장**이에요 (조직 > 카테고리 > 브랜드 > 상품). "
                   f"적재 범위는 {esc(_dmin)} ~ {esc(_dmax)}, {len(ddf):,}행이에요. "
                   "원장은 그날의 합계로 오지만 화면 값은 **일평균**이에요 — 합계로 보면 "
                   "며칠치인 이번 달이 전년 한 달과 맞붙어 가짜 급락이 떠요. "
                   "**객단가 = 거래액 ÷ 주문고객수**라 반품으로 고객수가 0 이하가 되면 "
                   "지어내지 않고 비워 둬요.")
    else:
        st.caption("MICRO 대시보드의 **구분06~09**(조직 > 카테고리 > 브랜드 > 상품) export예요. "
                   "거래액·고객수·객단가는 **일평균**이고 비중·상품CR은 비율이에요. "
                   "`*TOTAL`은 하위 항목의 단순 합이 아니라 파일이 준 값을 그대로 써요 — "
                   "거래액만 합이 상위와 일치하고, 고객수·상품UV는 유니크 값이라 조금 넘어요.")


def print_button(label="이 페이지 PDF 저장 / 인쇄"):
    # st.components.v1.html은 제거 예정(2026-06-01 이후) — 후속 API인 st.iframe 사용
    st.iframe(
        f"""<button onclick="window.parent.print()"
        style="float:right;background:#2E68B0;color:#fff;border:0;border-radius:6px;
        padding:7px 14px;font-size:13px;font-weight:600;cursor:pointer;
        font-family:'Pretendard',-apple-system,sans-serif">{label}</button>
        <div style="clear:both"></div>""", height=44)
    st.caption("버튼이 안 눌리면 Ctrl+P(Mac ⌘+P)를 누르고 대상을 'PDF로 저장'으로 바꿔 주세요.")

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
            help="주간 폴더를 zip으로 묶어 통째로 올려도 돼요. "
                 "전체관점 마스터(일·주·월)와 지표별 파일(가입율·가입자수·당일가입 첫구매율·비회원 트래픽), "
                 "MICRO 대시보드의 조직×카테고리(구분06×구분07) export, "
                 "브랜드·상품 결제 원장을 자동으로 인식해요.")
        st.markdown("---")
        # 목록·순서·이름은 전부 모듈 위쪽 `PAGE_ORDER` 한곳에서 나온다.
        page = st.radio("페이지", PAGES, key="wr_page")

    stored = load_store()
    expanded = expand_uploads(files) if files else []
    df_new = combine_files(tuple(expanded)) if expanded else pd.DataFrame()
    # 조직×카테고리는 축이 둘이라 별도 store에 쌓는다 (같은 업로드에서 같이 갈라 담긴다)
    oc_stored = load_orgcat_store()
    oc_new = combine_orgcat(tuple(expanded)) if expanded else pd.DataFrame(columns=ORGCAT_COLS)
    odf = merge_orgcat(oc_stored, oc_new)
    # **일별만 올려도 주·월이 나오게** 파생을 얹는다. 같은 실적인데 세 파일을 따로 받아
    # 올려야 했던 자리다. 파일이 준 값이 늘 이기고 파생은 빈자리만 메운다 — 주·월 파일을
    # 올리던 방식도 그대로 산다. 만든 기간은 `attrs`에 담겨 오니 **여기서 바로** 읽는다
    # (merge·concat에서 사라진다 — `af_rejected`와 같은 규칙).
    # **저장·백업은 파생 «전» 프레임(`odf_raw`)으로 한다.** 파생 행까지 저장하면 다음
    # 세션엔 그게 '파일이 준 값'으로 보여서, 일별을 고쳐 다시 올려도 옛 파생이 그 기간을
    # 차지한 채 안 갱신된다(파생은 파일 있는 기간을 비켜 가므로). 화면만 파생을 본다.
    odf_raw = odf
    odf = orgcat_derive_periods(odf_raw)
    _oc_made = odf.attrs.get("orgcat_derived", set()) if odf is not None else set()
    # 브랜드·상품 원장도 마찬가지로 별도 store — 카테고리 어휘가 MICRO와 달라 섞을 수 없다
    dt_stored = load_detail_store()
    dt_new = combine_detail(tuple(expanded)) if expanded else pd.DataFrame(columns=DETAIL_COLS)
    ddf = merge_detail(dt_stored, dt_new)

    has_any = (not stored.empty or not df_new.empty or not odf.empty or not ddf.empty)
    if files and df_new.empty and stored.empty:
        st.error("올린 파일에서 데이터를 읽지 못했어요. 파일명 형식을 확인해 주세요.")
        st.stop()
    if not has_any:
        st.markdown("## 📋 주간보고 통합 — 시작하기")
        st.caption("누적 데이터가 없어요. 원천 파일을 올리거나 예전 백업(ZIP)을 복원해 주세요.")
        cU, cR = st.columns(2)
        with cU:
            st.markdown("#### ① 원천 파일 업로드")
            source_upload_widget("wr_up_main")
        with cR:
            st.markdown("#### ② 또는 백업 복원")
            restore_widget("wr_restore_empty", label="백업 ZIP / CSV / JSON 올리기")
            st.caption("예전에 받은 `주간보고백업_*.zip`을 그대로 올리면 데이터와 메모가 통째로 복원돼요.")
        st.markdown("""
---
**인식되는 원천 파일**
- **마스터**: `전체관점 - 일자별/주별/월별 실적 (기본)`
- **지표별**: `월_가입율(일평균)`, `주_가입자수(일평균)`, `일_비회원 트래픽(일평균)`, `월_당일가입 첫구매율 (일평균)` …
- **조직×카테고리**: MICRO 대시보드 export (헤더에 `구분06`·`구분07`) — 일·주·월 각각
- **브랜드·상품 원장**: 헤더에 `결제_일자`·`BPU`·`대카테고리명`·`ADMIN브랜드명`·`상품명`·`거래액` — 브랜드·상품까지 파고들 수 있어요
- **앱푸시**: 파일명에 `PUSH`/`앱푸시`/`수신동의` 포함 또는 헤더가 앱푸시 형식이면 자동 인식
- **앱설치**: 앱 대시보드 export — **일별 파일**만 쌓아요 (`날짜,전체 설치,신규 설치,…`).
  주간·월간 파일은 인식은 하지만 저장하지 않아요 (주 경계가 달라요)
- 주간 폴더를 **zip으로 묶어 통째로** 올려도 돼요.
""")
        st.stop()

    # 누적 저장소와 병합 — 미리보기(저장 전까지 영구 반영 안 함)
    df = merge_store(stored, df_new)

    has_new = not df_new.empty or not oc_new.empty or not dt_new.empty
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
                    st.caption("인식 못 한 파일은 파일명(전체관점·월_가입율 등)과 형식을 확인해 주세요. "
                               "인식된 파일만 저장에 반영돼요.")
        if has_new:
            added, updated = upload_diff(stored, df_new)
            saved = st.session_state.get("wr_saved_sig") == sig
            if saved:
                st.success("저장됨 ✓ (누적 반영 완료)")
            else:
                _oc_note = (f"\n\n조직×카테고리 {len(oc_new):,}행도 함께 반영돼요."
                            if not oc_new.empty else "")
                _oc_note += (f"\n\n브랜드·상품 원장 {len(dt_new):,}행도 함께 반영돼요."
                             if not dt_new.empty else "")
                st.warning(f"새 데이터 감지 — 추가 {added}기간 · 갱신(겹침) {updated}기간"
                           + _oc_note + "\n\n**저장**을 눌러야 누적에 반영돼요.")
                if st.button("💾 저장 (누적 반영)", key="wr_commit",
                             type="primary", width="stretch"):
                    if not df_new.empty: save_store(df)
                    if not oc_new.empty: save_orgcat_store(odf_raw)
                    if not dt_new.empty: save_detail_store(ddf)
                    st.session_state["wr_saved_sig"] = sig
                    st.rerun()

    if df.empty:
        # 사이드바 기준 기간·차트 연도가 전부 마스터에서 나오므로 나머지 화면은 열 수 없다.
        # 조직·카테고리는 **자체 기간 선택**을 쓰니 그것만 올린 상태에서도 보여 준다.
        # 번호가 아니라 **이름**으로 가른다 — 예전엔 `page.startswith("09.")`였는데
        # 페이지를 재정렬하면서 조용히 안 열리게 됐다(증상이 '빈 화면'이라 안 드러난다).
        if page_name(page) == PAGE_ORGCAT_CH and not odf.empty:
            render_orgcat_channel_page(odf)
            st.stop()
        if page_name(page) == PAGE_ORGCAT and (not odf.empty or not ddf.empty):
            render_orgcat_page(odf, ddf)
            st.stop()
        _o8 = ", ".join(x for x in (f"조직×카테고리 {len(odf):,}행" if not odf.empty else "",
                                    f"브랜드·상품 원장 {len(ddf):,}행" if not ddf.empty else "")
                        if x)
        st.warning("첫구매(전체관점·지표별) 데이터가 없어요. 원천 파일을 올려 주세요."
                   + (f" ({_o8}은 저장돼 있어요 — "
                      f"「{esc(PAGE_ORGCAT)}」에서 볼 수 있어요)" if _o8 else ""))
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
        # 달은 라벨을 다시 파싱하지 말고 `sortkey`(=year*10000+월*100)에서 읽는다.
        # `re.match(...).group()`을 그대로 부르면 매치가 없는 라벨 하나가 사이드바를
        # 죽여 전 페이지가 같이 내려간다 — 그 사고를 두 번 냈다.
        _mrows = df[(df["gran"] == "월") & (df["year"] == ref_year)]["sortkey"].dropna()
        months_avail = sorted({int(k) // 100 % 100 for k in _mrows})
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
                                    key="wr_refw", format_func=month_trim,
                                    help="주간보고 대상 주차예요. 최신 주차가 진행 중이면 직전 완료 주차가 기본으로 잡혀요.")
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
                   else "⚠ ANTHROPIC_API_KEY가 없어요. Secrets에 추가해 주세요")

        st.markdown("---")
        st.markdown("**백업 · 복원**")
        saved_rows = len(load_store())
        pend = " · 저장 시 반영" if (has_new and not st.session_state.get("wr_saved_sig") == sig) else ""
        st.caption(f"누적 {saved_rows:,}행 저장됨 / 현재 보기 {len(df):,}행{pend} · "
                   f"조직×카테고리 {len(odf):,}행 · "
                   + (f"원장 {len(ddf):,}행 · " if not ddf.empty else "")
                   + f"메모 {len(st.session_state.wr_texts)}개")
        # **파생한 기간은 밝힌다.** 일별만 올려도 주·월이 뜨는데, 그게 파일에서 온 값인지
        # 만든 값인지 화면에서 갈리지 않으면 '주별 파일을 올렸던가?'를 되짚을 수가 없다.
        if _oc_made:
            _mw = sum(1 for g, _y, _l in _oc_made if g == "주")
            _mm = len(_oc_made) - _mw
            st.caption(f"↳ 조직×카테고리는 **일별에서 주 {_mw}개 · 월 {_mm}개**를 만들었어요. "
                       "파일로 올린 기간은 그대로 쓰고 빈자리만 채워요.")

        # 백업 파일은 **누를 때 만든다.** `st.download_button`은 data를 미리 받아야 해서,
        # 여기에 직렬화 결과를 그대로 넘기면 **매 리런마다** 원장 전체를 CSV로 찍고
        # 그 바이트를 브라우저로 보낸다. 실측 120만 행 원장에서 통합 ZIP 5.6초 +
        # 개별 CSV 3.4초(122MB)라, 백업을 안 받는 리런까지 9초씩 손해 보고 있었다.
        # 페이지·필터를 만질 때마다 내는 비용이라 08번만이 아니라 화면 전체가 느렸다.
        _bksig = (len(df), len(odf), len(ddf), len(st.session_state.wr_texts), sig)
        _bk = lazy_download("wr_bk_all", "⬇ 통합 백업 (ZIP · 데이터+메모)",
                            f"주간보고백업_{today_kst():%Y%m%d}.zip", "application/zip",
                            sig=_bksig, type="primary",
                            help="누적 데이터 CSV와 보고란·메모 JSON을 한 파일로 백업해요. "
                                 "재배포로 초기화돼도 이 ZIP을 '백업 복원'에 올리면 그대로 "
                                 "되살아나요. 파일은 누른 뒤에 만들어요.")
        if _bk:
            _bk(lambda: make_backup_zip(df, st.session_state.wr_texts,
                                        odf_raw, ddf))

        # 통합 복원 — zip/csv/json 자동 인식
        restore_widget("wr_restore")

        with st.expander("개별 백업 (데이터만 / 메모만)"):
            _d = today_kst()
            for _key, _lab, _fn, _mime, _make in (
                ("data", "⬇ 누적 데이터 (CSV)", f"wr_data_store_{_d:%Y%m%d}.csv", "text/csv",
                 lambda: df[STORE_COLS].to_csv(index=False).encode("utf-8-sig")),
                # 백업도 **파생 전**(`odf_raw`)이다 — 파생을 담아 두면 복원한 쪽에서
                # 그게 파일 값이 되어, 일별을 고쳐도 그 기간이 안 갱신된다.
                ("orgcat", "⬇ 조직×카테고리 (CSV)", f"wr_orgcat_store_{_d:%Y%m%d}.csv", "text/csv",
                 (lambda: odf_raw[ORGCAT_COLS].to_csv(index=False).encode("utf-8-sig"))
                 if not odf_raw.empty else None),
                # 원장은 gzip으로 준다 — 평문 CSV는 실측 122MB, gzip은 14MB다.
                # 복원(`_try_csv`)이 매직바이트로 gzip을 알아보므로 그대로 다시 올리면 된다.
                ("detail", "⬇ 브랜드·상품 원장 (CSV.GZ)",
                 f"wr_detail_store_{_d:%Y%m%d}.csv.gz", "application/gzip",
                 (lambda: gzip.compress(
                     ddf[DETAIL_COLS].to_csv(index=False).encode("utf-8-sig")))
                 if not ddf.empty else None),
                ("memo", "⬇ 보고란·메모 (JSON)", f"wr_insights_{_d:%Y%m%d}.json", "application/json",
                 lambda: json.dumps(st.session_state.wr_texts, ensure_ascii=False,
                                    indent=2).encode("utf-8")),
            ):
                if _make is None:
                    continue
                _go = lazy_download(f"wr_bk_{_key}", _lab, _fn, _mime, sig=_bksig)
                if _go:
                    _go(_make)

        # 초기화 — 2단계 확인 (실수 방지)
        if st.session_state.get("wr_confirm_clear"):
            st.warning("정말 초기화할까요? 누적 데이터·조직×카테고리·브랜드 상품 원장이 모두 지워지고 "
                       "되돌릴 수 없어요. 메모는 남아요. 초기화 전에 통합 백업을 받아두세요.")
            cc1, cc2 = st.columns(2)
            if cc1.button("삭제 확인", key="wr_clear_yes", type="primary", width="stretch"):
                if os.path.exists(DATA_STORE): os.remove(DATA_STORE)
                if os.path.exists(ORGCAT_STORE): os.remove(ORGCAT_STORE)
                if os.path.exists(DETAIL_STORE): os.remove(DETAIL_STORE)
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

    # 번호는 화면에만 붙는다 — 분기는 이름으로 가린다(`PAGE_ORDER` 옆 주석 참고).
    _pg = page_name(page)

    # ════════════ 주간보고 요약 ════════════
    if _pg == PAGE_SUMMARY:
        st.markdown(f"## 첫구매 주간보고 — {ref_year}년 {ref_month}월")
        wy, wlabel = week_ref(df, ref_year, ref_week)
        if wlabel:
            st.caption(f"기준 주차: {week_disp(wy, wlabel)}")

        # KPI 카드 (기준 주차 — 전주비·전월비·전년비 모두 주차 기준)
        if wlabel:
            cols = st.columns(4)
            kpi_metrics = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가", "가입자수"]
            py, plb = prev_label(df, "주", wy, wlabel)
            # 전월 동일 주차 (예: 06월 1주차 → 05월 1주차).
            # 5주차처럼 전월(또는 전년 동월)에 같은 주차가 없으면 그 달 마지막 주차로
            # 대체하고, 어떤 주와 비교했는지 뱃지에 같이 적는다.
            mom_y = mom_lbl = None
            yoy_lbl, mom_exact, yoy_exact = wlabel, True, True
            wm = re.match(r"(\d{1,2})월 (\d)주차", wlabel)
            if wm:
                mo, wk = int(wm.group(1)), int(wm.group(2))
                mom_y, mom_m = (wy, mo - 1) if mo > 1 else (wy - 1, 12)
                mom_lbl, mom_exact = week_like(df, mom_y, mom_m, wk)
                yoy_lbl, yoy_exact = week_like(df, wy - 1, mo, wk)
            mom_tag = "전월비" if (mom_exact or not mom_lbl) else f"전월비 · {mom_lbl}"
            yoy_tag = "전년비" if (yoy_exact or not yoy_lbl) else f"전년비 · {yoy_lbl}"
            for col, met in zip(cols, kpi_metrics):
                cur = pick(df, "주", met, "*TOTAL", wy, wlabel, "mtd")
                deltas = []
                prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
                deltas.append((fmt_delta(met, cur, prv), "전주비"))
                if wm:
                    mom = (pick(df, "주", met, "*TOTAL", mom_y, mom_lbl, "final")
                           if mom_lbl else np.nan)
                    deltas.append((fmt_delta(met, cur, mom), mom_tag))
                yoy = (pick(df, "주", met, "*TOTAL", wy - 1, yoy_lbl, "final")
                       if yoy_lbl else np.nan)
                deltas.append((fmt_delta(met, cur, yoy), yoy_tag))
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
            _sub = []
            if wm and not mom_exact and mom_lbl:
                _sub.append(f"전월에 {wm.group(2)}주차가 없어 **{mom_y}년 {mom_lbl}**(전월 마지막 주)와 비교")
            if wm and not yoy_exact and yoy_lbl:
                _sub.append(f"전년 동월에 {wm.group(2)}주차가 없어 **{wy-1}년 {yoy_lbl}**와 비교")
            if _sub:
                st.caption("ℹ️ " + " · ".join(_sub) + " — 주차 수가 달마다 4~5개로 달라서예요.")

            # 전전년까지 펼쳐보기 — 올해 흐름이 작년에도 있었는지(= 구조적인지) 확인용
            with st.expander(f"🔍 전년비 · 전전년비 비교 — {wy-2}년에도 같은 추세였는지",
                             expanded=False):
                t2, _l1, _l2, _e1, _e2 = yoy2_summary_table(df, wy, wlabel, METRICS7)
                _basis = []
                if not _e1 and _l1: _basis.append(f"전년 {wy-1}년 {_l1}")
                if not _e2 and _l2: _basis.append(f"전전년 {wy-2}년 {_l2}")
                st.caption(
                    f"기준: {week_disp(wy, wlabel)}"
                    + (f" · 대응 주차가 없어 {' · '.join(_basis)} 기준으로 비교" if _basis else "")
                    + " — **전년의 전년비**는 전년 동주가 그 전해 대비 어땠는지예요. "
                      "**추세**가 '2년 연속'이면 올해만의 현상이 아니라 이어져 온 흐름이에요.")
                wtable(style_delta_cols(t2), width="stretch")
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 실적 요약 — 전주비 (주차 기준)
        if wlabel:
            st.subheader("실적 요약 (일평균 · 전주비)")
            st.caption(f"기준 주차: {week_disp(wy, wlabel)} — 전주·전년 동주 대비")
            wtable(style_delta_cols(wow_summary_table(df, wy, wlabel, METRICS7)),
                         width="stretch", dl_name="실적 요약 (일평균 · 전주비)")
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 실적 요약 YoY 표
        st.subheader("실적 요약 (일평균 · 전년비)")
        tbl, (pm_y, pm_m) = yoy_summary_table(df, ref_year, ref_month, METRICS7)
        st.caption(f"전월({pm_m}월)은 월마감, 당월({ref_month}월)은 일마감(MTD) 기준 동일기간 비교. "
                   f"맨 끝 전월비(당월)은 {ref_year}년 {ref_month}월 ↔ {pm_y}년 {pm_m}월이에요 — "
                   "값이 일평균이라 기간 길이가 달라도 맞댈 수 있어요.")
        wtable(style_delta_cols(tbl), width="stretch", dl_name="실적 요약 (일평균 · 전년비)")

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
        for _mrow in (TREND_CHARTS[:3], TREND_CHARTS[3:]):
            for _col, _met in zip(st.columns(3), _mrow):
                with _col:
                    st.plotly_chart(yoy_chart(df, "주", _met, chart_years, h=280),
                                    width="stretch")

    # ════════════ 첫구매 퍼널별 상세 실적 ════════════
    elif _pg == PAGE_FUNNEL:
        _fwy, _fwlabel = week_ref(df, ref_year, ref_week)
        render_funnel_page(df, odf, ref_year, ref_month, _fwy, _fwlabel)

    # ════════════ 조직·카테고리·채널별 첫구매 상세 ════════════
    elif _pg == PAGE_ORGCAT_CH:
        render_orgcat_channel_page(odf)

    # ════════════ 채널별 실적 ════════════
    elif _pg == PAGE_CHANNEL:
        _cwy, _cwlabel = week_ref(df, ref_year, ref_week)
        render_channel_page(df, ref_year, ref_month, _cwy, _cwlabel, ch_sel)

    # ════════════ 앱푸시 동의 현황 ════════════
    elif _pg == PAGE_PUSH:
        render_push_page(df, ref_year, chart_years)

    # ════════════ 첫구매 고객 세그먼트 성과 ════════════
    # 예전엔 주·월 표 두 개가 박혀 있고 세그먼트를 하나씩 갈아 끼워야 했다 — 일 단위는
    # 아예 못 봤고, 세그먼트 사이의 이야기가 안 이어졌다. 「채널별 실적」과 **같은 함수**를
    # 쓴다(`render_axis_page`) — 축만 채널→세그먼트로 바뀐다. 복사해 두면 한쪽만 고쳐져
    # 조용히 갈린다.
    elif _pg == PAGE_SEGMENT:
        _swy, _swlabel = week_ref(df, ref_year, ref_week)
        _segs = [x for x in SEGMENTS if (df["segment"] == x).any()]
        if not _segs:
            st.markdown("## 첫구매 고객 세그먼트 성과")
            st.info("세그먼트 데이터가 없어요. 세그먼트별 실적 파일을 올려 주세요.")
            st.caption("세그먼트는 " + " · ".join(SEGMENTS) + "예요.")
        else:
            # 세그먼트 원천은 지표 이름이 마스터와 달라(거래액·고객수·객단가·CR·유입율)
            # `METRICS7`로 고정하면 표가 통째로 빈다 — **데이터가 준 지표**를 쓴다.
            _smets = (df[df["segment"].isin(_segs) & df["value"].notna()]
                      ["metric"].dropna().astype(str).unique().tolist())
            render_axis_page(df, ref_year, ref_month, _swy, _swlabel,
                             title="첫구매 고객 세그먼트 성과", axis="세그먼트",
                             kp="wr_sg", opts=_segs, pal=SEG_PAL,
                             dft=_segs, mets=_smets)
            with st.expander("📊 지표 산출식 및 용어 설명", expanded=False):
                st.markdown("""
**산출식**
* **거래액비중**: 첫구매 거래액 ÷ 전체 거래액
* **고객비중**: 첫구매 고객수 ÷ 전체 고객수
* **첫구매 객단가**: 첫구매 거래액 ÷ 첫구매 고객수
* **유입율**: DAU ÷ 유효회원수
* **CR (전환율)**: 첫구매 고객수 ÷ DAU

**용어 설명**
* **유효회원수**: 서비스에 정상적으로 가입해서 활동할 수 있는 전체 회원 수예요.
* **DAU (Daily Active Users)**: 하루 동안 서비스에 한 번 이상 방문해서 활동한 사용자 수예요.
""")

    # ════════════ 조직·카테고리별 실적 ════════════
    elif _pg == PAGE_ORGCAT:
        render_orgcat_page(odf, ddf)

    # ════════════ 통합 데이터·다운로드 ════════════
    elif _pg == PAGE_DOWNLOAD:
        st.markdown("## 통합 데이터 · 다운로드")
        # 원천이 셋으로 늘었다(마스터 · 조직×카테고리 · 결제 원장). 예전엔 마스터만
        # 보여 줘서 '원장을 올렸는데 어디 갔지'를 여기서 확인할 수가 없었다.
        st.caption("올린 파일을 모두 합친 누적 데이터예요. 원천마다 따로 쌓여요.")
        _stores = [("마스터 (지표 × 채널)", df, ["gran", "metric", "segment", "sortkey"]),
                   ("조직×카테고리 (MICRO)", odf, ["gran", "metric"] + ORGCAT_LV),
                   ("브랜드·상품 결제 원장", ddf, None)]
        _cards = []
        for _nm, _d, _ in _stores:
            if _d is None or _d.empty:
                _cards.append(f'<div class="kpi-card" style="flex:1"><div class="kpi-label">'
                              f'{esc(_nm)}</div><div class="kpi-value">–</div>'
                              f'<div class="kpi-delta na">아직 없어요</div></div>')
                continue
            _yr = (f"{int(_d['year'].min())}–{int(_d['year'].max())}년"
                   if "year" in _d.columns and _d["year"].notna().any() else "")
            _cards.append(f'<div class="kpi-card" style="flex:1"><div class="kpi-label">'
                          f'{esc(_nm)}</div><div class="kpi-value">{len(_d):,}행</div>'
                          f'<div class="kpi-delta">{esc(_yr)}</div></div>')
        st.markdown('<div style="display:flex;gap:10px;align-items:stretch">'
                    + "".join(_cards) + '</div>', unsafe_allow_html=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("미리보기")
        _pv = [n for n, d, _ in _stores if d is not None and not d.empty]
        if not _pv:
            st.info("아직 쌓인 데이터가 없어요. 사이드바에서 원천 파일을 올려 주세요.")
        else:
            guard_select("wr_dl_src", _pv)
            _pick = st.selectbox("원천", _pv, key="wr_dl_src")
            _d, _sortby = next((d, s) for n, d, s in _stores if n == _pick)
            _show = _d.sort_values([c for c in (_sortby or []) if c in _d.columns]) \
                if _sortby else _d
            st.caption(f"{len(_d):,}행 중 앞 2,000행이에요. 전체는 아래에서 받으세요.")
            wtable(_show.head(2000), width="stretch", height=420,
                   dl_name=f"{_pick} 미리보기")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("내려받기")
        # **누른 뒤에 만든다.** `st.download_button`은 data를 미리 받는 API라 그냥 넘기면
        # 받지도 않는 리런마다 전체를 CSV로 찍는다(원장은 120만 행이라 그것만 몇 초다).
        d1, d2 = st.columns(2)
        with d1:
            _sig = (len(df), len(odf) if odf is not None else 0,
                    len(ddf) if ddf is not None else 0)
            _go = lazy_download("wr_dl_master", "⬇ 마스터 long 데이터 (CSV)",
                                f"통합데이터_{today_kst():%Y%m%d}.csv", "text/csv",
                                sig=_sig, help="파일은 누른 뒤에 만들어요.")
            if _go:
                _go(lambda: df.to_csv(index=False).encode("utf-8-sig"))
        with d2:
            _gw = lazy_download("wr_dl_wb", "⬇ 엑셀 워크북 (첫구매_주간보고)",
                                f"첫구매_주간보고_{ref_year}{ref_month:02d}.xlsx",
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet",
                                sig=(_sig, ref_year, ref_month, tuple(chart_years)),
                                help="`첫구매_요약`(요약표·보고란·YoY 차트·채널표) + "
                                     "`월`·`주` 시트예요.")
            if _gw:
                _gw(lambda: build_workbook(df, st.session_state.wr_texts,
                                           ref_year, ref_month, chart_years))
        st.caption("원천별 개별 백업과 통합 ZIP은 **사이드바 「백업」**에 있어요 — "
                   "재배포로 초기화돼도 그 ZIP으로 그대로 되살릴 수 있어요.")

if st.runtime.exists():
    main()
