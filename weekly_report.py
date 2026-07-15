"""二쇨컙蹂닿퀬 ??泥リ뎄留??듯빀 ?ㅼ쟻 ??쒕낫??
?щ윭 ?먯쿇 ?묒?(?꾩껜愿??留덉뒪??+ 吏?쒕퀎 ?뚯씪)???낅줈?쒗븯硫??섎굹???듯빀 酉곕줈 醫낇빀?섍퀬,
?꾨뀈(YoY)쨌?꾩＜(WoW) 利앷컧?꾪솴怨?蹂닿퀬???媛뽰텣 二쇨컙蹂닿퀬 ?붾㈃??留뚮뱺??
?듯빀 寃곌낵??(??二?泥リ뎄留??붿빟 + 李⑦듃) ?묒? ?뚰겕遺곸쑝濡??ㅼ슫濡쒕뱶?????덈떎.
"""

import datetime
import io, json, os, re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

try:
    from streamlit_quill import st_quill
    HAS_QUILL = True
except Exception:
    HAS_QUILL = False

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# CONFIG
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
st.set_page_config(page_title="二쇨컙蹂닿퀬 ??泥リ뎄留??듯빀 ?ㅼ쟻", page_icon="?뱥",
                   layout="wide", initial_sidebar_state="expanded")

INSIGHT_FILE = "wr_insights.json"

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

  /* 寃뱀묠 諛⑹?: ?덉씠?꾩썐 釉붾줉???뺤쟻 諛곗튂?섍퀬 ?섏묠??洹몃?濡??몄텧 */
  [data-testid="stVerticalBlock"], [data-testid="stHorizontalBlock"],
  [data-testid="stVerticalBlockBorderWrapper"], [data-testid="column"],
  .element-container { position:static !important; transform:none !important;
    overflow:visible !important; }

  /* ?몄뇙 ??Plotly 李⑦듃 ?믪씠 遺뺢눼 ???꾨옒 ?붿냼媛 ?쒕ぉ ?꾨줈 諛??寃뱀튂??臾몄젣 李⑤떒 */
  .stPlotlyChart, .js-plotly-plot, [data-testid="stPlotlyChart"] {
    min-height:240px !important; break-inside:avoid; page-break-inside:avoid; }

  /* ?쒕ぉ???щ챸 諛곌꼍?쇰줈 ?ㅻⅨ ?붿냼? 寃뱀퀜 蹂댁씠吏 ?딅룄濡?*/
  h1, h2, h3, h4 { background:#fff !important; position:relative; z-index:1;
    page-break-after:avoid; break-after:avoid; }

  .stPlotlyChart, .report-box, table,
  [data-testid="stMetric"], [data-testid="column"] {
    break-inside:avoid; page-break-inside:avoid; }

  /* 利앷컧 ?됱긽(鍮④컯/珥덈줉) ?몄뇙???좎? */
  * { -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }
}
</style>
""", unsafe_allow_html=True)

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# ?됱긽 ?붾젅??(湲곗〈 first_purchase ? ?숈씪 怨꾩뿴)
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
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
    "吏곸젒": "blue", "愿묎퀬": "amber", "EP": "green", "PUSH": "purple",
    "?쒗쑕": "red", "釉뚮옖?쒓킅怨?: "teal", "誘몃뵒?댁빱癒몄뒪": "orange", "*TOTAL": "slate",
}
CHANNELS = ["吏곸젒", "愿묎퀬", "EP", "PUSH", "?쒗쑕", "釉뚮옖?쒓킅怨?, "誘몃뵒?댁빱癒몄뒪"]
YEAR_PAL = ["slate", "blue", "red", "green", "purple", "amber", "teal"]

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# 吏???뺤쓽
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
METRICS7 = ["泥リ뎄留?嫄곕옒??, "泥リ뎄留?怨좉컼??, "泥リ뎄留?媛앸떒媛",
            "鍮꾪쉶?먰듃?섑뵿", "媛?낆옄??, "媛?낆쑉", "?뱀씪媛?꿂R"]
PCT_METRICS = {"媛?낆쑉", "?뱀씪媛?꿂R", "?좎엯??, "CR", "嫄곕옒?〓퉬以?, "怨좉컼鍮꾩쨷"}
# 留덉뒪???뚯씪 吏????蹂닿퀬??吏??留ㅽ븨
MASTER_MAP = {"?쇳룊洹좉굅?섏븸": "泥リ뎄留?嫄곕옒??, "?쇳룊洹좉퀬媛앹닔": "泥リ뎄留?怨좉컼??,
              "?쇳룊洹좉컼?④?": "泥リ뎄留?媛앸떒媛"}
# 吏?쒕퀎 ?뚯씪紐???蹂닿퀬??吏??留ㅽ븨 (怨듬갚 ?쒓굅 ??留ㅼ묶)
METRIC_FILE_MAP = {
    "媛?낆쑉": "媛?낆쑉", "媛?낅쪧": "媛?낆쑉",
    "媛?낆옄??: "媛?낆옄??,
    "?뱀씪媛?낆껀援щℓ??: "?뱀씪媛?꿂R", "?뱀씪媛?꿂R": "?뱀씪媛?꿂R",
    "鍮꾪쉶?먰듃?섑뵿": "鍮꾪쉶?먰듃?섑뵿",
}

METRIC_UNIT = {
    "泥リ뎄留?嫄곕옒??: ("諛깅쭔??, 1e6), "泥リ뎄留?怨좉컼??: ("紐?, 1),
    "泥リ뎄留?媛앸떒媛": ("??, 1), "鍮꾪쉶?먰듃?섑뵿": ("紐?, 1),
    "媛?낆옄??: ("紐?, 1), "媛?낆쑉": ("%", 1), "?뱀씪媛?꿂R": ("%", 1),
    "?깊뫖?쒖닔?좊룞??: ("紐?, 1),
}

def fmt_value(metric, v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "??
    if metric in PCT_METRICS: return f"{v*100:.2f}%"
    if metric == "泥リ뎄留?嫄곕옒??: return f"{v/1e6:,.1f}諛깅쭔??
    if metric == "泥リ뎄留?媛앸떒媛": return f"{v:,.0f}??
    return f"{int(v):,}紐?  # 紐??⑥쐞(怨좉컼???????뚯닔??踰꾨┝

def fmt_delta(metric, cur, prev):
    """?꾨뀈鍮??꾩＜鍮?臾몄옄?? 鍮꾩쑉 吏?쒕뒗 %p 李⑥씠, 洹???利앷컧??""
    if cur is None or prev is None: return None
    if isinstance(cur, float) and np.isnan(cur): return None
    if isinstance(prev, float) and np.isnan(prev): return None
    if metric in PCT_METRICS:
        d = (cur - prev) * 100
        return f"??abs(d):.2f}%p" if d < 0 else f"+{d:.2f}%p"
    if prev == 0: return None
    d = (cur - prev) / prev * 100
    return f"??abs(d):.1f}%" if d < 0 else f"+{d:.1f}%"

def style_delta_cols(tbl):
    """利앷컧 而щ읆(??+)??鍮④컯/珥덈줉 ?됱긽 ?곸슜??Styler 諛섑솚"""
    delta_cols = [c for c in tbl.columns
                  if any(k in str(c) for k in ("?꾨뀈鍮?, "?꾩＜鍮?, "?꾩썡鍮?, "利앷컧"))]
    def _color(v):
        s = str(v)
        if s.startswith("??): return "color:#dc2626;font-weight:600"
        if s.startswith("+"): return "color:#16a34a;font-weight:600"
        return ""
    try:
        return tbl.style.map(_color, subset=delta_cols)
    except Exception:
        return tbl

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# ?뚯떛 ???먯쿇 ?뚯씪 ???듯빀 long DataFrame
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
YEAR_RE   = re.compile(r"^(20\d{2})(\.0)?$")
PERIOD_RE = re.compile(r"^\s*(\d{1,2}\s*??\s*\d\s*二쇱감)?|\d{1,2}/\d{1,2})\s*$")

def detect_file(name):
    """?뚯씪紐낆뿉??(kind, granularity ?뚰듃, metric) 媛먯?.
    ?⑥쐞(??二???? 留덉뒪???щ???parse_file?먯꽌 ?댁슜?쇰줈 理쒖쥌 ?먮퀎?섎?濡??ш린???뚰듃留?"""
    base = os.path.basename(name)
    base = re.sub(r"\.(xlsx|xls|csv)$", "", base, flags=re.I)
    key = base.replace(" ", "")
    # ?⑥쐞 ?뚰듃: ?쇱옄蹂?二쇰퀎/?붾퀎 ?ㅼ썙???먮뒗 ??/二?/?? ?묐몢??
    gran = None
    if "?쇱옄蹂? in key or "?곗씪由? in key: gran = "??
    elif "二쇰퀎" in key or "二쇨컙" in key:   gran = "二?
    elif "?붾퀎" in key or "?붽컙" in key:   gran = "??
    else:
        m = re.match(r"^(??二???[_\s]", base)
        if m: gran = m.group(1)
    if "?꾩껜愿?? in key or "留덉뒪?? in key:
        return "master", gran, None
    # 吏???ㅼ썙?쒕뒗 ?뚯씪紐??대뒓 ?꾩튂???덉뼱???몄떇
    for k, v in METRIC_FILE_MAP.items():
        if k.replace(" ", "") in key:
            return "metric", gran, v
    # ?묐몢?ы삎(??xxx)?몃뜲 紐⑤Ⅴ??吏?쒕㈃ ?뺣━???대쫫 洹몃?濡??ъ슜
    m = re.match(r"^(??二???[_\s]+(.+)$", base)
    if m:
        rest = re.sub(r"\(.*?\)", "", m.group(2)).strip()
        return "metric", m.group(1), rest or "湲고?"
    return None, gran, None

def _decode_text(data: bytes) -> str:
    for enc in ("utf-16", "utf-8-sig", "cp949", "utf-8"):
        try: return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError): pass
    return data.decode("utf-8", "replace")

def read_grid(name, data: bytes):
    """csv(UTF-16/???ы븿)쨌xlsx 紐⑤몢 2李⑥썝 ? 洹몃━?쒕줈 ?쎈뒗??""
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
    """? 媛???float (肄ㅻ쭏쨌% 泥섎━, %??鍮꾩쑉濡?蹂??"""
    if v is None: return np.nan
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "??): return np.nan
    pct = s.endswith("%")
    if pct: s = s[:-1]
    try: f = float(s)
    except ValueError: return np.nan
    return f / 100 if pct else f

def period_parts(gran, year, plabel):
    """湲곌컙 ?쇰꺼 ??(?쒖??쇰꺼, ?뺣젹??"""
    p = plabel.replace(" ", "")
    if gran == "??:
        m = re.match(r"(\d{1,2})??, p)
        if not m: return None
        mo = int(m.group(1))
        return f"{mo}??, year * 10000 + mo * 100
    if gran == "二?:
        m = re.match(r"(\d{1,2})??\d)二쇱감", p)
        if not m: return None
        mo, wk = int(m.group(1)), int(m.group(2))
        return f"{mo:02d}??{wk}二쇱감", year * 10000 + mo * 100 + wk
    m = re.match(r"(\d{1,2})/(\d{1,2})", p)
    if not m: return None
    mo, dd = int(m.group(1)), int(m.group(2))
    return f"{mo}/{dd}", year * 10000 + mo * 100 + dd

def parse_file(name, data: bytes) -> pd.DataFrame:
    # ?꾩쟻 ?곗씠??諛깆뾽 CSV(long ?щ㎎) ?ъ뾽濡쒕뱶 ??洹몃?濡?蹂듭썝
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

    # ?ㅻ뜑 ???먯깋 (??6??: ?곕룄 / 湲곌컙?쇰꺼 / 留덇컧援щ텇
    year_row = period_row = close_row = None
    for ri in range(min(6, len(rows))):
        cells = [_cell(c) for c in rows[ri]]
        nY = sum(1 for c in cells if YEAR_RE.match(c))
        nP = sum(1 for c in cells if PERIOD_RE.match(c))
        nC = sum(1 for c in cells if "留덇컧" in c)
        if year_row is None and nY >= 1 and nP == 0: year_row = ri
        if period_row is None and nP >= 2: period_row = ri
        if close_row is None and nC >= 2: close_row = ri
    if period_row is None: return pd.DataFrame()
    data_start = max(r for r in (year_row, period_row, close_row) if r is not None) + 1

    # ?? ?댁슜 湲곕컲 理쒖쥌 ?먮퀎 (?뚯씪紐낆? ?뚰듃??肉?
    # ?⑥쐞: 湲곌컙 ?쇰꺼 ?뺥깭媛 寃곗젙 (N??N二쇱감 ??二? N/N ???? N??????
    plabels = [_cell(c) for c in rows[period_row] if PERIOD_RE.match(_cell(c))]
    if any("二쇱감" in p for p in plabels):  gran = "二?
    elif any("/" in p for p in plabels):   gran = "??
    else:                                   gran = "??
    # 留덉뒪???щ?: ?ㅻ뜑??'援щ텇01'???덉쑝硫?留덉뒪???꾩껜愿?? 援ъ“
    if any("援щ텇01" in _cell(c) for ri in range(data_start) for c in rows[ri]):
        kind = "master"
    if kind is None or (kind == "metric" and not file_metric):
        return pd.DataFrame()

    ncols = max(len(r) for r in rows)
    def cell(ri, ci):
        return _cell(rows[ri][ci]) if ri is not None and ci < len(rows[ri]) else ""

    # 而щ읆蹂??곕룄(醫뚯륫 ffill)쨌湲곌컙쨌留덇컧
    col_year, cur_y = {}, None
    for ci in range(ncols):
        m = YEAR_RE.match(cell(year_row, ci)) if year_row is not None else None
        if m: cur_y = int(m.group(1))
        col_year[ci] = cur_y
    # 湲곌컙 ?쇰꺼??鍮??(蹂묓빀)???쇰쭏媛?MTD 而щ읆? 吏곸쟾 ?쇰꺼???댁뼱諛쏅뒗??
    data_cols, col_label, last_lbl = [], {}, None
    for ci in range(ncols):
        lbl = cell(period_row, ci)
        if PERIOD_RE.match(lbl) and col_year[ci]:
            last_lbl = lbl
            col_label[ci] = lbl
            data_cols.append(ci)
        elif (close_row is not None and "留덇컧" in cell(close_row, ci)
              and last_lbl and col_year[ci]):
            col_label[ci] = last_lbl
            data_cols.append(ci)

    seg_col = 1 if kind == "master" else 0
    records, cur_metric = [], None
    for ri in range(data_start, len(rows)):
        if kind == "master":
            m0 = cell(ri, 0)
            if m0 and m0 not in ("-", "??): cur_metric = m0
            metric = MASTER_MAP.get(cur_metric, cur_metric)
        else:
            metric = file_metric
        seg = cell(ri, seg_col)
        if not seg or seg in ("-", "??) or metric is None: continue
        for ci in data_cols:
            pp = period_parts(gran, col_year[ci], col_label[ci])
            if pp is None: continue
            label, sortkey = pp
            close = cell(close_row, ci) if close_row is not None else ""
            records.append({
                "gran": gran, "metric": metric, "segment": seg,
                "year": col_year[ci], "label": label, "sortkey": sortkey,
                "close": "mtd" if "?쇰쭏媛? in close and gran != "?? else "final",
                "value": _num(rows[ri][ci] if ci < len(rows[ri]) else None),
            })
    return pd.DataFrame(records)

def _zip_entry_name(info):
    """zip ???쒓? ?뚯씪紐?蹂듭썝 (UTF-8 ?뚮옒洹??놁쑝硫?cp437?뭖p949 ?ы빐??"""
    if info.flag_bits & 0x800:
        return info.filename
    for enc in ("cp949", "utf-8", "euc-kr"):
        try: return info.filename.encode("cp437").decode(enc)
        except (UnicodeDecodeError, UnicodeEncodeError): pass
    return info.filename

def expand_uploads(uploads):
    """?낅줈??紐⑸줉 ??(?대쫫, bytes) 紐⑸줉. zip? ??댁꽌 ?대? ?묒?/CSV瑜?爰쇰궦??""
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

def parse_push_file(name, data: bytes) -> pd.DataFrame:
    try:
        rows = read_grid(name, data)
    except Exception:
        return pd.DataFrame()
    if not rows or len(rows) < 8: return pd.DataFrame()
    if not any("Date" in str(c) for c in rows[0]): return pd.DataFrame()
    
    target_row = None
    in_new_section = False
    for ri, row in enumerate(rows):
        c0 = _cell(row[0] if len(row) > 0 else "")
        c1 = _cell(row[1] if len(row) > 1 else "")
        if c0 == "?좉퇋":
            in_new_section = True
        if in_new_section and c1 == "?좉퇋異붽?(+)":
            target_row = row
            break
            
    if not target_row: return pd.DataFrame()
    
    date_row = rows[1]
    
    # 1) ?좏슚???좎쭨 而щ읆??紐⑤몢 異붿텧?섍퀬 ??month)???뚯떛
    date_cols = []
    for ci in range(2, len(date_row)):
        d = _cell(date_row[ci])
        m = re.match(r"^(\d{1,2})/\d{1,2}$", d)
        if m:
            date_cols.append((ci, d, int(m.group(1))))
            
    if not date_cols:
        return pd.DataFrame()
        
    # 2) ?ㅼ뿉?쒕?????닚?쇰줈 ?쎌쑝硫댁꽌 ??year)媛 諛붾뚮뒗 吏??怨꾩궛
    # 留?留덉?留??곗씠?곕? 0?쇰줈 ?먭퀬, (1??<- 12??濡??섏뼱媛??뚮쭏???곕룄瑜?-1
    rel_years = {}
    current_rel_year = 0
    last_month = date_cols[-1][2]
    
    for ci, d, month in reversed(date_cols):
        # ??닚?쇰줈 ?쎌쓣 ???붿씠 1?먯꽌 12濡?而ㅼ?硫?(?ㅼ젣濡쒕뒗 12??-> 1?? ?닿? 諛붾?寃?
        if last_month < 6 and month > 6:
            current_rel_year -= 1
        elif last_month > 6 and month < 6:
            current_rel_year += 1
            
        rel_years[ci] = current_rel_year
        last_month = month

    records = []
    for ci, d, month in date_cols:
        val = _num(target_row[ci] if ci < len(target_row) else None)
        if pd.isna(val): continue
        records.append({
            "gran": "??, "metric": "?깊뫖?쒖닔?좊룞??, "segment": "*TOTAL",
            "year": rel_years[ci], "label": d, "sortkey": 0, "close": "final", "value": val
        })
    return pd.DataFrame(records)

@st.cache_data(show_spinner=False)
def combine_files(file_tuples) -> pd.DataFrame:
    """?낅줈???뚯씪?????듯빀 long DF. ?숈씪 ?ㅻ뒗 留덉?留??뚯씪 ?곗꽑"""
    frames = []
    push_frames = []
    for n, b in file_tuples:
        if "PUSH" in n.upper() and n.lower().endswith((".xlsx", ".xls")):
            pf = parse_push_file(n, b)
            if not pf.empty: push_frames.append(pf)
        else:
            df = parse_file(n, b)
            if not df.empty: frames.append(df)
            
    inferred_years = set()
    for f in frames:
        if "year" in f.columns:
            inferred_years.update(f["year"].dropna().unique())
            
    default_year = int(max(inferred_years)) if inferred_years else datetime.date.today().year
    
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

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# ?곗씠???꾩쟻 ??μ냼 ???낅줈?쒗븷 ?뚮쭏??蹂묓빀쨌???
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
DATA_STORE = "wr_data_store.csv"
KEY_COLS = ["gran", "metric", "segment", "year", "label", "close"]
STORE_COLS = KEY_COLS + ["sortkey", "value"]

def load_store() -> pd.DataFrame:
    if os.path.exists(DATA_STORE):
        try:
            d = pd.read_csv(DATA_STORE, encoding="utf-8-sig")
            if set(STORE_COLS) <= set(d.columns):
                return d[STORE_COLS]
        except Exception:
            pass
    return pd.DataFrame()

def save_store(df: pd.DataFrame):
    df[STORE_COLS].to_csv(DATA_STORE, index=False, encoding="utf-8-sig")

def merge_store(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """湲곗〈 ?꾩쟻 + ?좉퇋 ?낅줈??蹂묓빀 ??媛숈? (?⑥쐞쨌吏?쑣룹콈?먃룰린媛? ?ㅻ뒗 ?좉퇋 ?곗꽑"""
    if old is None or old.empty: return new
    if new is None or new.empty: return old
    return (pd.concat([old[STORE_COLS], new[STORE_COLS]], ignore_index=True)
            .drop_duplicates(subset=KEY_COLS, keep="last"))

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# ?뚯썝 ?ㅼ쟻 ?곗씠?곗뀑 (?좉퇋/湲곗〈 ?멸렇癒쇳듃 횞 9吏?? ???⑥쐞)
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
MEMBER_STORE = "wr_member_store.csv"
MEMBER_KEY = ["metric", "segment", "year", "label"]
MEMBER_COLS = MEMBER_KEY + ["sortkey", "value"]
MEMBER_SEGS = ["Total", "?뱀썡?좉퇋", "湲곌??낆떊洹?, "湲곗〈"]
MEMBER_METRICS = ["?좏슚?뚯썝??, "UV", "?붾갑臾몄쑉(%)", "援щℓ怨좉컼??, "CR(%)",
                  "媛앸떒媛", "嫄곕옒??(VAT?쒖쇅)", "?좏슚?뚯썝??嫄곕옒??, "?쒕룞怨좉컼??]
MEMBER_PCT = {"?붾갑臾몄쑉(%)", "CR(%)"}

def is_member_grid(rows):
    """?ㅻ뜑???뚯썝援щ텇/?뚯썝援щ텇?곸꽭 + ?좏슚?뚯썝??吏?쒓? ?덉쑝硫??뚯썝 ?ㅼ쟻 ?뚯씪"""
    head = " ".join(_cell(c) for ri in range(min(3, len(rows))) for c in rows[ri])
    col0 = " ".join(_cell(rows[ri][0]) for ri in range(len(rows)) if rows[ri])
    return ("?뚯썝援щ텇" in head and "?뚯썝援щ텇?곸꽭" in head
            and "?좏슚?뚯썝?? in col0)

def fmt_member(metric, v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "??
    if metric in MEMBER_PCT: return f"{v*100:.2f}%"
    if metric == "嫄곕옒??(VAT?쒖쇅)": return f"{v/1e8:,.1f}??
    if metric in ("媛앸떒媛", "?좏슚?뚯썝??嫄곕옒??): return f"{v:,.0f}??
    return f"{int(v):,}"

def parse_member_file(name, data: bytes) -> pd.DataFrame:
    # ?뚯썝 諛깆뾽 CSV 蹂듭썝
    if name.lower().endswith(".csv"):
        first = data[:300].decode("utf-8", "ignore").splitlines()[0] if data else ""
        if "metric" in first and "segment" in first and "sortkey" in first and "gran" not in first:
            try:
                d = pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
                if set(MEMBER_COLS) <= set(d.columns) and set(d["segment"]) <= set(MEMBER_SEGS) | {"Total"}:
                    return d[MEMBER_COLS]
            except Exception:
                pass
    try:
        rows = read_grid(name, data)
    except Exception:
        return pd.DataFrame()
    if not rows or not is_member_grid(rows): return pd.DataFrame()

    # ?ㅻ뜑: 0???곕룄(ffill), 1??李⑥썝紐??붾씪踰? ?곗씠?곕뒗 2?됰???
    year_row, label_row = rows[0], rows[1]
    # ???쇰꺼???쒖옉?섎뒗 而щ읆 = '媛?낆뿰李? ?ㅼ쓬 (蹂댄넻 7??
    data_c0 = 7
    col_year, cur = {}, None
    for ci in range(len(year_row)):
        m = re.search(r"(20\d{2})", _cell(year_row[ci]))
        if m: cur = int(m.group(1))
        col_year[ci] = cur
    data_cols = [ci for ci in range(data_c0, len(label_row))
                 if re.match(r"^\d{1,2}??", _cell(label_row[ci])) and col_year.get(ci)]

    records, cur_metric = [], None
    for ri in range(2, len(rows)):
        row = rows[ri]
        m0 = _cell(row[0]) if len(row) > 0 else ""
        if m0: cur_metric = m0
        if cur_metric not in MEMBER_METRICS: continue
        seg = _cell(row[2]) if len(row) > 2 else ""   # ?뚯썝援щ텇?곸꽭
        # ?깃툒/?깅퀎/?곕졊?/媛?낆뿰李?3~6????Total???됰쭔 (?꾩옱 遺꾪빐 ?놁쓬)
        dims = [_cell(row[ci]) if len(row) > ci else "" for ci in range(3, 7)]
        if any(d and d != "Total" for d in dims): continue
        if seg not in MEMBER_SEGS: continue
        for ci in data_cols:
            mo = int(re.match(r"(\d{1,2})??, _cell(label_row[ci])).group(1))
            yr = col_year[ci]
            records.append({"metric": cur_metric, "segment": seg, "year": yr,
                            "label": f"{mo}??, "sortkey": yr * 100 + mo,
                            "value": _num(row[ci] if ci < len(row) else None)})
    return pd.DataFrame(records)

def combine_member(file_tuples) -> pd.DataFrame:
    frames = [parse_member_file(n, b) for n, b in file_tuples]
    frames = [f for f in frames if not f.empty]
    if not frames: return pd.DataFrame()
    return (pd.concat(frames, ignore_index=True)
            .drop_duplicates(subset=MEMBER_KEY, keep="last"))

def load_member_store() -> pd.DataFrame:
    if os.path.exists(MEMBER_STORE):
        try:
            d = pd.read_csv(MEMBER_STORE, encoding="utf-8-sig")
            if set(MEMBER_COLS) <= set(d.columns): return d[MEMBER_COLS]
        except Exception:
            pass
    return pd.DataFrame()

def save_member_store(df): df[MEMBER_COLS].to_csv(MEMBER_STORE, index=False, encoding="utf-8-sig")

def merge_member(old, new):
    if old is None or old.empty: return new
    if new is None or new.empty: return old
    return (pd.concat([old[MEMBER_COLS], new[MEMBER_COLS]], ignore_index=True)
            .drop_duplicates(subset=MEMBER_KEY, keep="last"))

def _period_set(d, cols):
    if d is None or d.empty: return set()
    return set(map(tuple, d[cols].drop_duplicates().values.tolist()))

def upload_diff(stored, df_new, member_stored, member_new):
    """?낅줈???곗씠?곌? 湲곗〈 ?꾩쟻 ?鍮?異붽?/媛깆떊?섎뒗 湲곌컙 ??(?????誘몃━蹂닿린??"""
    cols_c, cols_m = ["gran", "year", "label"], ["year", "label"]
    nc, oc = _period_set(df_new, cols_c), _period_set(stored, cols_c)
    nm, om = _period_set(member_new, cols_m), _period_set(member_stored, cols_m)
    added = len(nc - oc) + len(nm - om)
    updated = len(nc & oc) + len(nm & om)
    return added, updated


def member_pick(mdf, metric, seg, year, mo):
    s = mdf[(mdf["metric"] == metric) & (mdf["segment"] == seg) &
            (mdf["year"] == year) & (mdf["label"] == f"{mo}??)]["value"].dropna()
    return s.iloc[-1] if len(s) else np.nan

def member_series(mdf, metric, seg, year):
    s = mdf[(mdf["metric"] == metric) & (mdf["segment"] == seg) & (mdf["year"] == year)]
    return s.sort_values("sortkey").set_index("label")["value"]


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# 議고쉶 ?ы띁 ??mtd(?뱀썡/?뱀＜ ?쇰쭏媛? vs final ?좏깮
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
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
    """???곕룄??湲곌컙?쇰꺼 ??媛?Series (sortkey ??"""
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
    """媛??理쒓렐 (year, label)"""
    sub = df[(df["gran"] == gran) & df["value"].notna()]
    if sub.empty: return None, None
    row = sub.loc[sub["sortkey"].idxmax()]
    return int(row["year"]), row["label"]

def prev_label(df, gran, year, label):
    """吏곸쟾 湲곌컙 (?곕룄 寃쎄퀎 ?ы븿)"""
    sub = (df[(df["gran"] == gran) & df["value"].notna()]
           [["year", "label", "sortkey"]].drop_duplicates().sort_values("sortkey"))
    keys = sub[["year", "label"]].apply(tuple, axis=1).tolist()
    try: i = keys.index((year, label))
    except ValueError: return None, None
    return keys[i - 1] if i > 0 else (None, None)

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# 蹂닿퀬? ?곸냽??
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def load_insights():
    if os.path.exists(INSIGHT_FILE):
        with open(INSIGHT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_insights(d):
    with open(INSIGHT_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def report_text_block(key, title, default="", regen=None, ai_fn=None):
    """?몄쭛 媛?ν븳 蹂닿퀬? ?띿뒪??諛뺤뒪 (JSON ???.
    regen ?띿뒪?몃? 二쇰㈃ '?먮룞 ?앹꽦'(?쒗뵆由? 踰꾪듉, ai_fn??二쇰㈃ 'AI ?앹꽦'(Claude) 踰꾪듉???щ떎.
    醫곸? 而щ읆 ?덉뿉?쒕룄 ??源⑥??꾨줉 踰꾪듉? 諛뺤뒪 ???꾨옒??諛곗튂?쒕떎."""
    store = st.session_state.wr_texts
    if not store.get(key): store[key] = default
    ekey = f"__wr_edit_{key}__"
    if ekey not in st.session_state: st.session_state[ekey] = False

    # ?쒕ぉ + ?≪뀡 踰꾪듉 (?쒕ぉ ??以? 踰꾪듉? 洹??꾨옒 ?뺤긽 ?덈퉬 而щ읆)
    st.markdown(f"**{title}**")

    if st.session_state[ekey]:
        if HAS_QUILL:
            # Word ?섏? 由ъ튂 ?먮뵒??(湲???ш린쨌?됀룰도寃뙿룰린?몄엫쨌諛묒쨪쨌紐⑸줉쨌?뺣젹 ??
            toolbar = [
                [{"size": ["small", False, "large", "huge"]}],
                ["bold", "italic", "underline", "strike"],
                [{"color": []}, {"background": []}],
                [{"list": "ordered"}, {"list": "bullet"}],
                [{"align": []}], ["clean"],
            ]
            # value(珥덇린媛???理쒖큹 吏꾩엯 ?먮뒗 store[key]媛 諛붾?寃쎌슦?먮쭔 二쇱엯?쒕떎.
            # 留?rerun留덈떎 value瑜??ㅼ떆 ?ｌ쑝硫???댄븨 以???κ컪?쇰줈 ?섎룎?꾧?
            # '怨꾩냽 由ы봽?덉떆'?섎뒗 踰꾧렇媛 諛쒖깮?쒕떎. (st_quill? ?낅젰留덈떎 rerun ?좊컻)
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
        if st.button("???, key=f"wr_save_{key}", type="primary",
                     use_container_width=True):
            store[key] = new if new is not None else store[key]
            all_d = load_insights(); all_d[key] = store[key]; save_insights(all_d)
            st.session_state[ekey] = False; st.rerun()
    else:
        st.markdown(f"<div class='report-box'>{store[key] or '?댁슜???낅젰?섏꽭??'}</div>",
                    unsafe_allow_html=True)
                    
    # AI 李멸퀬 硫붾え: ?곗씠?곗뿉 ???섏삤??諛곌꼍(?꾨줈紐⑥뀡쨌?대깽?맞룹씠?????곸쑝硫?
    # AI ?앹꽦 ??[諛곌꼍 硫붾え]濡?遺꾨━ 二쇱엯???먯씤쨌留λ씫 ?댁꽍???쒖슜?쒕떎.
    memo_val = ""
    if ai_fn is not None:
        mkey = f"{key}__memo"
        if mkey not in store: store[mkey] = ""
        with st.expander("?쭬 AI 李멸퀬 硫붾え (?꾨줈紐⑥뀡쨌?대깽?맞룹슫???댁뒋 ??諛곌꼍)",
                         expanded=bool(store[mkey])):
            memo_val = st.text_area(
                "?곗씠?곗뿉 ???섏삤??諛곌꼍???곸쑝硫?AI媛 ?먯씤쨌留λ씫 ?댁꽍???쒖슜?⑸땲?? "
                "(?섏튂???곗씠?곗뿉?쒕쭔 ?몄슜)",
                store[mkey], key=f"wr_memo_{key}", height=120)
            if st.button("硫붾え ???, key=f"wr_memosave_{key}",
                         use_container_width=True):
                store[mkey] = memo_val
                all_d = load_insights(); all_d[mkey] = memo_val; save_insights(all_d)
                st.rerun()
                    
    n = 2 + (1 if regen is not None else 0) + (1 if ai_fn is not None else 0)
    bcols = st.columns(n)
    bi = 0
    edit_on = st.session_state[ekey]
    if bcols[bi].button("?몄쭛" if not edit_on else "蹂닿린",
                        key=f"wr_edit_{key}", use_container_width=True):
        st.session_state[ekey] = not edit_on; st.rerun()
    bi += 1
    if regen is not None:
        if bcols[bi].button("?먮룞 ?앹꽦", key=f"wr_regen_{key}", use_container_width=True,
                            help="湲곗? 二쇱감 ?ㅼ쟻?쇰줈 ?쒗뵆由?臾멸뎄瑜?梨꾩썎?덈떎 (湲곗〈 ?댁슜 ?泥?"):
            store[key] = regen
            all_d = load_insights(); all_d[key] = regen; save_insights(all_d)
            st.session_state[ekey] = False; st.rerun()
        bi += 1
    if ai_fn is not None:
        if bcols[bi].button("AI ?앹꽦", key=f"wr_ai_{key}", use_container_width=True,
                            help="Claude媛 ?곗씠??+李멸퀬 硫붾え)瑜?蹂닿퀬 ?몄궗?댄듃 臾멸뎄瑜??묒꽦?⑸땲??(湲곗〈 ?댁슜 ?泥?"):
            store[mkey] = memo_val
            all_d = load_insights(); all_d[mkey] = memo_val; save_insights(all_d)
            with st.spinner("AI媛 ?몄궗?댄듃瑜??묒꽦 以묅?):
                text, err = ai_fn(memo_val)
            if err:
                st.error(err)
            else:
                store[key] = text
                all_d = load_insights(); all_d[key] = text; save_insights(all_d)
                st.session_state[ekey] = False; st.rerun()
    return store[key]

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# YoY ?붿빟??/ 異붿씠??鍮뚮뜑
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def month_label(n): return f"{n}??

def week_disp(year, label):
    """二쇱감 ?쒖떆: 2026??06??2二쇱감"""
    return f"{year}??{label}" if label else "-"

def wow_summary_table(df, wy, wlabel, metrics):
    """?꾩＜鍮??붿빟??(?ㅼ쟻 ?붿빟怨??숈씪 援ъ“): ?꾩＜쨌湲곗?二셋룹쟾二쇰퉬 + ?꾨뀈?숈＜쨌?꾨뀈鍮?""
    py, plb = prev_label(df, "二?, wy, wlabel)
    cols = [week_disp(py, plb), week_disp(wy, wlabel), "?꾩＜鍮?,
            week_disp(wy - 1, wlabel), "?꾨뀈鍮?]
    rows = []
    for met in metrics:
        cur = pick(df, "二?, met, "*TOTAL", wy, wlabel, "mtd")
        prv = pick(df, "二?, met, "*TOTAL", py, plb, "final") if plb else np.nan
        yoy = pick(df, "二?, met, "*TOTAL", wy - 1, wlabel, "final")
        rows.append({
            "援щ텇": met,
            cols[0]: fmt_value(met, prv), cols[1]: fmt_value(met, cur),
            cols[2]: fmt_delta(met, cur, prv) or "??,
            cols[3]: fmt_value(met, yoy), cols[4]: fmt_delta(met, cur, yoy) or "??,
        })
    return pd.DataFrame(rows).set_index("援щ텇")

def yoy_summary_table(df, ref_year, ref_month, metrics):
    """李멸퀬蹂?'?ㅼ쟻 ?붿빟' ?? ?꾩썡쨌?뱀썡 횞 (?꾨뀈, ?밸뀈, ?꾨뀈鍮?"""
    rows = []
    pm_y, pm_m = (ref_year, ref_month - 1) if ref_month > 1 else (ref_year - 1, 12)
    cols = [f"{pm_y-1}??{pm_m}??, f"{pm_y}??{pm_m}??, "?꾨뀈鍮??꾩썡)",
            f"{ref_year-1}??{ref_month}??, f"{ref_year}??{ref_month}??, "?꾨뀈鍮??뱀썡)"]
    for met in metrics:
        pm_prev = pick(df, "??, met, "*TOTAL", pm_y - 1, month_label(pm_m), "final")
        pm_cur  = pick(df, "??, met, "*TOTAL", pm_y,     month_label(pm_m), "final")
        cm_prev = pick(df, "??, met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
        cm_cur  = pick(df, "??, met, "*TOTAL", ref_year,     month_label(ref_month), "mtd")
        rows.append({
            "援щ텇": met,
            cols[0]: fmt_value(met, pm_prev), cols[1]: fmt_value(met, pm_cur),
            cols[2]: fmt_delta(met, pm_cur, pm_prev) or "??,
            cols[3]: fmt_value(met, cm_prev), cols[4]: fmt_value(met, cm_cur),
            cols[5]: fmt_delta(met, cm_cur, cm_prev) or "??,
        })
    return pd.DataFrame(rows).set_index("援щ텇"), (pm_y, pm_m)

def trend_table(df, gran, metrics, years, seg="*TOTAL"):
    """異붿씠?? ??吏?? ??(?곕룄, 湲곌컙)"""
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
    tbl = pd.DataFrame(out, index=pd.MultiIndex.from_tuples(columns, names=["?곕룄", "湲곌컙"])).T
    return tbl

def style_trend(tbl, metrics):
    # 理쒖떊 pandas??float 而щ읆??臾몄옄????낆쓣 湲덉??섎?濡?object濡?蹂?????щ㎎
    disp = tbl.astype(object).copy()
    for met in disp.index:
        disp.loc[met] = [fmt_value(met, v) for v in tbl.loc[met]]
    return disp

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# YoY ?쇱씤李⑦듃 (Plotly)
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
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
    if gran == "??:
        x_all = [month_label(i) for i in range(1, 13)]
    else:
        x_all = labels_sorted(df, gran, years)
    fig = go.Figure()
    for i, y in enumerate(sorted(years)):
        # ?곕룄留덈떎 ?녿뒗 二쇱감(5二쇱감 ????嫄대꼫?곌퀬 ?좎쓣 ?뉖뒗??
        s = series_by_label(df, gran, metric, seg, y, prefer="final").reindex(x_all).dropna()
        fig.add_trace(go.Scatter(
            x=s.index.tolist(), y=(s / div).tolist(), mode="lines+markers", name=str(y),
            line=dict(color=clr(YEAR_PAL[i % len(YEAR_PAL)]), width=2),
            marker=dict(size=5),
        ))
    gname = "?붾퀎" if gran == "?? else "二쇱감蹂?
    ly = base_layout(h, ysuffix=unit if unit == "%" else "",
                     title=f"{metric} {gname} 異붿씠 ({unit})")
    ly["xaxis"]["categoryorder"] = "array"
    ly["xaxis"]["categoryarray"] = x_all
    if gran == "二?: ly["xaxis"]["tickangle"] = -45; ly["xaxis"]["nticks"] = 20
    fig.update_layout(**ly)
    return fig

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# ?묒? ?뚰겕遺??대낫?닿린 (??二?泥リ뎄留??붿빟 + 李⑦듃)
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def build_workbook(df, texts, ref_year, ref_month, chart_years):
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_font = Font(bold=True, size=11)
    head_fill = PatternFill("solid", fgColor="EEF3FA")
    title_font = Font(bold=True, size=13)

    # ?? ?곗씠???쒗듃 (??二?: 吏??釉붾줉 ?ㅽ깮
    extra_metrics = [m for m in df["metric"].unique() if m not in METRICS7]
    for gran, sheet in (("??, "??), ("二?, "二?)):
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

    # ?? ?붿빟 ?쒗듃
    ws = wb.create_sheet("泥リ뎄留??붿빟", 0)
    wb.remove(wb["Sheet"])
    r = 1
    ws.cell(r, 1, f"泥リ뎄留?二쇨컙蹂닿퀬 ??{ref_year}??{ref_month}??湲곗?").font = Font(bold=True, size=15)
    r += 2

    # ?ㅼ쟻 ?붿빟 YoY ??
    ws.cell(r, 1, "?ㅼ쟻 ?붿빟 (?쇳룊洹?").font = title_font; r += 1
    tbl, _ = yoy_summary_table(df, ref_year, ref_month, METRICS7)
    ws.cell(r, 1, "援щ텇").font = head_font
    for j, cname in enumerate(tbl.columns):
        c = ws.cell(r, 2 + j, cname); c.font = head_font; c.fill = head_fill
    red_font = Font(color="DC2626")
    green_font = Font(color="16A34A")
    for i, met in enumerate(tbl.index):
        ws.cell(r + 1 + i, 1, met).font = head_font
        for j, cname in enumerate(tbl.columns):
            c = ws.cell(r + 1 + i, 2 + j, tbl.loc[met, cname])
            s = str(c.value)
            if "?꾨뀈鍮? in str(cname):
                if s.startswith("??): c.font = red_font
                elif s.startswith("+"): c.font = green_font
    r += len(tbl) + 3

    # 蹂닿퀬?
    ws.cell(r, 1, "?꾩＜ 二쇱슂 吏???꾪솴").font = title_font
    ws.cell(r, 8, "湲덉＜ 吏묓뻾 ?댁슜 ?붿빟").font = title_font
    c1 = ws.cell(r + 1, 1, texts.get("wr_metrics_summary", ""))
    c2 = ws.cell(r + 1, 8, texts.get("wr_exec_summary", ""))
    for c in (c1, c2): c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 8, end_column=6)
    ws.merge_cells(start_row=r + 1, start_column=8, end_row=r + 8, end_column=13)
    r += 11

    # 李⑦듃 ?곗씠??釉붾줉 + ?쇱씤李⑦듃 (?붾퀎/二쇱감蹂?횞 嫄곕옒?≤룰퀬媛앹닔쨌媛앸떒媛)
    chart_metrics = ["泥リ뎄留?嫄곕옒??, "泥リ뎄留?怨좉컼??, "泥リ뎄留?媛앸떒媛"]
    for gran, gname in (("??, "?붾퀎"), ("二?, "二쇱감蹂?)):
        x_all = ([month_label(i) for i in range(1, 13)] if gran == "??
                 else labels_sorted(df, gran, chart_years))
        if not x_all: continue
        anchor_row = r
        for k, met in enumerate(chart_metrics):
            ws.cell(r, 1, f"{met} {gname} (李⑦듃 ?곗씠??").font = head_font
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
                ch.title = f"{met} {gname} 異붿씠"
                ch.height, ch.width = 7.5, 13
                data = Reference(ws, min_col=1, min_row=r + 1,
                                 max_col=1 + len(x_all), max_row=r + nrow)
                cats = Reference(ws, min_col=2, min_row=r, max_col=1 + len(x_all))
                ch.add_data(data, titles_from_data=True, from_rows=True)
                ch.set_categories(cats)
                ws.add_chart(ch, f"{get_column_letter(2 + len(x_all) + 1 + (k % 3) * 8)}{anchor_row}")
            r += nrow + 2
        r += 14

    # 梨꾨꼸蹂??ㅼ쟻 (?뱀썡 YoY)
    ws.cell(r, 1, f"梨꾨꼸蹂??ㅼ쟻 ??{ref_year}??{ref_month}??(?꾨뀈鍮?").font = title_font; r += 1
    for met in chart_metrics:
        ws.cell(r, 1, met).font = head_font
        heads = [f"{ref_year-1}??{ref_month}??, f"{ref_year}??{ref_month}??, "?꾨뀈鍮?]
        for j, hd in enumerate(heads):
            c = ws.cell(r, 2 + j, hd); c.font = head_font; c.fill = head_fill
        segs = ["*TOTAL"] + CHANNELS
        for i, seg in enumerate(segs):
            pv = pick(df, "??, met, seg, ref_year - 1, month_label(ref_month), "mtd")
            cv = pick(df, "??, met, seg, ref_year, month_label(ref_month), "mtd")
            ws.cell(r + 1 + i, 1, seg)
            ws.cell(r + 1 + i, 2, None if np.isnan(pv) else pv).number_format = "#,##0"
            ws.cell(r + 1 + i, 3, None if np.isnan(cv) else cv).number_format = "#,##0"
            dcell = ws.cell(r + 1 + i, 4, fmt_delta(met, cv, pv) or "??)
            ds = str(dcell.value)
            if ds.startswith("??): dcell.font = red_font
            elif ds.startswith("+"): dcell.font = green_font
        r += len(segs) + 3

    ws.column_dimensions["A"].width = 16
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# ?먮룞 蹂닿퀬 珥덉븞
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def _delta_html(d):
    """利앷컧 臾몄옄?????됱긽 span (??떊????鍮④컯 / ?좎옣 + 珥덈줉)"""
    if not d: return ""
    if d.startswith("??): return f'<span style="color:#dc2626;font-weight:700">{d}</span>'
    return f'<span style="color:#16a34a;font-weight:700">{d}</span>'

def auto_draft(df, ref_year, ref_month, ref_week=None):
    """湲곗? 二쇱감(?놁쑝硫?湲곗? ?? ?ㅼ쟻?쇰줈 蹂닿퀬 臾멸뎄 ?먮룞 ?앹꽦 (HTML, ??떊??鍮④컯)"""
    lines = []
    for met in METRICS7:
        if ref_week:
            cur = pick(df, "二?, met, "*TOTAL", ref_year, ref_week, "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            py, plb = prev_label(df, "二?, ref_year, ref_week)
            prv = pick(df, "二?, met, "*TOTAL", py, plb, "final") if plb else np.nan
            yoy = pick(df, "二?, met, "*TOTAL", ref_year - 1, ref_week, "final")
            parts = [f" - {met} ??{fmt_value(met, cur)}"]
            d_w = fmt_delta(met, cur, prv)
            d_y = fmt_delta(met, cur, yoy)
            if d_w: parts.append(f"?꾩＜鍮?{_delta_html(d_w)}")
            if d_y: parts.append(f"?꾨뀈鍮?{_delta_html(d_y)}")
            lines.append(", ".join(parts))
        else:
            cur = pick(df, "??, met, "*TOTAL", ref_year, month_label(ref_month), "mtd")
            prv = pick(df, "??, met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            d = fmt_delta(met, cur, prv)
            tail = f", ?꾨뀈鍮?{_delta_html(d)}" if d else ""
            lines.append(f" - {met} ??{fmt_value(met, cur)}" + tail)
    return "<br>".join(lines)

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# AI ?몄궗?댄듃 ?앹꽦 (Claude API) ??紐⑤뜽 援먯껜 媛??
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
AI_MODELS = {
    "Claude Sonnet 4.6 (洹좏삎쨌湲곕낯)": "claude-sonnet-4-6",
    "Claude Opus 4.8 (理쒓퀬 ?덉쭏)": "claude-opus-4-8",
    "Claude Haiku 4.5 (鍮좊쫫쨌???": "claude-haiku-4-5",
}

def _anthropic_key():
    """Streamlit secrets ?먮뒗 ?섍꼍蹂?섏뿉??API ??議고쉶"""
    try:
        if "ANTHROPIC_API_KEY" in st.secrets:
            return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY")

def _ai_metric_facts(df, ref_year, ref_month, ref_week=None):
    """紐⑤뜽???섍만 吏?쑣룹쬆媛??붿빟 ?띿뒪??""
    rows = []
    for met in METRICS7:
        if ref_week:
            cur = pick(df, "二?, met, "*TOTAL", ref_year, ref_week, "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            py, plb = prev_label(df, "二?, ref_year, ref_week)
            prv = pick(df, "二?, met, "*TOTAL", py, plb, "final") if plb else np.nan
            yoy = pick(df, "二?, met, "*TOTAL", ref_year - 1, ref_week, "final")
            wm = re.match(r"(\d{1,2})??(\d)二쇱감", ref_week)
            mom = np.nan
            if wm:
                mo, wk = int(wm.group(1)), int(wm.group(2))
                my, mm = (ref_year, mo - 1) if mo > 1 else (ref_year - 1, 12)
                mom = pick(df, "二?, met, "*TOTAL", my, f"{mm:02d}??{wk}二쇱감", "final")
            rows.append(f"- {met}: {fmt_value(met, cur)} "
                        f"(?꾩＜鍮?{fmt_delta(met, cur, prv) or '??}, "
                        f"?꾩썡鍮?{fmt_delta(met, cur, mom) or '??}, "
                        f"?꾨뀈鍮?{fmt_delta(met, cur, yoy) or '??})")
        else:
            cur = pick(df, "??, met, "*TOTAL", ref_year, month_label(ref_month), "mtd")
            prv = pick(df, "??, met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            rows.append(f"- {met}: {fmt_value(met, cur)} (?꾨뀈鍮?{fmt_delta(met, cur, prv) or '??})")
    return "\n".join(rows)

def ai_generate_insight(df, ref_year, ref_month, ref_week, model,
                        focus="?꾩＜ 二쇱슂 吏???꾪솴", memo=""):
    """Claude API濡?蹂닿퀬 ?몄궗?댄듃 ?앹꽦 ??HTML(??떊??鍮④컯) 諛섑솚. (?띿뒪?? ?먮윭) ?쒗뵆.
    memo: ?ъ슜?먭? ?곸? ?뺤꽦 諛곌꼍(?꾨줈紐⑥뀡쨌?대깽????. ?섏튂? 遺꾨━??[諛곌꼍 硫붾え]濡?二쇱엯."""
    key = _anthropic_key()
    if not key:
        return None, ("ANTHROPIC_API_KEY媛 ?ㅼ젙?섏? ?딆븯?듬땲?? "
                      "Streamlit Cloud ??Settings ??Secrets??ANTHROPIC_API_KEY瑜?異붽??섏꽭??")
    try:
        import anthropic
    except ImportError:
        return None, "anthropic ?⑦궎吏媛 ?ㅼ튂?섏? ?딆븯?듬땲?? requirements.txt 諛섏쁺 ???щ같?ы븯?몄슂."

    period = f"{ref_year}??{ref_week}" if ref_week else f"{ref_year}??{ref_month}??
    facts = _ai_metric_facts(df, ref_year, ref_month, ref_week)
    system = (
        "?뱀떊? LF紐?CRM 泥リ뎄留?蹂닿퀬?쒕? ?묒꽦?섎뒗 ?곗씠??遺꾩꽍媛?낅땲?? "
        "?쒓뎅???ㅻТ 蹂닿퀬 臾멸뎄瑜??묒꽦???? 湲?臾몄옣??遺덈┸(??...? ...?쇰줈 ...)???쇳븯怨?"
        "諛섎뱶??'- '(硫붿씤 ?붿빟)? '??'(?몃? ?섏튂/?댁꽍)瑜??ъ슜?섎뒗 怨꾩링??遺덈┸ 援ъ“濡?異쒕젰?섏꽭??\n\n"
        "?덉떆:\n"
        "- 泥リ뎄留?嫄곕옒??諛?怨좉컼???숇컲 媛먯냼\n"
        "??嫄곕옒??88.9諛깅쭔??(?꾩＜鍮???.5% 쨌 ?꾨뀈鍮???4.9%)\n"
        "??怨좉컼??639紐?(?꾩＜鍮???.7% 쨌 ?꾨뀈鍮???4.1%)\n"
        "???꾨줈紐⑥뀡 醫낅즺 ?ы뙆濡??좎엯 ?鍮??꾪솚 ?⑥쑉???議고븳 寃껋쑝濡?蹂댁엫\n\n"
        "[?뺤젙 ?섏튂]???덈뒗 ?レ옄留??몄슜?섍퀬 ?섏튂瑜?吏?대궡吏 留덉꽭?? "
        "[諛곌꼍 硫붾え]???먯씤 異붿젙/?댁꽍?먮쭔 ?쒖슜?섎ŉ(?⑥젙吏볦? 留먭퀬 '~濡?蹂댁엫' ???ъ슜), 鍮꾩뼱?덉쑝硫??섏튂 ?⑺듃留?湲곗옱?섏꽭?? "
        "異쒕젰? HTML濡쒕쭔 (<br> 濡?以꾨컮轅?. "
        "利앷컧 ?섏튂 以???떊??媛먯냼)? <span style=\"color:#dc2626;font-weight:700\">??/span>, "
        "?좎옣(利앷?)? <span style=\"color:#16a34a;font-weight:700\">??/span>濡?媛먯떥?몄슂. "
        "?쒕줎쨌留븐쓬留??놁씠 諛붾줈 怨꾩링??遺덈┸留?異쒕젰?섏꽭??"
    )
    memo_block = (memo or "").strip() or "(?놁쓬)"
    user = (f"[湲곗?: {period}]\n\n"
            f"[?뺤젙 ?섏튂]\n{facts}\n\n"
            f"[諛곌꼍 硫붾え]\n{memo_block}\n\n"
            f"???먮즺濡?'{focus}' 臾멸뎄瑜??묒꽦?섏꽭??")
    try:
        client = anthropic.Anthropic(api_key=key)
        resp = client.messages.create(
            model=model, max_tokens=2000, system=system,
            messages=[{"role": "user", "content": user}],
        )
        text = "".join(b.text for b in resp.content if b.type == "text").strip()
        return (text or None), (None if text else "鍮??묐떟??諛섑솚?먯뒿?덈떎.")
    except anthropic.AuthenticationError:
        return None, "API ???몄쬆???ㅽ뙣?덉뒿?덈떎. ?ㅻ? ?뺤씤?섏꽭??"
    except anthropic.RateLimitError:
        return None, "?붿껌??留롮븘 ?쇱떆?곸쑝濡??쒗븳?먯뒿?덈떎. ?좎떆 ???ㅼ떆 ?쒕룄?섏꽭??"
    except Exception as e:
        return None, f"?앹꽦 以??ㅻ쪟: {e}"

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# ?뚯썝 ?ㅼ쟻 ?섏씠吏
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def member_delta(metric, cur, prev):
    if cur is None or prev is None: return None
    if isinstance(cur, float) and np.isnan(cur): return None
    if isinstance(prev, float) and np.isnan(prev): return None
    if metric in MEMBER_PCT:
        d = (cur - prev) * 100
        return f"??abs(d):.2f}%p" if d < 0 else f"+{d:.2f}%p"
    if prev == 0: return None
    d = (cur - prev) / prev * 100
    return f"??abs(d):.1f}%" if d < 0 else f"+{d:.1f}%"

def render_member_page(mdf, chart_years=None):
    st.markdown("## ?뚯썝 ?ㅼ쟻 ???좉퇋/湲곗〈 ?멸렇癒쇳듃")
    if mdf.empty:
        st.info("?뚯썝 ?ㅼ쟻 ?곗씠?곌? ?놁뒿?덈떎. ?뚯썝 ?ㅼ쟻 ?묒????낅줈?쒗빐二쇱꽭??")
        return
    years = sorted(mdf["year"].dropna().unique().astype(int))
    last_y = years[-1]
    last_mo = int(mdf[mdf["year"] == last_y]["sortkey"].max() % 100)
    st.caption(f"理쒖떊: {last_y}??{last_mo}??쨌 {years[0]}??years[-1]}??)

    # KPI 移대뱶 ???좏슚?뚯썝??援щℓ怨좉컼??嫄곕옒??CR (Total, ?꾩썡쨌?꾨뀈鍮?
    pm_y, pm_m = (last_y, last_mo - 1) if last_mo > 1 else (last_y - 1, 12)
    kpis = ["?좏슚?뚯썝??, "援щℓ怨좉컼??, "嫄곕옒??(VAT?쒖쇅)", "CR(%)"]
    cols = st.columns(4)
    for col, met in zip(cols, kpis):
        cur = member_pick(mdf, met, "Total", last_y, last_mo)
        mom = member_pick(mdf, met, "Total", pm_y, pm_m)
        yoy = member_pick(mdf, met, "Total", last_y - 1, last_mo)
        pills = ""
        for d, lab in [(member_delta(met, cur, mom), "?꾩썡鍮?),
                       (member_delta(met, cur, yoy), "?꾨뀈鍮?)]:
            if d:
                neg = d.startswith("??)
                pills += (f'<div class="kpi-delta {"down" if neg else "up"}">'
                          f'{"" if neg else "??"}{d} ({lab})</div>')
            else:
                pills += f'<div class="kpi-delta na">??({lab})</div>'
        col.markdown(f'<div class="kpi-card"><div class="kpi-label">{met} ({last_mo}??</div>'
                     f'<div class="kpi-value">{fmt_member(met, cur)}</div>{pills}</div>',
                     unsafe_allow_html=True)
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # ?좉퇋/湲곗〈 援ъ꽦鍮?(?좏슚?뚯썝??理쒖떊??
    st.subheader(f"?뚯썝 援ъ꽦 ??{last_y}??{last_mo}??(?좏슚?뚯썝??")
    comp = {s: member_pick(mdf, "?좏슚?뚯썝??, s, last_y, last_mo)
            for s in ["?뱀썡?좉퇋", "湲곌??낆떊洹?, "湲곗〈"]}
    total = member_pick(mdf, "?좏슚?뚯썝??, "Total", last_y, last_mo)
    c1, c2 = st.columns([1, 1])
    with c1:
        fig = go.Figure(go.Pie(
            labels=list(comp.keys()), values=[v for v in comp.values()],
            hole=0.55, marker=dict(colors=[clr("blue"), clr("teal"), clr("slate")]),
            textinfo="label+percent"))
        fig.update_layout(**{**base_layout(280, title="?좏슚?뚯썝 援ъ꽦鍮?), "showlegend": False})
                            st.plotly_chart(fig, use_container_width=True)

            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
            st.subheader("월별 추이표 (합계 및 평균 동의율)")
            
            # 일별 데이터를 월별 합계로 집계하여 전년비 표 생성
            push_metrics = ["가입자수", "앱푸시수신동의", "동의율"]
            sub = df[(df["gran"] == "일") & (df["year"].isin(chart_years)) & (df["metric"].isin(["가입자수", "앱푸시수신동의"]))].copy()
            
            if not sub.empty:
                # 4/24 이상치 제거 반영 (집계 시에도 제외되도록)
                glitch_dates = ["4/24", "04/24", "2026/04/24", "2026-04-24", "4-24", "04-24"]
                sub = sub[~sub["label"].astype(str).str.strip().isin(glitch_dates)]
                
                sub["month"] = sub["label"].apply(lambda x: int(str(x).split('/')[0]) if '/' in str(x) else 0)
                grp = sub.groupby(["year", "month", "metric"])["value"].sum().unstack("metric").reset_index()
                
                if "가입자수" in grp.columns and "앱푸시수신동의" in grp.columns:
                    grp["동의율"] = grp["앱푸시수신동의"] / grp["가입자수"]
                    grp.loc[grp["동의율"] > 1.0, "동의율"] = np.nan
                else:
                    grp["동의율"] = np.nan
                    
                cols = []
                out_tbl = {}
                for y in chart_years:
                    for m in range(1, 13):
                        row = grp[(grp["year"] == y) & (grp["month"] == m)]
                        if not row.empty:
                            cols.append((y, f"{m}월"))
                            for met in push_metrics:
                                if met not in out_tbl: out_tbl[met] = []
                                out_tbl[met].append(row[met].values[0] if met in row.columns else np.nan)
                                
                if cols:
                    tbl_df = pd.DataFrame(out_tbl, index=pd.MultiIndex.from_tuples(cols, names=["연도", "기간"])).T
                    st.dataframe(style_trend(tbl_df, push_metrics), use_container_width=True)
    with c2:
        rows = []
        for s in ["Total", "?뱀썡?좉퇋", "湲곌??낆떊洹?, "湲곗〈"]:
            v = member_pick(mdf, "?좏슚?뚯썝??, s, last_y, last_mo)
            share = (v / total * 100) if (total and not np.isnan(v)) else np.nan
            rows.append({"?멸렇癒쇳듃": s, "?좏슚?뚯썝??: fmt_member("?좏슚?뚯썝??, v),
                         "鍮꾩쨷": "?? if np.isnan(share) else f"{share:.1f}%"})
        st.dataframe(pd.DataFrame(rows).set_index("?멸렇癒쇳듃"), use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 吏???좏깮 ???멸렇癒쇳듃蹂?YoY 異붿씠
    met = st.selectbox("吏???좏깮", MEMBER_METRICS, key="wr_mem_metric")
    cyears = chart_years or years[-2:]
    cyears = [y for y in cyears if y in years] or years[-2:]
    div = 1e8 if met == "嫄곕옒??(VAT?쒖쇅)" else 1
    unit = "?? if met == "嫄곕옒??(VAT?쒖쇅)" else ("%" if met in MEMBER_PCT else "")
    if met in MEMBER_PCT: div = 0.01

    st.subheader(f"{met} ???멸렇癒쇳듃蹂???異붿씠 ({last_y}??")
    x_all = [f"{i}?? for i in range(1, 13)]
    fig = go.Figure()
    seg_color = {"Total": "slate", "?뱀썡?좉퇋": "blue", "湲곌??낆떊洹?: "teal", "湲곗〈": "amber"}
    for s in MEMBER_SEGS:
        ser = member_series(mdf, met, s, last_y).reindex(x_all).dropna()
        if ser.empty: continue
        fig.add_trace(go.Scatter(x=ser.index.tolist(), y=(ser / div).tolist(),
                                 mode="lines+markers", name=s,
                                 line=dict(color=clr(seg_color[s]), width=2), marker=dict(size=5)))
    ly = base_layout(320, ysuffix=unit if unit == "%" else "", title=f"{met} ({unit})")
    ly["xaxis"]["categoryorder"] = "array"; ly["xaxis"]["categoryarray"] = x_all
    fig.update_layout(**ly)
                        st.plotly_chart(fig, use_container_width=True)

            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
            st.subheader("월별 추이표 (합계 및 평균 동의율)")
            
            # 일별 데이터를 월별 합계로 집계하여 전년비 표 생성
            push_metrics = ["가입자수", "앱푸시수신동의", "동의율"]
            sub = df[(df["gran"] == "일") & (df["year"].isin(chart_years)) & (df["metric"].isin(["가입자수", "앱푸시수신동의"]))].copy()
            
            if not sub.empty:
                # 4/24 이상치 제거 반영 (집계 시에도 제외되도록)
                glitch_dates = ["4/24", "04/24", "2026/04/24", "2026-04-24", "4-24", "04-24"]
                sub = sub[~sub["label"].astype(str).str.strip().isin(glitch_dates)]
                
                sub["month"] = sub["label"].apply(lambda x: int(str(x).split('/')[0]) if '/' in str(x) else 0)
                grp = sub.groupby(["year", "month", "metric"])["value"].sum().unstack("metric").reset_index()
                
                if "가입자수" in grp.columns and "앱푸시수신동의" in grp.columns:
                    grp["동의율"] = grp["앱푸시수신동의"] / grp["가입자수"]
                    grp.loc[grp["동의율"] > 1.0, "동의율"] = np.nan
                else:
                    grp["동의율"] = np.nan
                    
                cols = []
                out_tbl = {}
                for y in chart_years:
                    for m in range(1, 13):
                        row = grp[(grp["year"] == y) & (grp["month"] == m)]
                        if not row.empty:
                            cols.append((y, f"{m}월"))
                            for met in push_metrics:
                                if met not in out_tbl: out_tbl[met] = []
                                out_tbl[met].append(row[met].values[0] if met in row.columns else np.nan)
                                
                if cols:
                    tbl_df = pd.DataFrame(out_tbl, index=pd.MultiIndex.from_tuples(cols, names=["연도", "기간"])).T
                    st.dataframe(style_trend(tbl_df, push_metrics), use_container_width=True)

    # Total ?꾨뀈 鍮꾧탳
    st.subheader(f"{met} ??Total ?꾨뀈 鍮꾧탳")
    fig2 = go.Figure()
    for i, y in enumerate(sorted(cyears)):
        ser = member_series(mdf, met, "Total", y).reindex(x_all).dropna()
        if ser.empty: continue
        fig2.add_trace(go.Scatter(x=ser.index.tolist(), y=(ser / div).tolist(),
                                  mode="lines+markers", name=str(y),
                                  line=dict(color=clr(YEAR_PAL[i % len(YEAR_PAL)]), width=2),
                                  marker=dict(size=5)))
    ly2 = base_layout(320, ysuffix=unit if unit == "%" else "", title=f"{met} Total ({unit})")
    ly2["xaxis"]["categoryorder"] = "array"; ly2["xaxis"]["categoryarray"] = x_all
    fig2.update_layout(**ly2)
    st.plotly_chart(fig2, use_container_width=True)

    # ?멸렇癒쇳듃 횞 ????(理쒖떊 ?곕룄)
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader(f"{met} ???멸렇癒쇳듃 횞 ??({last_y}??")
    rows = []
    for s in MEMBER_SEGS:
        ser = member_series(mdf, met, s, last_y)
        row = {"?멸렇癒쇳듃": s}
        for lb in x_all:
            if lb in ser.index and not (isinstance(ser[lb], float) and np.isnan(ser[lb])):
                row[lb] = fmt_member(met, ser[lb])
        rows.append(row)
    st.dataframe(pd.DataFrame(rows).set_index("?멸렇癒쇳듃"), use_container_width=True)

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# ?섏씠吏 PDF ???(釉뚮씪?곗? ?몄뇙 ??PDF, 李⑦듃 ?ы븿)
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def print_button(label="???섏씠吏 PDF ???/ ?몄뇙"):
    components.html(
        f"""<button onclick="window.parent.print()"
        style="float:right;background:#2E68B0;color:#fff;border:0;border-radius:6px;
        padding:7px 14px;font-size:13px;font-weight:600;cursor:pointer;
        font-family:'Pretendard',-apple-system,sans-serif">{label}</button>
        <div style="clear:both"></div>""", height=44)
    st.caption("踰꾪듉???숈옉?섏? ?딆쑝硫?Ctrl+P(Mac ??P) ????곸쓣 'PDF濡?????쇰줈 ?몄뇙?섏꽭??")

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# 硫붿씤 ??
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def main():
    if "wr_texts" not in st.session_state:
        st.session_state.wr_texts = load_insights()

    with st.sidebar:
        st.markdown("## ?뱥 二쇨컙蹂닿퀬 ?듯빀")
        files = st.file_uploader(
            "?먯쿇 ?묒?/CSV/ZIP ?낅줈??(蹂듭닔 ?좏깮)",
            type=["xlsx", "xls", "csv", "zip"], accept_multiple_files=True, key="wr_up",
            help="二쇨컙 ?대뜑瑜?zip?쇰줈 臾띠뼱 ?듭㎏濡??щ젮???⑸땲?? "
                 "?꾩껜愿??留덉뒪????二??? + 吏?쒕퀎 ?뚯씪(媛?낆쑉쨌媛?낆옄?샕룸떦?쇨???泥リ뎄留ㅼ쑉쨌鍮꾪쉶???몃옒?????먮룞 ?몄떇?⑸땲??")
        st.markdown("---")
        PAGES = ["01. 二쇨컙蹂닿퀬 ?붿빟", "02. ?붾퀎 異붿씠", "03. 二쇱감蹂?異붿씠",
                 "04. 梨꾨꼸蹂??ㅼ쟻", "05. ?뚯썝 ?ㅼ쟻", "06. ?듯빀 ?곗씠?걔룸떎?대줈??, "07. ?깊뫖???숈쓽 ?꾪솴"]
        page = st.radio("?섏씠吏", PAGES, key="wr_page")

    stored = load_store()
    member_stored = load_member_store()
    expanded = expand_uploads(files) if files else []
    df_new = combine_files(tuple(expanded)) if expanded else pd.DataFrame()
    member_new = combine_member(tuple(expanded)) if expanded else pd.DataFrame()

    has_any = (not stored.empty or not member_stored.empty
               or not df_new.empty or not member_new.empty)
    if files and df_new.empty and member_new.empty and stored.empty and member_stored.empty:
        st.error("?낅줈?쒗븳 ?뚯씪?먯꽌 ?곗씠?곕? ?쎌? 紐삵뻽?듬땲?? ?뚯씪紐??뺤떇???뺤씤?댁＜?몄슂.")
        st.stop()
    if not has_any:
        st.info("?몚 ?ъ씠?쒕컮?먯꽌 二쇨컙 ?대뜑???묒?/CSV/ZIP ?뚯씪?ㅼ쓣 ?낅줈?쒗빐二쇱꽭??")
        st.markdown("""
- **留덉뒪??*: `?꾩껜愿??- ?쇱옄蹂?二쇰퀎/?붾퀎 ?ㅼ쟻 (湲곕낯)`
- **吏?쒕퀎**: `??媛?낆쑉(?쇳룊洹?`, `二?媛?낆옄???쇳룊洹?`, `??鍮꾪쉶???몃옒???쇳룊洹?`, `???뱀씪媛??泥リ뎄留ㅼ쑉 (?쇳룊洹?` ??
- **?뚯썝 ?ㅼ쟻**: ?좉퇋/湲곗〈 ?뚯썝蹂?9媛?吏???붾퀎 ?뚯씪 (?좏슚?뚯썝?샕톃V쨌援щℓ怨좉컼?샕텰R ??
- 二쇨컙 ?대뜑瑜?**zip?쇰줈 臾띠뼱 ?듭㎏濡?* ?щ젮???⑸땲?? ?뚯씪 ?댁슜?쇰줈 ?⑥쐞쨌吏?쒕? ?먮룞 媛먯??⑸땲??
- ?낅줈????**???*???꾨Ⅴ硫??꾩쟻?⑸땲?? ?ㅼ쓬??湲곌컙???ㅻⅨ ?뚯씪???щ━硫?**寃뱀튂??湲곌컙? 理쒖떊媛믪쑝濡?媛깆떊**?섍퀬 ?섎㉧吏???댁뼱遺숈뒿?덈떎.
""")
        st.stop()

    # ?꾩쟻 ??μ냼? 蹂묓빀 ??誘몃━蹂닿린(????꾧퉴吏 ?곴뎄 諛섏쁺 ????
    df = merge_store(stored, df_new)
    mdf = merge_member(member_stored, member_new)

    has_new = (not df_new.empty) or (not member_new.empty)
    sig = tuple(sorted((n, len(b)) for n, b in expanded))
    with st.sidebar:
        if has_new:
            added, updated = upload_diff(stored, df_new, member_stored, member_new)
            saved = st.session_state.get("wr_saved_sig") == sig
            if saved:
                st.success("??λ맖 ??(?꾩쟻 諛섏쁺 ?꾨즺)")
            else:
                st.warning(f"???곗씠??媛먯? ??異붽? {added}湲곌컙 쨌 媛깆떊(寃뱀묠) {updated}湲곌컙\n\n"
                           "**???* ?뚮윭???꾩쟻??諛섏쁺?⑸땲??")
                if st.button("?뮶 ???(?꾩쟻 諛섏쁺)", key="wr_commit",
                             type="primary", use_container_width=True):
                    if not df_new.empty: save_store(df)
                    if not member_new.empty: save_member_store(mdf)
                    st.session_state["wr_saved_sig"] = sig
                    st.rerun()

    # 肄붿뼱 ?곗씠?곌? 鍮꾩뼱 ?뚯썝 ?곗씠?곕쭔 ?덉쓣 ?? ?뚯썝 ?섏씠吏濡??덈궡
    if df.empty:
        st.warning("泥リ뎄留??꾩껜愿??吏?쒕퀎) ?곗씠?곌? ?놁뒿?덈떎. **05. ?뚯썝 ?ㅼ쟻** ?섏씠吏瑜??댁슜?섏꽭??")
        if not mdf.empty:
            print_button()
            render_member_page(mdf)
        st.stop()

    # ?? ?몄떇 寃곌낵 + ?꾪꽣
    years_all = sorted(df["year"].dropna().unique().astype(int))
    ly, llabel = latest_period(df, "??)
    ref_year_default = ly or years_all[-1]
    with st.sidebar:
        st.markdown("---")
        src = f"?몄떇???뚯씪 {len(expanded)}媛? if expanded else "?꾩쟻 ?곗씠???ъ슜 以?
        st.caption(f"{src} 쨌 吏??{df['metric'].nunique()}醫?쨌 "
                   f"{years_all[0]}??years_all[-1]}??)
        st.markdown("**湲곗? 湲곌컙**")
        ref_year = st.selectbox("湲곗? ?곕룄", years_all[::-1],
                                index=years_all[::-1].index(ref_year_default), key="wr_refy")
        months_avail = sorted({int(re.match(r"(\d+)??, l).group(1))
                               for l in df[(df["gran"] == "??) & (df["year"] == ref_year)]["label"]})
        ref_month = st.selectbox("湲곗? ??, months_avail[::-1], key="wr_refm")
        weeks_avail = (df[(df["gran"] == "二?) & (df["year"] == ref_year) & df["value"].notna()]
                       [["label", "sortkey"]].drop_duplicates()
                       .sort_values("sortkey")["label"].tolist())
        if weeks_avail:
            # 理쒖떊 二쇱감媛 吏꾪뻾 以??쇰쭏媛??곗씠?곗씠嫄곕굹 ?ㅻ뒛???랁븳 二쇱감)?대㈃ 吏곸쟾 二쇰? 湲곕낯媛믪쑝濡?
            latest_w = weeks_avail[-1]
            today = datetime.date.today()
            cur_week_lbl = f"{today.month:02d}??{(today.day - 1) // 7 + 1}二쇱감"
            is_partial = not df[(df["gran"] == "二?) & (df["year"] == ref_year) &
                                (df["label"] == latest_w) & (df["close"] == "mtd")].empty
            in_progress = is_partial or (ref_year == today.year and latest_w == cur_week_lbl)
            default_week = (weeks_avail[-2] if in_progress and len(weeks_avail) >= 2
                            else latest_w)
            ref_week = st.selectbox("湲곗? 二쇱감", weeks_avail[::-1],
                                    index=weeks_avail[::-1].index(default_week),
                                    key="wr_refw",
                                    help="二쇨컙蹂닿퀬 ???二쇱감. 理쒖떊 二쇱감媛 吏꾪뻾 以묒씠硫?吏곸쟾 ?꾨즺 二쇱감媛 湲곕낯媛믪엯?덈떎.")
        else:
            ref_week = None
        st.markdown("**李⑦듃 ?곕룄**")
        default_yrs = years_all[-2:] if len(years_all) >= 2 else years_all
        chart_years = st.multiselect("鍮꾧탳 ?곕룄", years_all, default=default_yrs, key="wr_cyrs")
        st.markdown("**梨꾨꼸**")
        ch_sel = st.multiselect("梨꾨꼸 ?좏깮", CHANNELS, default=CHANNELS, key="wr_ch")
        if not chart_years: chart_years = default_yrs

        st.markdown("---")
        st.markdown("**AI ?몄궗?댄듃 紐⑤뜽**")
        ai_label = st.selectbox("紐⑤뜽 ?좏깮", list(AI_MODELS.keys()), key="wr_ai_model_label")
        st.session_state["wr_ai_model"] = AI_MODELS[ai_label]
        st.caption("??API ???ㅼ젙?? if _anthropic_key()
                   else "??ANTHROPIC_API_KEY 誘몄꽕????Secrets??異붽??섏꽭??)

        st.markdown("---")
        st.markdown("**?꾩쟻 ?곗씠??*")
        saved_rows = len(load_store())
        pend = " 쨌 ?????諛섏쁺" if (has_new and not st.session_state.get("wr_saved_sig") == sig) else ""
        st.caption(f"??λ맖 {saved_rows:,}??/ ?꾩옱 蹂닿린 {len(df):,}??pend}")
        st.download_button("燧??꾩쟻 ?곗씠??諛깆뾽 (CSV)",
                           df[STORE_COLS].to_csv(index=False).encode("utf-8-sig"),
                           "wr_data_store.csv", "text/csv", use_container_width=True,
                           help="???щ같?????꾩쟻 ?곗씠?곌? 珥덇린?붾맆 ???덉쑝??二쇨린?곸쑝濡?諛깆뾽?섏꽭?? ??CSV瑜??ㅼ떆 ?낅줈?쒗븯硫?蹂듭썝?⑸땲??")
        if not mdf.empty:
            st.caption(f"?뚯썝 ?ㅼ쟻 {len(mdf):,}??)
            st.download_button("?뮶 ?뚯썝 ?ㅼ쟻 諛깆뾽 (CSV)",
                               mdf[MEMBER_COLS].to_csv(index=False).encode("utf-8-sig"),
                               "wr_member_store.csv", "text/csv", use_container_width=True)
        if st.button("?뿊 ?꾩쟻 ?곗씠??珥덇린??, key="wr_clear_store", use_container_width=True):
            if os.path.exists(MEMBER_STORE): os.remove(MEMBER_STORE)
            if os.path.exists(DATA_STORE): os.remove(DATA_STORE)
            st.cache_data.clear()
            st.rerun()

        st.markdown("---")
        st.markdown("**蹂닿퀬?쨌硫붾え**")
        st.download_button(
            "燧?蹂닿퀬?쨌硫붾え 諛깆뾽 (JSON)",
            json.dumps(st.session_state.wr_texts, ensure_ascii=False, indent=2).encode("utf-8"),
            "wr_insights.json", "application/json", use_container_width=True,
            help="紐⑤뱺 蹂닿퀬?쨌?몄궗?댄듃 硫붾え瑜?諛깆뾽?⑸땲?? ?щ같?щ줈 珥덇린?붾뤌?????뚯씪??蹂듭썝?섎㈃ ?섏궡?꾨궔?덈떎.")
        restore = st.file_uploader("硫붾え 蹂듭썝 (JSON ?낅줈??", type=["json"], key="wr_restore_memo")
        if restore is not None:
            try:
                data = json.loads(restore.getvalue().decode("utf-8"))
                if isinstance(data, dict) and st.session_state.get("wr_restored") != restore.name:
                    merged = {**st.session_state.wr_texts, **data}  # ?낅줈??媛??곗꽑
                    st.session_state.wr_texts = merged
                    all_d = load_insights(); all_d.update(merged); save_insights(all_d)
                    st.session_state["wr_restored"] = restore.name
                    st.success(f"{len(data)}媛?硫붾え 蹂듭썝????); st.rerun()
            except Exception as e:
                st.error(f"蹂듭썝 ?ㅽ뙣: {e}")

    texts = st.session_state.wr_texts
    print_button()

    # ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧 01. 二쇨컙蹂닿퀬 ?붿빟 ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
    if page == "01. 二쇨컙蹂닿퀬 ?붿빟":
        st.markdown(f"## 泥リ뎄留?二쇨컙蹂닿퀬 ??{ref_year}??{ref_month}??)
        wy, wlabel = (ref_year, ref_week) if ref_week else latest_period(df, "二?)
        if wlabel:
            st.caption(f"湲곗? 二쇱감: {week_disp(wy, wlabel)}")

        # KPI 移대뱶 (湲곗? 二쇱감 ???꾩＜鍮꽷룹쟾?붾퉬쨌?꾨뀈鍮?紐⑤몢 二쇱감 湲곗?)
        if wlabel:
            cols = st.columns(4)
            kpi_metrics = ["泥リ뎄留?嫄곕옒??, "泥リ뎄留?怨좉컼??, "泥リ뎄留?媛앸떒媛", "媛?낆옄??]
            py, plb = prev_label(df, "二?, wy, wlabel)
            # ?꾩썡 ?숈씪 二쇱감 (?? 06??1二쇱감 ??05??1二쇱감)
            mom_y = mom_lbl = None
            wm = re.match(r"(\d{1,2})??(\d)二쇱감", wlabel)
            if wm:
                mo, wk = int(wm.group(1)), int(wm.group(2))
                mom_y, mom_m = (wy, mo - 1) if mo > 1 else (wy - 1, 12)
                mom_lbl = f"{mom_m:02d}??{wk}二쇱감"
            for col, met in zip(cols, kpi_metrics):
                cur = pick(df, "二?, met, "*TOTAL", wy, wlabel, "mtd")
                deltas = []
                prv = pick(df, "二?, met, "*TOTAL", py, plb, "final") if plb else np.nan
                deltas.append((fmt_delta(met, cur, prv), "?꾩＜鍮?))
                if mom_lbl:
                    mom = pick(df, "二?, met, "*TOTAL", mom_y, mom_lbl, "final")
                    deltas.append((fmt_delta(met, cur, mom), "?꾩썡鍮?))
                yoy = pick(df, "二?, met, "*TOTAL", wy - 1, wlabel, "final")
                deltas.append((fmt_delta(met, cur, yoy), "?꾨뀈鍮?))
                pills = ""
                for d, lab in deltas:
                    if d:
                        neg = d.startswith("??)
                        cls = "down" if neg else "up"
                        prefix = "" if neg else "??"
                        pills += f'<div class="kpi-delta {cls}">{prefix}{d} ({lab})</div>'
                    else:
                        pills += f'<div class="kpi-delta na">??({lab})</div>'
                col.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">{met} ({week_disp(wy, wlabel)})</div>'
                    f'<div class="kpi-value">{fmt_value(met, cur)}</div>{pills}</div>',
                    unsafe_allow_html=True)
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # ?ㅼ쟻 ?붿빟 ???꾩＜鍮?(二쇱감 湲곗?)
        if wlabel:
            st.subheader("?ㅼ쟻 ?붿빟 (?쇳룊洹?쨌 ?꾩＜鍮?")
            st.caption(f"湲곗? 二쇱감: {week_disp(wy, wlabel)} ???꾩＜쨌?꾨뀈 ?숈＜ ?鍮?)
            st.dataframe(style_delta_cols(wow_summary_table(df, wy, wlabel, METRICS7)),
                         use_container_width=True)
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # ?ㅼ쟻 ?붿빟 YoY ??
        st.subheader("?ㅼ쟻 ?붿빟 (?쇳룊洹?쨌 ?꾨뀈鍮?")
        tbl, (pm_y, pm_m) = yoy_summary_table(df, ref_year, ref_month, METRICS7)
        st.caption(f"?꾩썡({pm_m}??? ?붾쭏媛? ?뱀썡({ref_month}??? ?쇰쭏媛?MTD) 湲곗? ?숈씪湲곌컙 鍮꾧탳")
        st.dataframe(style_delta_cols(tbl), use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 蹂닿퀬?
        draft = auto_draft(df, ref_year, ref_month, ref_week=wlabel)
        ai_model = st.session_state.get("wr_ai_model", "claude-sonnet-4-6")
        cL, cR = st.columns(2)
        with cL:
            report_text_block("wr_metrics_summary", "?꾩＜ 二쇱슂 吏???꾪솴",
                              default=draft, regen=draft,
                              ai_fn=lambda memo: ai_generate_insight(df, ref_year, ref_month,
                                                                wlabel, ai_model, memo=memo))
        with cR:
            report_text_block("wr_exec_summary", "湲덉＜ 吏묓뻾 ?댁슜 ?붿빟")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # ?듭떖 李⑦듃 (二쇱감蹂?YoY 3醫?
        st.subheader("二쇱감蹂?異붿씠 ???꾨뀈 鍮꾧탳")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["泥リ뎄留?嫄곕옒??, "泥リ뎄留?怨좉컼??, "泥リ뎄留?媛앸떒媛"]):
            with col:
                st.plotly_chart(yoy_chart(df, "二?, met, chart_years, h=280),
                                use_container_width=True)

    # ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧 02. ?붾퀎 異붿씠 ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
    elif page == "02. ?붾퀎 異붿씠":
        st.markdown("## ?붾퀎 異붿씠")
        st.subheader("?붾퀎 異붿씠 李⑦듃 ???꾨뀈 鍮꾧탳")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["泥リ뎄留?嫄곕옒??, "泥リ뎄留?怨좉컼??, "泥リ뎄留?媛앸떒媛"]):
            with col:
                st.plotly_chart(yoy_chart(df, "??, met, chart_years, h=280),
                                use_container_width=True)
        c4, c5, c6 = st.columns(3)
        for col, met in zip((c4, c5, c6), ["鍮꾪쉶?먰듃?섑뵿", "媛?낆옄??, "媛?낆쑉"]):
            with col:
                st.plotly_chart(yoy_chart(df, "??, met, chart_years, h=280),
                                use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("?붾퀎 異붿씠??(?쇳룊洹?")
        tbl = trend_table(df, "??, METRICS7, chart_years)
        st.dataframe(style_trend(tbl, METRICS7), use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        ai_model = st.session_state.get("wr_ai_model", "claude-sonnet-4-6")
        report_text_block(
            f"wr_month_memo_{ref_year}_{ref_month}",
            f"{ref_year}??{ref_month}???≪뀡쨌?댁뒋?ы빆",
            ai_fn=lambda memo: ai_generate_insight(df, ref_year, ref_month, None, ai_model,
                                              focus=f"{ref_year}??{ref_month}???≪뀡쨌?댁뒋 諛??몄궗?댄듃",
                                              memo=memo))

    # ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧 03. 二쇱감蹂?異붿씠 ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
    elif page == "03. 二쇱감蹂?異붿씠":
        st.markdown("## 二쇱감蹂?異붿씠")
        st.subheader("二쇱감蹂?異붿씠 李⑦듃 ???꾨뀈 鍮꾧탳")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["泥リ뎄留?嫄곕옒??, "泥リ뎄留?怨좉컼??, "泥リ뎄留?媛앸떒媛"]):
            with col:
                st.plotly_chart(yoy_chart(df, "二?, met, chart_years, h=280),
                                use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"二쇱감蹂?異붿씠????{ref_year}??)
        tbl = trend_table(df, "二?, METRICS7, [ref_year])
        if not tbl.empty:
            recent = tbl.columns[-16:]
            st.dataframe(style_trend(tbl[recent], METRICS7), use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("?꾩＜鍮?WoW)쨌?꾨뀈鍮?YoY) 利앷컧")
        wy, wlabel = (ref_year, ref_week) if ref_week else latest_period(df, "二?)
        if wlabel:
            st.caption(f"湲곗? 二쇱감: {week_disp(wy, wlabel)}")
            st.dataframe(style_delta_cols(wow_summary_table(df, wy, wlabel, METRICS7)),
                         use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        wy2, wlabel2 = (ref_year, ref_week) if ref_week else latest_period(df, "二?)
        ai_model = st.session_state.get("wr_ai_model", "claude-sonnet-4-6")
        report_text_block(
            f"wr_week_memo_{wy2}_{wlabel2}",
            f"{wy2}??{wlabel2} ?≪뀡쨌?댁뒋?ы빆" if wlabel2 else "二쇱감蹂??≪뀡쨌?댁뒋?ы빆",
            ai_fn=lambda memo: ai_generate_insight(df, ref_year, ref_month, wlabel2, ai_model,
                                              focus=f"{wlabel2} 二쇱감 ?≪뀡쨌?댁뒋 諛??몄궗?댄듃",
                                              memo=memo))

    # ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧 04. 梨꾨꼸蹂??ㅼ쟻 ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
    elif page == "04. 梨꾨꼸蹂??ㅼ쟻":
        st.markdown("## 梨꾨꼸(BPU)蹂??ㅼ쟻")
        avail = [m for m in METRICS7 if (df["metric"] == m).any()]
        met = st.selectbox("吏???좏깮", avail, key="wr_chmet")

        st.subheader(f"{met} ??{ref_year}??{ref_month}??梨꾨꼸蹂??꾨뀈鍮?)
        rows = []
        for seg in ["*TOTAL"] + [c for c in CHANNELS if c in ch_sel]:
            pv = pick(df, "??, met, seg, ref_year - 1, month_label(ref_month), "mtd")
            cv = pick(df, "??, met, seg, ref_year, month_label(ref_month), "mtd")
            rows.append({"梨꾨꼸": seg,
                         f"{ref_year-1}??{ref_month}??: fmt_value(met, pv),
                         f"{ref_year}??{ref_month}??: fmt_value(met, cv),
                         "?꾨뀈鍮?: fmt_delta(met, cv, pv) or "??})
        st.dataframe(style_delta_cols(pd.DataFrame(rows).set_index("梨꾨꼸")),
                     use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"{met} ??梨꾨꼸蹂???異붿씠 ({ref_year}??")
        unit, div = METRIC_UNIT.get(met, ("", 1))
        if met in PCT_METRICS: div, unit = 0.01, "%"
        fig = go.Figure()
        x = [month_label(i) for i in range(1, 13)]
        for seg in [c for c in CHANNELS if c in ch_sel]:
            s = series_by_label(df, "??, met, seg, ref_year).reindex(x).dropna()
            if s.empty: continue
            fig.add_trace(go.Scatter(
                x=s.index.tolist(), y=(s / div).tolist(), mode="lines+markers", name=seg,
                line=dict(color=clr(CHANNEL_PAL.get(seg, "blue")), width=1.8),
                marker=dict(size=4)))
        ly = base_layout(340, ysuffix=unit if unit == "%" else "",
                         title=f"{met} 梨꾨꼸蹂?({unit})")
        fig.update_layout(**ly)
                            st.plotly_chart(fig, use_container_width=True)

            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
            st.subheader("월별 추이표 (합계 및 평균 동의율)")
            
            # 일별 데이터를 월별 합계로 집계하여 전년비 표 생성
            push_metrics = ["가입자수", "앱푸시수신동의", "동의율"]
            sub = df[(df["gran"] == "일") & (df["year"].isin(chart_years)) & (df["metric"].isin(["가입자수", "앱푸시수신동의"]))].copy()
            
            if not sub.empty:
                # 4/24 이상치 제거 반영 (집계 시에도 제외되도록)
                glitch_dates = ["4/24", "04/24", "2026/04/24", "2026-04-24", "4-24", "04-24"]
                sub = sub[~sub["label"].astype(str).str.strip().isin(glitch_dates)]
                
                sub["month"] = sub["label"].apply(lambda x: int(str(x).split('/')[0]) if '/' in str(x) else 0)
                grp = sub.groupby(["year", "month", "metric"])["value"].sum().unstack("metric").reset_index()
                
                if "가입자수" in grp.columns and "앱푸시수신동의" in grp.columns:
                    grp["동의율"] = grp["앱푸시수신동의"] / grp["가입자수"]
                    grp.loc[grp["동의율"] > 1.0, "동의율"] = np.nan
                else:
                    grp["동의율"] = np.nan
                    
                cols = []
                out_tbl = {}
                for y in chart_years:
                    for m in range(1, 13):
                        row = grp[(grp["year"] == y) & (grp["month"] == m)]
                        if not row.empty:
                            cols.append((y, f"{m}월"))
                            for met in push_metrics:
                                if met not in out_tbl: out_tbl[met] = []
                                out_tbl[met].append(row[met].values[0] if met in row.columns else np.nan)
                                
                if cols:
                    tbl_df = pd.DataFrame(out_tbl, index=pd.MultiIndex.from_tuples(cols, names=["연도", "기간"])).T
                    st.dataframe(style_trend(tbl_df, push_metrics), use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"{met} ??梨꾨꼸 횞 ????({ref_year}??")
        rows = []
        for seg in ["*TOTAL"] + [c for c in CHANNELS if c in ch_sel]:
            s = series_by_label(df, "??, met, seg, ref_year)
            row = {"梨꾨꼸": seg}
            for lb in [month_label(i) for i in range(1, 13)]:
                if lb in s.index and not np.isnan(s[lb]):
                    row[lb] = fmt_value(met, s[lb])
            rows.append(row)
        st.dataframe(pd.DataFrame(rows).set_index("梨꾨꼸"), use_container_width=True)

    # ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧 05. ?뚯썝 ?ㅼ쟻 ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
    elif page == "05. ?뚯썝 ?ㅼ쟻":
        render_member_page(mdf, chart_years)

    # ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧 06. ?듯빀 ?곗씠?걔룸떎?대줈???먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
    elif page == "06. ?듯빀 ?곗씠?걔룸떎?대줈??:
        st.markdown("## ?듯빀 ?곗씠??쨌 ?ㅼ슫濡쒕뱶")
        st.caption("?낅줈?쒗븳 紐⑤뱺 ?뚯씪???⑹튇 ?듯빀 long ?곗씠?곗엯?덈떎.")
        st.dataframe(df.sort_values(["gran", "metric", "segment", "sortkey"]).head(2000),
                     use_container_width=True, height=420)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("?듯빀 ?뚰겕遺??ㅼ슫濡쒕뱶")
        st.caption("`泥リ뎄留??붿빟`(?붿빟?쑣룸낫怨좊?쨌YoY 李⑦듃쨌梨꾨꼸?? + `??쨌`二?(?듯빀 ?곗씠?? 3媛??쒗듃")
        if st.button("?뱿 ?묒? ?뚰겕遺??앹꽦", key="wr_build"):
            with st.spinner("?뚰겕遺??앹꽦 以묅?):
                xls = build_workbook(df, st.session_state.wr_texts,
                                     ref_year, ref_month, chart_years)
            st.download_button(
                "?ㅼ슫濡쒕뱶 ??泥リ뎄留?二쇨컙蹂닿퀬.xlsx", xls,
                file_name=f"泥リ뎄留?二쇨컙蹂닿퀬_{ref_year}{ref_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("?듯빀 long ?곗씠??CSV", csv, "?듯빀?곗씠??csv", "text/csv")
        if not mdf.empty:
            st.download_button("?뚯썝 ?ㅼ쟻 ?곗씠??CSV",
                               mdf.to_csv(index=False).encode("utf-8-sig"),
                               "?뚯썝?ㅼ쟻?곗씠??csv", "text/csv")

    # ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧 07. ?깊뫖???숈쓽 ?꾪솴 ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
    elif page == "07. ?깊뫖???숈쓽 ?꾪솴":
        st.markdown("## ?깊뫖???숈쓽 ?꾪솴 (?쇱옄蹂?")
        df_daily = df[df["gran"] == "??]
        
        if df_daily.empty:
            st.info("?쇱옄蹂??곗씠?곌? ?놁뒿?덈떎. PUSH(7) 諛???媛?낆옄???뚯씪???낅줈?쒗빐二쇱꽭??")
        else:
            st.subheader(f"{ref_year}???좉퇋媛???鍮??깊뫖???섏떊?숈쓽??異붿씠")
            dates = labels_sorted(df, "??, [ref_year])
            
            if not dates:
                st.info(f"{ref_year}???쇱옄蹂??곗씠?곌? ?놁뒿?덈떎.")
            else:
                rows = []
                for d in dates:
                    push_val = pick(df, "??, "?깊뫖?쒖닔?좊룞??, "*TOTAL", ref_year, d, "final")
                    join_val = pick(df, "??, "媛?낆옄??, "*TOTAL", ref_year, d, "final")
                    
                    # 4/24 ?댁긽移??섎뱶肄붾뵫 ?쒓굅 (鍮꾩젙???ㅽ뙆?댄겕)
                    if str(d).strip() in ["4/24", "04/24", "2026/04/24", "2026-04-24", "4-24", "04-24"]:
                        push_val = np.nan
                        
                    rate = np.nan
                    if not np.isnan(push_val) and not np.isnan(join_val) and join_val > 0:
                        rate = push_val / join_val
                        # ?댁긽移??쒓굅: 100% 珥덇낵 ???곗씠??湲由ъ튂濡?媛꾩＜
                        if rate > 1.0:
                            rate = np.nan
                        
                    rows.append({
                        "?좎쭨": d,
                        "?깊뫖?쒖닔?좊룞??: push_val,
                        "媛?낆옄??: join_val,
                        "?숈쓽??: rate
                    })
                
                res_df = pd.DataFrame(rows)
                
                fig = go.Figure()
                fig.add_trace(go.Bar(x=res_df["?좎쭨"], y=res_df["媛?낆옄??], name="媛?낆옄??, marker_color=clr("slate"), opacity=0.6, yaxis="y1"))
                fig.add_trace(go.Bar(x=res_df["?좎쭨"], y=res_df["?깊뫖?쒖닔?좊룞??], name="?깊뫖?쒖닔?좊룞??, marker_color=clr("blue"), opacity=0.8, yaxis="y1"))
                fig.add_trace(go.Scatter(x=res_df["?좎쭨"], y=res_df["?숈쓽??]*100, name="?숈쓽??%)", mode="lines+markers", line=dict(color=clr("red"), width=2), yaxis="y2"))
                
                max_rate = max(res_df["?숈쓽??].dropna()*100, default=10)
                if pd.isna(max_rate) or max_rate == 0:
                    max_rate = 10
                
                fig.update_layout(
                    title="?쇱옄蹂??좉퇋媛??諛??깊뫖???섏떊?숈쓽 ?꾪솴",
                    xaxis=dict(type="category", tickangle=-45),
                    yaxis=dict(title="紐?, gridcolor="#f1f5f9"),
                    yaxis2=dict(title="%", overlaying="y", side="right", range=[0, max_rate * 1.2], gridcolor="rgba(0,0,0,0)"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, bgcolor="rgba(255,255,255,0.8)"),
                    barmode="group",
                    height=450,
                    margin=dict(l=10, r=10, t=40, b=10),
                    paper_bgcolor="rgba(248,249,252,0)", plot_bgcolor="rgba(248,249,252,0)"
                )
                
                                    st.plotly_chart(fig, use_container_width=True)

            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
            st.subheader("월별 추이표 (합계 및 평균 동의율)")
            
            # 일별 데이터를 월별 합계로 집계하여 전년비 표 생성
            push_metrics = ["가입자수", "앱푸시수신동의", "동의율"]
            sub = df[(df["gran"] == "일") & (df["year"].isin(chart_years)) & (df["metric"].isin(["가입자수", "앱푸시수신동의"]))].copy()
            
            if not sub.empty:
                # 4/24 이상치 제거 반영 (집계 시에도 제외되도록)
                glitch_dates = ["4/24", "04/24", "2026/04/24", "2026-04-24", "4-24", "04-24"]
                sub = sub[~sub["label"].astype(str).str.strip().isin(glitch_dates)]
                
                sub["month"] = sub["label"].apply(lambda x: int(str(x).split('/')[0]) if '/' in str(x) else 0)
                grp = sub.groupby(["year", "month", "metric"])["value"].sum().unstack("metric").reset_index()
                
                if "가입자수" in grp.columns and "앱푸시수신동의" in grp.columns:
                    grp["동의율"] = grp["앱푸시수신동의"] / grp["가입자수"]
                    grp.loc[grp["동의율"] > 1.0, "동의율"] = np.nan
                else:
                    grp["동의율"] = np.nan
                    
                cols = []
                out_tbl = {}
                for y in chart_years:
                    for m in range(1, 13):
                        row = grp[(grp["year"] == y) & (grp["month"] == m)]
                        if not row.empty:
                            cols.append((y, f"{m}월"))
                            for met in push_metrics:
                                if met not in out_tbl: out_tbl[met] = []
                                out_tbl[met].append(row[met].values[0] if met in row.columns else np.nan)
                                
                if cols:
                    tbl_df = pd.DataFrame(out_tbl, index=pd.MultiIndex.from_tuples(cols, names=["연도", "기간"])).T
                    st.dataframe(style_trend(tbl_df, push_metrics), use_container_width=True)
                
                st.markdown("### ?곸꽭 ?곗씠??)
                disp_df = res_df.copy()
                disp_df["?깊뫖?쒖닔?좊룞??] = disp_df["?깊뫖?쒖닔?좊룞??].apply(lambda x: f"{int(x):,}" if not pd.isna(x) else "??)
                disp_df["媛?낆옄??] = disp_df["媛?낆옄??].apply(lambda x: f"{int(x):,}" if not pd.isna(x) else "??)
                disp_df["?숈쓽??] = disp_df["?숈쓽??].apply(lambda x: f"{x*100:.1f}%" if not pd.isna(x) else "??)
                st.dataframe(disp_df.set_index("?좎쭨"), use_container_width=True)


if st.runtime.exists():
    main()

