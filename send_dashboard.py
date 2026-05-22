import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from scipy import stats
import io, json, os, datetime

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="발송 분석 대시보드", page_icon="📨",
                   layout="wide", initial_sidebar_state="expanded")

INSIGHT_FILE = "insights_store.json"

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.vg{background:rgba(72,187,120,.08);border-left:3px solid #48bb78;border-radius:0 8px 8px 0;padding:12px 16px;margin:8px 0;color:#86efac;font-size:13px;line-height:1.65}
.vr{background:rgba(245,101,101,.08);border-left:3px solid #f56565;border-radius:0 8px 8px 0;padding:12px 16px;margin:8px 0;color:#fca5a5;font-size:13px;line-height:1.65}
.va{background:rgba(237,137,54,.08);border-left:3px solid #ed8936;border-radius:0 8px 8px 0;padding:12px 16px;margin:8px 0;color:#fcd34d;font-size:13px;line-height:1.65}
.vb{background:rgba(79,143,255,.08);border-left:3px solid #4f8fff;border-radius:0 8px 8px 0;padding:12px 16px;margin:8px 0;color:#93c5fd;font-size:13px;line-height:1.65}
.sdiv{border-top:1px solid #e2e8f0;margin:22px 0}
.stat-label{font-size:11px;color:#545c6a;margin-bottom:3px;font-weight:500;letter-spacing:.04em;text-transform:uppercase}
.appendix{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin-top:12px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 색상
# ══════════════════════════════════════════════════════
PALETTE = {
    "blue":   ("rgba(79,143,255,1)",  "rgba(79,143,255,0.15)"),
    "red":    ("rgba(245,101,101,1)", "rgba(245,101,101,0.15)"),
    "green":  ("rgba(72,187,120,1)",  "rgba(72,187,120,0.15)"),
    "amber":  ("rgba(237,137,54,1)",  "rgba(237,137,54,0.15)"),
    "purple": ("rgba(159,122,234,1)", "rgba(159,122,234,0.15)"),
    "teal":   ("rgba(56,178,172,1)",  "rgba(56,178,172,0.15)"),
    "pink":   ("rgba(236,72,153,1)",  "rgba(236,72,153,0.15)"),
}
def clr(n): return PALETTE[n][0]
def cbg(n): return PALETTE[n][1]

def base_layout(h=280, ysuffix="", title=""):
    return dict(paper_bgcolor="rgba(248,249,252,0)", plot_bgcolor="rgba(248,249,252,0)",
                font=dict(color="#475569",size=11), margin=dict(l=10,r=10,t=36,b=10),
                height=h, showlegend=False, title=dict(text=title,font=dict(color="#c0c8d8",size=13)),
                xaxis=dict(gridcolor="rgba(0,0,0,0)",linecolor="#e2e8f0",
                           tickfont=dict(color="#64748b",size=11)),
                yaxis=dict(gridcolor="#f1f5f9",linecolor="#e2e8f0",
                           tickfont=dict(color="#64748b",size=11),ticksuffix=ysuffix))

# ══════════════════════════════════════════════════════
# 인사이트 저장소
# ══════════════════════════════════════════════════════
def load_insights():
    if os.path.exists(INSIGHT_FILE):
        with open(INSIGHT_FILE,"r",encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_insights(data):
    with open(INSIGHT_FILE,"w",encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if "insights" not in st.session_state:
    st.session_state.insights = load_insights()

# ══════════════════════════════════════════════════════
# 통계 헬퍼
# ══════════════════════════════════════════════════════
def linreg(x, y):
    mask = ~np.isnan(x.astype(float)) & ~np.isnan(y.astype(float))
    if mask.sum() < 5: return dict(slope=np.nan, r2=np.nan, p=np.nan)
    sl, _, r, p, _ = stats.linregress(x[mask], y[mask])
    return dict(slope=sl, r2=r**2, p=p)

def dow_residual(df, col):
    vals = df[col].values.astype(float)
    res  = vals.copy()
    for d in range(7):
        idx = df["dow"].values == d
        if idx.sum() > 0: res[idx] -= np.nanmean(res[idx])
    return res

def sig_label(p):
    if np.isnan(p): return "–"
    if p < 0.001: return "p<0.001 ★★★"
    if p < 0.01:  return "p<0.01 ★★"
    if p < 0.05:  return "p<0.05 ★"
    return "유의하지 않음 (ns)"

def pct(a, b):
    return (b-a)/a*100 if a and not np.isnan(a) else 0

# ══════════════════════════════════════════════════════
# 파싱
# ══════════════════════════════════════════════════════
ALL_METRICS = {
    "perSend":     ["인당발송건수"],
    "revenue":     ["거래액"],
    "rps":         ["발송건당거래액"],
    "totalSend":   ["총발송건수"],
    "customers":   ["유니크발송고객수"],
    "ctr":         ["CTR"],
    "uniqueInflow":["유니크유입"],
    "totalInflow": ["총유입"],
    "visitPerPerson":["인당방문횟수"],
    "purchaseCust":["구매고객수"],
    "purchaseCnt": ["구매건수"],
    "purchasePerPerson":["인당구매건수"],
    "avgOrderVal": ["객단가"],
    "unitPrice":   ["건단가"],
    "mRevenue":    ["M당거래액"],
    "pointM":      ["적립M"],
}

METRIC_LABELS = {
    "perSend":"인당 발송 건수","revenue":"거래액","rps":"발송건당 거래액",
    "totalSend":"총 발송 건수","customers":"유니크 발송 고객수","ctr":"CTR",
    "uniqueInflow":"유니크 유입","totalInflow":"총 유입",
    "visitPerPerson":"인당 방문 횟수","purchaseCust":"구매 고객수",
    "purchaseCnt":"구매 건수","purchasePerPerson":"인당 구매 건수",
    "avgOrderVal":"객단가","unitPrice":"건단가",
    "mRevenue":"M당 거래액","pointM":"적립M",
    "purchaseRate":"구매전환율(CR)","rpc":"고객당 매출",
}

METRIC_FORMAT = {
    "perSend":("{:.2f}건","건"), "revenue":("{:.0f}원","억원"),
    "rps":("{:.0f}원","원"), "totalSend":("{:.0f}건","M건"),
    "customers":("{:.0f}명","만명"), "ctr":("{:.2%}","%"),
    "uniqueInflow":("{:.0f}명","만명"), "totalInflow":("{:.0f}명","만명"),
    "visitPerPerson":("{:.2f}회","회"), "purchaseCust":("{:.0f}명","명"),
    "purchaseCnt":("{:.0f}건","건"), "purchasePerPerson":("{:.3f}건","건"),
    "avgOrderVal":("{:.0f}원","원"), "unitPrice":("{:.0f}원","원"),
    "mRevenue":("{:.0f}원","원"), "pointM":("{:.0f}",""),
    "purchaseRate":("{:.3%}","%"), "rpc":("{:.0f}원","원"),
}

@st.cache_data(show_spinner=False)
def parse_xlsx(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    ws  = pd.read_excel(xls, header=None, sheet_name=xls.sheet_names[0])
    date_row   = ws.iloc[1, 2:]
    metric_col = ws.iloc[3:, 0].astype(str).str.strip()

    dates = []
    for v in date_row:
        if pd.isnull(v): continue
        try:
            d = pd.Timestamp("1899-12-30")+pd.Timedelta(days=int(v)) if isinstance(v,(int,float)) else pd.Timestamp(v)
            dates.append(d)
        except: dates.append(pd.NaT)

    def get_m(keywords):
        for kw in keywords:
            match = metric_col[metric_col.str.replace(" ","") == kw.replace(" ","")]
            if not match.empty:
                vals = ws.iloc[match.index[0], 2:2+len(dates)].values
                return pd.to_numeric(vals, errors="coerce")
        return np.full(len(dates), np.nan)

    data = {k: get_m(v) for k, v in ALL_METRICS.items()}
    data["date"] = dates
    df = pd.DataFrame(data).dropna(subset=["perSend","revenue"]).copy()
    df["date"]    = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)
    df["t"]       = np.arange(len(df))
    df["dow"]     = df["date"].dt.dayofweek
    df["month"]   = df["date"].dt.to_period("M").astype(str)
    df["quarter"] = df["date"].dt.to_period("Q").astype(str)
    df["year"]    = df["date"].dt.year
    df["week"]    = df["date"].dt.isocalendar().week.astype(int)
    # 파생 지표
    df["purchaseRate"] = df["purchaseCust"] / df["customers"]
    df["rpc"]          = df["revenue"] / df["customers"]
    return df

@st.cache_data(show_spinner=False)
def compute(file_bytes):
    df = parse_xlsx(file_bytes)

    # Aggregations
    monthly   = df.groupby("month",   sort=True).agg(n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}).reset_index()
    quarterly = df.groupby("quarter", sort=True).agg(n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}).reset_index()
    yearly    = df.groupby("year",    sort=True).agg(n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}).reset_index()

    # Buckets
    BINS = [0,2.0,2.5,3.0,3.5,4.0,4.5,99]
    LBLS = ["~2.0건","2.0~2.5","2.5~3.0","3.0~3.5","3.5~4.0","4.0~4.5","4.5건+"]
    df["bucket"] = pd.cut(df["perSend"], bins=BINS, labels=LBLS)
    buckets = df.groupby("bucket", observed=True).agg(n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}).reset_index()
    buckets = buckets[buckets["n"] >= 30].reset_index(drop=True)

    # Quintile
    df_s = df.sort_values("totalSend").reset_index(drop=True)
    sz   = len(df_s)//5
    quintile = pd.DataFrame([
        dict(label=["Q1 최소","Q2","Q3","Q4","Q5 최대"][i],
             **{k: df_s.iloc[i*sz:(i+1)*sz][k].mean() for k in list(ALL_METRICS)+["purchaseRate","rpc"]})
        for i in range(5)
    ])

    # DoW stats
    DOW = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}
    dow_comp = []
    for d in [0,1,2,3,4]:
        sub = df[df["dow"]==d]
        if len(sub)<20: continue
        med = sub["perSend"].median()
        lo, hi = sub[sub["perSend"]<=med], sub[sub["perSend"]>med]
        dow_comp.append(dict(dow=DOW[d], median=med,
            lowRev=lo["revenue"].mean(), highRev=hi["revenue"].mean(),
            diff=lo["revenue"].mean()-hi["revenue"].mean(),
            lowRps=lo["rps"].mean(), highRps=hi["rps"].mean(),
            lowCtr=lo["ctr"].mean(), highCtr=hi["ctr"].mean()))
    dow_corr = []
    for d in range(7):
        sub = df[df["dow"]==d]
        if len(sub)<20: continue
        dow_corr.append(dict(dow=DOW[d],
            corrRevenue=float(np.corrcoef(sub["perSend"],sub["revenue"])[0,1]),
            corrRps=float(np.corrcoef(sub["perSend"],sub["rps"])[0,1])))

    # Regression
    t = df["t"].values.astype(float)
    reg = {k: linreg(t, dow_residual(df,k) if k not in ["perSend"] else df[k].values)
           for k in ["perSend","ctr","purchaseRate","rps","revenue"]}

    meta = dict(start=str(df["date"].min().date()), end=str(df["date"].max().date()), days=len(df))
    return dict(df=df, monthly=monthly, quarterly=quarterly, yearly=yearly,
                buckets=buckets, quintile=quintile,
                dow_comp=dow_comp, dow_corr=dow_corr, reg=reg, meta=meta)

# ══════════════════════════════════════════════════════
# 인사이트 에디터
# ══════════════════════════════════════════════════════
def insight_editor(page_key, default_lines):
    store = st.session_state.insights
    if page_key not in store:
        store[page_key] = [{"text": t, "color": "vb", "bold": False} for t in default_lines]

    items = store[page_key]
    st.markdown("#### 💡 인사이트 메모 <span style='font-size:11px;color:#545c6a'>(클릭 편집 · 서버 자동 저장)</span>", unsafe_allow_html=True)

    COLOR_OPTIONS = {"파란색(vb)":"vb","초록색(vg)":"vg","빨간색(vr)":"vr","주황색(va)":"va"}
    to_delete = []

    for i, item in enumerate(items):
        with st.expander(f"{'**' if item.get('bold') else ''}{item['text'][:60]}{'...' if len(item['text'])>60 else ''}{'**' if item.get('bold') else ''}", expanded=False):
            col1, col2, col3 = st.columns([5,2,1])
            new_text = col1.text_area("내용", item["text"], key=f"ins_txt_{page_key}_{i}", height=80, label_visibility="collapsed")
            new_color = col2.selectbox("색상", list(COLOR_OPTIONS.keys()),
                index=list(COLOR_OPTIONS.values()).index(item.get("color","vb")),
                key=f"ins_col_{page_key}_{i}", label_visibility="collapsed")
            new_bold = col2.checkbox("볼드", value=item.get("bold",False), key=f"ins_bold_{page_key}_{i}")
            if col3.button("🗑", key=f"ins_del_{page_key}_{i}"):
                to_delete.append(i)
            items[i] = {"text":new_text, "color":COLOR_OPTIONS[new_color], "bold":new_bold}

        clss = items[i]["color"]
        txt  = f"<strong>{items[i]['text']}</strong>" if items[i]["bold"] else items[i]["text"]
        st.markdown(f"<div class='{clss}'>{txt}</div>", unsafe_allow_html=True)

    for idx in sorted(to_delete, reverse=True):
        items.pop(idx)

    if st.button("＋ 인사이트 추가", key=f"ins_add_{page_key}"):
        items.append({"text":"새 인사이트를 입력하세요.", "color":"vb", "bold":False})
        st.rerun()

    store[page_key] = items
    st.session_state.insights = store
    save_insights(store)


# ══════════════════════════════════════════════════════
# 편집 가능 텍스트 시스템
# ══════════════════════════════════════════════════════
if "editable_texts" not in st.session_state:
    st.session_state.editable_texts = load_insights().get("__texts__", {})

def editable_text(key, default, tag="p", style=""):
    """클릭하면 인라인 편집 가능한 텍스트"""
    texts = st.session_state.editable_texts
    if key not in texts:
        texts[key] = default
    edit_key = f"__editing_{key}__"
    if edit_key not in st.session_state:
        st.session_state[edit_key] = False

    if st.session_state[edit_key]:
        col1, col2 = st.columns([10,1])
        new_val = col1.text_area("", texts[key], key=f"ta_{key}", height=80, label_visibility="collapsed")
        if col2.button("✓", key=f"save_{key}"):
            texts[key] = new_val
            st.session_state[edit_key] = False
            # 저장
            all_data = load_insights()
            all_data["__texts__"] = texts
            save_insights(all_data)
            st.session_state.editable_texts = texts
            st.rerun()
    else:
        rendered = texts[key]
        st.markdown(f'<{tag} style="cursor:pointer;{style}" title="클릭하여 편집">{rendered}</{tag}>', unsafe_allow_html=True)
        if st.button("✏️", key=f"edit_{key}", help="편집"):
            st.session_state[edit_key] = True
            st.rerun()
    return texts[key]

def verdict_box(key, default_text, default_color="vg", emoji=""):
    """편집 가능한 verdict 박스"""
    store = st.session_state.insights
    vkey  = f"__verdict_{key}__"
    ckey  = f"__vcolor_{key}__"
    ekey  = f"__editing_v_{key}__"

    if vkey not in store: store[vkey] = default_text
    if ckey not in store: store[ckey] = default_color
    if ekey not in st.session_state: st.session_state[ekey] = False

    COLOR_MAP = {
        "초록(✅)":"vg","빨강(❌)":"vr","주황(⚠️)":"va","파랑(ℹ️)":"vb"
    }
    COLOR_REVERSE = {"vg":"초록(✅)","vr":"빨강(❌)","va":"주황(⚠️)","vb":"파랑(ℹ️)"}

    if st.session_state[ekey]:
        c1,c2,c3 = st.columns([6,2,1])
        new_text  = c1.text_area("내용 편집", store[vkey], key=f"vta_{key}", height=100, label_visibility="collapsed")
        cur_color = COLOR_REVERSE.get(store[ckey],"초록(✅)")
        new_color = c2.selectbox("색상", list(COLOR_MAP.keys()),
                                  index=list(COLOR_MAP.keys()).index(cur_color),
                                  key=f"vcol_{key}", label_visibility="collapsed")
        if c3.button("✓", key=f"vsave_{key}"):
            store[vkey]  = new_text
            store[ckey]  = COLOR_MAP[new_color]
            st.session_state[ekey] = False
            all_data = load_insights()
            all_data.update(store)
            save_insights(all_data)
            st.session_state.insights = store
            st.rerun()
    else:
        clss = store[ckey]
        txt  = store[vkey].replace("\n","<br>")
        col_a, col_b = st.columns([20,1])
        col_a.markdown(f'<div class="{clss}">{txt}</div>', unsafe_allow_html=True)
        if col_b.button("✏️", key=f"vedit_{key}", help="편집"):
            st.session_state[ekey] = True
            st.rerun()


# ══════════════════════════════════════════════════════
# 기간 비교 헬퍼
# ══════════════════════════════════════════════════════
COMPARE_OPTS = ["전년비 (YoY)", "전분기비 (QoQ)", "전월비 (MoM)", "전주비 (WoW)", "전체 기간 처음↔끝"]

def get_compare_periods(df_all, mode):
    """비교 기준 기간과 현재 기간 반환 → (df_current, df_prev, label_cur, label_prev)"""
    df_all = df_all.sort_values("date")
    last_date = df_all["date"].max()

    if mode == "전년비 (YoY)":
        cur  = df_all[df_all["date"].dt.year == last_date.year]
        prev = df_all[df_all["date"].dt.year == last_date.year - 1]
        return cur, prev, str(last_date.year)+"년", str(last_date.year-1)+"년"
    elif mode == "전분기비 (QoQ)":
        cur_q  = pd.Period(last_date, "Q")
        prev_q = cur_q - 1
        cur  = df_all[df_all["date"].dt.to_period("Q") == cur_q]
        prev = df_all[df_all["date"].dt.to_period("Q") == prev_q]
        return cur, prev, str(cur_q), str(prev_q)
    elif mode == "전월비 (MoM)":
        cur_m  = pd.Period(last_date, "M")
        prev_m = cur_m - 1
        cur  = df_all[df_all["date"].dt.to_period("M") == cur_m]
        prev = df_all[df_all["date"].dt.to_period("M") == prev_m]
        return cur, prev, str(cur_m), str(prev_m)
    elif mode == "전주비 (WoW)":
        cur_w  = last_date.isocalendar()[1]
        prev_w = cur_w - 1
        cur_y  = last_date.year
        cur  = df_all[(df_all["date"].dt.isocalendar().week == cur_w) & (df_all["date"].dt.year == cur_y)]
        prev = df_all[(df_all["date"].dt.isocalendar().week == prev_w) & (df_all["date"].dt.year == cur_y)]
        return cur, prev, f"{cur_y}W{cur_w}", f"{cur_y}W{prev_w}"
    else:  # 전체
        mid = df_all["date"].median()
        cur  = df_all[df_all["date"] >= mid]
        prev = df_all[df_all["date"] <  mid]
        return cur, prev, "후반기", "전반기"

def compare_table(df_cur, df_prev, label_cur, label_prev, keys=None):
    """기간 비교 테이블 생성"""
    if keys is None:
        keys = ["perSend","revenue","rps","totalSend","customers","ctr",
                "purchaseRate","purchaseCust","purchaseCnt","avgOrderVal","unitPrice",
                "uniqueInflow","totalInflow","visitPerPerson","rpc"]
    rows = []
    for k in keys:
        try:
            v_cur  = df_cur[k].mean()  if len(df_cur)  else np.nan
            v_prev = df_prev[k].mean() if len(df_prev) else np.nan
            chg = pct(v_prev, v_cur) if not np.isnan(v_prev) and v_prev!=0 else np.nan
            rows.append({
                "지표": METRIC_LABELS.get(k, k),
                label_prev: fmt_val(k, v_prev),
                label_cur:  fmt_val(k, v_cur),
                "증감율": f"{chg:+.1f}%" if not np.isnan(chg) else "–",
                "_chg": chg, "_k": k,
            })
        except: pass
    df_out = pd.DataFrame(rows)
    return df_out

def styled_compare(df_tbl):
    """증감율에 색상 적용"""
    positive_good = {"revenue","purchaseCust","purchaseCnt","ctr","purchaseRate",
                     "rpc","avgOrderVal","uniqueInflow","totalInflow"}
    def color_row(row):
        chg = row.get("_chg", np.nan)
        k   = row.get("_k","")
        if np.isnan(chg): return [""] * len(row)
        good = (chg > 0) == (k in positive_good)
        color = "color: #16a34a; font-weight:600" if good else "color: #dc2626; font-weight:600"
        result = [""] * len(row)
        result[list(row.index).index("증감율")] = color
        return result

    display = df_tbl.drop(columns=["_chg","_k"], errors="ignore")
    try:
        return display.style.apply(color_row, axis=1)
    except:
        return display


# ══════════════════════════════════════════════════════
# Appendix
# ══════════════════════════════════════════════════════
def show_appendix(df_raw, label="근거 데이터"):
    with st.expander(f"📎 Appendix — {label}", expanded=False):
        st.markdown("<div class='appendix'>", unsafe_allow_html=True)
        st.dataframe(df_raw.reset_index(drop=True), use_container_width=True)
        csv = df_raw.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇ CSV 다운로드", csv, file_name=f"{label}.csv", mime="text/csv")
        st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 통계 해설
# ══════════════════════════════════════════════════════
def stat_explainer():
    with st.expander("📖 통계 용어 해설 — 기초 지식 없이도 이해하기", expanded=False):
        verdict_box("stat_r2",
            "R² (결정계수) — \"이 지표가 얼마나 일관된 패턴을 보이는가\"\n0~1 사이 숫자. 0.3 이상이면 뚜렷한 경향, 0.1~0.3은 약한 경향, 0.1 미만은 거의 패턴 없음.\n예: R²=0.286 → 시간 흐름에 따른 하락 패턴이 28.6% 설명됨 — 꽤 뚜렷한 경향",
            "vb")
        verdict_box("stat_p",
            "p값 (유의확률) — \"이 패턴이 우연일 확률\"\np<0.001 ★★★ → 우연일 확률 0.1% 미만 = 거의 확실한 경향\np<0.01 ★★ → 우연일 확률 1% 미만 = 신뢰할 수 있는 경향\np<0.05 ★ → 우연일 확률 5% 미만 = 통계적으로 유의함\np≥0.05 ns → 우연일 수도 있음 = 근거 부족",
            "vg")
        verdict_box("stat_corr",
            "상관계수 (Correlation) — \"두 지표가 얼마나 함께 움직이는가\"\n-1~+1 사이. 음수 = 한쪽이 오르면 다른 쪽 내려감, 양수 = 같이 움직임.\n-0.4 이하면 꽤 강한 음의 관계. 예: 발송↑ → 건당거래액↓",
            "va")
        verdict_box("stat_slope",
            "선형회귀 기울기 (slope) — \"하루에 얼마씩 변하는가\"\n예: slope=-0.0028%p/일 → CTR이 매일 0.0028%p씩 하락 = 1년이면 약 1%p 하락",
            "vb")

# ══════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📨 발송 분석")
    uploaded = st.file_uploader("엑셀 업로드", type=["xlsx","xls"], label_visibility="collapsed")

    if uploaded:
        file_bytes = uploaded.read()
        G    = compute(file_bytes)
        meta = G["meta"]
        df   = G["df"]
        st.success(f"✅ {meta['days']}일 데이터")
        st.caption(f"{meta['start']} ~ {meta['end']}")

        # 날짜 필터
        st.markdown("---")
        st.markdown("**📅 기간 필터**")
        date_min = pd.to_datetime(meta["start"])
        date_max = pd.to_datetime(meta["end"])
        d1 = st.date_input("시작일", date_min, min_value=date_min, max_value=date_max)
        d2 = st.date_input("종료일", date_max, min_value=date_min, max_value=date_max)
        mask_date = (df["date"] >= pd.to_datetime(d1)) & (df["date"] <= pd.to_datetime(d2))
        df_f = df[mask_date].copy()

        # 요일 필터
        st.markdown("**📆 요일 필터**")
        DOW_KR = ["월","화","수","목","금","토","일"]
        sel_dow = st.multiselect("요일 선택", DOW_KR, default=DOW_KR)
        dow_map = {v:i for i,v in enumerate(["월","화","수","목","금","토","일"])}
        sel_dow_idx = [dow_map[d] for d in sel_dow]
        df_f = df_f[df_f["dow"].isin(sel_dow_idx)]

        st.markdown("---")
        st.caption(f"필터 적용 후: {len(df_f)}일")
    else:
        G = None; df_f = None

    st.markdown("---")
    PAGE_LIST = [
        "🏠 전체 요약",
        "📨 발송 빈도 효율",
        "📈 피로도 시계열",
        "🔬 인과 검증",
        "📊 지표 상관 분석",
        "📅 요일별 패턴",
        "🎯 발송 최적 구간",
        "📉 한계수익 분석",
    ]
    page = st.radio("분석 주제", PAGE_LIST, label_visibility="collapsed")

    if G:
        st.markdown("---")
        lines = ["# 발송 분석 인사이트"]
        for pk, items in st.session_state.insights.items():
            if pk.startswith("__"): continue
            if not isinstance(items, list): continue
            lines.append(f"\n## {pk}")
            for it in items:
                txt = it.get("text", str(it)) if isinstance(it, dict) else str(it)
                lines.append(f"- {txt}")
        st.download_button("📋 인사이트 내보내기", "\n".join(lines),
                           file_name="insights.txt", mime="text/plain", use_container_width=True)

# ══════════════════════════════════════════════════════
# 업로드 전 안내
# ══════════════════════════════════════════════════════
if G is None:
    st.title("발송 분석 대시보드")
    st.markdown("왼쪽에서 **MTD 발송 상세 엑셀 파일**을 업로드하면 분석이 시작됩니다.")
    c1,c2,c3,c4 = st.columns(4)
    for cw,title,items in [
        (c1,"발송 지표",["인당 발송 건수","총 발송 건수","유니크 고객수"]),
        (c2,"효율 지표",["CTR","발송건당 거래액","구매전환율(CR)"]),
        (c3,"유입 지표",["유니크 유입","총 유입","인당 방문 횟수"]),
        (c4,"구매 지표",["구매 고객수","구매 건수","객단가"]),
    ]:
        with cw: st.info(f"**{title}**\n\n"+"\n\n".join(f"• {i}" for i in items))
    st.stop()

# ── 공통 데이터 ──
quarterly = G["quarterly"]
monthly   = G["monthly"]
buckets   = G["buckets"]
quintile  = G["quintile"]
dow_comp  = G["dow_comp"]
dow_corr  = G["dow_corr"]
reg       = G["reg"]

# 필터 적용 재집계
def reagg(df_filt, grp):
    if len(df_filt) == 0: return pd.DataFrame()
    return df_filt.groupby(grp, sort=True).agg(
        n=("revenue","count"),
        **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}
    ).reset_index()

monthly_f   = reagg(df_f, "month")
quarterly_f = reagg(df_f, "quarter")
df_f["bucket_f"] = pd.cut(df_f["perSend"], bins=[0,2,2.5,3,3.5,4,4.5,99],
                           labels=["~2.0","2.0~2.5","2.5~3.0","3.0~3.5","3.5~4.0","4.0~4.5","4.5+"])
buckets_f = df_f.groupby("bucket_f", observed=True).agg(
    n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}
).reset_index()
buckets_f = buckets_f[buckets_f["n"]>=10].reset_index(drop=True)

fq = quarterly_f.iloc[0] if len(quarterly_f) else quarterly.iloc[0]
lq = quarterly_f.iloc[-1] if len(quarterly_f) else quarterly.iloc[-1]

ALL_METRIC_KEYS = list(ALL_METRICS.keys()) + ["purchaseRate","rpc"]
metric_select_list = [METRIC_LABELS.get(k,k) for k in ALL_METRIC_KEYS]
def key_from_label(lbl): return next((k for k,v in METRIC_LABELS.items() if v==lbl), None)

def fmt_val(k, v):
    if k == "revenue": return f"{v/1e8:.3f}억"
    if k in ("totalSend","customers","uniqueInflow","totalInflow"):
        return f"{v/1e4:.1f}만" if v>10000 else f"{v:.0f}"
    if k in ("ctr","purchaseRate"): return f"{v*100:.2f}%"
    fmt = METRIC_FORMAT.get(k, ("{:.2f}",""))[0]
    try: return fmt.format(v)
    except: return str(round(v,2))

def metric_color(k):
    cmap = {"perSend":"amber","revenue":"blue","rps":"green","totalSend":"teal",
            "customers":"blue","ctr":"red","purchaseRate":"purple",
            "rpc":"green","purchaseCust":"purple","purchaseCnt":"teal",
            "avgOrderVal":"amber","unitPrice":"amber"}
    return cmap.get(k,"blue")

# ══════════════════════════════════════════════════════
# PAGE: 전체 요약
# ══════════════════════════════════════════════════════
if page == "🏠 전체 요약":
    _t = editable_text("title_overview", "전체 요약", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"{meta['start']} ~ {meta['end']} ({meta['days']}일 전체 / 필터 적용: {len(df_f)}일)")

    kpi_keys = ["perSend","revenue","ctr","purchaseRate","rps","purchaseCust","avgOrderVal","totalSend"]
    cols = st.columns(4)
    for i, k in enumerate(kpi_keys):
        v_now = df_f[k].mean() if len(df_f) else np.nan
        v_all = G["df"][k].mean()
        delta = pct(v_all, v_now)
        inv = k not in ["revenue","purchaseCust","avgOrderVal","ctr","purchaseRate","rps","rpc"]
        cols[i%4].metric(METRIC_LABELS[k], fmt_val(k,v_now),
                         f"{delta:+.1f}% vs 전체평균",
                         delta_color="inverse" if inv else "normal")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 주요 지표 트렌드 차트 (지표 선택)
    m1, m2 = st.columns([2,1])
    sel_metric_label = m1.selectbox("추이 지표 선택", metric_select_list, index=0)
    agg_unit = m2.selectbox("집계 단위", ["월별","분기별","연도별"])
    sel_k = key_from_label(sel_metric_label)

    agg_df = {"월별":monthly_f,"분기별":quarterly_f,"연도별":reagg(df_f,"year")}[agg_unit]
    x_col  = {"월별":"month","분기별":"quarter","연도별":"year"}[agg_unit]

    if sel_k and len(agg_df):
        yv = agg_df[sel_k].tolist()
        cn = metric_color(sel_k)
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=agg_df[x_col].astype(str).tolist(), y=yv,
            mode="lines+markers", line=dict(color=clr(cn),width=2),
            marker=dict(size=5,color=clr(cn)),
            fill="tozeroy", fillcolor=cbg(cn),
        ))
        layout = base_layout(260, title=f"{sel_metric_label} 추이")
        layout["xaxis"]["tickangle"] = -30
        fig.update_layout(**layout)
        st.plotly_chart(fig, use_container_width=True)
        show_appendix(agg_df[[x_col,"n",sel_k]].rename(columns={sel_k:sel_metric_label}), f"전체요약_{sel_metric_label}")

    # 핵심 변화 요약
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("핵심 변화 요약")
    cc1,cc2 = st.columns(2)
    for col_w, keys in [(cc1,["perSend","ctr","purchaseRate","rps"]),
                        (cc2,["revenue","purchaseCust","avgOrderVal","totalSend"])]:
        rows = []
        for k in keys:
            if len(quarterly_f) >= 2:
                chg = pct(fq[k], lq[k])
                rows.append({"지표":METRIC_LABELS[k], f"{fq['quarter']}":fmt_val(k,fq[k]),
                              f"{lq['quarter']}":fmt_val(k,lq[k]), "변화율":f"{chg:+.1f}%"})
        if rows: col_w.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


    # ── 기간 비교 전체 테이블 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("📊 기간 비교 전체 지표")
    cmp_mode = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_ov")
    df_cur, df_prev, lbl_cur, lbl_prev = get_compare_periods(G["df"], cmp_mode)
    if len(df_cur) and len(df_prev):
        tbl_all = compare_table(df_cur, df_prev, lbl_cur, lbl_prev)
        st.dataframe(styled_compare(tbl_all), use_container_width=True, hide_index=True)
        show_appendix(tbl_all.drop(columns=["_chg","_k"],errors="ignore"), f"기간비교_{cmp_mode}")
    else:
        st.info("비교 가능한 이전 기간 데이터가 없습니다.")

    stat_explainer()
    insight_editor("전체요약", [
        f"분석 기간 {meta['start']}~{meta['end']} ({meta['days']}일)",
        "인당 발송 건수 증가와 CTR·구매율 하락이 동기간 동반 발생",
        "발송건당 거래액은 지속적으로 감소 추세",
    ])

# ══════════════════════════════════════════════════════
# PAGE: 발송 빈도 효율
# ══════════════════════════════════════════════════════
elif page == "📨 발송 빈도 효율":
    _t = editable_text("title_freq", "발송 빈도 효율 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")

    q1v, q5v = quintile.iloc[0], quintile.iloc[4]
    verdict_box("freq_v1",
        f'❌ "많이 보낼수록 매출이 오른다" — 기각\nQ1(평균 {q1v["totalSend"]/1e6:.2f}M건) 거래액 {q1v["revenue"]/1e8:.3f}억 vs Q5({q5v["totalSend"]/1e6:.2f}M건) {q5v["revenue"]/1e8:.3f}억. 2배 이상 더 보내도 매출은 오히려 낮습니다.',
        "vr")
    verdict_box("freq_v2",
        "⚠️ \"발송 줄이면 매출 오른다\" — 과잉 주장\n요일 통제 후 방향이 일관되지 않아 인과 주장 불가입니다.",
        "va")
    verdict_box("freq_v3",
        "✅ 입증 가능: 과잉 발송은 비용만 늘린다\n발송건당 거래액은 구간 높아질수록 단조 감소합니다.",
        "vg")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("인당 발송 구간별 성과 (30일+ 구간)")

    sel_m2 = st.selectbox("지표", metric_select_list, index=metric_select_list.index("발송건당 거래액"))
    sel_k2 = key_from_label(sel_m2)
    cn2    = metric_color(sel_k2) if sel_k2 else "blue"

    if sel_k2 and len(buckets_f):
        labels = buckets_f["bucket_f"].astype(str).tolist()
        yv2    = buckets_f[sel_k2].tolist()
        text2  = [fmt_val(sel_k2, v) for v in yv2]
        fig3 = go.Figure(go.Bar(
            x=labels, y=yv2,
            marker_color=cbg(cn2), marker_line_color=clr(cn2), marker_line_width=1.5,
            text=text2, textposition="outside", textfont=dict(color="#64748b",size=11),
        ))
        layout3 = base_layout(300, title=f"구간별 {sel_m2}")
        for i,row in buckets_f.iterrows():
            fig3.add_annotation(x=str(row["bucket_f"]),y=0,text=f"n={row['n']}일",
                showarrow=False,yanchor="top",yshift=-16,font=dict(size=10,color="#94a3b8"))
        fig3.update_layout(**layout3)
        st.plotly_chart(fig3, use_container_width=True)
        show_appendix(buckets_f[["bucket_f","n",sel_k2]].rename(columns={"bucket_f":"구간",sel_k2:sel_m2}), f"구간별_{sel_m2}")

    # 5분위 이중축
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("총 발송량 5분위 vs 지표 (이중축)")
    ma, mb = st.columns(2)
    ym1_label = ma.selectbox("좌축 지표", metric_select_list, index=metric_select_list.index("거래액"), key="q5_l")
    ym2_label = mb.selectbox("우축 지표", metric_select_list, index=metric_select_list.index("발송건당 거래액"), key="q5_r")
    ym1, ym2 = key_from_label(ym1_label), key_from_label(ym2_label)
    if ym1 and ym2:
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(x=list(quintile["label"]),y=list(quintile[ym1]),
            name=ym1_label,yaxis="y",marker_color=cbg("blue"),
            marker_line_color=clr("blue"),marker_line_width=1.5))
        fig4.add_trace(go.Scatter(x=list(quintile["label"]),y=list(quintile[ym2]),
            name=ym2_label,yaxis="y2",mode="lines+markers",
            line=dict(color=clr("amber"),width=2),marker=dict(size=6,color=clr("amber"))))
        layout4 = base_layout(280, title="5분위 비교")
        layout4["showlegend"] = True
        layout4["legend"]  = dict(orientation="h",y=1.05,bgcolor="rgba(0,0,0,0)",font=dict(color="#64748b"))
        layout4["yaxis"]["title"] = ym1_label
        layout4["yaxis2"] = dict(overlaying="y",side="right",title=ym2_label,
                                  tickfont=dict(color="#64748b",size=11),gridcolor="rgba(0,0,0,0)")
        fig4.update_layout(**layout4)
        st.plotly_chart(fig4, use_container_width=True)
        show_appendix(quintile[["label",ym1,ym2]].rename(columns={"label":"5분위",ym1:ym1_label,ym2:ym2_label}), "5분위분석")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("📊 기간 비교 — 발송 효율")
    cmp_mode_fr = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_fr")
    df_cur_fr, df_prev_fr, lbl_cur_fr, lbl_prev_fr = get_compare_periods(G["df"], cmp_mode_fr)
    if len(df_cur_fr) and len(df_prev_fr):
        tbl_fr = compare_table(df_cur_fr, df_prev_fr, lbl_cur_fr, lbl_prev_fr,
                               keys=["perSend","totalSend","customers","rps","revenue",
                                     "ctr","purchaseRate","purchaseCnt"])
        st.dataframe(styled_compare(tbl_fr), use_container_width=True, hide_index=True)
        show_appendix(tbl_fr.drop(columns=["_chg","_k"],errors="ignore"), f"발송효율비교_{cmp_mode_fr}")

    stat_explainer()
    insight_editor("발송빈도효율", [
        "발송건당 거래액은 구간 높아질수록 단조 감소 — 추가 발송의 한계 기여 감소",
        "총 발송량 Q1(최소) 구간의 거래액이 Q5(최대)보다 높음",
        "과잉 발송은 비용만 늘리고 매출 기여는 없음",
    ])

# ══════════════════════════════════════════════════════
# PAGE: 피로도 시계열
# ══════════════════════════════════════════════════════
elif page == "📈 피로도 시계열":
    _t = editable_text("title_ts", "피로도 시계열 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption("요일 통제 후 잔차 회귀 기준")

    sc = pct(fq["perSend"], lq["perSend"])
    cc = pct(fq["ctr"], lq["ctr"])

    verdict_box("ts_v1",
        f"✅ 피로도 누적 가설 — 요일 통제 후 p<0.001로 입증됨\n인당 발송 증가({sc:+.0f}%)와 CTR 하락({cc:+.0f}%)이 같은 기간 동반 발생. 요일 효과 제거 후에도 모든 효율 지표 하락 추세가 통계적으로 유의합니다.",
        "vg")

    # 지표 선택 트렌드
    sel_ts = st.selectbox("분석 지표", metric_select_list, index=0, key="ts_sel")
    ts_k   = key_from_label(sel_ts)
    ts_agg = st.selectbox("집계 단위", ["월별","분기별"], key="ts_agg")
    ts_df  = monthly_f if ts_agg=="월별" else quarterly_f
    x_col  = "month" if ts_agg=="월별" else "quarter"

    if ts_k and len(ts_df):
        cn_ts = metric_color(ts_k)
        fig5 = go.Figure(go.Scatter(
            x=ts_df[x_col].astype(str).tolist(), y=ts_df[ts_k].tolist(),
            mode="lines+markers", line=dict(color=clr(cn_ts),width=2),
            marker=dict(size=5,color=clr(cn_ts)),
            fill="tozeroy", fillcolor=cbg(cn_ts),
        ))
        layout5 = base_layout(260, title=f"{sel_ts} 추이 ({ts_agg})")
        layout5["xaxis"]["tickangle"] = -30
        fig5.update_layout(**layout5)
        st.plotly_chart(fig5, use_container_width=True)

    # 이중축: 발송 vs 선택 지표
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("인당 발송 vs 효율 지표 역방향 확인")
    dual_sel = st.selectbox("비교 지표 (우축)", metric_select_list,
                             index=metric_select_list.index("CTR"), key="dual_sel")
    dual_k = key_from_label(dual_sel)
    if dual_k and len(quarterly_f):
        fig6 = go.Figure()
        fig6.add_trace(go.Scatter(x=quarterly_f["quarter"].astype(str).tolist(),
            y=quarterly_f["perSend"].tolist(), name="인당발송(건)", yaxis="y",
            mode="lines+markers",line=dict(color=clr("blue"),width=2),
            marker=dict(size=5,color=clr("blue"))))
        fig6.add_trace(go.Scatter(x=quarterly_f["quarter"].astype(str).tolist(),
            y=quarterly_f[dual_k].tolist(), name=dual_sel, yaxis="y2",
            mode="lines+markers",line=dict(color=clr("red"),width=2,dash="dot"),
            marker=dict(size=5,color=clr("red"))))
        layout6 = base_layout(280, title="발송 vs 효율 (이중축)")
        layout6["showlegend"] = True
        layout6["legend"] = dict(orientation="h",y=1.05,bgcolor="rgba(0,0,0,0)",font=dict(color="#64748b"))
        layout6["yaxis"]["title"]    = "인당발송(건)"
        layout6["yaxis"]["ticksuffix"] = "건"
        layout6["yaxis2"] = dict(overlaying="y",side="right",title=dual_sel,
                                  tickfont=dict(color="#64748b",size=11),gridcolor="rgba(0,0,0,0)")
        fig6.update_layout(**layout6)
        st.plotly_chart(fig6, use_container_width=True)
        show_appendix(quarterly_f[["quarter","perSend",dual_k]].rename(
            columns={"quarter":"분기","perSend":"인당발송",dual_k:dual_sel}), "피로도이중축")

    # 회귀 통계
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("통계 유의성 (요일 통제 후 잔차 회귀)")
    reg_cols = st.columns(len(reg))
    for i,(k,v) in enumerate(reg.items()):
        lbl = METRIC_LABELS.get(k,k)
        inv = k not in ["perSend","revenue"]
        reg_cols[i].metric(f"{lbl} 추세",
            f"{v['slope']:+.5f}/일" if not np.isnan(v['slope']) else "–",
            f"R²={v['r2']:.3f}  {sig_label(v['p'])}",
            delta_color="inverse" if inv else "normal")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("📊 기간 비교 — 효율 지표")
    cmp_mode_ts = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_ts")
    df_cur_ts, df_prev_ts, lbl_cur_ts, lbl_prev_ts = get_compare_periods(G["df"], cmp_mode_ts)
    if len(df_cur_ts) and len(df_prev_ts):
        tbl_ts = compare_table(df_cur_ts, df_prev_ts, lbl_cur_ts, lbl_prev_ts,
                               keys=["perSend","ctr","purchaseRate","rps","revenue","rpc"])
        st.dataframe(styled_compare(tbl_ts), use_container_width=True, hide_index=True)
        show_appendix(tbl_ts.drop(columns=["_chg","_k"],errors="ignore"), f"피로도비교_{cmp_mode_ts}")

    stat_explainer()
    insight_editor("피로도시계열", [
        "인당 발송이 늘어난 기간과 효율이 하락한 기간이 일치함 (요일 통제 후 p<0.001)",
        "피로도 외 대안 가설(수신 모수 확대, 메시지 품질 변화) 배제 불가 — A/B 테스트 필요",
        "경영진 대응: '지난 2년간 발송 63% 증가, 클릭률 35% 하락 — 통계적으로 유의한 패턴'",
    ])

# ══════════════════════════════════════════════════════
# PAGE: 인과 검증
# ══════════════════════════════════════════════════════
elif page == "🔬 인과 검증":
    _t = editable_text("title_causal", "인과 검증", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    verdict_box("causal_v1",
        "❌ \"발송 줄이면 매출 오른다\" — 인과 근거 없음\n요일 통제 후 평일 중 절반 이상에서 방향이 일관되지 않습니다.",
        "vr")
    verdict_box("causal_v2",
        "⚠️ \"많이 보내면 매출 오른다\"도 — 입증 불가\n요일별 상관계수 방향이 혼재합니다.",
        "va")
    verdict_box("causal_v3",
        "✅ 입증되는 것: 건당 효율은 일관되게 악화\n모든 요일에서 발송 많은 날의 건당 거래액이 낮습니다.",
        "vg")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("요일 통제 후 — 같은 요일 내 비교")
    dc_rows = []
    for r in dow_comp:
        dc_rows.append({"요일":r["dow"]+"요일","기준":f"≤{r['median']:.1f}건",
            "적은날 거래액":f"{r['lowRev']/1e8:.3f}억","많은날 거래액":f"{r['highRev']/1e8:.3f}억",
            "차이":f"{'▲' if r['diff']>0 else '▼'}{abs(r['diff'])/1e6:.1f}M ({'적은날↑' if r['diff']>0 else '많은날↑'})",
            "건당거래(적↔많)":f"{r['lowRps']:.0f}원↔{r['highRps']:.0f}원"})
    dc_df = pd.DataFrame(dc_rows)
    st.dataframe(dc_df, use_container_width=True, hide_index=True)
    show_appendix(dc_df, "요일통제비교")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("요일 통제 후 상관계수")
    dc_corr = pd.DataFrame(dow_corr)
    fig7 = go.Figure()
    for sign, cn in [(True,"blue"),(False,"red")]:
        sub = dc_corr[dc_corr["corrRevenue"]>=0] if sign else dc_corr[dc_corr["corrRevenue"]<0]
        if len(sub):
            fig7.add_trace(go.Bar(x=list(sub["dow"]),y=list(sub["corrRevenue"].round(3)),
                name="vs 거래액(양)" if sign else "vs 거래액(음)",
                showlegend=True, marker_color=cbg(cn),
                marker_line_color=clr(cn), marker_line_width=1.5))
    fig7.add_trace(go.Scatter(x=list(dc_corr["dow"]),y=list(dc_corr["corrRps"].round(3)),
        name="vs 건당거래액",mode="lines+markers",line=dict(color=clr("red"),width=2),
        marker=dict(size=6,color=clr("red"))))
    layout7 = base_layout(260, title="상관계수 (발송 vs 지표)")
    layout7["showlegend"] = True
    layout7["legend"] = dict(orientation="h",y=1.1,bgcolor="rgba(0,0,0,0)",font=dict(color="#64748b"))
    layout7["yaxis"]["range"] = [-1, 0.5]
    layout7["yaxis"]["title"] = "상관계수"
    layout7["shapes"] = [dict(type="line",x0=-0.5,x1=len(dc_corr)-0.5,y0=0,y1=0,
                               line=dict(color="#2a2d3a",width=1,dash="dot"))]
    fig7.update_layout(**layout7)
    st.plotly_chart(fig7, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    verdict_box("causal_mgmt2",
        "경영진 대응 — 정확한 주장:\n발송을 줄이면 매출이 오른다는 보장은 없습니다. 하지만 현재 발송 수준은 추가 비용 대비 매출 기여가 없습니다. 이 비용을 타겟팅 개선이나 다른 채널에 투자하는 것이 더 효과적입니다.",
        "vg")
    verdict_box("causal_ab2",
        "다음 단계: A/B 테스트\n통제군: 현행 발송 유지 | 실험군: 인당 2.5건 이하 제한\n4주 이상, 거래액·CTR·구매율·수신거부율 측정",
        "vb")
    st.subheader("📊 기간 비교 — 인과 관련 지표")
    cmp_mode_ca = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_ca")
    df_cur_ca, df_prev_ca, lbl_cur_ca, lbl_prev_ca = get_compare_periods(G["df"], cmp_mode_ca)
    if len(df_cur_ca) and len(df_prev_ca):
        tbl_ca = compare_table(df_cur_ca, df_prev_ca, lbl_cur_ca, lbl_prev_ca,
                               keys=["perSend","revenue","rps","ctr","purchaseRate","rpc"])
        st.dataframe(styled_compare(tbl_ca), use_container_width=True, hide_index=True)

    stat_explainer()
    insight_editor("인과검증", [
        "요일 통제 후 발송 vs 거래액 상관계수: 방향 불일치 → 인과 주장 불가",
        "발송 vs 건당거래액: 모든 요일에서 음의 상관(-0.4~-0.8) → 일관된 패턴",
        "올바른 주장: '과잉 발송은 효율 없이 비용만 늘린다'",
    ])

# ══════════════════════════════════════════════════════
# PAGE: 지표 상관 분석
# ══════════════════════════════════════════════════════
elif page == "📊 지표 상관 분석":
    _t = editable_text("title_corr", "지표 상관 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption("두 지표 간의 관계를 산점도와 상관계수로 확인합니다.")

    ca, cb = st.columns(2)
    xm_label = ca.selectbox("X축 지표", metric_select_list, index=metric_select_list.index("인당 발송 건수"), key="corr_x")
    ym_label = cb.selectbox("Y축 지표", metric_select_list, index=metric_select_list.index("발송건당 거래액"), key="corr_y")
    xk = key_from_label(xm_label); yk = key_from_label(ym_label)

    if xk and yk and len(df_f) > 5:
        xv = df_f[xk].dropna(); yv_s = df_f[yk].dropna()
        common = df_f[[xk,yk]].dropna()
        corr_val = np.corrcoef(common[xk], common[yk])[0,1]
        reg_r = linreg(common[xk].values, common[yk].values)

        c1,c2,c3 = st.columns(3)
        c1.metric("상관계수", f"{corr_val:.3f}", "음의 관계" if corr_val<0 else "양의 관계")
        c2.metric("R²", f"{reg_r['r2']:.3f}", sig_label(reg_r["p"]))
        c3.metric("표본 수", f"{len(common)}일")

        fig8 = go.Figure(go.Scatter(
            x=common[xk].tolist(), y=common[yk].tolist(),
            mode="markers", marker=dict(color=clr("blue"),size=5,opacity=0.6),
            text=df_f.loc[common.index,"date"].dt.strftime("%Y-%m-%d").tolist(),
        ))
        # 추세선
        xfit = np.linspace(common[xk].min(), common[xk].max(), 100)
        yfit = reg_r["slope"]*xfit + linreg(common[xk].values, common[yk].values).get("slope",0)
        fig8.add_trace(go.Scatter(x=xfit.tolist(), y=(reg_r["slope"]*xfit).tolist(),
            mode="lines", line=dict(color=clr("red"),width=1.5,dash="dot"), name="추세선"))
        layout8 = base_layout(300, title=f"{xm_label} vs {ym_label} (r={corr_val:.3f})")
        layout8["xaxis"]["title"] = xm_label
        layout8["yaxis"]["title"] = ym_label
        fig8.update_layout(**layout8)
        st.plotly_chart(fig8, use_container_width=True)
        show_appendix(common.rename(columns={xk:xm_label,yk:ym_label}), f"산점도_{xm_label}_{ym_label}")

    # 상관 매트릭스
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("주요 지표 상관 매트릭스")
    mat_keys = ["perSend","revenue","rps","ctr","purchaseRate","purchaseCnt","avgOrderVal"]
    mat_labels = [METRIC_LABELS[k] for k in mat_keys]
    if len(df_f) > 10:
        corr_mat = df_f[mat_keys].corr().round(3)
        corr_mat.columns = mat_labels; corr_mat.index = mat_labels
        fig9 = go.Figure(go.Heatmap(
            z=corr_mat.values.tolist(), x=mat_labels, y=mat_labels,
            colorscale=[[0,"rgba(245,101,101,0.8)"],[0.5,"rgba(248,249,252,1)"],[1,"rgba(79,143,255,0.8)"]],
            zmid=0, text=corr_mat.round(2).values.tolist(), texttemplate="%{text}",
            showscale=True,
        ))
        layout9 = base_layout(380, title="지표 간 상관 매트릭스")
        layout9["xaxis"]["tickangle"] = -30
        fig9.update_layout(**layout9)
        st.plotly_chart(fig9, use_container_width=True)
        show_appendix(corr_mat.reset_index().rename(columns={"index":"지표"}), "상관매트릭스")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("📊 기간 비교 — 전 지표")
    cmp_mode_co = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_co")
    df_cur_co, df_prev_co, lbl_cur_co, lbl_prev_co = get_compare_periods(G["df"], cmp_mode_co)
    if len(df_cur_co) and len(df_prev_co):
        tbl_co = compare_table(df_cur_co, df_prev_co, lbl_cur_co, lbl_prev_co)
        st.dataframe(styled_compare(tbl_co), use_container_width=True, hide_index=True)
        show_appendix(tbl_co.drop(columns=["_chg","_k"],errors="ignore"), f"전지표비교_{cmp_mode_co}")

    stat_explainer()
    insight_editor("지표상관분석", [
        "인당 발송 건수와 발송건당 거래액은 강한 음의 상관관계",
        "CTR과 구매율은 높은 양의 상관 — 클릭이 구매로 연결되는 구조",
    ])

# ══════════════════════════════════════════════════════
# PAGE: 요일별 패턴
# ══════════════════════════════════════════════════════
elif page == "📅 요일별 패턴":
    _t = editable_text("title_dow", "요일별 패턴 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")

    sel_dow_m = st.selectbox("분석 지표", metric_select_list, index=metric_select_list.index("거래액"), key="dow_m")
    dow_k = key_from_label(sel_dow_m)
    DOW_KR = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}

    if dow_k and len(df_f):
        dow_agg = df_f.groupby("dow")[dow_k].mean().reset_index()
        dow_agg["요일"] = dow_agg["dow"].map(DOW_KR)

        fig10 = go.Figure(go.Bar(
            x=dow_agg["요일"].tolist(), y=dow_agg[dow_k].tolist(),
            marker_color=cbg("blue"), marker_line_color=clr("blue"), marker_line_width=1.5,
            text=[fmt_val(dow_k,v) for v in dow_agg[dow_k]], textposition="outside",
            textfont=dict(color="#64748b",size=11),
        ))
        layout10 = base_layout(280, title=f"요일별 {sel_dow_m} 평균")
        fig10.update_layout(**layout10)
        st.plotly_chart(fig10, use_container_width=True)

    # 요일 × 지표 히트맵
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("요일 × 지표 히트맵")
    heat_keys = ["perSend","revenue","rps","ctr","purchaseRate","avgOrderVal"]
    heat_labels = [METRIC_LABELS[k] for k in heat_keys]
    if len(df_f):
        heat_df = df_f.groupby("dow")[heat_keys].mean()
        heat_norm = (heat_df - heat_df.mean()) / (heat_df.std() + 1e-9)
        heat_norm.index = [DOW_KR[i] for i in heat_norm.index]
        heat_norm.columns = heat_labels
        fig11 = go.Figure(go.Heatmap(
            z=heat_norm.values.tolist(), x=heat_labels, y=list(heat_norm.index),
            colorscale=[[0,"rgba(245,101,101,0.8)"],[0.5,"rgba(248,249,252,1)"],[1,"rgba(79,143,255,0.8)"]],
            zmid=0, text=heat_df.round(2).values.tolist(), texttemplate="%{text}",
        ))
        layout11 = base_layout(280, title="요일 × 지표 (표준화)")
        fig11.update_layout(**layout11)
        st.plotly_chart(fig11, use_container_width=True)
        show_appendix(heat_df.reset_index().rename(columns={"dow":"요일"}), "요일별히트맵")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("📊 기간 비교 — 요일별 성과")
    cmp_mode_dw = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_dw")
    df_cur_dw, df_prev_dw, lbl_cur_dw, lbl_prev_dw = get_compare_periods(G["df"], cmp_mode_dw)
    DOW_KR2 = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}
    if len(df_cur_dw) and len(df_prev_dw) and dow_k:
        dw_rows = []
        for d in range(7):
            vc = df_cur_dw[df_cur_dw["dow"]==d][dow_k].mean()
            vp = df_prev_dw[df_prev_dw["dow"]==d][dow_k].mean()
            chg = pct(vp, vc) if not np.isnan(vp) and vp!=0 else np.nan
            dw_rows.append({"요일":DOW_KR2[d]+"요일", lbl_prev_dw:fmt_val(dow_k,vp),
                             lbl_cur_dw:fmt_val(dow_k,vc), "증감율":f"{chg:+.1f}%" if not np.isnan(chg) else "–"})
        st.dataframe(pd.DataFrame(dw_rows), use_container_width=True, hide_index=True)

    stat_explainer()
    insight_editor("요일별패턴", [
        "특정 요일에 발송이 집중되어 있는지, 그 요일의 효율은 어떤지 확인 필요",
        "요일 패턴은 발송 빈도 분석의 교란변수 — 반드시 통제 후 해석",
    ])

# ══════════════════════════════════════════════════════
# PAGE: 발송 최적 구간
# ══════════════════════════════════════════════════════
elif page == "🎯 발송 최적 구간":
    _t = editable_text("title_opt", "발송 최적 구간 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption("거래액·효율·구매율을 종합한 최적 발송 구간을 찾습니다.")

    if len(buckets_f):
        # 종합 점수: 각 지표를 0-1 정규화 후 가중 합산
        score_keys = ["revenue","rps","ctr","purchaseRate"]
        score_labels = [METRIC_LABELS[k] for k in score_keys]
        weights = {}
        st.subheader("지표별 가중치 설정")
        wcols = st.columns(len(score_keys))
        for i,(k,lbl) in enumerate(zip(score_keys,score_labels)):
            weights[k] = wcols[i].slider(lbl, 0.0, 1.0, 0.25, 0.05, key=f"w_{k}")

        # 정규화 점수
        scored = buckets_f.copy()
        for k in score_keys:
            mn, mx = scored[k].min(), scored[k].max()
            scored[f"norm_{k}"] = (scored[k]-mn)/(mx-mn+1e-9)
        scored["score"] = sum(scored[f"norm_{k}"]*weights[k] for k in score_keys)
        scored["score"] = scored["score"] / (sum(weights.values())+1e-9)

        best_idx = scored["score"].idxmax()
        best = scored.iloc[best_idx]
        verdict_box("opt_best",
            f"🎯 종합 최적 구간: {best['bucket_f']}\n가중 점수 {best['score']:.3f} / 거래액 {best.revenue/1e8:.3f}억 / 건당거래 {best.rps:.0f}원 / CTR {best.ctr*100:.2f}% / 구매율 {best.purchaseRate*100:.3f}%",
            "vg")

        fig12 = go.Figure()
        fig12.add_trace(go.Bar(
            x=scored["bucket_f"].astype(str).tolist(),
            y=scored["score"].round(3).tolist(),
            marker_color=[clr("green") if i==best_idx else cbg("blue") for i in range(len(scored))],
            marker_line_color=clr("blue"), marker_line_width=1.5,
            text=[f"{v:.3f}" for v in scored["score"]],
            textposition="outside", textfont=dict(color="#64748b",size=11),
        ))
        layout12 = base_layout(280, title="구간별 종합 점수 (가중 합산)")
        fig12.update_layout(**layout12)
        st.plotly_chart(fig12, use_container_width=True)
        show_appendix(scored[["bucket_f","n","revenue","rps","ctr","purchaseRate","score"]].rename(
            columns={"bucket_f":"구간","revenue":"거래액","rps":"건당거래","ctr":"CTR",
                     "purchaseRate":"구매율","score":"종합점수"}), "최적구간분석")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("📊 기간 비교 — 구간 효과")
    cmp_mode_op = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_op")
    df_cur_op, df_prev_op, lbl_cur_op, lbl_prev_op = get_compare_periods(G["df"], cmp_mode_op)
    if len(df_cur_op) and len(df_prev_op):
        tbl_op = compare_table(df_cur_op, df_prev_op, lbl_cur_op, lbl_prev_op,
                               keys=["perSend","revenue","rps","ctr","purchaseRate","avgOrderVal"])
        st.dataframe(styled_compare(tbl_op), use_container_width=True, hide_index=True)

    stat_explainer()
    insight_editor("발송최적구간", [
        "단순 거래액 기준이 아닌 효율·구매율 종합 시 최적 구간이 달라질 수 있음",
        "가중치를 조정해 경영진 우선순위에 맞는 최적 구간 제시 가능",
    ])

# ══════════════════════════════════════════════════════
# PAGE: 한계수익 분석
# ══════════════════════════════════════════════════════
elif page == "📉 한계수익 분석":
    _t = editable_text("title_lm", "한계수익 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption("발송 1건 추가 시 거래액·클릭이 얼마나 증가/감소하는지 분석합니다.")

    sel_lm = st.selectbox("분석 지표", metric_select_list, index=metric_select_list.index("발송건당 거래액"), key="lm_sel")
    lm_k   = key_from_label(sel_lm)

    if lm_k and len(buckets_f) >= 2:
        bk = buckets_f.copy()
        bk["marginal"] = bk[lm_k].diff()
        bk["bucket_str"] = bk["bucket_f"].astype(str)

        c1_lm, c2_lm = st.columns(2)
        # 절대값
        fig13 = go.Figure(go.Bar(
            x=bk["bucket_str"].tolist(), y=bk[lm_k].tolist(),
            marker_color=cbg("green"), marker_line_color=clr("green"), marker_line_width=1.5,
            text=[fmt_val(lm_k,v) for v in bk[lm_k]], textposition="outside",
            textfont=dict(color="#64748b",size=11),
        ))
        layout13 = base_layout(260, title=f"구간별 {sel_lm}")
        fig13.update_layout(**layout13)
        c1_lm.plotly_chart(fig13, use_container_width=True)

        # 한계 변화
        fig14 = go.Figure(go.Bar(
            x=bk["bucket_str"].tolist(), y=bk["marginal"].tolist(),
            marker_color=[cbg("green") if v>0 else cbg("red") for v in bk["marginal"].fillna(0)],
            marker_line_color=[clr("green") if v>0 else clr("red") for v in bk["marginal"].fillna(0)],
            marker_line_width=1.5,
            text=[f"{v:+.1f}" if not np.isnan(v) else "" for v in bk["marginal"]],
            textposition="outside", textfont=dict(color="#64748b",size=11),
        ))
        layout14 = base_layout(260, title=f"한계 변화량 (구간 간 차이)")
        layout14["shapes"] = [dict(type="line",x0=-0.5,x1=len(bk)-0.5,y0=0,y1=0,
                                    line=dict(color="#2a2d3a",width=1,dash="dot"))]
        fig14.update_layout(**layout14)
        c2_lm.plotly_chart(fig14, use_container_width=True)
        show_appendix(bk[["bucket_str","n",lm_k,"marginal"]].rename(
            columns={"bucket_str":"구간",lm_k:sel_lm,"marginal":"한계변화"}), "한계수익분석")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("📊 기간 비교 — 한계 효율")
    cmp_mode_lm = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_lm")
    df_cur_lm, df_prev_lm, lbl_cur_lm, lbl_prev_lm = get_compare_periods(G["df"], cmp_mode_lm)
    if len(df_cur_lm) and len(df_prev_lm):
        tbl_lm = compare_table(df_cur_lm, df_prev_lm, lbl_cur_lm, lbl_prev_lm,
                               keys=["perSend","rps","ctr","purchaseRate","revenue","rpc"])
        st.dataframe(styled_compare(tbl_lm), use_container_width=True, hide_index=True)

    stat_explainer()
    insight_editor("한계수익분석", [
        "발송건당 거래액은 구간이 높아질수록 체감 감소 — 추가 발송의 한계 기여가 0에 수렴",
        "한계수익이 0 이하로 떨어지는 구간부터는 발송이 오히려 역효과",
        "경영진 설득 포인트: 지금 발송 수준은 음의 한계수익 구간에 있음",
    ])
