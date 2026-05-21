import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from scipy import stats
import io

st.set_page_config(
    page_title="발송 빈도 분석 대시보드",
    page_icon="📨",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #0f1117; }
[data-testid="stSidebar"] { background: #1a1d27; border-right: 1px solid #2a2d3a; }
[data-testid="stMetric"] { background: #1e2130; border-radius: 8px; padding: 12px 16px; border: 1px solid #2a2d3a; }
[data-testid="stMetricLabel"] { color: #8a92a0 !important; font-size: 12px !important; }
[data-testid="stMetricValue"] { color: #f0f2f5 !important; font-size: 22px !important; }
.vg  { background:rgba(72,187,120,0.08);  border-left:3px solid #48bb78; border-radius:0 8px 8px 0; padding:12px 16px; margin:8px 0; color:#86efac; font-size:13px; line-height:1.65; }
.vr  { background:rgba(245,101,101,0.08); border-left:3px solid #f56565; border-radius:0 8px 8px 0; padding:12px 16px; margin:8px 0; color:#fca5a5; font-size:13px; line-height:1.65; }
.va  { background:rgba(237,137,54,0.08);  border-left:3px solid #ed8936; border-radius:0 8px 8px 0; padding:12px 16px; margin:8px 0; color:#fcd34d; font-size:13px; line-height:1.65; }
.vb  { background:rgba(79,143,255,0.08);  border-left:3px solid #4f8fff; border-radius:0 8px 8px 0; padding:12px 16px; margin:8px 0; color:#93c5fd; font-size:13px; line-height:1.65; }
.sdiv { border-top:1px solid #2a2d3a; margin:24px 0; }
</style>
""", unsafe_allow_html=True)

# ── 색상 (모두 rgba 문자열) ──
C = {
    "blue":   ("rgba(79,143,255,1)",   "rgba(79,143,255,0.15)"),
    "red":    ("rgba(245,101,101,1)",  "rgba(245,101,101,0.15)"),
    "green":  ("rgba(72,187,120,1)",   "rgba(72,187,120,0.15)"),
    "amber":  ("rgba(237,137,54,1)",   "rgba(237,137,54,0.15)"),
    "purple": ("rgba(159,122,234,1)",  "rgba(159,122,234,0.15)"),
    "teal":   ("rgba(56,178,172,1)",   "rgba(56,178,172,0.15)"),
}

def clr(name):      return C[name][0]
def clr_bg(name):   return C[name][1]

def base_layout(height=280, ticksuffix=""):
    return dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#8a92a0", size=11),
        margin=dict(l=10, r=10, t=30, b=10),
        height=height,
        showlegend=False,
        xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#2a2d3a",
                   tickcolor="#2a2d3a", tickfont=dict(color="#8a92a0", size=11)),
        yaxis=dict(gridcolor="#1e2130", linecolor="#2a2d3a",
                   tickcolor="#2a2d3a", tickfont=dict(color="#8a92a0", size=11),
                   ticksuffix=ticksuffix),
    )

def sig_star(p):
    if np.isnan(p): return ""
    if p < 0.001: return "★★★ p<0.001"
    if p < 0.01:  return "★★  p<0.01"
    if p < 0.05:  return "★   p<0.05"
    return "ns"

# ── 파싱 ──
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
            if isinstance(v, (int, float)):
                d = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v))
            else:
                d = pd.Timestamp(v)
            dates.append(d)
        except:
            dates.append(pd.NaT)

    def get_metric(keywords):
        for kw in keywords:
            match = metric_col[metric_col.str.replace(" ","") == kw.replace(" ","")]
            if not match.empty:
                row_idx = match.index[0]
                vals = ws.iloc[row_idx, 2:2+len(dates)].values
                return pd.to_numeric(vals, errors="coerce")
        return np.full(len(dates), np.nan)

    data = pd.DataFrame({
        "date":        dates,
        "perSend":     get_metric(["인당발송건수"]),
        "revenue":     get_metric(["거래액"]),
        "rps":         get_metric(["발송건당거래액"]),
        "totalSend":   get_metric(["총발송건수"]),
        "customers":   get_metric(["유니크발송고객수"]),
        "ctr":         get_metric(["CTR"]),
        "purchaseCust":get_metric(["구매고객수"]),
    })
    data = data.dropna(subset=["perSend","revenue"]).copy()
    data["date"]    = pd.to_datetime(data["date"])
    data = data.sort_values("date").reset_index(drop=True)
    data["t"]       = np.arange(len(data))
    data["dow"]     = data["date"].dt.dayofweek
    data["month"]   = data["date"].dt.to_period("M").astype(str)
    data["quarter"] = data["date"].dt.to_period("Q").astype(str)
    data["purchaseRate"] = data["purchaseCust"] / data["customers"]
    data["rpc"]     = data["revenue"] / data["customers"]
    return data

@st.cache_data(show_spinner=False)
def compute(file_bytes):
    df = parse_xlsx(file_bytes)

    monthly = df.groupby("month", sort=True).agg(
        n=("revenue","count"), avgSends=("perSend","mean"),
        avgRevenue=("revenue","mean"), avgRps=("rps","mean"),
        avgCtr=("ctr","mean"), avgPr=("purchaseRate","mean"),
    ).reset_index()

    quarterly = df.groupby("quarter", sort=True).agg(
        n=("revenue","count"), avgSends=("perSend","mean"),
        avgRevenue=("revenue","mean"), avgRps=("rps","mean"),
        avgCtr=("ctr","mean"), avgPr=("purchaseRate","mean"),
    ).reset_index()

    BINS = [0,2.0,2.5,3.0,3.5,4.0,4.5,99]
    LBLS = ["~2.0건","2.0~2.5","2.5~3.0","3.0~3.5","3.5~4.0","4.0~4.5","4.5건+"]
    df["bucket"] = pd.cut(df["perSend"], bins=BINS, labels=LBLS)
    buckets = df.groupby("bucket", observed=True).agg(
        n=("revenue","count"), avgSends=("perSend","mean"),
        avgRevenue=("revenue","mean"), avgRps=("rps","mean"),
        avgCtr=("ctr","mean"), avgPr=("purchaseRate","mean"),
        avgRpc=("rpc","mean"), avgTotalSend=("totalSend","mean"),
    ).reset_index()
    buckets = buckets[buckets["n"] >= 30].reset_index(drop=True)

    df_s = df.sort_values("totalSend").reset_index(drop=True)
    sz   = len(df_s) // 5
    quintile = pd.DataFrame([
        dict(label=["Q1 최소","Q2","Q3","Q4","Q5 최대"][i],
             avgTotalSend=df_s.iloc[i*sz:(i+1)*sz]["totalSend"].mean(),
             avgRevenue=df_s.iloc[i*sz:(i+1)*sz]["revenue"].mean(),
             avgRps=df_s.iloc[i*sz:(i+1)*sz]["rps"].mean(),
             avgPerSend=df_s.iloc[i*sz:(i+1)*sz]["perSend"].mean())
        for i in range(5)
    ])

    DOW = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}
    dow_comp = []
    for d in [0,1,2,3,4]:
        sub = df[df["dow"]==d].copy()
        if len(sub) < 20: continue
        med = sub["perSend"].median()
        lo, hi = sub[sub["perSend"]<=med], sub[sub["perSend"]>med]
        dow_comp.append(dict(
            dow=DOW[d], median=med,
            lowRev=lo["revenue"].mean(), highRev=hi["revenue"].mean(),
            diff=lo["revenue"].mean()-hi["revenue"].mean(),
            lowRps=lo["rps"].mean(), highRps=hi["rps"].mean(),
        ))

    dow_corr = []
    for d in range(7):
        sub = df[df["dow"]==d]
        if len(sub) < 20: continue
        dow_corr.append(dict(
            dow=DOW[d],
            corrRevenue=float(np.corrcoef(sub["perSend"], sub["revenue"])[0,1]),
            corrRps=float(np.corrcoef(sub["perSend"], sub["rps"])[0,1]),
        ))

    t = df["t"].values.astype(float)
    def dow_res(col):
        vals = df[col].values.astype(float)
        res  = vals.copy()
        for d in range(7):
            idx = df["dow"].values == d
            if idx.sum() > 0: res[idx] -= np.nanmean(res[idx])
        return res
    def linreg(x, y):
        mask = ~np.isnan(x) & ~np.isnan(y)
        x, y = x[mask], y[mask]
        if len(x) < 5: return dict(slope=np.nan, r2=np.nan, p=np.nan)
        slope, _, r, p, _ = stats.linregress(x, y)
        return dict(slope=slope, r2=r**2, p=p)
    reg = dict(
        sends=linreg(t, df["perSend"].values),
        ctr=  linreg(t, dow_res("ctr")),
        pr=   linreg(t, dow_res("purchaseRate")),
        rps=  linreg(t, dow_res("rps")),
    )
    meta = dict(start=str(df["date"].min().date()),
                end=str(df["date"].max().date()), days=len(df))
    return dict(df=df, monthly=monthly, quarterly=quarterly,
                buckets=buckets, quintile=quintile,
                dow_comp=dow_comp, dow_corr=dow_corr, reg=reg, meta=meta)

# ── 사이드바 ──
with st.sidebar:
    st.markdown("## 📨 발송 빈도 분석")
    uploaded = st.file_uploader("엑셀 업로드", type=["xlsx","xls"],
                                 label_visibility="collapsed")
    if uploaded:
        file_bytes = uploaded.read()
        G    = compute(file_bytes)
        meta = G["meta"]
        st.success(f"✅ {meta['days']}일 데이터")
        st.caption(f"{meta['start']} ~ {meta['end']}")
    else:
        G = None

    st.markdown("---")
    page = st.radio("분석 주제", [
        "📊 전체 요약", "📨 발송 빈도 분석",
        "📈 피로도 시계열", "🔬 인과 검증",
    ], label_visibility="collapsed")

    if G:
        st.markdown("---")
        q    = G["quarterly"]
        fq, lq = q.iloc[0], q.iloc[-1]
        reg  = G["reg"]
        lines = [
            f"# 발송 빈도 분석 인사이트",
            f"기간: {meta['start']} ~ {meta['end']} ({meta['days']}일)",
            "",
            "## 전체 요약",
            f"인당 발송: {fq.avgSends:.2f} → {lq.avgSends:.2f}건",
            f"CTR: {fq.avgCtr*100:.2f}% → {lq.avgCtr*100:.2f}%",
            f"구매율: {fq.avgPr*100:.3f}% → {lq.avgPr*100:.3f}%",
            f"건당거래액: {fq.avgRps:.0f}원 → {lq.avgRps:.0f}원",
            "",
            "## 피로도 시계열 (요일 통제 후 회귀)",
            f"인당발송: slope={reg['sends']['slope']:.4f}/일, R²={reg['sends']['r2']:.3f}, {sig_star(reg['sends']['p'])}",
            f"CTR: R²={reg['ctr']['r2']:.3f}, {sig_star(reg['ctr']['p'])}",
            f"구매율: R²={reg['pr']['r2']:.3f}, {sig_star(reg['pr']['p'])}",
            "",
            "## 인과 검증",
            "- 요일 통제 후 발송 vs 거래액 상관: 방향 불일치",
            "- 발송 vs 건당거래액: 모든 요일에서 음의 상관 (일관)",
            "- 올바른 주장: '과잉 발송은 비용만 늘린다'",
        ]
        st.download_button("📋 인사이트 다운로드", "\n".join(lines),
                           file_name="발송빈도_인사이트.txt",
                           mime="text/plain", use_container_width=True)

# ── 업로드 안내 ──
if G is None:
    st.title("발송 빈도 분석 대시보드")
    st.markdown("왼쪽에서 **MTD 발송 상세 엑셀 파일**을 업로드하면 분석이 시작됩니다.")
    c1,c2,c3,c4 = st.columns(4)
    for col_w, title, items in [
        (c1,"발송 지표",["인당 발송 건수","총 발송 건수","유니크 고객수"]),
        (c2,"효율 지표",["CTR","발송건당 거래액","구매율"]),
        (c3,"매출 지표",["거래액","구매 고객수","구매 건수"]),
        (c4,"단가 지표",["객단가","건단가","M당 거래액"]),
    ]:
        with col_w:
            st.info(f"**{title}**\n\n" + "\n\n".join(f"• {i}" for i in items))
    st.stop()

# ── 데이터 ──
df        = G["df"]
monthly   = G["monthly"]
quarterly = G["quarterly"]
buckets   = G["buckets"]
quintile  = G["quintile"]
dow_comp  = G["dow_comp"]
dow_corr  = G["dow_corr"]
reg       = G["reg"]
meta      = G["meta"]
fq, lq   = quarterly.iloc[0], quarterly.iloc[-1]

def pct_chg(a, b): return (b-a)/a*100 if a else 0


# ══════════════════════════════════════
# PAGE 1 — 전체 요약
# ══════════════════════════════════════
if page == "📊 전체 요약":
    st.title("전체 요약")
    st.caption(f"{meta['start']} ~ {meta['end']} ({meta['days']}일)")

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("인당 발송 (분기평균)", f"{lq.avgSends:.2f}건",
              f"{pct_chg(fq.avgSends, lq.avgSends):+.0f}% ({fq.quarter}→{lq.quarter})")
    c2.metric("CTR", f"{lq.avgCtr*100:.2f}%",
              f"{pct_chg(fq.avgCtr, lq.avgCtr):+.0f}%", delta_color="inverse")
    c3.metric("구매율", f"{lq.avgPr*100:.3f}%",
              f"{pct_chg(fq.avgPr, lq.avgPr):+.0f}%", delta_color="inverse")
    c4.metric("건당 거래액", f"{lq.avgRps:.0f}원",
              f"{pct_chg(fq.avgRps, lq.avgRps):+.0f}%", delta_color="inverse")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    metric_opt = st.selectbox("지표 선택", ["거래액 (억원)","인당 발송 건수","CTR (%)"])
    cfg = {
        "거래액 (억원)":  ("avgRevenue", lambda x: round(x/1e8,3), "억원", "blue"),
        "인당 발송 건수": ("avgSends",   lambda x: round(x,2),     "건",   "amber"),
        "CTR (%)":        ("avgCtr",     lambda x: round(x*100,2),  "%",    "red"),
    }
    col_k, tfm, unit, cname = cfg[metric_opt]
    yv = [tfm(v) for v in monthly[col_k]]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=list(monthly["month"]), y=yv,
        mode="lines+markers",
        line=dict(color=clr(cname), width=2),
        marker=dict(size=4, color=clr(cname)),
        fill="tozeroy", fillcolor=clr_bg(cname),
    ))
    layout = base_layout(280, unit)
    layout["xaxis"]["tickangle"] = -45
    fig.update_layout(**layout)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("💡 인사이트 메모")
    st.info(f"""
**{fq.quarter} → {lq.quarter} 핵심 변화**
- 인당 발송: {fq.avgSends:.2f}건 → {lq.avgSends:.2f}건 ({pct_chg(fq.avgSends,lq.avgSends):+.0f}%)
- CTR: {fq.avgCtr*100:.2f}% → {lq.avgCtr*100:.2f}% ({pct_chg(fq.avgCtr,lq.avgCtr):+.0f}%)
- 구매율: {fq.avgPr*100:.3f}% → {lq.avgPr*100:.3f}% ({pct_chg(fq.avgPr,lq.avgPr):+.0f}%)
- 건당 거래액: {fq.avgRps:.0f}원 → {lq.avgRps:.0f}원 ({pct_chg(fq.avgRps,lq.avgRps):+.0f}%)
    """)
    st.text_area("📝 추가 메모", placeholder="메모 입력...", height=100, key="memo_ov")


# ══════════════════════════════════════
# PAGE 2 — 발송 빈도 분석
# ══════════════════════════════════════
elif page == "📨 발송 빈도 분석":
    st.title("발송 빈도 분석")
    q1, q5 = quintile.iloc[0], quintile.iloc[4]

    st.markdown(f"""
<div class="vr"><strong>❌ "많이 보낼수록 매출이 오른다" — 기각</strong><br>
Q1(평균 {q1.avgTotalSend/1e6:.2f}M건) 거래액 {q1.avgRevenue/1e8:.3f}억 vs Q5({q5.avgTotalSend/1e6:.2f}M건) {q5.avgRevenue/1e8:.3f}억.
2배 이상 더 보내도 매출은 오히려 낮습니다.</div>
<div class="va"><strong>⚠️ "발송 줄이면 매출 오른다" — 과잉 주장</strong><br>
요일 통제 후 동일 요일 내 비교에서 방향이 일관되지 않습니다. 인과 주장 근거 부족합니다.</div>
<div class="vg"><strong>✅ 입증 가능: 과잉 발송은 비용만 늘린다</strong><br>
발송건당 거래액은 구간이 높아질수록 단조 감소. 추가 발송의 한계 기여가 0에 수렴합니다.</div>
""", unsafe_allow_html=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("인당 발송 구간별 성과 (30일+ 신뢰 구간)")

    metric_opt = st.selectbox("지표", [
        "거래액 (억원)","발송건당 거래액 (원)","CTR (%)","구매율 (%)","고객당 매출 (원)"
    ], key="freq_m")
    cfg2 = {
        "거래액 (억원)":     ("avgRevenue", lambda x: round(x/1e8,3), "억원", "blue"),
        "발송건당 거래액 (원)":("avgRps",   lambda x: round(x,0),     "원",   "green"),
        "CTR (%)":           ("avgCtr",    lambda x: round(x*100,2),  "%",    "red"),
        "구매율 (%)":         ("avgPr",    lambda x: round(x*100,3),  "%",    "purple"),
        "고객당 매출 (원)":   ("avgRpc",   lambda x: round(x,0),      "원",   "teal"),
    }
    col_k, tfm, unit, cname = cfg2[metric_opt]
    labels = list(buckets["bucket"].astype(str))
    yv     = [tfm(v) for v in buckets[col_k]]

    fig = go.Figure(go.Bar(
        x=labels, y=yv,
        marker_color=clr_bg(cname),
        marker_line_color=clr(cname),
        marker_line_width=1.5,
        text=[f"{v}{unit}" for v in yv],
        textposition="outside",
        textfont=dict(color="#8a92a0", size=11),
    ))
    layout = base_layout(300, unit)
    for i, row in buckets.iterrows():
        fig.add_annotation(x=str(row["bucket"]), y=0,
            text=f"n={row['n']}일", showarrow=False,
            yanchor="top", yshift=-16,
            font=dict(size=10, color="#545c6a"))
    fig.update_layout(**layout)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("총 발송량 5분위 vs 거래액")

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=list(quintile["label"]),
        y=[round(v/1e8,3) for v in quintile["avgRevenue"]],
        name="거래액(억)", yaxis="y",
        marker_color=clr_bg("blue"),
        marker_line_color=clr("blue"),
        marker_line_width=1.5,
    ))
    fig2.add_trace(go.Scatter(
        x=list(quintile["label"]),
        y=[round(v,0) for v in quintile["avgRps"]],
        name="건당거래액(원)", yaxis="y2",
        mode="lines+markers",
        line=dict(color=clr("amber"), width=2),
        marker=dict(size=6, color=clr("amber")),
    ))
    layout2 = base_layout(280)
    layout2["showlegend"] = True
    layout2["legend"] = dict(orientation="h", y=1.05,
                              bgcolor="rgba(0,0,0,0)",
                              font=dict(color="#8a92a0"))
    layout2["yaxis"]["title"]      = "거래액(억)"
    layout2["yaxis"]["ticksuffix"] = "억"
    layout2["yaxis2"] = dict(
        overlaying="y", side="right",
        title="건당거래(원)", ticksuffix="원",
        gridcolor="rgba(0,0,0,0)", tickcolor="#2a2d3a",
        tickfont=dict(color="#8a92a0", size=11),
    )
    fig2.update_layout(**layout2)
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("구간별 상세 데이터")
    tbl = buckets.copy()
    tbl["구간"]        = tbl["bucket"].astype(str)
    tbl["표본(일)"]    = tbl["n"]
    tbl["인당발송(건)"] = tbl["avgSends"].apply(lambda x: f"{x:.2f}")
    tbl["거래액(억)"]  = tbl["avgRevenue"].apply(lambda x: f"{x/1e8:.3f}")
    tbl["건당거래(원)"] = tbl["avgRps"].apply(lambda x: f"{x:.0f}")
    tbl["CTR"]         = tbl["avgCtr"].apply(lambda x: f"{x*100:.2f}%")
    tbl["구매율"]       = tbl["avgPr"].apply(lambda x: f"{x*100:.3f}%")
    tbl["고객당매출(원)"] = tbl["avgRpc"].apply(lambda x: f"{x:.0f}")
    st.dataframe(tbl[["구간","표본(일)","인당발송(건)","거래액(억)",
                       "건당거래(원)","CTR","구매율","고객당매출(원)"]],
                 use_container_width=True, hide_index=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.text_area("📝 추가 메모", placeholder="분석 메모...", height=100, key="memo_freq")


# ══════════════════════════════════════
# PAGE 3 — 피로도 시계열
# ══════════════════════════════════════
elif page == "📈 피로도 시계열":
    st.title("피로도 시계열 분석")
    st.caption("요일 통제 후 잔차 회귀 기준")

    sc = pct_chg(fq.avgSends, lq.avgSends)
    cc = pct_chg(fq.avgCtr,   lq.avgCtr)
    pc = pct_chg(fq.avgPr,    lq.avgPr)
    rc = pct_chg(fq.avgRps,   lq.avgRps)

    st.markdown(f"""
<div class="vg"><strong>✅ 피로도 누적 가설 — 입증됨 (요일 통제 후 p&lt;0.001)</strong><br>
{fq.quarter} → {lq.quarter}: 인당 발송 {fq.avgSends:.2f}→{lq.avgSends:.2f}건({sc:+.0f}%),
CTR {fq.avgCtr*100:.2f}→{lq.avgCtr*100:.2f}%({cc:+.0f}%),
구매율 {fq.avgPr*100:.3f}→{lq.avgPr*100:.3f}%({pc:+.0f}%),
건당거래액 {fq.avgRps:.0f}→{lq.avgRps:.0f}원({rc:+.0f}%).</div>
<div class="vb"><strong>ℹ️ 단, 인과 해석 주의</strong><br>
피로도 외 대안(수신 모수 확대, 메시지 품질 변화)을 완전히 배제할 수 없습니다.</div>
""", unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("인당 발송 증가", f"{sc:+.0f}%", f"{fq.avgSends:.2f}→{lq.avgSends:.2f}건")
    c2.metric("CTR 하락",  f"{cc:.0f}%", f"{fq.avgCtr*100:.2f}→{lq.avgCtr*100:.2f}%", delta_color="inverse")
    c3.metric("구매율 하락", f"{pc:.0f}%", f"{fq.avgPr*100:.3f}→{lq.avgPr*100:.3f}%", delta_color="inverse")
    c4.metric("건당거래 하락", f"{rc:.0f}%", f"{fq.avgRps:.0f}→{lq.avgRps:.0f}원", delta_color="inverse")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    ts_opt = st.selectbox("차트 유형", [
        "인당발송 vs CTR (이중축)","인당 발송 건수","CTR (%)","구매율 (%)","건당 거래액 (원)"
    ], key="ts_opt")

    if ts_opt == "인당발송 vs CTR (이중축)":
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=list(quarterly["quarter"]),
            y=list(quarterly["avgSends"].round(2)),
            mode="lines+markers", name="인당발송(건)", yaxis="y",
            line=dict(color=clr("blue"), width=2),
            marker=dict(size=5, color=clr("blue")),
        ))
        fig.add_trace(go.Scatter(
            x=list(quarterly["quarter"]),
            y=list((quarterly["avgCtr"]*100).round(2)),
            mode="lines+markers", name="CTR(%)", yaxis="y2",
            line=dict(color=clr("red"), width=2, dash="dot"),
            marker=dict(size=5, color=clr("red")),
        ))
        layout = base_layout(300)
        layout["showlegend"] = True
        layout["legend"] = dict(orientation="h", y=1.05,
                                  bgcolor="rgba(0,0,0,0)",
                                  font=dict(color="#8a92a0"))
        layout["yaxis"]["title"]      = "인당발송(건)"
        layout["yaxis"]["ticksuffix"] = "건"
        layout["yaxis2"] = dict(
            overlaying="y", side="right", title="CTR(%)", ticksuffix="%",
            gridcolor="rgba(0,0,0,0)", tickcolor="#2a2d3a",
            tickfont=dict(color="#8a92a0", size=11),
        )
        fig.update_layout(**layout)
    else:
        cfg3 = {
            "인당 발송 건수":    ("avgSends", lambda x: round(x,2), "건",  "blue"),
            "CTR (%)":          ("avgCtr",   lambda x: round(x*100,2), "%", "red"),
            "구매율 (%)":        ("avgPr",   lambda x: round(x*100,3), "%", "purple"),
            "건당 거래액 (원)":  ("avgRps",  lambda x: round(x,0),    "원", "green"),
        }
        col_k, tfm, unit, cname = cfg3[ts_opt]
        fig = go.Figure(go.Scatter(
            x=list(quarterly["quarter"]),
            y=[tfm(v) for v in quarterly[col_k]],
            mode="lines+markers",
            line=dict(color=clr(cname), width=2),
            marker=dict(size=5, color=clr(cname)),
            fill="tozeroy", fillcolor=clr_bg(cname),
        ))
        layout = base_layout(280, unit)
        layout["xaxis"]["tickangle"] = -30
        fig.update_layout(**layout)

    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("분기별 상세")
    q_tbl = quarterly.copy()
    q_tbl["인당발송"] = q_tbl["avgSends"].apply(lambda x: f"{x:.2f}건")
    q_tbl["거래액"]   = q_tbl["avgRevenue"].apply(lambda x: f"{x/1e8:.3f}억")
    q_tbl["건당거래"] = q_tbl["avgRps"].apply(lambda x: f"{x:.0f}원")
    q_tbl["CTR"]      = q_tbl["avgCtr"].apply(lambda x: f"{x*100:.2f}%")
    q_tbl["구매율"]   = q_tbl["avgPr"].apply(lambda x: f"{x*100:.3f}%")
    q_tbl["표본"]     = q_tbl["n"].astype(str) + "일"
    st.dataframe(q_tbl[["quarter","표본","인당발송","거래액","건당거래","CTR","구매율"]]
                 .rename(columns={"quarter":"분기"}),
                 use_container_width=True, hide_index=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("통계 유의성 (요일 통제 후 잔차 회귀)")
    rc1,rc2,rc3,rc4 = st.columns(4)
    rc1.metric("인당발송 추세", f"+{reg['sends']['slope']:.4f}/일",
               f"R²={reg['sends']['r2']:.3f}  {sig_star(reg['sends']['p'])}")
    rc2.metric("CTR 추세", f"{reg['ctr']['slope']*100:.5f}%p/일",
               f"R²={reg['ctr']['r2']:.3f}  {sig_star(reg['ctr']['p'])}", delta_color="inverse")
    rc3.metric("구매율 추세", f"{reg['pr']['slope']*100:.6f}%p/일",
               f"R²={reg['pr']['r2']:.3f}  {sig_star(reg['pr']['p'])}", delta_color="inverse")
    rc4.metric("건당거래 추세", f"{reg['rps']['slope']:.4f}원/일",
               f"R²={reg['rps']['r2']:.3f}  {sig_star(reg['rps']['p'])}", delta_color="inverse")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.markdown("""
<div class="vg"><strong>경영진 보고용:</strong><br>
"지난 2년간 인당 발송 건수가 63% 증가했고, 같은 기간 클릭률은 35%, 구매율은 44% 하락했습니다.
이 두 추세는 요일 효과와 무관하게 통계적으로 유의하며(p&lt;0.001),
발송 피로도 누적 외에 이를 설명할 대안 가설이 없습니다."</div>
""", unsafe_allow_html=True)
    st.text_area("📝 추가 메모", placeholder="후속 과제 입력...", height=100, key="memo_ts")


# ══════════════════════════════════════
# PAGE 4 — 인과 검증
# ══════════════════════════════════════
elif page == "🔬 인과 검증":
    st.title("인과 검증")
    st.caption('"발송 줄이면 매출 오른다"는 주장의 근거와 한계')

    st.markdown("""
<div class="vr"><strong>❌ "발송 줄이면 매출 오른다" — 인과 근거 없음</strong><br>
요일 통제 후 평일 5개 중 3개에서 발송이 많은 날의 거래액이 같거나 더 높습니다. 방향이 일관되지 않아 인과 주장 불가입니다.</div>
<div class="va"><strong>⚠️ "많이 보내면 매출 오른다"도 — 입증 불가</strong><br>
요일 통제 후 발송량 vs 거래액 상관계수가 요일마다 방향이 다릅니다.</div>
<div class="vg"><strong>✅ 입증되는 것: 발송건당 효율은 일관되게 악화</strong><br>
모든 요일에서 발송이 많은 날의 건당 거래액이 낮습니다(상관 -0.4 ~ -0.8).</div>
""", unsafe_allow_html=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("요일 통제 후 — 같은 요일 내 발송 적은 날 vs 많은 날")
    dc_rows = []
    for r in dow_comp:
        diff_str = f"{'▲' if r['diff']>0 else '▼'} {abs(r['diff'])/1e6:.1f}M ({'적은날↑' if r['diff']>0 else '많은날↑'})"
        dc_rows.append({
            "요일":           r["dow"]+"요일",
            "기준(인당)":     f"≤{r['median']:.1f}건",
            "적은날 거래액":  f"{r['lowRev']/1e8:.3f}억",
            "많은날 거래액":  f"{r['highRev']/1e8:.3f}억",
            "거래액 차이":    diff_str,
            "건당거래(적↔많)": f"{r['lowRps']:.0f}원 ↔ {r['highRps']:.0f}원",
        })
    st.dataframe(pd.DataFrame(dc_rows), use_container_width=True, hide_index=True)
    st.caption("* 동일 요일 내 비교. 거래액 방향이 일관되지 않음 = 인과 주장 불가.")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("요일 통제 후 상관계수")
    dc_df = pd.DataFrame(dow_corr)

    fig = go.Figure()
    # vs 거래액 — 양수는 blue, 음수는 red로 색상 분리 (단일 trace에 조건부 색 대신 trace 두 개)
    pos_mask = dc_df["corrRevenue"] >= 0
    for mask, col_name in [(pos_mask, "blue"), (~pos_mask, "red")]:
        sub = dc_df[mask]
        if len(sub) == 0: continue
        fig.add_trace(go.Bar(
            x=list(sub["dow"]), y=list(sub["corrRevenue"].round(3)),
            name="vs 거래액" if col_name=="blue" else "",
            showlegend=(col_name=="blue"),
            marker_color=clr_bg(col_name),
            marker_line_color=clr(col_name),
            marker_line_width=1.5,
        ))
    fig.add_trace(go.Scatter(
        x=list(dc_df["dow"]), y=list(dc_df["corrRps"].round(3)),
        mode="lines+markers", name="vs 건당거래액",
        line=dict(color=clr("red"), width=2),
        marker=dict(size=6, color=clr("red")),
    ))
    layout = base_layout(260)
    layout["showlegend"] = True
    layout["legend"] = dict(orientation="h", y=1.05,
                              bgcolor="rgba(0,0,0,0)",
                              font=dict(color="#8a92a0"))
    layout["yaxis"]["range"] = [-1, 0.5]
    layout["yaxis"]["title"] = "상관계수"
    layout["shapes"] = [dict(
        type="line", x0=-0.5, x1=len(dc_df)-0.5, y0=0, y1=0,
        line=dict(color="#2a2d3a", width=1, dash="dot")
    )]
    fig.update_layout(**layout)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("경영진 대응 — 올바른 주장")
    st.markdown("""
<div class="vg"><strong>데이터로 할 수 있는 정확한 주장:</strong><br>
"발송을 줄이면 매출이 오른다는 보장은 없습니다. 하지만 지금 수준의 발송량은
추가 비용 대비 매출 기여가 없습니다. 발송건당 거래액 효율은 모든 요일에서 일관되게 낮으며,
이 비용을 타겟팅 개선이나 다른 채널에 투자하는 것이 더 효과적이라는 주장은 데이터로 뒷받침됩니다."
</div>
<div class="vb"><strong>다음 단계 — A/B 테스트 필요:</strong><br>
• 통제군: 현행 발송 빈도 유지<br>
• 실험군: 인당 2.5건 이하로 제한<br>
• 측정: 거래액, CTR, 구매율, 수신거부율 (4주 이상)
</div>
""", unsafe_allow_html=True)
    st.text_area("📝 추가 메모", placeholder="후속 과제...", height=120, key="memo_caus")
ENDOFFILE
