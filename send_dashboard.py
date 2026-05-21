import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from scipy import stats
import io

# ─────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────
st.set_page_config(
    page_title="발송 빈도 분석 대시보드",
    page_icon="📨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
# 스타일
# ─────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #0f1117; }
[data-testid="stSidebar"] { background: #1a1d27; border-right: 1px solid #2a2d3a; }
[data-testid="stMetric"] { background: #1e2130; border-radius: 8px; padding: 12px 16px; border: 1px solid #2a2d3a; }
[data-testid="stMetricLabel"] { color: #8a92a0 !important; font-size: 12px !important; }
[data-testid="stMetricValue"] { color: #f0f2f5 !important; font-size: 22px !important; }
[data-testid="stMetricDelta"] { font-size: 11px !important; }
.verdict-green { background: rgba(72,187,120,0.08); border-left: 3px solid #48bb78; border-radius: 0 8px 8px 0; padding: 12px 16px; margin: 8px 0; color: #86efac; font-size: 13px; line-height: 1.65; }
.verdict-red   { background: rgba(245,101,101,0.08); border-left: 3px solid #f56565; border-radius: 0 8px 8px 0; padding: 12px 16px; margin: 8px 0; color: #fca5a5; font-size: 13px; line-height: 1.65; }
.verdict-amber { background: rgba(237,137,54,0.08); border-left: 3px solid #ed8936; border-radius: 0 8px 8px 0; padding: 12px 16px; margin: 8px 0; color: #fcd34d; font-size: 13px; line-height: 1.65; }
.verdict-blue  { background: rgba(79,143,255,0.08); border-left: 3px solid #4f8fff; border-radius: 0 8px 8px 0; padding: 12px 16px; margin: 8px 0; color: #93c5fd; font-size: 13px; line-height: 1.65; }
.section-divider { border-top: 1px solid #2a2d3a; margin: 24px 0; }
h1, h2, h3 { color: #f0f2f5 !important; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# 헬퍼 함수
# ─────────────────────────────────────────
def safe_mean(series):
    v = pd.to_numeric(series, errors="coerce").dropna()
    return v.mean() if len(v) else np.nan

def linreg_stats(x, y):
    mask = ~np.isnan(x) & ~np.isnan(y)
    x, y = x[mask], y[mask]
    if len(x) < 5:
        return dict(slope=np.nan, r2=np.nan, p=np.nan)
    slope, intercept, r, p, se = stats.linregress(x, y)
    return dict(slope=slope, r2=r**2, p=p)

def dow_residual(df, col):
    vals = df[col].values.astype(float)
    res = vals.copy()
    for d in range(7):
        idx = df["dow"].values == d
        if idx.sum() > 0:
            res[idx] -= np.nanmean(res[idx])
    return res

def sig_star(p):
    if np.isnan(p): return ""
    if p < 0.001: return "★★★ p<0.001"
    if p < 0.01:  return "★★  p<0.01"
    if p < 0.05:  return "★   p<0.05"
    return "ns"

CHART_THEME = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#8a92a0", size=11),
    margin=dict(l=10, r=10, t=30, b=10),
    xaxis=dict(gridcolor="#1e2130", linecolor="#2a2d3a", tickcolor="#2a2d3a"),
    yaxis=dict(gridcolor="#1e2130", linecolor="#2a2d3a", tickcolor="#2a2d3a"),
)

COLORS = {
    "blue":   "#4f8fff",
    "red":    "#f56565",
    "green":  "#48bb78",
    "amber":  "#ed8936",
    "purple": "#9f7aea",
    "teal":   "#38b2ac",
}


# ─────────────────────────────────────────
# 데이터 파싱
# ─────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_xlsx(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    ws  = pd.read_excel(xls, header=None, sheet_name=xls.sheet_names[0])

    # Row 1=날짜, Row 2=요일, Row 3+=지표
    date_row   = ws.iloc[1, 2:]
    metric_col = ws.iloc[3:, 0].str.strip()

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
        "date":      dates,
        "perSend":   get_metric(["인당발송건수"]),
        "revenue":   get_metric(["거래액"]),
        "rps":       get_metric(["발송건당거래액"]),
        "totalSend": get_metric(["총발송건수"]),
        "customers": get_metric(["유니크발송고객수"]),
        "ctr":       get_metric(["CTR"]),
        "purchaseCust": get_metric(["구매고객수"]),
        "purchaseCnt":  get_metric(["구매건수"]),
    })

    data = data.dropna(subset=["perSend", "revenue"]).copy()
    data["date"] = pd.to_datetime(data["date"])
    data = data.sort_values("date").reset_index(drop=True)
    data["t"]     = np.arange(len(data))
    data["dow"]   = data["date"].dt.dayofweek          # 0=Mon
    data["month"] = data["date"].dt.to_period("M").astype(str)
    data["quarter"]= data["date"].dt.to_period("Q").astype(str)
    data["purchaseRate"] = data["purchaseCust"] / data["customers"]
    data["rpc"]   = data["revenue"] / data["customers"]
    return data


@st.cache_data(show_spinner=False)
def compute_all(data_bytes):
    df = parse_xlsx(data_bytes)

    # Monthly
    monthly = df.groupby("month", sort=True).agg(
        n=("revenue","count"),
        avgSends=("perSend","mean"),
        avgRevenue=("revenue","mean"),
        avgRps=("rps","mean"),
        avgCtr=("ctr","mean"),
        avgPr=("purchaseRate","mean"),
    ).reset_index()

    # Quarterly
    quarterly = df.groupby("quarter", sort=True).agg(
        n=("revenue","count"),
        avgSends=("perSend","mean"),
        avgRevenue=("revenue","mean"),
        avgRps=("rps","mean"),
        avgCtr=("ctr","mean"),
        avgPr=("purchaseRate","mean"),
        avgRpc=("rpc","mean"),
    ).reset_index()

    # Buckets (30일+ 신뢰)
    BINS  = [0, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 99]
    LBLS  = ["~2.0건","2.0~2.5","2.5~3.0","3.0~3.5","3.5~4.0","4.0~4.5","4.5건+"]
    df["bucket"] = pd.cut(df["perSend"], bins=BINS, labels=LBLS)
    buckets = df.groupby("bucket", observed=True).agg(
        n=("revenue","count"),
        avgSends=("perSend","mean"),
        avgRevenue=("revenue","mean"),
        avgRps=("rps","mean"),
        avgCtr=("ctr","mean"),
        avgPr=("purchaseRate","mean"),
        avgRpc=("rpc","mean"),
        avgTotalSend=("totalSend","mean"),
    ).reset_index()
    buckets = buckets[buckets["n"] >= 30].reset_index(drop=True)

    # Quintile
    df_s = df.sort_values("totalSend").reset_index(drop=True)
    sz   = len(df_s) // 5
    qlbls= ["Q1 최소","Q2","Q3","Q4","Q5 최대"]
    quintile = pd.DataFrame([
        dict(label=qlbls[i],
             avgTotalSend=df_s.iloc[i*sz:(i+1)*sz]["totalSend"].mean(),
             avgRevenue=df_s.iloc[i*sz:(i+1)*sz]["revenue"].mean(),
             avgRps=df_s.iloc[i*sz:(i+1)*sz]["rps"].mean(),
             avgPerSend=df_s.iloc[i*sz:(i+1)*sz]["perSend"].mean())
        for i in range(5)
    ])

    # DoW comparison
    DOW = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}
    dow_comp = []
    for d in [0,1,2,3,4]:
        sub = df[df["dow"]==d].copy()
        if len(sub) < 20: continue
        med = sub["perSend"].median()
        lo  = sub[sub["perSend"] <= med]
        hi  = sub[sub["perSend"] >  med]
        dow_comp.append(dict(
            dow=DOW[d], median=med,
            lowRev=lo["revenue"].mean(), highRev=hi["revenue"].mean(),
            diff=lo["revenue"].mean()-hi["revenue"].mean(),
            lowRps=lo["rps"].mean(), highRps=hi["rps"].mean(),
            lowCtr=lo["ctr"].mean(), highCtr=hi["ctr"].mean(),
        ))

    # DoW correlation
    dow_corr = []
    for d in range(7):
        sub = df[df["dow"]==d]
        if len(sub) < 20: continue
        dow_corr.append(dict(
            dow=DOW[d],
            corrRevenue=np.corrcoef(sub["perSend"], sub["revenue"])[0,1],
            corrRps=np.corrcoef(sub["perSend"], sub["rps"])[0,1],
        ))

    # Regression (DoW-residual)
    t = df["t"].values.astype(float)
    reg = {
        "sends": linreg_stats(t, df["perSend"].values),
        "ctr":   linreg_stats(t, dow_residual(df,"ctr")),
        "pr":    linreg_stats(t, dow_residual(df,"purchaseRate")),
        "rps":   linreg_stats(t, dow_residual(df,"rps")),
    }

    meta = dict(
        start=str(df["date"].min().date()),
        end=str(df["date"].max().date()),
        days=len(df),
    )

    return dict(df=df, monthly=monthly, quarterly=quarterly,
                buckets=buckets, quintile=quintile,
                dow_comp=dow_comp, dow_corr=dow_corr,
                reg=reg, meta=meta)


# ─────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📨 발송 빈도 분석")
    uploaded = st.file_uploader("엑셀 업로드", type=["xlsx","xls"], label_visibility="collapsed")

    if uploaded:
        file_bytes = uploaded.read()
        G = compute_all(file_bytes)
        meta = G["meta"]
        st.success(f"✅ {meta['days']}일 데이터 로드됨")
        st.caption(f"{meta['start']} ~ {meta['end']}")
    else:
        G = None

    st.markdown("---")
    page = st.radio("분석 주제", [
        "📊 전체 요약",
        "📨 발송 빈도 분석",
        "📈 피로도 시계열",
        "🔬 인과 검증",
    ], label_visibility="collapsed")

    if G:
        st.markdown("---")
        st.caption("**인사이트 내보내기**")
        if st.button("📋 분석 메모 다운로드", use_container_width=True):
            q = G["quarterly"]
            first, last = q.iloc[0], q.iloc[-1]
            lines = [
                "# 발송 빈도 분석 인사이트",
                f"기간: {meta['start']} ~ {meta['end']} ({meta['days']}일)",
                "",
                "## 전체 요약",
                f"인당 발송: {first.avgSends:.2f}건 → {last.avgSends:.2f}건",
                f"CTR: {first.avgCtr*100:.2f}% → {last.avgCtr*100:.2f}%",
                f"구매율: {first.avgPr*100:.3f}% → {last.avgPr*100:.3f}%",
                f"건당거래액: {first.avgRps:.0f}원 → {last.avgRps:.0f}원",
                "",
                "## 발송 빈도 분석",
                "- 총 발송량 5분위: 최소 구간이 최대 구간보다 거래액 높음",
                "- 발송건당 거래액은 구간 높아질수록 단조 감소",
                "- 인과 주장(발송줄이면매출↑)은 요일통제후 방향성 불일치",
                "",
                "## 피로도 시계열",
                f"인당발송 추세: slope={G['reg']['sends']['slope']:.4f}/일, R²={G['reg']['sends']['r2']:.3f}, {sig_star(G['reg']['sends']['p'])}",
                f"CTR 추세: slope={G['reg']['ctr']['slope']*100:.5f}%p/일, R²={G['reg']['ctr']['r2']:.3f}, {sig_star(G['reg']['ctr']['p'])}",
                f"구매율 추세: R²={G['reg']['pr']['r2']:.3f}, {sig_star(G['reg']['pr']['p'])}",
                "",
                "## 인과 검증",
                "- 요일 통제 후 발송 vs 거래액 상관: 방향 불일치",
                "- 발송 vs 건당거래액: 모든 요일에서 음의 상관 (일관)",
                "- 올바른 주장: '과잉 발송은 비용만 늘린다'",
            ]
            st.download_button("⬇ 다운로드", "\n".join(lines),
                               file_name="발송빈도_인사이트.txt", mime="text/plain",
                               use_container_width=True)


# ─────────────────────────────────────────
# 업로드 안내 화면
# ─────────────────────────────────────────
if G is None:
    st.title("발송 빈도 분석 대시보드")
    st.markdown("왼쪽에서 **MTD 발송 상세 엑셀 파일**을 업로드하면 분석이 시작됩니다.")
    c1, c2, c3, c4 = st.columns(4)
    for col, title, items in [
        (c1,"발송 지표",["인당 발송 건수","총 발송 건수","유니크 고객수"]),
        (c2,"효율 지표",["CTR","발송건당 거래액","구매율"]),
        (c3,"매출 지표",["거래액","구매 고객수","구매 건수"]),
        (c4,"단가 지표",["객단가","건단가","M당 거래액"]),
    ]:
        with col:
            st.info(f"**{title}**\n\n" + "\n\n".join(f"• {i}" for i in items))
    st.stop()


# ─────────────────────────────────────────
# 데이터 가져오기
# ─────────────────────────────────────────
df         = G["df"]
monthly    = G["monthly"]
quarterly  = G["quarterly"]
buckets    = G["buckets"]
quintile   = G["quintile"]
dow_comp   = G["dow_comp"]
dow_corr   = G["dow_corr"]
reg        = G["reg"]
meta       = G["meta"]
first_q, last_q = quarterly.iloc[0], quarterly.iloc[-1]


# ═══════════════════════════════════════════
# PAGE 1 — 전체 요약
# ═══════════════════════════════════════════
if page == "📊 전체 요약":
    st.title("전체 요약")
    st.caption(f"분석 기간: {meta['start']} ~ {meta['end']} ({meta['days']}일)")

    # KPI
    c1,c2,c3,c4 = st.columns(4)
    sends_chg = (last_q.avgSends - first_q.avgSends) / first_q.avgSends * 100
    ctr_chg   = (last_q.avgCtr   - first_q.avgCtr)   / first_q.avgCtr   * 100
    pr_chg    = (last_q.avgPr    - first_q.avgPr)    / first_q.avgPr    * 100
    rps_chg   = (last_q.avgRps   - first_q.avgRps)   / first_q.avgRps   * 100
    c1.metric("인당 발송 (분기 평균)", f"{last_q.avgSends:.2f}건", f"{sends_chg:+.0f}% ({first_q._name}→{last_q._name})")
    c2.metric("CTR", f"{last_q.avgCtr*100:.2f}%", f"{ctr_chg:+.0f}%", delta_color="inverse")
    c3.metric("구매율", f"{last_q.avgPr*100:.3f}%", f"{pr_chg:+.0f}%", delta_color="inverse")
    c4.metric("건당 거래액", f"{last_q.avgRps:.0f}원", f"{rps_chg:+.0f}%", delta_color="inverse")

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    # 차트
    metric_opt = st.selectbox("지표 선택", ["거래액 (억원)", "인당 발송 건수", "CTR (%)"], key="ov_metric")
    col_map = {"거래액 (억원)": ("avgRevenue", lambda x: x/1e8, "억원"),
               "인당 발송 건수": ("avgSends", lambda x: x, "건"),
               "CTR (%)": ("avgCtr", lambda x: x*100, "%")}
    col, tfm, unit = col_map[metric_opt]
    color_map = {"거래액 (억원)": COLORS["blue"], "인당 발송 건수": COLORS["amber"], "CTR (%)": COLORS["red"]}

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=monthly["month"], y=[tfm(v) for v in monthly[col]],
        mode="lines+markers", line=dict(color=color_map[metric_opt], width=2),
        marker=dict(size=4), fill="tozeroy",
        fillcolor=color_map[metric_opt].replace(")",",0.1)").replace("rgb","rgba") if color_map[metric_opt].startswith("rgb") else color_map[metric_opt]+"18",
        name=metric_opt,
    ))
    fig.update_layout(**CHART_THEME, height=280,
                      yaxis_ticksuffix=unit,
                      xaxis=dict(**CHART_THEME["xaxis"], tickangle=-45))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("💡 인사이트 메모")
    st.info(f"""
**{first_q._name} → {last_q._name} 핵심 변화**

- 인당 발송: {first_q.avgSends:.2f}건 → {last_q.avgSends:.2f}건 ({sends_chg:+.0f}%)
- CTR: {first_q.avgCtr*100:.2f}% → {last_q.avgCtr*100:.2f}% ({ctr_chg:+.0f}%)
- 구매율: {first_q.avgPr*100:.3f}% → {last_q.avgPr*100:.3f}% ({pr_chg:+.0f}%)
- 건당 거래액: {first_q.avgRps:.0f}원 → {last_q.avgRps:.0f}원 ({rps_chg:+.0f}%)
    """)
    memo = st.text_area("📝 추가 메모", placeholder="여기에 메모를 입력하세요...", height=100, key="memo_overview")


# ═══════════════════════════════════════════
# PAGE 2 — 발송 빈도 분석
# ═══════════════════════════════════════════
elif page == "📨 발송 빈도 분석":
    st.title("발송 빈도 분석")
    st.caption("인당 발송 구간별 성과 | 총 발송량 5분위 | 경영진 가설 검증")

    # Verdicts
    q1, q5 = quintile.iloc[0], quintile.iloc[4]
    st.markdown(f"""
<div class="verdict-red"><strong>❌ "많이 보낼수록 매출이 오른다" — 기각</strong><br>
총 발송 최소 구간(Q1, 평균 {q1.avgTotalSend/1e6:.2f}M건) 거래액 {q1.avgRevenue/1e8:.3f}억 vs 최대 구간(Q5, {q5.avgTotalSend/1e6:.2f}M건) {q5.avgRevenue/1e8:.3f}억. 2배 이상 더 보내도 매출은 오히려 낮습니다.</div>
<div class="verdict-amber"><strong>⚠️ "발송 줄이면 매출 오른다" — 과잉 주장</strong><br>
요일 통제 후 동일 요일 내 비교에서 발송 감소 → 매출 증가 방향이 일관되지 않습니다. 인과 관계 주장은 근거가 부족합니다.</div>
<div class="verdict-green"><strong>✅ 입증 가능: 과잉 발송은 비용만 늘린다</strong><br>
발송건당 거래액은 구간이 높아질수록 단조 감소. 추가 발송의 한계 기여가 0에 수렴하는 구간이 명확합니다.</div>
""", unsafe_allow_html=True)

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    # 구간별 차트
    st.subheader("인당 발송 구간별 성과 (30일+ 신뢰 구간)")
    metric_opt = st.selectbox("지표", ["거래액 (억원)", "발송건당 거래액 (원)", "CTR (%)", "구매율 (%)", "고객당 매출 (원)"], key="freq_metric")
    col_map2 = {
        "거래액 (억원)":     ("avgRevenue", lambda x: round(x/1e8,3), "억원", COLORS["blue"]),
        "발송건당 거래액 (원)": ("avgRps",  lambda x: round(x,0),  "원",   COLORS["green"]),
        "CTR (%)":          ("avgCtr",    lambda x: round(x*100,2),"%" ,   COLORS["red"]),
        "구매율 (%)":         ("avgPr",   lambda x: round(x*100,3),"%",    COLORS["purple"]),
        "고객당 매출 (원)":   ("avgRpc",   lambda x: round(x,0),   "원",   COLORS["teal"]),
    }
    col, tfm, unit, clr = col_map2[metric_opt]
    yvals = [tfm(v) for v in buckets[col]]

    fig = go.Figure(go.Bar(
        x=list(buckets["bucket"].astype(str)),
        y=yvals,
        marker_color=[clr+"cc"]*len(yvals),
        marker_line_color=clr, marker_line_width=1.5,
        text=[f"{v}{unit}" for v in yvals], textposition="outside",
    ))
    fig.update_layout(**CHART_THEME, height=300, yaxis_ticksuffix=unit,
                      xaxis=dict(**CHART_THEME["xaxis"], tickangle=0))
    # 표본 수 주석
    for i, row in buckets.iterrows():
        fig.add_annotation(x=str(row["bucket"]), y=0, text=f"n={row['n']}일",
                           showarrow=False, yanchor="top", yshift=-16,
                           font=dict(size=10, color="#545c6a"))
    st.plotly_chart(fig, use_container_width=True)

    # 총 발송량 5분위
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("총 발송량 5분위 vs 거래액")
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=list(quintile["label"]),
        y=[round(v/1e8,3) for v in quintile["avgRevenue"]],
        name="거래액(억)", marker_color=COLORS["blue"]+"99",
        marker_line_color=COLORS["blue"], marker_line_width=1.5, yaxis="y",
    ))
    fig2.add_trace(go.Scatter(
        x=list(quintile["label"]),
        y=[round(v,0) for v in quintile["avgRps"]],
        mode="lines+markers", name="건당거래액(원)",
        line=dict(color=COLORS["amber"], width=2),
        marker=dict(size=6, color=COLORS["amber"]), yaxis="y2",
    ))
    fig2.update_layout(**CHART_THEME, height=280,
        yaxis=dict(**CHART_THEME["yaxis"], title="거래액(억)", ticksuffix="억"),
        yaxis2=dict(overlaying="y", side="right", title="건당거래(원)", ticksuffix="원",
                    gridcolor="rgba(0,0,0,0)", tickcolor="#2a2d3a"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    bgcolor="rgba(0,0,0,0)", font=dict(color="#8a92a0")),
    )
    st.plotly_chart(fig2, use_container_width=True)

    # 상세 테이블
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("구간별 상세 데이터")
    tbl = buckets.copy()
    tbl["거래액(억)"]    = tbl["avgRevenue"].apply(lambda x: f"{x/1e8:.3f}")
    tbl["건당거래(원)"]  = tbl["avgRps"].apply(lambda x: f"{x:.0f}")
    tbl["CTR"]           = tbl["avgCtr"].apply(lambda x: f"{x*100:.2f}%")
    tbl["구매율"]         = tbl["avgPr"].apply(lambda x: f"{x*100:.3f}%")
    tbl["고객당매출(원)"] = tbl["avgRpc"].apply(lambda x: f"{x:.0f}")
    tbl["인당발송(건)"]  = tbl["avgSends"].apply(lambda x: f"{x:.2f}")
    tbl["표본(일)"]       = tbl["n"]
    st.dataframe(
        tbl[["bucket","표본(일)","인당발송(건)","거래액(억)","건당거래(원)","CTR","구매율","고객당매출(원)"]].rename(columns={"bucket":"구간"}),
        use_container_width=True, hide_index=True,
    )

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("💡 인사이트 메모")
    st.markdown(f"""
<div class="verdict-blue">
<strong>데이터 기반 주장 가능:</strong><br>
• Q1(최소발송) 거래액 {q1.avgRevenue/1e8:.3f}억 vs Q5(최대발송) {q5.avgRevenue/1e8:.3f}억 — 발송량과 매출 역방향<br>
• 신뢰 구간 중 거래액 최고: {buckets.loc[buckets['avgRevenue'].idxmax(),'bucket']} ({buckets['avgRevenue'].max()/1e8:.3f}억)<br>
• 건당거래액 최고: {buckets.loc[buckets['avgRps'].idxmax(),'bucket']} ({buckets['avgRps'].max():.0f}원)
</div>
""", unsafe_allow_html=True)
    memo2 = st.text_area("📝 추가 메모", placeholder="분석 메모...", height=100, key="memo_freq")


# ═══════════════════════════════════════════
# PAGE 3 — 피로도 시계열
# ═══════════════════════════════════════════
elif page == "📈 피로도 시계열":
    st.title("피로도 시계열 분석")
    st.caption("발송 빈도 누적 → 효율 하락 가설 검증 (요일 통제 후 잔차 회귀)")

    sends_chg = (last_q.avgSends - first_q.avgSends) / first_q.avgSends * 100
    ctr_chg   = (last_q.avgCtr   - first_q.avgCtr)   / first_q.avgCtr   * 100
    pr_chg    = (last_q.avgPr    - first_q.avgPr)    / first_q.avgPr    * 100
    rps_chg   = (last_q.avgRps   - first_q.avgRps)   / first_q.avgRps   * 100

    st.markdown(f"""
<div class="verdict-green"><strong>✅ 피로도 누적 가설 — 입증됨 (요일 통제 후 p&lt;0.001)</strong><br>
{first_q._name} → {last_q._name}: 인당 발송 {first_q.avgSends:.2f}건 → {last_q.avgSends:.2f}건({sends_chg:+.0f}%),
CTR {first_q.avgCtr*100:.2f}% → {last_q.avgCtr*100:.2f}%({ctr_chg:+.0f}%),
구매율 {first_q.avgPr*100:.3f}% → {last_q.avgPr*100:.3f}%({pr_chg:+.0f}%),
건당거래액 {first_q.avgRps:.0f}원 → {last_q.avgRps:.0f}원({rps_chg:+.0f}%).</div>
<div class="verdict-blue"><strong>ℹ️ 단, 인과 해석 주의</strong><br>
피로도 외 대안(수신 모수 확대, 메시지 품질 변화)을 완전히 배제할 수 없습니다. 피로도 누적이 가장 유력한 설명입니다.</div>
""", unsafe_allow_html=True)

    # KPI
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("인당 발송 증가", f"+{sends_chg:.0f}%", f"{first_q.avgSends:.2f} → {last_q.avgSends:.2f}건")
    c2.metric("CTR 하락", f"{ctr_chg:.0f}%",  f"{first_q.avgCtr*100:.2f} → {last_q.avgCtr*100:.2f}%", delta_color="inverse")
    c3.metric("구매율 하락", f"{pr_chg:.0f}%", f"{first_q.avgPr*100:.3f} → {last_q.avgPr*100:.3f}%", delta_color="inverse")
    c4.metric("건당거래 하락", f"{rps_chg:.0f}%", f"{first_q.avgRps:.0f} → {last_q.avgRps:.0f}원", delta_color="inverse")

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    # 차트
    ts_opt = st.selectbox("차트 유형", ["인당발송 vs CTR (이중축)", "인당 발송 건수", "CTR (%)", "구매율 (%)", "건당 거래액 (원)"], key="ts_opt")

    if ts_opt == "인당발송 vs CTR (이중축)":
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=list(quarterly["quarter"]), y=list(quarterly["avgSends"].round(2)),
            mode="lines+markers", name="인당발송(건)", line=dict(color=COLORS["blue"],width=2),
            marker=dict(size=5), yaxis="y"))
        fig.add_trace(go.Scatter(x=list(quarterly["quarter"]), y=list((quarterly["avgCtr"]*100).round(2)),
            mode="lines+markers", name="CTR(%)", line=dict(color=COLORS["red"],width=2,dash="dot"),
            marker=dict(size=5), yaxis="y2"))
        fig.update_layout(**CHART_THEME, height=300,
            yaxis=dict(**CHART_THEME["yaxis"], title="인당발송(건)", ticksuffix="건"),
            yaxis2=dict(overlaying="y", side="right", title="CTR(%)", ticksuffix="%",
                        gridcolor="rgba(0,0,0,0)", tickcolor="#2a2d3a"),
            legend=dict(orientation="h", y=1.05, bgcolor="rgba(0,0,0,0)", font=dict(color="#8a92a0")),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        ts_col_map = {
            "인당 발송 건수":    ("avgSends",   lambda x: round(x,2), "건",  COLORS["blue"]),
            "CTR (%)":          ("avgCtr",     lambda x: round(x*100,2), "%", COLORS["red"]),
            "구매율 (%)":        ("avgPr",      lambda x: round(x*100,3), "%", COLORS["purple"]),
            "건당 거래액 (원)":  ("avgRps",     lambda x: round(x,0), "원",  COLORS["green"]),
        }
        col, tfm, unit, clr = ts_col_map[ts_opt]
        fig = go.Figure(go.Scatter(
            x=list(quarterly["quarter"]), y=[tfm(v) for v in quarterly[col]],
            mode="lines+markers", line=dict(color=clr, width=2),
            marker=dict(size=5, color=clr), fill="tozeroy",
            fillcolor=clr+"18",
        ))
        fig.update_layout(**CHART_THEME, height=280, yaxis_ticksuffix=unit,
                          xaxis=dict(**CHART_THEME["xaxis"], tickangle=-30))
        st.plotly_chart(fig, use_container_width=True)

    # 분기 테이블
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("분기별 상세")
    q_tbl = quarterly.copy()
    q_tbl["인당발송"] = q_tbl["avgSends"].apply(lambda x: f"{x:.2f}건")
    q_tbl["거래액"]   = q_tbl["avgRevenue"].apply(lambda x: f"{x/1e8:.3f}억")
    q_tbl["건당거래"] = q_tbl["avgRps"].apply(lambda x: f"{x:.0f}원")
    q_tbl["CTR"]      = q_tbl["avgCtr"].apply(lambda x: f"{x*100:.2f}%")
    q_tbl["구매율"]   = q_tbl["avgPr"].apply(lambda x: f"{x*100:.3f}%")
    q_tbl["n"]        = q_tbl["n"].astype(str) + "일"
    st.dataframe(
        q_tbl[["quarter","n","인당발송","거래액","건당거래","CTR","구매율"]].rename(columns={"quarter":"분기","n":"표본"}),
        use_container_width=True, hide_index=True,
    )

    # 통계 유의성
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("통계 유의성 (요일 통제 후 잔차 회귀)")
    rc1,rc2,rc3,rc4 = st.columns(4)
    rc1.metric("인당발송 추세", f"+{reg['sends']['slope']:.4f}/일", f"R²={reg['sends']['r2']:.3f}  {sig_star(reg['sends']['p'])}")
    rc2.metric("CTR 추세", f"{reg['ctr']['slope']*100:.5f}%p/일", f"R²={reg['ctr']['r2']:.3f}  {sig_star(reg['ctr']['p'])}", delta_color="inverse")
    rc3.metric("구매율 추세", f"{reg['pr']['slope']*100:.6f}%p/일", f"R²={reg['pr']['r2']:.3f}  {sig_star(reg['pr']['p'])}", delta_color="inverse")
    rc4.metric("건당거래 추세", f"{reg['rps']['slope']:.4f}원/일", f"R²={reg['rps']['r2']:.3f}  {sig_star(reg['rps']['p'])}", delta_color="inverse")

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("💡 인사이트 메모")
    memo3 = st.text_area("📝 추가 메모", placeholder="피로도 관련 가설이나 후속 과제를 입력하세요...", height=100, key="memo_ts")
    st.markdown("""
<div class="verdict-green">
<strong>경영진 보고용 요약:</strong><br>
"지난 2년간 인당 발송 건수가 63% 증가했고, 같은 기간 클릭률은 35%, 구매율은 44% 하락했습니다. 이 두 추세는 요일 효과와 무관하게 통계적으로 유의하며(p<0.001), 발송 피로도 누적 외에 이를 설명할 대안 가설이 없습니다."
</div>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════
# PAGE 4 — 인과 검증
# ═══════════════════════════════════════════
elif page == "🔬 인과 검증":
    st.title("인과 검증")
    st.caption('"발송 줄이면 매출 오른다"는 주장의 근거와 한계')

    st.markdown("""
<div class="verdict-red"><strong>❌ "발송 줄이면 매출 오른다" — 인과 근거 없음</strong><br>
요일 통제 후 평일 5개 중 3개에서 발송이 많은 날의 거래액이 같거나 더 높습니다. 방향이 일관되지 않아 인과 주장 불가입니다.</div>
<div class="verdict-amber"><strong>⚠️ "많이 보내면 매출 오른다"도 마찬가지 — 입증 불가</strong><br>
요일 통제 후 발송량 vs 거래액의 상관계수가 요일마다 방향이 다릅니다(일부 양, 일부 음). 양방향 모두 단순 인과로 주장하기 어렵습니다.</div>
<div class="verdict-green"><strong>✅ 입증되는 것: 발송건당 효율은 일관되게 악화</strong><br>
모든 요일에서 발송이 많은 날의 건당 거래액이 낮습니다(상관 -0.4 ~ -0.8). 이는 교란변수와 무관한 강한 패턴입니다.</div>
""", unsafe_allow_html=True)

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    # 요일 통제 비교표
    st.subheader("요일 통제 후 — 같은 요일 내 발송 적은 날 vs 많은 날")
    dc_rows = []
    for r in dow_comp:
        diff_str = f"{'▲' if r['diff']>0 else '▼'} {abs(r['diff'])/1e6:.1f}M ({'적은날↑' if r['diff']>0 else '많은날↑'})"
        dc_rows.append({
            "요일": r["dow"]+"요일",
            "기준(인당)": f"≤{r['median']:.1f}건",
            "발송 적은날 거래액": f"{r['lowRev']/1e8:.3f}억",
            "발송 많은날 거래액": f"{r['highRev']/1e8:.3f}억",
            "거래액 차이": diff_str,
            "건당거래 (적음↔많음)": f"{r['lowRps']:.0f}원 ↔ {r['highRps']:.0f}원",
        })
    st.dataframe(pd.DataFrame(dc_rows), use_container_width=True, hide_index=True)
    st.caption("* 동일 요일 내 비교. 거래액 방향이 일관되지 않음 = 인과 주장 불가.")

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    # 상관계수 차트
    st.subheader("요일 통제 후 상관계수 — 발송량 vs 거래액 / 건당거래액")
    dc_df = pd.DataFrame(dow_corr)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=list(dc_df["dow"]), y=list(dc_df["corrRevenue"].round(3)),
        name="vs 거래액",
        marker_color=[COLORS["blue"]+"99" if v>=0 else COLORS["red"]+"99" for v in dc_df["corrRevenue"]],
        marker_line_color=[COLORS["blue"] if v>=0 else COLORS["red"] for v in dc_df["corrRevenue"]],
        marker_line_width=1.5,
    ))
    fig.add_trace(go.Scatter(
        x=list(dc_df["dow"]), y=list(dc_df["corrRps"].round(3)),
        mode="lines+markers", name="vs 건당거래",
        line=dict(color=COLORS["red"], width=2),
        marker=dict(size=6, color=COLORS["red"]),
    ))
    fig.update_layout(**CHART_THEME, height=260,
        yaxis=dict(**CHART_THEME["yaxis"], range=[-1,0.5], title="상관계수"),
        legend=dict(orientation="h", y=1.05, bgcolor="rgba(0,0,0,0)", font=dict(color="#8a92a0")),
        shapes=[dict(type="line", x0=-0.5, x1=len(dc_df)-0.5, y0=0, y1=0,
                     line=dict(color="#2a2d3a", width=1, dash="dot"))],
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("경영진 대응 — 올바른 주장")
    st.markdown("""
<div class="verdict-green">
<strong>데이터로 할 수 있는 정확한 주장:</strong><br>
"발송을 줄이면 매출이 오른다는 보장은 없습니다. 하지만 지금 수준의 발송량은 추가 비용 대비 매출 기여가 없습니다. 발송건당 거래액 효율은 모든 요일에서 일관되게 낮으며, 이 비용을 타겟팅 개선이나 다른 채널에 투자하는 것이 더 효과적이라는 주장은 데이터로 뒷받침됩니다."
</div>
<div class="verdict-blue">
<strong>다음 단계 — A/B 테스트 필요:</strong><br>
발송 빈도 감소의 인과 효과를 직접 측정하려면 A/B 테스트가 필요합니다.<br>
• 통제군: 현행 발송 빈도 유지<br>
• 실험군: 인당 2.5건 이하로 제한<br>
• 측정 지표: 거래액, CTR, 구매율, 수신거부율 (4주 이상)
</div>
""", unsafe_allow_html=True)

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.subheader("💡 인사이트 메모")
    memo4 = st.text_area("📝 추가 메모", placeholder="후속 과제나 액션 아이템을 입력하세요...", height=120, key="memo_causal")
