"""LF몰 전체 카테고리 키워드 리서치 대시보드 (Semrush 실데이터)

대상   : lfmall.co.kr (LF몰)
경쟁사 : W컨셉 · 한섬 · SSF샵 · SI빌리지
데이터 : data/lfmall_keyword_research.csv
         · 검색량(MSV) = Semrush kr DB phrase_these 실측 (919개 카테고리 키워드)
         · 순위/난이도 = 5사 키워드갭(domain_domains) 병합 — 패션 카테고리 위주
         · Status(LF몰 기준): Strong/Weak/Missing/공백(5사 모두 미보유)/미수집

갱신: python build_keyword_data.py  →  data/lfmall_keyword_research.csv 재생성
실행: streamlit run keyword_research_dashboard.py
"""

import io
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="LF몰 키워드 리서치", page_icon="🔎",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.card{border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin:8px 0;line-height:1.6;background:#ffffff}
.sdiv{border-top:1px solid #e2e8f0;margin:20px 0}
.tag{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:6px;margin-right:6px;color:white}
.cap{color:#64748b;font-size:13px}
h1,h2,h3{color:#1e293b}
</style>
""", unsafe_allow_html=True)

STATUS_COLOR = {"Strong": "#48bb78", "Weak": "#ed8936", "Missing": "#f56565",
                "공백": "#4f8fff", "미수집": "#cbd5e1"}
SITES = ["LF몰", "W컨셉", "한섬", "SSF샵", "SI빌리지"]

@st.cache_data
def load():
    df = pd.read_csv("data/lfmall_keyword_research.csv")
    return df

df = load()

def base_layout(h=320, title="", showlegend=False):
    return dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#475569", size=12), margin=dict(l=10, r=10, t=42, b=10),
                height=h, showlegend=showlegend,
                title=dict(text=title, font=dict(color="#94a3b8", size=13)),
                legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1,
                            font=dict(size=11)),
                xaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                           tickfont=dict(color="#64748b", size=11)),
                yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                           tickfont=dict(color="#64748b", size=11)))

def to_excel(d):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        d.to_excel(w, index=False, sheet_name="키워드리서치")
    return buf.getvalue()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ══════════════════════════════════════════════════════
# HEADER + KPI
# ══════════════════════════════════════════════════════
st.markdown("## 🔎 LF몰 전체 카테고리 키워드 리서치")
st.markdown(
    "<div class='cap'>대상 <b>lfmall.co.kr</b> · 경쟁사 <b>W컨셉·한섬·SSF샵·SI빌리지</b> · "
    f"카테고리 키워드 <b>{len(df):,}</b>개 · Semrush 한국(kr) DB 검색량 실측</div>",
    unsafe_allow_html=True)
st.write("")

have_vol = int((df["검색량"] > 0).sum())
have_pos = int((df["Status"] != "미수집").sum())
n_missing = int((df["Status"] == "Missing").sum())
n_white = int((df["Status"] == "공백").sum())

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("카테고리 키워드", f"{len(df):,}개", f"검색량 보유 {have_vol}개")
k2.metric("총 검색량(MSV)", f"{int(df['검색량'].sum()):,}", "월 합계")
k3.metric("섹션", f"{df['섹션'].nunique()}개", "카테고리 그룹")
k4.metric("🔴 Missing(선점)", f"{n_missing}개", "경쟁사만 보유")
k5.metric("🔵 공백(화이트스페이스)", f"{n_white}개", "5사 모두 미보유")

st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🔎 키워드 리서치")
    st.caption("LF몰 전체 카테고리 · Semrush kr")
    st.markdown("---")
    st.markdown("**Status (LF몰 기준)**")
    st.markdown(
        f"<span class='tag' style='background:{STATUS_COLOR['Strong']}'>Strong</span> 경쟁사보다 앞섬<br>"
        f"<span class='tag' style='background:{STATUS_COLOR['Weak']}'>Weak</span> 경쟁사보다 열위<br>"
        f"<span class='tag' style='background:{STATUS_COLOR['Missing']}'>Missing</span> 경쟁사만 보유<br>"
        f"<span class='tag' style='background:{STATUS_COLOR['공백']}'>공백</span> 5사 모두 미보유<br>"
        f"<span class='tag' style='background:{STATUS_COLOR['미수집']};color:#475569'>미수집</span> 순위 미집계",
        unsafe_allow_html=True)
    st.markdown("---")
    st.caption("검색량은 Semrush 실측(phrase_these). 순위/난이도는 5사 키워드갭에서 "
               "확보된 패션 키워드에 한해 병합. 비패션 카테고리의 경쟁사 순위는 "
               "추가 SERP 수집이 필요.")

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["🚀 우선순위 타겟", "🗂️ 섹션 분석", "🎯 경쟁사 Status", "📋 전체 데이터(엑셀)", "📖 가이드"])

# ──────────────────────────────────────────────────────
# TAB 1 — 우선순위 타겟
# ──────────────────────────────────────────────────────
with tab1:
    st.markdown("##### 검색량 기반 선점 우선순위")
    st.caption("우선순위 = 검색량 × 난이도할인 × 상태가중(Missing/공백 우대). "
               "고검색량·미보유 카테고리부터 페이지를 만들어 선점.")
    c0, c1 = st.columns([1, 1])
    sec_f = c0.multiselect("섹션 필터", sorted(df["섹션"].unique()))
    topn = c1.slider("상위 N개", 10, 60, 30, step=5)
    d = df if not sec_f else df[df["섹션"].isin(sec_f)]
    d = d.sort_values("우선순위", ascending=False).head(topn)

    cc1, cc2 = st.columns([1.1, 1])
    with cc1:
        dd = d.sort_values("우선순위", ascending=True)
        fig = go.Figure(go.Bar(
            x=dd["우선순위"], y=dd["키워드"], orientation="h",
            marker=dict(color=[STATUS_COLOR.get(s, "#cbd5e1") for s in dd["Status"]]),
            customdata=dd[["검색량", "섹션", "Status"]],
            hovertemplate="<b>%{y}</b><br>우선순위 %{x:,}<br>검색량 %{customdata[0]:,} · "
                          "%{customdata[1]} · %{customdata[2]}<extra></extra>"))
        fig.update_layout(**base_layout(h=max(420, topn * 16), title="우선순위 (색=Status)"))
        st.plotly_chart(fig, use_container_width=True)
    with cc2:
        st.dataframe(
            d[["키워드", "섹션", "검색량", "Status", "우선순위"]],
            use_container_width=True, hide_index=True, height=max(420, topn * 16),
            column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d"),
                           "우선순위": st.column_config.ProgressColumn(
                               "우선순위", format="%d", min_value=0,
                               max_value=int(df["우선순위"].max()))})

# ──────────────────────────────────────────────────────
# TAB 2 — 섹션 분석
# ──────────────────────────────────────────────────────
with tab2:
    st.markdown("##### 섹션별 검색 수요 & 키워드 분포")
    sec = (df.groupby("섹션").agg(키워드수=("키워드", "count"), 총검색량=("검색량", "sum"))
             .reset_index().sort_values("총검색량", ascending=False))
    c1, c2 = st.columns([1, 1])
    with c1:
        dd = sec.sort_values("총검색량", ascending=True)
        fig = go.Figure(go.Bar(x=dd["총검색량"], y=dd["섹션"], orientation="h",
                               marker_color="#4f8fff",
                               text=[f"{int(v):,}" for v in dd["총검색량"]],
                               textposition="outside"))
        fig.update_layout(**base_layout(h=560, title="섹션별 총 검색량"))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.treemap(df[df["검색량"] > 0], path=[px.Constant("전체"), "섹션", "키워드"],
                         values="검색량", color="섹션")
        fig.update_traces(marker=dict(line=dict(color="white", width=1)),
                          hovertemplate="<b>%{label}</b><br>검색량 %{value:,}<extra></extra>")
        fig.update_layout(margin=dict(l=4, r=4, t=10, b=4), height=560)
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(sec, use_container_width=True, hide_index=True,
                 column_config={"총검색량": st.column_config.ProgressColumn(
                     "총검색량", format="%d", min_value=0, max_value=int(sec["총검색량"].max()))})

# ──────────────────────────────────────────────────────
# TAB 3 — 경쟁사 Status
# ──────────────────────────────────────────────────────
with tab3:
    st.markdown("##### 경쟁사 대비 Status (순위 확보된 키워드)")
    st.caption("5사 키워드갭에서 순위가 확인된 키워드의 LF몰 경쟁 위치. "
               "Missing=경쟁사만 보유(선점 기회), 공백=5사 모두 미보유(무경쟁 화이트스페이스).")
    known = df[df["Status"] != "미수집"].copy()
    cc1, cc2 = st.columns([1, 1.4])
    with cc1:
        sc = known["Status"].value_counts()
        fig = go.Figure(go.Bar(x=sc.index, y=sc.values,
                               marker_color=[STATUS_COLOR.get(s, "#cbd5e1") for s in sc.index],
                               text=sc.values, textposition="outside"))
        fig.update_layout(**base_layout(h=320, title="Status 분포(순위 확보분)"))
        st.plotly_chart(fig, use_container_width=True)
    with cc2:
        st.dataframe(
            known.sort_values("검색량", ascending=False)[
                ["키워드", "섹션", "검색량", "난이도"] + SITES + ["Status"]],
            use_container_width=True, hide_index=True, height=320,
            column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d")})
    st.info("비패션 카테고리(가전·뷰티·리빙·식품 등)의 경쟁사 순위는 미수집 상태입니다. "
            "전체 919개 키워드의 5사 SERP 순위를 모두 채우려면 Semrush 추가 수집이 필요합니다.")

# ──────────────────────────────────────────────────────
# TAB 4 — 전체 데이터 + 엑셀/CSV 다운로드
# ──────────────────────────────────────────────────────
with tab4:
    st.markdown("##### 전체 키워드 리서치 데이터 (필터 + 다운로드)")
    f0, f1, f2, f3 = st.columns([1.3, 1.3, 1, 1])
    fsec = f0.multiselect("섹션", sorted(df["섹션"].unique()))
    fstat = f1.multiselect("Status", list(STATUS_COLOR.keys()))
    fmsv = f2.slider("최소 검색량", 0, int(df["검색량"].max()), 0, step=1000)
    q = f3.text_input("키워드 검색", "")
    fdf = df[df["검색량"] >= fmsv]
    if fsec:
        fdf = fdf[fdf["섹션"].isin(fsec)]
    if fstat:
        fdf = fdf[fdf["Status"].isin(fstat)]
    if q:
        fdf = fdf[fdf["키워드"].str.contains(q, case=False, na=False)]
    fdf = fdf.sort_values("검색량", ascending=False)

    st.markdown(f"**{len(fdf):,}개** 키워드")
    d1, d2 = st.columns(2)
    d1.download_button("⬇️ 엑셀(.xlsx) 다운로드", to_excel(fdf),
                       "lfmall_keyword_research.xlsx", XLSX_MIME, use_container_width=True)
    d2.download_button("⬇️ CSV 다운로드", fdf.to_csv(index=False).encode("utf-8-sig"),
                       "lfmall_keyword_research.csv", "text/csv", use_container_width=True)
    st.dataframe(
        fdf[["키워드", "섹션", "검색량", "난이도"] + SITES + ["Status", "우선순위"]],
        use_container_width=True, hide_index=True, height=520,
        column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d"),
                       "우선순위": st.column_config.NumberColumn("우선순위", format="%d")})

# ──────────────────────────────────────────────────────
# TAB 5 — 가이드
# ──────────────────────────────────────────────────────
with tab5:
    st.markdown(
        "<div class='card'>"
        "<b>분석 대상</b> — LF몰(lfmall.co.kr)의 전체 카테고리 분류 키워드 919개를 "
        "Semrush 한국(kr) DB로 검색량 실측하고, 경쟁 4사(W컨셉·한섬·SSF샵·SI빌리지)와 "
        "비교했습니다.</div>"
        "<div class='card'>"
        "<b>지표</b><br>"
        "• 검색량(MSV) = Google 월별 검색량 · 도메인 숫자 = 순위(1~100, 0=없음)<br>"
        f"• <span class='tag' style='background:{STATUS_COLOR['Strong']}'>Strong</span> LF몰이 경쟁사보다 앞섬 · "
        f"<span class='tag' style='background:{STATUS_COLOR['Weak']}'>Weak</span> 열위 · "
        f"<span class='tag' style='background:{STATUS_COLOR['Missing']}'>Missing</span> 경쟁사만 보유<br>"
        f"• <span class='tag' style='background:{STATUS_COLOR['공백']}'>공백</span> 5사 모두 미보유(무경쟁 선점지) · "
        f"<span class='tag' style='background:{STATUS_COLOR['미수집']};color:#475569'>미수집</span> 순위 미집계</div>"
        "<div class='card'>"
        "<b>프로그래매틱 SEO 활용</b> — 검색량이 크고 경쟁사가 미보유(Missing/공백)인 "
        "카테고리부터 허브-스포크 템플릿 페이지를 만들어 선점. 우선순위 탭이 그 순서를 제시합니다.</div>",
        unsafe_allow_html=True)
    st.caption("*검색량은 Semrush 추정치. 순위/난이도는 5사 키워드갭에서 확보된 패션 키워드 위주로 "
               "병합되어 일부 비패션 키워드는 '미수집'입니다.")

st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)
st.caption(f"카테고리 키워드 {len(df):,}개 · LF몰 vs W컨셉·한섬·SSF샵·SI빌리지 · Semrush kr DB")
