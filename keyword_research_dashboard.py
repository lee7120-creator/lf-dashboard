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

def load():
    # 캐시 미사용: @st.cache_data가 이전 배포의 구버전 DataFrame을 물려주면
    # 신규 컬럼(순위 등)이 없어 KeyError가 나므로, 매 실행 새로 읽는다.
    df = pd.read_csv("data/lfmall_keyword_research.csv", encoding="utf-8-sig")
    # 구버전 CSV 호환 방어: 컬럼 누락 시 기본값 채움
    for col, default in [("순위", 0), ("우선순위", ""), ("패션", "N"),
                         ("섹션", "기타"), ("Status", "미수집")]:
        if col not in df.columns:
            df[col] = default
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
n_missing = int((df["Status"] == "Missing").sum())
n_fashion = int((df["순위"] > 0).sum())

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("카테고리 키워드", f"{len(df):,}개", f"검색량 보유 {have_vol}개")
k2.metric("총 검색량(MSV)", f"{int(df['검색량'].sum()):,}", "월 합계")
k3.metric("섹션", f"{df['섹션'].nunique()}개", "카테고리 그룹")
k4.metric("🔴 Missing(선점)", f"{n_missing}개", "경쟁사만 보유")
k5.metric("🎯 패션 우선순위", f"{n_fashion}개", "1순위~ 부여")

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
    st.markdown("##### 패션 카테고리 선점 우선순위 (1순위 = 최우선)")
    st.caption("패션·뷰티·골프 카테고리 키워드를 대상으로, √검색량·난이도·경쟁상태를 종합해 "
               "1순위부터 서수로 부여. 가전·리빙·식품 등 비패션 일반어는 제외됩니다.")
    rank_df = df[df["순위"] > 0]
    c0, c1 = st.columns([1, 1])
    sec_f = c0.multiselect("섹션 필터", sorted(rank_df["섹션"].unique()))
    topn = c1.slider("상위 N개", 10, 60, 30, step=5)
    d = rank_df if not sec_f else rank_df[rank_df["섹션"].isin(sec_f)]
    d = d.sort_values("순위").head(topn)

    cc1, cc2 = st.columns([1.1, 1])
    with cc1:
        dd = d.sort_values("순위", ascending=False)
        fig = go.Figure(go.Bar(
            x=dd["검색량"], y=dd["키워드"], orientation="h",
            marker=dict(color=[STATUS_COLOR.get(s, "#cbd5e1") for s in dd["Status"]]),
            text=dd["우선순위"], textposition="outside",
            customdata=dd[["우선순위", "섹션", "Status"]],
            hovertemplate="<b>%{y}</b><br>%{customdata[0]} · 검색량 %{x:,}<br>"
                          "%{customdata[1]} · %{customdata[2]}<extra></extra>"))
        fig.update_layout(**base_layout(h=max(420, topn * 16),
                                        title="패션 우선순위 (막대=검색량, 색=Status)"))
        fig.update_xaxes(range=[0, dd["검색량"].max() * 1.18])
        st.plotly_chart(fig, use_container_width=True)
    with cc2:
        st.dataframe(
            d[["우선순위", "키워드", "섹션", "검색량", "Status"]],
            use_container_width=True, hide_index=True, height=max(420, topn * 16),
            column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d")})

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
        column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d")})

# ──────────────────────────────────────────────────────
# TAB 5 — 가이드
# ──────────────────────────────────────────────────────
with tab5:
    st.markdown("##### 📅 데이터 기간 기준")
    st.markdown(
        "<div class='card'>"
        "• <b>출처/시점</b> — Semrush 한국(<b>kr</b>) DB, <b>2026-06-22 스냅샷</b>. "
        "별도 기간을 지정하지 않아 <b>가장 최신 가용 데이터</b>가 반영됨.<br>"
        "• <b>검색량은 단일 월이 아니라 \"최근 12개월 평균 월간 검색량\"</b>(rolling 12-month average)입니다. "
        "즉 한 달에 평균 몇 번 검색되는지를 직전 1년으로 평균낸 값.<br>"
        "• 순위·난이도도 같은 2026-06-22 시점 기준.<br>"
        "• ⚠️ <b>계절 키워드 주의</b> — 패딩·수영복처럼 시즌성이 강한 키워드는 성수기 단월 검색량이 "
        "연평균보다 훨씬 큽니다(현재 값은 연평균).</div>",
        unsafe_allow_html=True)

    st.markdown("##### 📖 용어 설명")
    st.markdown(
        "<div class='card'>"
        "• <b>검색량(MSV, Monthly Search Volume)</b> — 한 달 평균 검색 횟수(최근 12개월 평균)<br>"
        "• <b>난이도(KD, Keyword Difficulty)</b> — 0~100, 상위노출 난이도. <b>낮을수록 쉬움</b>. 빈칸 = 미집계<br>"
        "• <b>도메인 순위(LF몰/W컨셉/한섬/SSF샵/SI빌리지)</b> — 그 키워드 검색 시 해당 사이트의 순위(1~100위). "
        "<b>0 = 100위 밖(노출 없음)</b><br>"
        "• <b>섹션</b> — 키워드를 카테고리 그룹(의류·가방·뷰티…)으로 자동 분류한 것<br>"
        "• <b>우선순위(N순위)</b> — 페이지를 만들어 공략할 순서. <b>1순위 = 최우선</b><br>"
        f"• <b>Status</b> — "
        f"<span class='tag' style='background:{STATUS_COLOR['Strong']}'>Strong</span> 경쟁사보다 앞섬 / "
        f"<span class='tag' style='background:{STATUS_COLOR['Weak']}'>Weak</span> 열위 / "
        f"<span class='tag' style='background:{STATUS_COLOR['Missing']}'>Missing</span> 경쟁사만 보유(선점기회) / "
        f"<span class='tag' style='background:{STATUS_COLOR['공백']}'>공백</span> 5사 모두 미보유(무경쟁) / "
        f"<span class='tag' style='background:{STATUS_COLOR['미수집']};color:#475569'>미수집</span> 순위 미집계"
        "</div>",
        unsafe_allow_html=True)

    st.markdown("##### 🧮 연산 기준")
    st.markdown(
        "<div class='card'>"
        "<b>① 우선순위 점수</b><br>"
        "<code>점수 = √(검색량) × 난이도할인 × 상태가중</code><br>"
        "• <b>√(검색량)</b> — 제곱근. 검색량이 매우 큰 키워드(뮬·시계·크림 등)가 점수를 독점하지 않도록 "
        "영향력을 <b>완화</b>합니다.<br>"
        "• <b>난이도할인 = (100 − KD) / 100</b> — 쉬울수록 1에 가까움. 난이도 미상은 40으로 가정(→0.6).<br>"
        "• <b>상태가중</b> — Missing 1.0 · 공백 0.9 · 미수집 0.65 · Weak 0.6 · Strong 0.4(10위 밖이면 0.5). "
        "선점기회일수록 우대.<br>"
        "• 이 점수 내림차순으로 <b>1순위·2순위…</b> 부여. 대상은 <b>패션 + 뷰티 + 골프 섹션 & 검색량&gt;0</b> "
        "키워드뿐(가전·리빙·식품 등 비패션 일반어는 제외 — '야구' 같은 거품 방지).</div>"
        "<div class='card'>"
        "<b>② Status 판정(LF몰 기준)</b><br>"
        "• LF몰=0 &amp; 경쟁사도 모두 0 → <b>공백</b><br>"
        "• LF몰=0 &amp; 경쟁사 중 하나라도 순위 있음 → <b>Missing</b><br>"
        "• LF몰 순위 &lt; 경쟁사 최고순위(또는 경쟁사 없음) → <b>Strong</b><br>"
        "• 그 외(경쟁사가 더 앞섬) → <b>Weak</b> · 순위 데이터 자체가 없으면 → <b>미수집</b></div>",
        unsafe_allow_html=True)

    st.markdown("##### 🎯 활용 & 한계")
    st.markdown(
        "<div class='card'>"
        "• <b>활용</b> — 1순위부터 허브-스포크 템플릿 페이지를 만들어 선점. 특히 경쟁사 미보유"
        "(Missing/공백) 상위 키워드가 가장 빠른 선점 기회.<br>"
        "• <b>한계</b> — 순위/난이도는 5사 키워드갭에서 확보된 키워드에 한해 병합되어, 비패션 등 "
        "상당수는 <b>미수집</b>입니다. 전체 919개의 5사 SERP 순위를 모두 채우려면 Semrush 추가 수집이 "
        "필요합니다. 검색량은 Semrush 추정치로 실제와 차이가 있을 수 있습니다.</div>",
        unsafe_allow_html=True)

st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)
st.caption(f"카테고리 키워드 {len(df):,}개 · LF몰 vs W컨셉·한섬·SSF샵·SI빌리지 · Semrush kr DB")
