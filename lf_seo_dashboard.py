"""LF몰 SEO 통합 대시보드

사이드바에서 3개 뷰 전환:
  1) 경쟁사 분석   — LF몰 vs W컨셉·한섬·SSF샵·SI빌리지 (패션 62개, Semrush)
  2) 키워드 리서치 — 전체 카테고리 919개 (Semrush 검색량) + 엑셀 다운로드
  3) 네이버 쇼핑   — 네이버 검색광고/데이터랩 지표 (키 입력 후 활성)

데이터:
  data/lfmall_keyword_research.csv     (키워드 리서치)
  data/naver_keyword_metrics.csv       (네이버, 있으면 자동 표시)
  경쟁사 62개는 본 파일에 내장

실행: streamlit run lf_seo_dashboard.py
"""

import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="LF몰 SEO 대시보드", page_icon="🧭",
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

# ── 공통 ──
STATUS_COLOR = {"Strong": "#48bb78", "Weak": "#ed8936", "Missing": "#f56565",
                "공백": "#4f8fff", "미수집": "#cbd5e1"}
PALETTE = {"blue": "#4f8fff", "red": "#f56565", "amber": "#ed8936",
           "green": "#48bb78", "purple": "#9f7aea", "slate": "#64748b"}
SITE_COLOR = {"LF몰": PALETTE["blue"], "W컨셉": PALETTE["amber"], "한섬": PALETTE["purple"],
              "SSF샵": PALETTE["red"], "SI빌리지": PALETTE["slate"]}
SITES = ["LF몰", "W컨셉", "한섬", "SSF샵", "SI빌리지"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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


def to_excel(d, sheet="data"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        d.to_excel(w, index=False, sheet_name=sheet)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════
# VIEW 1 — 경쟁사 분석 (패션 62개)
# ══════════════════════════════════════════════════════════════════
SUBJECT, COMPETITORS = "LF몰", ["W컨셉", "한섬", "SSF샵", "SI빌리지"]
COMP_DATA = [
    ("패딩", 33100, 21, 0, 11, 0, 0, 0, "아우터"), ("코트", 27100, 23, 0, 0, 0, 18, 0, "아우터"),
    ("재킷", 22200, 32, 0, 0, 49, 8, 0, "아우터"), ("모피", 22200, 29, 5, 0, 0, 0, 0, "아우터"),
    ("가디건", 12100, 20, 0, 0, 0, 14, 0, "아우터"), ("무스탕", 12100, 20, 13, 2, 0, 0, 0, "아우터"),
    ("트렌치 코트", 12100, 27, 0, 0, 0, 7, 0, "아우터"), ("경량 패딩", 12100, 16, 5, 0, 0, 20, 0, "아우터"),
    ("바람막이", 14800, 19, 0, 0, 0, 20, 0, "아우터"), ("블레이저", 12100, 28, 0, 0, 0, 29, 0, "아우터"),
    ("블루종", 6600, 28, 15, 5, 65, 9, 0, "아우터"),
    ("블라우스", 22200, 24, 0, 2, 0, 16, 0, "상의"), ("셔츠", 22200, 29, 0, 23, 0, 24, 0, "상의"),
    ("니트", 18100, 19, 43, 2, 0, 8, 0, "상의"), ("시스루", 18100, 27, 0, 0, 0, 9, 0, "상의"),
    ("티셔츠", 12100, 16, 0, 5, 0, 32, 0, "상의"), ("후드 집업", 12100, 14, 0, 5, 0, 23, 0, "상의"),
    ("크롭 티", 9900, 17, 32, 2, 0, 3, 0, "상의"), ("탱크 탑", 8100, 23, 0, 0, 0, 21, 0, "상의"),
    ("오프 숄더", 8100, 17, 0, 3, 0, 0, 0, "상의"),
    ("청바지", 74000, 23, 0, 0, 0, 27, 0, "하의"), ("바지", 40500, 19, 0, 0, 0, 14, 0, "하의"),
    ("돌핀 팬츠", 22200, 27, 0, 0, 0, 5, 0, "하의"), ("데님", 22200, 22, 0, 0, 0, 46, 0, "하의"),
    ("미니 스커트", 14800, 31, 18, 0, 0, 19, 0, "하의"), ("버뮤다 팬츠", 14800, 18, 20, 0, 0, 24, 0, "하의"),
    ("슬랙스", 14800, 16, 0, 14, 0, 0, 0, "하의"), ("치마", 14800, 24, 0, 0, 0, 43, 0, "하의"),
    ("반바지", 8100, 16, 0, 0, 0, 24, 0, "하의"),
    ("드레스", 22200, 31, 35, 0, 0, 0, 0, "드레스/수영"), ("비키니", 60500, 35, 0, 5, 0, 0, 0, "드레스/수영"),
    ("모노 키니", 12100, 12, 23, 2, 0, 0, 0, "드레스/수영"),
    ("부츠", 12100, 26, 0, 3, 0, 48, 0, "신발"), ("구두", 9900, 39, 0, 0, 0, 30, 0, "신발"),
    ("로퍼", 12100, 20, 0, 19, 0, 0, 0, "신발"), ("샌들", 6600, 21, 30, 11, 0, 0, 0, "신발"),
    ("더비 슈즈", 8100, 17, 10, 6, 0, 0, 0, "신발"),
    ("가방", 33100, 35, 0, 5, 0, 0, 0, "가방"), ("크로스 백", 12100, 16, 0, 2, 0, 0, 0, "가방"),
    ("백팩", 9900, 20, 0, 0, 0, 24, 0, "가방"), ("토트 백", 9900, 19, 0, 2, 0, 46, 0, "가방"),
    ("메신저 백", 6600, 23, 0, 20, 0, 29, 0, "가방"), ("에코 백", 8100, 10, 26, 3, 0, 0, 0, "가방"),
    ("모자", 18100, 18, 0, 0, 0, 12, 0, "액세서리"), ("지갑", 14800, 20, 0, 0, 0, 30, 0, "액세서리"),
    ("키링", 14800, 18, 0, 0, 0, 49, 0, "액세서리"), ("넥타이", 9900, 32, 0, 0, 0, 21, 0, "액세서리"),
    ("벨트", 6600, 26, 0, 0, 67, 0, 0, "액세서리"), ("뿔테 안경", 9900, 20, 17, 0, 0, 19, 0, "액세서리"),
    ("반다나", 9900, 23, 12, 6, 0, 24, 0, "액세서리"), ("목도리", 8100, 20, 8, 3, 0, 0, 0, "액세서리"),
    ("비니", 8100, 27, 0, 0, 0, 23, 0, "액세서리"), ("카드 지갑", 8100, 19, 23, 11, 0, 42, 0, "액세서리"),
    ("니삭스", 6600, 31, 0, 2, 0, 12, 0, "액세서리"),
    ("속옷", 18100, 23, 0, 0, 0, 19, 0, "언더웨어"), ("여자 속옷", 8100, 14, 9, 26, 0, 0, 0, "언더웨어"),
    ("여자 팬티", 9900, 20, 2, 0, 0, 0, 0, "언더웨어"), ("티 팬티", 12100, 29, 23, 0, 0, 19, 0, "언더웨어"),
    ("잠옷", 8100, 19, 19, 0, 0, 0, 0, "언더웨어"),
    ("정장", 18100, 20, 0, 0, 26, 15, 0, "정장/포멀"), ("세미 정장", 6600, 21, 41, 0, 0, 0, 0, "정장/포멀"),
    ("버건디", 9900, 29, 0, 7, 0, 0, 0, "컬러"),
]


@st.cache_data
def comp_df():
    df = pd.DataFrame(COMP_DATA, columns=["키워드", "MSV", "KD"] + SITES + ["카테고리"])

    def status(r):
        if r[SUBJECT] == 0:
            return "Missing"
        rk = [r[c] for c in COMPETITORS if r[c] > 0]
        return "Strong" if (not rk or r[SUBJECT] < min(rk)) else "Weak"

    def tier(r):
        s, kd, msv, gw = r["Status"], r["KD"], r["MSV"], r[SUBJECT]
        if s == "Missing":
            if kd <= 25 and msv >= 10000:
                return "🥇 즉시 선점"
            return "🥈 선점 후보" if (kd <= 25 or msv >= 10000) else "🥉 롱테일 선점"
        if s == "Weak":
            return "🔧 최적화(탈환)"
        return "🔧 최적화(순위↑)" if (s == "Strong" and gw > 10) else "✅ 방어"

    def prio(r):
        sf = {"Missing": 1.0, "Weak": 0.65}.get(
            r["Status"], 0.45 if r[SUBJECT] > 10 else 0.15)
        return int(round(r["MSV"] * (100 - r["KD"]) / 100 * sf))

    df["Status"] = df.apply(status, axis=1)
    df["타겟 등급"] = df.apply(tier, axis=1)
    df["우선순위"] = df.apply(prio, axis=1)
    return df


def render_competitor():
    df = comp_df()
    st.markdown("## 🥊 경쟁사 분석 — LF몰 vs W컨셉·한섬·SSF샵·SI빌리지")
    st.markdown("<div class='cap'>패션 카테고리 키워드 62개 · Semrush 한국(kr) DB · "
                "순위 1~100(0=없음) · Status는 LF몰 기준</div>", unsafe_allow_html=True)
    st.write("")
    n_missing = int((df["Status"] == "Missing").sum())
    prime = df[(df["Status"] == "Missing") & (df["KD"] <= 25) & (df["MSV"] >= 10000)]
    k = st.columns(5)
    k[0].metric("분석 키워드", f"{len(df)}개")
    k[1].metric("🟢 Strong", int((df["Status"] == "Strong").sum()))
    k[2].metric("🟠 Weak", int((df["Status"] == "Weak").sum()))
    k[3].metric("🔴 Missing", n_missing, "경쟁사만 보유")
    k[4].metric("🥇 즉시 선점", f"{len(prime)}개", "고MSV·저난이도")
    st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

    t1, t2, t3, t4 = st.tabs(["🎯 기회 매트릭스", "🚀 선점 타겟", "🔍 도메인 커버리지", "📋 진단"])

    with t1:
        st.markdown("##### 기회 매트릭스 — 좌상단(고검색량·저난이도)이 황금존")
        fig = go.Figure()
        fig.add_shape(type="rect", x0=0, x1=25, y0=10000, y1=df["MSV"].max() * 1.05,
                      fillcolor="rgba(72,187,120,0.08)", line=dict(width=0), layer="below")
        for s in ["Missing", "Weak", "Strong"]:
            d = df[df["Status"] == s]
            fig.add_trace(go.Scatter(
                x=d["KD"], y=d["MSV"], mode="markers", name=s,
                marker=dict(size=(d["MSV"] / 1800 + 8), color=STATUS_COLOR[s],
                            opacity=0.78, line=dict(color="white", width=1)),
                customdata=d[["키워드", "타겟 등급", SUBJECT]],
                hovertemplate="<b>%{customdata[0]}</b><br>MSV %{y:,} · KD %{x}"
                              "<br>LF몰 %{customdata[2]}위<br>%{customdata[1]}<extra></extra>"))
        fig.update_layout(**base_layout(h=440, showlegend=True))
        fig.update_xaxes(title="Keyword Difficulty (낮을수록 쉬움) →", autorange="reversed")
        fig.update_yaxes(title="검색량(MSV)")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("##### 🥇 즉시 선점 타겟 (Missing·KD≤25·MSV≥10,000)")
        st.dataframe(prime.sort_values("MSV", ascending=False)[
            ["키워드", "카테고리", "MSV", "KD"] + COMPETITORS],
            use_container_width=True, hide_index=True,
            column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")})

    with t2:
        st.markdown("##### 선점 우선순위 (MSV×난이도할인×상태가중)")
        topn = st.slider("상위 N개", 10, 50, 25, step=5, key="comp_topn")
        rk = df.sort_values("우선순위", ascending=False).head(topn)
        c1, c2 = st.columns([1.1, 1])
        with c1:
            dd = rk.sort_values("우선순위", ascending=True)
            fig = go.Figure(go.Bar(x=dd["우선순위"], y=dd["키워드"], orientation="h",
                                   marker=dict(color=[STATUS_COLOR[s] for s in dd["Status"]]),
                                   customdata=dd[["MSV", "KD", "Status", "타겟 등급"]],
                                   hovertemplate="<b>%{y}</b><br>우선순위 %{x:,}<br>"
                                   "MSV %{customdata[0]:,}·KD %{customdata[1]}<br>"
                                   "%{customdata[2]}·%{customdata[3]}<extra></extra>"))
            fig.update_layout(**base_layout(h=max(420, topn * 17), title="우선순위(색=Status)"))
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.dataframe(rk[["키워드", "타겟 등급", "MSV", "KD", SUBJECT, "Status", "우선순위"]],
                         use_container_width=True, hide_index=True, height=max(420, topn * 17),
                         column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")})

    with t3:
        st.markdown("##### 도메인별 커버리지 (이 62개 패션 키워드 기준)")
        rows = []
        for c in SITES:
            m = df[c] > 0
            rows.append((c, int(m.sum()), int(df[c].between(1, 10).sum()),
                         round(df.loc[m, c].mean(), 1) if m.any() else 0))
        cov = pd.DataFrame(rows, columns=["도메인", "노출 키워드", "1페이지(1~10위)", "평균 순위"])
        c1, c2 = st.columns([1.2, 1])
        with c1:
            fig = go.Figure()
            fig.add_trace(go.Bar(name="노출", x=cov["도메인"], y=cov["노출 키워드"],
                                 marker_color=PALETTE["slate"], text=cov["노출 키워드"],
                                 textposition="outside"))
            fig.add_trace(go.Bar(name="1페이지", x=cov["도메인"], y=cov["1페이지(1~10위)"],
                                 marker_color=PALETTE["green"], text=cov["1페이지(1~10위)"],
                                 textposition="outside"))
            fig.update_layout(**base_layout(h=340, showlegend=True), barmode="group")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.dataframe(cov, use_container_width=True, hide_index=True, height=230)
            st.caption("SSF샵·W컨셉이 패션 일반 키워드를 가장 넓게 선점. LF몰은 노출이 적어 "
                       "**Missing 선점 여지가 큼**.")

    with t4:
        st.markdown(
            "<div class='card'>🏆 <b>SSF샵</b>이 SEO 리더, <b>W컨셉</b>이 효율왕, "
            "<b>LF몰</b>은 키워드는 많은데 순위가 낮은 '잠자는 거인'. → 이미 5~20위에 걸린 키워드를 "
            "상위로 끌어올리는 것이 가장 빠른 ROI.</div>"
            "<div class='card'><b>Status</b> — "
            f"<span class='tag' style='background:{STATUS_COLOR['Strong']}'>Strong</span> 경쟁사보다 앞섬 / "
            f"<span class='tag' style='background:{STATUS_COLOR['Weak']}'>Weak</span> 열위 / "
            f"<span class='tag' style='background:{STATUS_COLOR['Missing']}'>Missing</span> 경쟁사만 보유(선점기회)</div>",
            unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# VIEW 2 — 키워드 리서치 (919개)
# ══════════════════════════════════════════════════════════════════
def load_kw():
    df = pd.read_csv("data/lfmall_keyword_research.csv", encoding="utf-8-sig")
    for col, dv in [("순위", 0), ("우선순위", ""), ("패션", "N"),
                    ("섹션", "기타"), ("Status", "미수집")]:
        if col not in df.columns:
            df[col] = dv
    return df


def render_keyword():
    df = load_kw()
    st.markdown("## 🔎 키워드 리서치 — 전체 카테고리 919개")
    st.markdown("<div class='cap'>대상 lfmall.co.kr · 경쟁사 W컨셉·한섬·SSF샵·SI빌리지 · "
                "Semrush 한국(kr) DB 검색량 실측</div>", unsafe_allow_html=True)
    st.write("")
    k = st.columns(5)
    k[0].metric("카테고리 키워드", f"{len(df):,}개", f"검색량 보유 {int((df['검색량'] > 0).sum())}개")
    k[1].metric("총 검색량", f"{int(df['검색량'].sum()):,}")
    k[2].metric("섹션", f"{df['섹션'].nunique()}개")
    k[3].metric("🔴 Missing", int((df["Status"] == "Missing").sum()), "경쟁사만 보유")
    k[4].metric("🎯 패션 우선순위", int((df["순위"] > 0).sum()), "1순위~ 부여")
    st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

    t1, t2, t3, t4, t5 = st.tabs(
        ["🚀 우선순위", "🗂️ 섹션 분석", "🎯 경쟁사 Status", "📋 전체 데이터(엑셀)", "📖 가이드"])

    with t1:
        st.markdown("##### 패션 카테고리 선점 우선순위 (1순위 = 최우선)")
        st.caption("패션·뷰티·골프 키워드 대상, √검색량·난이도·경쟁상태 종합 서수. "
                   "가전·리빙·식품 등 비패션 일반어는 제외.")
        rank_df = df[df["순위"] > 0]
        c0, c1 = st.columns(2)
        sec_f = c0.multiselect("섹션 필터", sorted(rank_df["섹션"].unique()), key="kw_secf")
        topn = c1.slider("상위 N개", 10, 60, 30, step=5, key="kw_topn")
        d = (rank_df if not sec_f else rank_df[rank_df["섹션"].isin(sec_f)]).sort_values("순위").head(topn)
        c1a, c2a = st.columns([1.1, 1])
        with c1a:
            dd = d.sort_values("순위", ascending=False)
            fig = go.Figure(go.Bar(x=dd["검색량"], y=dd["키워드"], orientation="h",
                                   marker=dict(color=[STATUS_COLOR.get(s, "#cbd5e1") for s in dd["Status"]]),
                                   text=dd["우선순위"], textposition="outside",
                                   customdata=dd[["우선순위", "섹션", "Status"]],
                                   hovertemplate="<b>%{y}</b><br>%{customdata[0]}·검색량 %{x:,}<br>"
                                   "%{customdata[1]}·%{customdata[2]}<extra></extra>"))
            fig.update_layout(**base_layout(h=max(420, topn * 16), title="패션 우선순위(막대=검색량)"))
            fig.update_xaxes(range=[0, dd["검색량"].max() * 1.18])
            st.plotly_chart(fig, use_container_width=True)
        with c2a:
            st.dataframe(d[["우선순위", "키워드", "섹션", "검색량", "Status"]],
                         use_container_width=True, hide_index=True, height=max(420, topn * 16),
                         column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d")})

    with t2:
        st.markdown("##### 섹션별 검색 수요 & 분포")
        sec = (df.groupby("섹션").agg(키워드수=("키워드", "count"), 총검색량=("검색량", "sum"))
                 .reset_index().sort_values("총검색량", ascending=False))
        c1, c2 = st.columns(2)
        with c1:
            dd = sec.sort_values("총검색량", ascending=True)
            fig = go.Figure(go.Bar(x=dd["총검색량"], y=dd["섹션"], orientation="h",
                                   marker_color="#4f8fff",
                                   text=[f"{int(v):,}" for v in dd["총검색량"]], textposition="outside"))
            fig.update_layout(**base_layout(h=560, title="섹션별 총 검색량"))
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.treemap(df[df["검색량"] > 0], path=[px.Constant("전체"), "섹션", "키워드"],
                             values="검색량", color="섹션")
            fig.update_traces(marker=dict(line=dict(color="white", width=1)),
                              hovertemplate="<b>%{label}</b><br>검색량 %{value:,}<extra></extra>")
            fig.update_layout(margin=dict(l=4, r=4, t=10, b=4), height=560)
            st.plotly_chart(fig, use_container_width=True)

    with t3:
        st.markdown("##### 경쟁사 대비 Status (순위 확보된 키워드)")
        known = df[df["Status"] != "미수집"]
        c1, c2 = st.columns([1, 1.4])
        with c1:
            sc = known["Status"].value_counts()
            fig = go.Figure(go.Bar(x=sc.index, y=sc.values,
                                   marker_color=[STATUS_COLOR.get(s, "#cbd5e1") for s in sc.index],
                                   text=sc.values, textposition="outside"))
            fig.update_layout(**base_layout(h=320, title="Status 분포(순위 확보분)"))
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.dataframe(known.sort_values("검색량", ascending=False)[
                ["키워드", "섹션", "검색량", "난이도"] + SITES + ["Status"]],
                use_container_width=True, hide_index=True, height=320,
                column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d")})
        st.info("비패션 카테고리의 경쟁사 순위는 미수집 상태입니다. 전체 919개 5사 SERP는 추가 수집 필요.")

    with t4:
        st.markdown("##### 전체 키워드 리서치 데이터 (필터 + 다운로드)")
        f0, f1, f2, f3 = st.columns([1.3, 1.3, 1, 1])
        fsec = f0.multiselect("섹션", sorted(df["섹션"].unique()), key="kw_fsec")
        fstat = f1.multiselect("Status", list(STATUS_COLOR.keys()), key="kw_fstat")
        fmsv = f2.slider("최소 검색량", 0, int(df["검색량"].max()), 0, step=1000, key="kw_fmsv")
        q = f3.text_input("키워드 검색", "", key="kw_q")
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
        d1.download_button("⬇️ 엑셀(.xlsx)", to_excel(fdf, "키워드리서치"),
                           "lfmall_keyword_research.xlsx", XLSX_MIME, use_container_width=True)
        d2.download_button("⬇️ CSV", fdf.to_csv(index=False).encode("utf-8-sig"),
                           "lfmall_keyword_research.csv", "text/csv", use_container_width=True)
        st.dataframe(fdf[["키워드", "섹션", "검색량", "난이도"] + SITES + ["Status", "우선순위"]],
                     use_container_width=True, hide_index=True, height=520,
                     column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d")})

    with t5:
        st.markdown("##### 📅 기간 기준")
        st.markdown("<div class='card'>검색량 = Semrush 한국(kr) DB, <b>2026-06-22 스냅샷</b>, "
                    "<b>최근 12개월 평균 월간 검색량</b>(단월 아님). 계절 키워드(패딩·수영복)는 "
                    "성수기 단월이 연평균보다 큼.</div>", unsafe_allow_html=True)
        st.markdown("##### 🧮 연산 기준")
        st.markdown(
            "<div class='card'><code>우선순위 점수 = √검색량 × 난이도할인 × 상태가중</code><br>"
            "• √검색량 — 초고볼륨(뮬·시계·크림) 지배력 완화<br>"
            "• 난이도할인 = (100−KD)/100 (미상은 40 가정)<br>"
            "• 상태가중: Missing 1.0·공백 0.9·미수집 0.65·Weak 0.6·Strong 0.4<br>"
            "• 점수 내림차순 1순위~. 대상 = 패션+뷰티+골프 섹션 & 검색량&gt;0</div>"
            "<div class='card'><b>Status 판정</b> — LF몰=0&amp;경쟁사도 0→공백 / LF몰=0&amp;경쟁사 있음→Missing / "
            "LF몰&lt;경쟁사 최고순위→Strong / 그 외→Weak / 순위데이터 없음→미수집</div>",
            unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# VIEW 3 — 네이버 쇼핑
# ══════════════════════════════════════════════════════════════════
NAVER_CSV = "data/naver_keyword_metrics.csv"
TREND_COLOR = {"급상승": "#e11d48", "상승": "#f97316", "유지": "#94a3b8",
               "하락": "#3b82f6", "급하락": "#1d4ed8", "": "#cbd5e1"}


def render_naver():
    st.markdown("## 🛒 네이버 쇼핑·검색")
    st.markdown("<div class='cap'>네이버 검색광고 API(월간 검색량) + 데이터랩(12개월 추이) 기반. "
                "Semrush(구글)로는 안 잡히는 한국 실수요를 보정.</div>", unsafe_allow_html=True)
    st.write("")

    if not os.path.exists(NAVER_CSV):
        st.warning("아직 네이버 데이터가 없습니다. 아래 절차로 키를 넣고 수집하면 이 탭이 활성화됩니다.")
        st.markdown(
            "<div class='card'><b>① 키 발급 (개인계정 가능)</b><br>"
            "• 검색광고 API — searchad.naver.com → 도구 → API 사용 관리 "
            "(<code>API_KEY</code>·<code>SECRET</code>·<code>CUSTOMER_ID</code>)<br>"
            "• 데이터랩/쇼핑인사이트 — developers.naver.com → 애플리케이션 등록 "
            "(<code>CLIENT_ID</code>·<code>CLIENT_SECRET</code>)</div>"
            "<div class='card'><b>② 프로젝트 루트 <code>.env</code> 에 입력</b><br>"
            "<code>NAVER_AD_API_KEY=...</code><br><code>NAVER_AD_SECRET=...</code><br>"
            "<code>NAVER_AD_CUSTOMER_ID=...</code><br><code>NAVER_CLIENT_ID=...</code><br>"
            "<code>NAVER_CLIENT_SECRET=...</code></div>"
            "<div class='card'><b>③ 수집 실행</b><br>"
            "<code>python naver_keyword_data.py</code> → data/naver_keyword_metrics.csv 생성<br>"
            "<code>python build_keyword_data.py</code> → 키워드 리서치 검색량을 네이버로 갱신</div>",
            unsafe_allow_html=True)
        return

    nv = pd.read_csv(NAVER_CSV, encoding="utf-8-sig")
    nv = nv.sort_values("네이버검색량", ascending=False)
    has_trend = nv["추이"].astype(str).str.len().gt(0).any() if "추이" in nv else False

    k = st.columns(4)
    k[0].metric("수집 키워드", f"{len(nv):,}개")
    k[1].metric("네이버 검색량 합", f"{int(nv['네이버검색량'].sum()):,}")
    k[2].metric("모바일 비중",
                f"{round(nv['모바일'].sum() / max(1, nv[['PC', '모바일']].sum().sum()) * 100)}%")
    rising = int(nv["추이"].isin(["상승", "급상승"]).sum()) if has_trend else 0
    k[3].metric("📈 상승 키워드", f"{rising}개")
    st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

    has_monthly = "월별추이" in nv.columns and nv["월별추이"].astype(str).str.len().gt(0).any()

    def parse_series(s):
        try:
            return [float(x) for x in str(s).split("|") if x != ""]
        except ValueError:
            return []

    t1, t2, t3 = st.tabs(["📈 검색량 TOP", "📅 기간 추이", "📋 전체(엑셀)"])

    with t1:
        topn = st.slider("상위 N개", 10, 60, 30, step=5, key="nv_topn")
        d = nv.head(topn).sort_values("네이버검색량")
        fig = go.Figure(go.Bar(x=d["네이버검색량"], y=d["키워드"], orientation="h",
                               marker_color=PALETTE["green"],
                               customdata=d[["PC", "모바일"]],
                               hovertemplate="<b>%{y}</b><br>네이버 %{x:,}"
                               "<br>PC %{customdata[0]:,}·모바일 %{customdata[1]:,}<extra></extra>"))
        fig.update_layout(**base_layout(h=max(420, topn * 16), title="네이버 월간 검색량 TOP"))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("PC+모바일 합산 월간 검색수. Semrush(구글)와 비교하면 한국 실수요 차이를 알 수 있습니다.")

    with t2:
        if not has_monthly:
            st.info("데이터랩 키(CLIENT_ID/SECRET) 추가 후 재수집하면 **기간 슬라이더 + 추이 라인차트**가 "
                    "활성화됩니다. (검색광고 키만으로는 절대 검색량만 수집)")
        else:
            months = st.slider("📅 기간 (최근 N개월)", 3, 12, 12, step=1, key="nv_months")
            nv2 = nv.copy()
            nv2["_s"] = nv2["월별추이"].map(parse_series)

            def grow(s):
                w = s[-months:]
                if len(w) < 2:
                    return 0
                n = max(1, len(w) // 3)
                e = sum(w[:n]) / n
                l = sum(w[-n:]) / n
                return round((l - e) / e * 100) if e else 0

            nv2["기간성장%"] = nv2["_s"].map(grow)
            up = nv2[nv2["기간성장%"] >= 5].sort_values("기간성장%", ascending=False)
            st.markdown(f"##### 최근 {months}개월 상승 키워드 ({len(up)}개)")
            sel = st.multiselect("추이 비교 키워드", nv2["키워드"].tolist(),
                                 default=up["키워드"].head(6).tolist(), key="nv_sel")
            if sel:
                fig = go.Figure()
                for kw in sel:
                    s = nv2.loc[nv2["키워드"] == kw, "_s"].iloc[0][-months:]
                    fig.add_trace(go.Scatter(y=s, x=list(range(-len(s) + 1, 1)),
                                             mode="lines+markers", name=kw))
                fig.update_layout(**base_layout(h=360, showlegend=True,
                                                title=f"최근 {months}개월 상대 검색추이(데이터랩 지수)"))
                fig.update_xaxes(title="개월 전 (0 = 최근月)")
                st.plotly_chart(fig, use_container_width=True)
            st.dataframe(up[["키워드", "네이버검색량", "기간성장%", "추이지수"]].head(60),
                         use_container_width=True, hide_index=True, height=320,
                         column_config={"네이버검색량": st.column_config.NumberColumn("네이버검색량", format="%d")})

    with t3:
        st.markdown(f"##### 전체 네이버 지표 ({len(nv):,}개)")
        st.download_button("⬇️ 엑셀(.xlsx)", to_excel(nv, "네이버지표"),
                           "naver_keyword_metrics.xlsx", XLSX_MIME)
        st.dataframe(nv, use_container_width=True, hide_index=True, height=480,
                     column_config={"네이버검색량": st.column_config.NumberColumn("네이버검색량", format="%d")})


# ══════════════════════════════════════════════════════════════════
# 라우팅
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🧭 LF몰 SEO 대시보드")
    view = st.radio("보기 선택", ["🥊 경쟁사 분석", "🔎 키워드 리서치", "🛒 네이버 쇼핑"])
    st.markdown("---")
    st.caption("Semrush 한국(kr) DB + 네이버 검색광고/데이터랩. "
               "데이터는 추정치로 실제와 차이가 있을 수 있습니다.")

if view.startswith("🥊"):
    render_competitor()
elif view.startswith("🔎"):
    render_keyword()
else:
    render_naver()
