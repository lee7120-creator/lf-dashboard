"""LF몰 프로그래매틱 SEO 키워드 선점 분석 대시보드

대상       : lfmall.co.kr (LF몰)
경쟁사     : wconcept.co.kr(W컨셉) · thehandsome.com(한섬) ·
             ssfshop.com(SSF샵) · sivillage.com(SI빌리지)
관점       : Programmatic SEO — 경쟁사 대비 미보유(Missing)·열위(Weak) 키워드를
             고MSV·저Difficulty 우선순위로 선점할 타겟 도출

지표 정의
  · Search Volume(MSV) : Google 기준 월별 검색량 (높을수록 매력)
  · Keyword Difficulty : 키워드 선점 난이도 0~100 (낮을수록 유리)
  · 도메인 컬럼 숫자     : 검색 시 순위(1~100). 0 = 순위 없음
  · Status
      - Strong  : LF몰 순위가 경쟁 4사보다 높음(앞섬)
      - Weak    : LF몰이 순위는 있으나 경쟁사보다 낮음
      - Missing : LF몰 순위 없음(0) → 신규 페이지 선점 기회
      - ※ Strong이라도 순위가 10위 밖이면 SEO 최적화 필요

데이터 출처
  · Semrush 한국(kr) DB 실데이터 (organic_research / domain_domains 키워드 갭)
  · 스냅샷 2026-06-22, 검색량순 상위에서 패션 카테고리 키워드만 선별
  · 브랜드명·가전·뷰티·주얼리·골프 등 비(非)패션 키워드는 제외
  · 갱신 시: Semrush MCP로 domain_domains 재실행
    domains='*|or|lfmall.co.kr|+|or|wconcept.co.kr|+|or|thehandsome.com|+|or|ssfshop.com|+|or|sivillage.com'

실행: streamlit run programmatic_seo_dashboard.py
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="LF몰 프로그래매틱 SEO 선점 분석", page_icon="🎯",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.card{border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin:8px 0;line-height:1.65;background:#ffffff}
.sdiv{border-top:1px solid #e2e8f0;margin:22px 0}
.tag{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:6px;margin-right:6px;color:white}
.cap{color:#64748b;font-size:13px}
h1,h2,h3{color:#1e293b}
</style>
""", unsafe_allow_html=True)

SUBJECT = "LF몰"
COMPETITORS = ["W컨셉", "한섬", "SSF샵", "SI빌리지"]
ALL_SITES = [SUBJECT] + COMPETITORS
STATUS_COLOR = {"Strong": "#48bb78", "Weak": "#ed8936", "Missing": "#f56565"}
PALETTE = {"blue": "#4f8fff", "red": "#f56565", "amber": "#ed8936",
           "green": "#48bb78", "purple": "#9f7aea", "slate": "#64748b"}
SITE_COLOR = {"LF몰": PALETTE["blue"], "W컨셉": PALETTE["amber"], "한섬": PALETTE["purple"],
              "SSF샵": PALETTE["red"], "SI빌리지": PALETTE["slate"]}

def base_layout(h=320, title="", showlegend=False):
    return dict(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#475569", size=12), margin=dict(l=10, r=10, t=44, b=10),
        height=h, showlegend=showlegend,
        title=dict(text=title, font=dict(color="#94a3b8", size=13)),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1,
                    font=dict(size=11)),
        xaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11)),
        yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11)),
    )

# ══════════════════════════════════════════════════════
# DATA — Semrush kr DB 실데이터 (패션 카테고리 키워드만 선별, 2026-06-22)
# (키워드, MSV, KD, LF몰, W컨셉, 한섬, SSF샵, SI빌리지, 카테고리)  순위 0=없음
# ══════════════════════════════════════════════════════
DATA = [
    # 아우터
    ("패딩", 33100, 21, 0, 11, 0, 0, 0, "아우터"),
    ("코트", 27100, 23, 0, 0, 0, 18, 0, "아우터"),
    ("재킷", 22200, 32, 0, 0, 49, 8, 0, "아우터"),
    ("모피", 22200, 29, 5, 0, 0, 0, 0, "아우터"),
    ("가디건", 12100, 20, 0, 0, 0, 14, 0, "아우터"),
    ("무스탕", 12100, 20, 13, 2, 0, 0, 0, "아우터"),
    ("트렌치 코트", 12100, 27, 0, 0, 0, 7, 0, "아우터"),
    ("경량 패딩", 12100, 16, 5, 0, 0, 20, 0, "아우터"),
    ("바람막이", 14800, 19, 0, 0, 0, 20, 0, "아우터"),
    ("블레이저", 12100, 28, 0, 0, 0, 29, 0, "아우터"),
    ("블루종", 6600, 28, 15, 5, 65, 9, 0, "아우터"),
    # 상의
    ("블라우스", 22200, 24, 0, 2, 0, 16, 0, "상의"),
    ("셔츠", 22200, 29, 0, 23, 0, 24, 0, "상의"),
    ("니트", 18100, 19, 43, 2, 0, 8, 0, "상의"),
    ("시스루", 18100, 27, 0, 0, 0, 9, 0, "상의"),
    ("티셔츠", 12100, 16, 0, 5, 0, 32, 0, "상의"),
    ("후드 집업", 12100, 14, 0, 5, 0, 23, 0, "상의"),
    ("크롭 티", 9900, 17, 32, 2, 0, 3, 0, "상의"),
    ("탱크 탑", 8100, 23, 0, 0, 0, 21, 0, "상의"),
    ("오프 숄더", 8100, 17, 0, 3, 0, 0, 0, "상의"),
    # 하의
    ("청바지", 74000, 23, 0, 0, 0, 27, 0, "하의"),
    ("바지", 40500, 19, 0, 0, 0, 14, 0, "하의"),
    ("돌핀 팬츠", 22200, 27, 0, 0, 0, 5, 0, "하의"),
    ("데님", 22200, 22, 0, 0, 0, 46, 0, "하의"),
    ("미니 스커트", 14800, 31, 18, 0, 0, 19, 0, "하의"),
    ("버뮤다 팬츠", 14800, 18, 20, 0, 0, 24, 0, "하의"),
    ("슬랙스", 14800, 16, 0, 14, 0, 0, 0, "하의"),
    ("치마", 14800, 24, 0, 0, 0, 43, 0, "하의"),
    ("반바지", 8100, 16, 0, 0, 0, 24, 0, "하의"),
    # 드레스/수영
    ("드레스", 22200, 31, 35, 0, 0, 0, 0, "드레스/수영"),
    ("비키니", 60500, 35, 0, 5, 0, 0, 0, "드레스/수영"),
    ("모노 키니", 12100, 12, 23, 2, 0, 0, 0, "드레스/수영"),
    # 신발
    ("부츠", 12100, 26, 0, 3, 0, 48, 0, "신발"),
    ("구두", 9900, 39, 0, 0, 0, 30, 0, "신발"),
    ("로퍼", 12100, 20, 0, 19, 0, 0, 0, "신발"),
    ("샌들", 6600, 21, 30, 11, 0, 0, 0, "신발"),
    ("더비 슈즈", 8100, 17, 10, 6, 0, 0, 0, "신발"),
    # 가방
    ("가방", 33100, 35, 0, 5, 0, 0, 0, "가방"),
    ("크로스 백", 12100, 16, 0, 2, 0, 0, 0, "가방"),
    ("백팩", 9900, 20, 0, 0, 0, 24, 0, "가방"),
    ("토트 백", 9900, 19, 0, 2, 0, 46, 0, "가방"),
    ("메신저 백", 6600, 23, 0, 20, 0, 29, 0, "가방"),
    ("에코 백", 8100, 10, 26, 3, 0, 0, 0, "가방"),
    # 액세서리
    ("모자", 18100, 18, 0, 0, 0, 12, 0, "액세서리"),
    ("지갑", 14800, 20, 0, 0, 0, 30, 0, "액세서리"),
    ("키링", 14800, 18, 0, 0, 0, 49, 0, "액세서리"),
    ("넥타이", 9900, 32, 0, 0, 0, 21, 0, "액세서리"),
    ("벨트", 6600, 26, 0, 0, 67, 0, 0, "액세서리"),
    ("뿔테 안경", 9900, 20, 17, 0, 0, 19, 0, "액세서리"),
    ("반다나", 9900, 23, 12, 6, 0, 24, 0, "액세서리"),
    ("목도리", 8100, 20, 8, 3, 0, 0, 0, "액세서리"),
    ("비니", 8100, 27, 0, 0, 0, 23, 0, "액세서리"),
    ("카드 지갑", 8100, 19, 23, 11, 0, 42, 0, "액세서리"),
    ("니삭스", 6600, 31, 0, 2, 0, 12, 0, "액세서리"),
    # 언더웨어
    ("속옷", 18100, 23, 0, 0, 0, 19, 0, "언더웨어"),
    ("여자 속옷", 8100, 14, 9, 26, 0, 0, 0, "언더웨어"),
    ("여자 팬티", 9900, 20, 2, 0, 0, 0, 0, "언더웨어"),
    ("티 팬티", 12100, 29, 23, 0, 0, 19, 0, "언더웨어"),
    ("잠옷", 8100, 19, 19, 0, 0, 0, 0, "언더웨어"),
    # 정장/포멀
    ("정장", 18100, 20, 0, 0, 26, 15, 0, "정장/포멀"),
    ("세미 정장", 6600, 21, 41, 0, 0, 0, 0, "정장/포멀"),
    # 컬러
    ("버건디", 9900, 29, 0, 7, 0, 0, 0, "컬러"),
]

df = pd.DataFrame(DATA, columns=["키워드", "MSV", "KD"] + ALL_SITES + ["카테고리"])

# ── Status / 등급 / 우선순위 ──
def calc_status(row):
    gw = row[SUBJECT]
    if gw == 0:
        return "Missing"
    ranked = [row[c] for c in COMPETITORS if row[c] > 0]
    if not ranked or gw < min(ranked):
        return "Strong"
    return "Weak"

def calc_tier(row):
    s, kd, msv, gw = row["Status"], row["KD"], row["MSV"], row[SUBJECT]
    if s == "Missing":
        if kd <= 25 and msv >= 10000:
            return "🥇 즉시 선점"
        if kd <= 25 or msv >= 10000:
            return "🥈 선점 후보"
        return "🥉 롱테일 선점"
    if s == "Weak":
        return "🔧 최적화(탈환)"
    if s == "Strong" and gw > 10:
        return "🔧 최적화(순위↑)"
    return "✅ 방어"

def calc_priority(row):
    diff_factor = (100 - row["KD"]) / 100
    s, gw = row["Status"], row[SUBJECT]
    sf = (1.0 if s == "Missing" else 0.65 if s == "Weak"
          else 0.45 if (s == "Strong" and gw > 10) else 0.15)
    return int(round(row["MSV"] * diff_factor * sf))

def best_competitor(row):
    ranked = [(row[c], c) for c in COMPETITORS if row[c] > 0]
    return min(ranked) if ranked else (0, "-")

df["Status"] = df.apply(calc_status, axis=1)
df["타겟 등급"] = df.apply(calc_tier, axis=1)
df["우선순위"] = df.apply(calc_priority, axis=1)
df["LF_1P"] = df[SUBJECT].between(1, 10)

# ══════════════════════════════════════════════════════
# HEADER + KPI
# ══════════════════════════════════════════════════════
st.markdown("## 🎯 LF몰 프로그래매틱 SEO 키워드 선점 분석")
st.markdown(
    "<div class='cap'>대상 <b>lfmall.co.kr</b> · 경쟁사 "
    "<b>W컨셉 · 한섬 · SSF샵 · SI빌리지</b> · 패션 카테고리 키워드 "
    f"<b>{len(df)}</b>개 · Semrush 한국(kr) DB 실데이터</div>",
    unsafe_allow_html=True,
)
st.write("")

n_strong = int((df["Status"] == "Strong").sum())
n_weak = int((df["Status"] == "Weak").sum())
n_missing = int((df["Status"] == "Missing").sum())
prime = df[(df["Status"] == "Missing") & (df["KD"] <= 25) & (df["MSV"] >= 10000)]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("분석 키워드", f"{len(df)}개", f"총 MSV {df['MSV'].sum():,}")
k2.metric("🟢 Strong", f"{n_strong}개", "앞서는 키워드")
k3.metric("🟠 Weak", f"{n_weak}개", "탈환 대상", delta_color="inverse")
k4.metric("🔴 Missing", f"{n_missing}개", "선점 기회", delta_color="off")
k5.metric("🥇 즉시 선점 타겟", f"{len(prime)}개", "고MSV·저난이도·미보유")

st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🎯 프로그래매틱 SEO")
    st.caption("LF몰 키워드 선점 분석")
    st.markdown("---")
    st.markdown("**Status 정의** (LF몰 기준)")
    st.markdown(
        f"<span class='tag' style='background:{STATUS_COLOR['Strong']}'>Strong</span> 경쟁사보다 앞섬<br>"
        f"<span class='tag' style='background:{STATUS_COLOR['Weak']}'>Weak</span> 순위 있으나 열위<br>"
        f"<span class='tag' style='background:{STATUS_COLOR['Missing']}'>Missing</span> 순위 없음 → 선점 기회",
        unsafe_allow_html=True,
    )
    st.markdown("---")
    st.markdown("**비교 대상**")
    for s in ALL_SITES:
        st.markdown(f"<span style='color:{SITE_COLOR[s]};font-size:18px'>●</span> <b>{s}</b>",
                    unsafe_allow_html=True)
    st.markdown("---")
    st.caption("Semrush kr DB 실데이터 · 패션 카테고리 키워드만 선별. "
               "브랜드명·가전·뷰티 등 비패션 키워드는 제외.")

# ══════════════════════════════════════════════════════
# 탭
# ══════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["🎯 기회 매트릭스", "🚀 선점 타겟", "🗂️ 카테고리 클러스터",
     "📊 Status 분석", "🔍 도메인 커버리지", "📖 가이드"]
)

# ──────────────────────────────────────────────────────
# TAB 1 — 기회 매트릭스
# ──────────────────────────────────────────────────────
with tab1:
    st.markdown("##### 키워드 기회 매트릭스 — 좌상단(고검색량·저난이도)이 황금존")
    fig = go.Figure()
    fig.add_shape(type="rect", x0=0, x1=25, y0=10000, y1=df["MSV"].max() * 1.05,
                  fillcolor="rgba(72,187,120,0.08)", line=dict(width=0), layer="below")
    fig.add_annotation(x=12.5, y=df["MSV"].max(), text="🥇 황금존 (선점 우선)",
                       showarrow=False, font=dict(color="#48bb78", size=12))
    for s in ["Missing", "Weak", "Strong"]:
        d = df[df["Status"] == s]
        fig.add_trace(go.Scatter(
            x=d["KD"], y=d["MSV"], mode="markers", name=s,
            marker=dict(size=(d["MSV"] / 1800 + 8), color=STATUS_COLOR[s],
                        opacity=0.78, line=dict(color="white", width=1)),
            customdata=d[["키워드", "타겟 등급", SUBJECT]],
            hovertemplate=("<b>%{customdata[0]}</b><br>MSV %{y:,} · KD %{x}"
                           "<br>LF몰 순위 %{customdata[2]}<br>%{customdata[1]}<extra></extra>"),
        ))
    fig.update_layout(**base_layout(h=440, showlegend=True))
    fig.update_xaxes(title="Keyword Difficulty (낮을수록 쉬움) →", autorange="reversed")
    fig.update_yaxes(title="Monthly Search Volume")
    st.plotly_chart(fig, use_container_width=True)
    st.caption("💡 x축은 난이도 역방향(왼쪽=쉬움). **왼쪽 위 = 검색량 크고 쉬운 키워드**. "
               "빨간 점(Missing)이 황금존에 있으면 = 페이지만 만들면 바로 선점 가능.")

    st.markdown("##### 🥇 즉시 선점 타겟 (Missing · KD≤25 · MSV≥10,000)")
    st.dataframe(
        prime.sort_values("MSV", ascending=False)[
            ["키워드", "카테고리", "MSV", "KD"] + COMPETITORS],
        use_container_width=True, hide_index=True,
        column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")},
    )

# ──────────────────────────────────────────────────────
# TAB 2 — 선점 타겟
# ──────────────────────────────────────────────────────
with tab2:
    st.markdown("##### 프로그래매틱 SEO 선점 우선순위")
    st.caption("우선순위 = MSV × (난이도 할인) × 상태가중(Missing 1.0 / Weak 0.65 / Strong>10위 0.45).")
    topn = st.slider("상위 N개", 10, 50, 25, step=5)
    ranked = df.sort_values("우선순위", ascending=False).head(topn)

    c1, c2 = st.columns([1.15, 1])
    with c1:
        dd = ranked.sort_values("우선순위", ascending=True)
        fig = go.Figure(go.Bar(
            x=dd["우선순위"], y=dd["키워드"], orientation="h",
            marker=dict(color=[STATUS_COLOR[s] for s in dd["Status"]]),
            customdata=dd[["MSV", "KD", "Status", "타겟 등급"]],
            hovertemplate=("<b>%{y}</b><br>우선순위 %{x:,}<br>MSV %{customdata[0]:,} · "
                           "KD %{customdata[1]}<br>%{customdata[2]} · %{customdata[3]}<extra></extra>"),
        ))
        fig.update_layout(**base_layout(h=max(420, topn * 17), title="우선순위 점수 (색=Status)"))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.dataframe(
            ranked[["키워드", "타겟 등급", "MSV", "KD", SUBJECT, "Status", "우선순위"]],
            use_container_width=True, hide_index=True, height=max(420, topn * 17),
            column_config={
                "MSV": st.column_config.NumberColumn("MSV", format="%d"),
                "우선순위": st.column_config.ProgressColumn(
                    "우선순위", format="%d", min_value=0,
                    max_value=int(df["우선순위"].max())),
            },
        )
    tier_order = ["🥇 즉시 선점", "🥈 선점 후보", "🥉 롱테일 선점",
                  "🔧 최적화(탈환)", "🔧 최적화(순위↑)", "✅ 방어"]
    summ = (df.groupby("타겟 등급")
              .agg(키워드수=("키워드", "count"), 총MSV=("MSV", "sum"))
              .reindex([t for t in tier_order if t in df["타겟 등급"].values])
              .reset_index())
    st.markdown("##### 타겟 등급별 요약")
    st.dataframe(summ, use_container_width=True, hide_index=True,
                 column_config={"총MSV": st.column_config.NumberColumn("총MSV", format="%d")})

# ──────────────────────────────────────────────────────
# TAB 3 — 카테고리 클러스터
# ──────────────────────────────────────────────────────
with tab3:
    st.markdown("##### 카테고리 클러스터 — 허브·스포크 페이지 아키텍처")
    st.caption("프로그래매틱 SEO는 카테고리(허브) 아래 키워드별 페이지(스포크)를 묶는다. "
               "블록 크기=MSV, 색=Status. 빨간 블록이 많은 카테고리 = 선점 여지 큰 클러스터.")
    fig = px.treemap(
        df, path=[px.Constant("전체"), "카테고리", "키워드"], values="MSV",
        color="Status", color_discrete_map={**STATUS_COLOR, "(?)": "#cbd5e1"},
        custom_data=["KD", SUBJECT],
    )
    fig.update_traces(
        marker=dict(line=dict(color="white", width=1)),
        hovertemplate="<b>%{label}</b><br>MSV %{value:,}<br>KD %{customdata[0]}<extra></extra>",
    )
    fig.update_layout(margin=dict(l=8, r=8, t=10, b=8), height=460)
    st.plotly_chart(fig, use_container_width=True)

    cat = (df.groupby("카테고리")
             .agg(키워드수=("키워드", "count"), 총MSV=("MSV", "sum"), 평균KD=("KD", "mean"),
                  Missing=("Status", lambda s: (s == "Missing").sum()),
                  Weak=("Status", lambda s: (s == "Weak").sum()),
                  Strong=("Status", lambda s: (s == "Strong").sum()))
             .reset_index().sort_values("총MSV", ascending=False))
    cat["평균KD"] = cat["평균KD"].round(0).astype(int)
    cat["선점여지%"] = ((cat["Missing"] + cat["Weak"]) / cat["키워드수"] * 100).round(0).astype(int)
    st.markdown("##### 카테고리별 선점 여지")
    st.dataframe(
        cat, use_container_width=True, hide_index=True,
        column_config={
            "총MSV": st.column_config.ProgressColumn("총MSV", format="%d", min_value=0,
                                                     max_value=int(cat["총MSV"].max())),
            "선점여지%": st.column_config.ProgressColumn("선점여지%", format="%d%%",
                                                       min_value=0, max_value=100),
        },
    )

# ──────────────────────────────────────────────────────
# TAB 4 — Status 분석
# ──────────────────────────────────────────────────────
with tab4:
    c0, c1, c2, c3 = st.columns([1.2, 1.2, 1, 1])
    sel_status = c0.multiselect("Status", ["Strong", "Weak", "Missing"],
                                default=["Missing", "Weak", "Strong"])
    sel_cat = c1.multiselect("카테고리", sorted(df["카테고리"].unique()))
    min_msv = c2.slider("최소 MSV", 0, int(df["MSV"].max()), 0, step=1000)
    max_kd = c3.slider("최대 KD", 0, 50, 50, step=1)
    q = st.text_input("키워드 검색", "")

    fdf = df[df["Status"].isin(sel_status) & (df["MSV"] >= min_msv) & (df["KD"] <= max_kd)]
    if sel_cat:
        fdf = fdf[fdf["카테고리"].isin(sel_cat)]
    if q:
        fdf = fdf[fdf["키워드"].str.contains(q, case=False, na=False)]

    cc1, cc2 = st.columns(2)
    sc = df["Status"].value_counts().reindex(["Strong", "Weak", "Missing"]).fillna(0)
    sm = df.groupby("Status")["MSV"].sum().reindex(["Strong", "Weak", "Missing"]).fillna(0)
    with cc1:
        fig = go.Figure(go.Bar(x=sc.index, y=sc.values,
                               marker_color=[STATUS_COLOR[s] for s in sc.index],
                               text=sc.values.astype(int), textposition="outside"))
        fig.update_layout(**base_layout(h=260, title="Status별 키워드 수"))
        st.plotly_chart(fig, use_container_width=True)
    with cc2:
        fig = go.Figure(go.Bar(x=sm.index, y=sm.values,
                               marker_color=[STATUS_COLOR[s] for s in sm.index],
                               text=[f"{int(v):,}" for v in sm.values], textposition="outside"))
        fig.update_layout(**base_layout(h=260, title="Status별 총 검색량(MSV)"))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown(f"##### 키워드 상세 ({len(fdf)}개)")
    st.dataframe(
        fdf[["키워드", "카테고리", "MSV", "KD"] + ALL_SITES + ["Status", "타겟 등급"]]
        .sort_values("MSV", ascending=False),
        use_container_width=True, hide_index=True, height=460,
        column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")},
    )
    st.download_button("⬇️ 필터된 키워드 CSV 내려받기",
                       fdf.to_csv(index=False).encode("utf-8-sig"),
                       "lfmall_keywords.csv", "text/csv")

# ──────────────────────────────────────────────────────
# TAB 5 — 도메인 커버리지
# ──────────────────────────────────────────────────────
with tab5:
    st.markdown("##### 도메인별 커버리지 — 이 패션 키워드셋에서 누가 얼마나 노출되나")
    rows = []
    for col in ALL_SITES:
        ranked_mask = df[col] > 0
        p1 = int(df[col].between(1, 10).sum())
        avg = df.loc[ranked_mask, col].mean()
        rows.append((col, int(ranked_mask.sum()), p1, round(avg, 1) if ranked_mask.any() else 0))
    cov = pd.DataFrame(rows, columns=["도메인", "노출 키워드", "1페이지(1~10위)", "평균 순위"])

    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig = go.Figure()
        fig.add_trace(go.Bar(name="노출 키워드", x=cov["도메인"], y=cov["노출 키워드"],
                             marker_color=PALETTE["slate"],
                             text=cov["노출 키워드"], textposition="outside"))
        fig.add_trace(go.Bar(name="1페이지(1~10위)", x=cov["도메인"], y=cov["1페이지(1~10위)"],
                             marker_color=PALETTE["green"],
                             text=cov["1페이지(1~10위)"], textposition="outside"))
        fig.update_layout(**base_layout(h=340, showlegend=True), barmode="group")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.dataframe(cov, use_container_width=True, hide_index=True, height=230)
        st.caption("**SSF샵·W컨셉이 패션 일반 키워드를 가장 넓게 선점** → 핵심 경쟁자. "
                   "한섬·SI빌리지는 이 키워드셋 노출이 적음. LF몰은 노출 자체가 적어 "
                   "**Missing 선점 여지가 매우 큼**.")

    st.markdown("##### LF몰 vs 최강 경쟁사 — 키워드별 맞대결")
    duel = df.copy()
    bc = duel.apply(best_competitor, axis=1, result_type="expand")
    duel["경쟁사 최고순위"] = bc[0]
    duel["경쟁사"] = bc[1]
    def winner(r):
        lf, bcr = r[SUBJECT], r["경쟁사 최고순위"]
        if lf > 0 and (bcr == 0 or lf < bcr):
            return "LF몰 우위"
        if bcr > 0:
            return "경쟁사 우위"
        return "-"
    duel["우위"] = duel.apply(winner, axis=1)
    win = int((duel["우위"] == "LF몰 우위").sum())
    lose = int((duel["우위"] == "경쟁사 우위").sum())
    st.caption(f"전체 {len(duel)}개 중 LF몰 우위 **{win}개** / 경쟁사 우위 **{lose}개** "
               "(Missing 다수 → 경쟁사 우위가 많음 = 선점 기회).")
    st.dataframe(
        duel[["키워드", "MSV", "KD", SUBJECT, "경쟁사", "경쟁사 최고순위", "Status", "우위"]]
        .sort_values("MSV", ascending=False),
        use_container_width=True, hide_index=True, height=340,
        column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d"),
                       SUBJECT: st.column_config.NumberColumn("LF몰 순위")},
    )

# ──────────────────────────────────────────────────────
# TAB 6 — 가이드
# ──────────────────────────────────────────────────────
with tab6:
    st.markdown("##### 분석 대상")
    st.markdown(
        "<div class='card'>LF몰(lfmall.co.kr)과 경쟁 4사(W컨셉·한섬·SSF샵·SI빌리지)가 "
        "보유한 키워드 중 <b>패션 관련 키워드</b>만 선별해 정리했습니다. "
        "(Semrush kr DB 실데이터, 브랜드명·가전·뷰티·주얼리 등 비패션 키워드 제외)</div>",
        unsafe_allow_html=True,
    )
    st.markdown("##### 지표 설명")
    st.markdown(
        "<div class='card'>"
        "• <b>Search Volume(MSV)</b> = Google 기준 월별 검색량, <b>Difficulty</b> = 선점 난이도<br>"
        "• MSV가 높되 Difficulty가 낮은 키워드를 공략하는 것이 유리<br>"
        "• 각 도메인 숫자 = 검색 시 순위(1~100). <b>0 = 순위 없음</b><br>"
        "• 1~10위를 선점해 SERP 1페이지 노출이 이상적 목표<br>"
        f"• <span class='tag' style='background:{STATUS_COLOR['Strong']}'>Strong</span> "
        "LF몰 순위가 경쟁 4사보다 높음 / "
        f"<span class='tag' style='background:{STATUS_COLOR['Weak']}'>Weak</span> 경쟁사보다 낮음 / "
        f"<span class='tag' style='background:{STATUS_COLOR['Missing']}'>Missing</span> 순위 없음<br>"
        "• <b>Strong이라도 순위가 10위 밖이면 SEO 최적화 필요</b>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.markdown("##### 프로그래매틱 SEO 실행 로드맵")
    st.markdown(
        "<div class='card'>"
        "1️⃣ <b>황금존 Missing 선점</b> — 청바지·바지·코트·셔츠·블라우스·티셔츠·후드집업 등 "
        "고MSV·저KD·미보유 키워드부터 카테고리 템플릿 페이지 생성<br>"
        "2️⃣ <b>허브-스포크 구조</b> — 카테고리(하의/아우터/상의…) 허브 아래 키워드별 페이지 연결, "
        "서브폴더 URL(`/category/`)로 도메인 권위 집중<br>"
        "3️⃣ <b>Weak 탈환</b> — 니트·크롭티·카드지갑 등 순위는 있으나 경쟁사에 밀리는 키워드 콘텐츠 강화<br>"
        "4️⃣ <b>Strong 방어·순위↑</b> — 드레스·모피·세미정장 등 10위 밖 Strong 키워드 온페이지 최적화<br>"
        "5️⃣ <b>품질 관리</b> — 페이지마다 고유 가치(상품 큐레이션·코디·사이즈 가이드)로 thin content 회피"
        "</div>",
        unsafe_allow_html=True,
    )
    st.caption("*순위/검색량은 Semrush 추정치로 실제 수치와 차이가 있을 수 있습니다.")

st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)
st.caption(f"패션 키워드 {len(df)}개 · LF몰 vs W컨셉·한섬·SSF샵·SI빌리지 · Semrush kr DB 실데이터 (2026-06-22)")
