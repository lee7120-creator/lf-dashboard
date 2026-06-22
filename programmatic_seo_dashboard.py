"""굿웨어몰 프로그래매틱 SEO 키워드 선점 분석 대시보드

대상       : goodwearmall.com (굿웨어몰)
경쟁사     : uniqlo.com(유니클로) · lfmall.co.kr(LF몰) · elandmall.co.kr(이랜드몰)
관점       : Programmatic SEO — 경쟁사 대비 미보유(Missing)·열위(Weak) 키워드를
             고MSV·저Difficulty 우선순위로 선점할 타겟 도출

지표 정의
  · Search Volume(MSV) : Google 기준 월별 검색량 (높을수록 매력)
  · Keyword Difficulty : 키워드 선점 난이도 0~100 (낮을수록 유리)
  · 도메인 컬럼 숫자     : 검색 시 순위(1~100). 0 = 순위 없음
  · Status
      - Strong  : 굿웨어몰 순위가 경쟁 3사보다 높음(앞섬)
      - Weak    : 굿웨어몰이 순위는 있으나 경쟁사보다 낮음
      - Missing : 굿웨어몰 순위 없음(0) → 신규 페이지 선점 기회
      - ※ Strong이라도 순위가 10위 밖이면 SEO 최적화 필요

데이터 출처
  · 키워드 보유량/순위는 내부 SEO 툴 추정치 (실제 수치와 차이 있을 수 있음)
  · 갱신 시: 동일 포맷(키워드, MSV, KD, 4개 도메인 순위)으로 아래 DATA 교체
    Semrush MCP로 확장하려면 organic_research → domain_organic / domain_domains 활용

실행: streamlit run programmatic_seo_dashboard.py
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="굿웨어몰 프로그래매틱 SEO 선점 분석", page_icon="🎯",
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

DOMAINS = ["goodwearmall.com", "uniqlo.com", "lfmall.co.kr", "elandmall.co.kr"]
DOMAIN_KR = {"goodwearmall.com": "굿웨어몰", "uniqlo.com": "유니클로",
             "lfmall.co.kr": "LF몰", "elandmall.co.kr": "이랜드몰"}
STATUS_COLOR = {"Strong": "#48bb78", "Weak": "#ed8936", "Missing": "#f56565"}

PALETTE = {"blue": "#4f8fff", "red": "#f56565", "amber": "#ed8936",
           "green": "#48bb78", "purple": "#9f7aea", "slate": "#64748b"}

def base_layout(h=320, ysuffix="", xsuffix="", title="", showlegend=False):
    return dict(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#475569", size=12), margin=dict(l=10, r=10, t=44, b=10),
        height=h, showlegend=showlegend,
        title=dict(text=title, font=dict(color="#94a3b8", size=13)),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1,
                    font=dict(size=11)),
        xaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11), ticksuffix=xsuffix),
        yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11), ticksuffix=ysuffix),
    )

# ══════════════════════════════════════════════════════
# DATA — 오가닉 키워드 상세 리포트 (내부 SEO 툴 추정치)
# (키워드, MSV, Difficulty, 굿웨어몰, 유니클로, LF몰, 이랜드몰, 카테고리)
# ══════════════════════════════════════════════════════
DATA = [
    ("청바지", 33100, 22, 0, 18, 0, 0, "하의"),
    ("패딩", 33100, 25, 43, 14, 0, 0, "아우터"),
    ("바지", 27100, 17, 0, 8, 0, 0, "하의"),
    ("수영복", 27100, 32, 7, 0, 0, 0, "수영/비치"),
    ("미니 스커트", 22200, 21, 0, 24, 0, 0, "하의"),
    ("스니커즈", 22200, 23, 0, 55, 0, 0, "신발"),
    ("스타킹", 22200, 24, 0, 0, 0, 20, "언더웨어"),
    ("레깅스", 18100, 31, 0, 12, 0, 41, "하의"),
    ("버뮤다 팬츠", 18100, 27, 23, 4, 0, 0, "하의"),
    ("코트", 18100, 20, 0, 18, 0, 0, "아우터"),
    ("후드 티", 18100, 19, 57, 33, 0, 0, "상의"),
    ("흰색", 18100, 40, 0, 78, 0, 0, "컬러"),
    ("래쉬 가드", 14800, 18, 15, 0, 0, 0, "수영/비치"),
    ("모자", 14800, 19, 0, 18, 0, 0, "액세서리"),
    ("선글라스", 14800, 20, 0, 18, 0, 0, "액세서리"),
    ("셔츠", 14800, 15, 0, 5, 0, 0, "상의"),
    ("정장", 14800, 20, 10, 6, 0, 0, "정장/포멀"),
    ("치마", 14800, 23, 0, 31, 0, 0, "하의"),
    ("파카", 14800, 37, 0, 7, 0, 0, "아우터"),
    ("니트", 12100, 17, 18, 6, 0, 0, "상의"),
    ("블라우스", 12100, 18, 12, 13, 0, 0, "상의"),
    ("크로스 백", 12100, 19, 0, 16, 0, 0, "가방"),
    ("크롭 컷", 12100, 29, 0, 30, 0, 0, "상의"),
    ("티셔츠", 12100, 17, 9, 1, 0, 0, "상의"),
    ("가디건", 9900, 24, 0, 5, 0, 0, "아우터"),
    ("구두", 9900, 22, 0, 15, 0, 0, "신발"),
    ("네이비", 9900, 39, 0, 0, 44, 0, "컬러"),
    ("데님", 9900, 19, 33, 14, 0, 0, "하의"),
    ("맨투맨", 9900, 23, 5, 0, 0, 0, "상의"),
    ("바람막이", 9900, 16, 69, 0, 0, 0, "아우터"),
    ("반바지", 9900, 17, 0, 1, 0, 0, "하의"),
    ("백팩", 9900, 15, 0, 22, 0, 0, "가방"),
    ("브랜드", 9900, 36, 49, 0, 0, 0, "일반"),
    ("블레이저", 9900, 22, 0, 13, 0, 0, "아우터"),
    ("슬랙스", 9900, 19, 5, 7, 0, 0, "하의"),
    ("양말", 9900, 18, 0, 5, 0, 0, "언더웨어"),
    ("여자 수영복", 9900, 19, 44, 43, 0, 0, "수영/비치"),
    ("크롭 티", 9900, 19, 6, 36, 0, 0, "상의"),
    ("후 리스", 9900, 23, 41, 4, 0, 0, "아우터"),
    ("후드 집업", 9900, 15, 0, 14, 0, 0, "상의"),
    ("나시", 8100, 20, 72, 69, 0, 0, "상의"),
    ("남자 패션", 8100, 25, 62, 0, 0, 0, "일반"),
    ("명품", 8100, 46, 0, 0, 0, 36, "일반"),
    ("버건디", 8100, 34, 0, 75, 6, 0, "컬러"),
    ("베레모", 8100, 27, 0, 0, 14, 0, "액세서리"),
    ("여자 속옷", 8100, 17, 72, 0, 10, 0, "언더웨어"),
    ("잠옷", 8100, 21, 55, 0, 0, 0, "언더웨어"),
    ("조거 팬츠", 8100, 19, 19, 12, 0, 0, "하의"),
    ("트렌치 코트", 8100, 25, 0, 15, 0, 0, "아우터"),
    ("후드", 8100, 20, 0, 16, 0, 0, "상의"),
    ("남자 셔츠", 6600, 17, 4, 7, 0, 0, "상의"),
    ("남자 쇼핑몰", 6600, 27, 80, 70, 0, 65, "일반"),
    ("남자 수영복", 6600, 16, 47, 49, 0, 0, "수영/비치"),
    ("롱 패딩", 6600, 19, 16, 0, 0, 0, "아우터"),
    ("반팔 티", 6600, 14, 73, 27, 0, 0, "상의"),
    ("베이지 색", 6600, 35, 0, 91, 0, 0, "컬러"),
    ("벨트", 6600, 23, 5, 4, 0, 0, "액세서리"),
    ("샌들", 6600, 25, 0, 53, 0, 0, "신발"),
    ("세미 정장", 6600, 19, 7, 32, 5, 0, "정장/포멀"),
    ("스웨터", 6600, 21, 15, 8, 0, 0, "상의"),
    ("슬링 백", 6600, 14, 0, 60, 0, 0, "가방"),
    ("와이드 팬츠", 6600, 14, 0, 3, 0, 0, "하의"),
    ("청자켓 코디", 6600, 27, 45, 0, 0, 0, "아우터"),
    ("카고 바지", 6600, 19, 19, 8, 0, 0, "하의"),
    ("탱크 탑", 6600, 21, 75, 11, 0, 0, "상의"),
    ("토트 백", 6600, 23, 0, 89, 0, 0, "가방"),
    ("회색", 6600, 32, 0, 82, 0, 0, "컬러"),
    ("경량 패딩", 5400, 19, 0, 6, 0, 0, "아우터"),
    ("나일론", 5400, 27, 0, 8, 0, 0, "소재"),
    ("남자 옷", 5400, 27, 67, 61, 0, 50, "일반"),
    ("무테 안경", 5400, 21, 0, 0, 19, 0, "액세서리"),
    ("브라 탑", 5400, 16, 3, 1, 0, 0, "언더웨어"),
    ("블루종", 5400, 24, 7, 5, 0, 0, "아우터"),
    ("어그 부츠", 5400, 23, 0, 0, 0, 14, "신발"),
    ("자켓", 5400, 15, 80, 0, 0, 0, "아우터"),
    ("조끼", 5400, 22, 0, 22, 0, 0, "상의"),
    ("차콜", 5400, 13, 0, 0, 22, 0, "컬러"),
    ("청자켓", 5400, 19, 1, 0, 0, 0, "아우터"),
    ("코르셋", 5400, 30, 0, 0, 8, 0, "언더웨어"),
    ("항공 점퍼", 5400, 17, 8, 0, 0, 0, "아우터"),
    ("등산화", 4400, 14, 0, 0, 0, 40, "신발"),
    ("러닝 플러스", 4400, 21, 73, 0, 0, 0, "일반"),
    ("린넨", 4400, 14, 0, 3, 0, 0, "소재"),
    ("린넨 셔츠", 4400, 19, 20, 6, 0, 0, "상의"),
    ("민소매", 4400, 22, 5, 0, 15, 0, "상의"),
    ("반팔 셔츠", 4400, 15, 10, 16, 0, 0, "상의"),
    ("버뮤다", 4400, 17, 4, 16, 60, 0, "하의"),
    ("뷔스티에", 4400, 22, 10, 0, 0, 0, "상의"),
    ("스웨이드", 4400, 24, 0, 26, 0, 0, "소재"),
    ("스카프", 4400, 23, 0, 56, 0, 0, "액세서리"),
]

df = pd.DataFrame(DATA, columns=["키워드", "MSV", "KD", "굿웨어몰", "유니클로",
                                 "LF몰", "이랜드몰", "카테고리"])

# ── Status / 등급 / 우선순위 점수 계산 ──
def calc_status(row):
    gw = row["굿웨어몰"]
    if gw == 0:
        return "Missing"
    comps = [row["유니클로"], row["LF몰"], row["이랜드몰"]]
    ranked = [c for c in comps if c > 0]
    if not ranked or gw < min(ranked):
        return "Strong"
    return "Weak"

def calc_tier(row):
    s, kd, msv, gw = row["Status"], row["KD"], row["MSV"], row["굿웨어몰"]
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
    # 프로그래매틱 SEO 우선순위 = MSV × 난이도할인 × 상태가중
    diff_factor = (100 - row["KD"]) / 100
    s, gw = row["Status"], row["굿웨어몰"]
    if s == "Missing":
        sf = 1.0
    elif s == "Weak":
        sf = 0.65
    elif s == "Strong" and gw > 10:
        sf = 0.45
    else:
        sf = 0.15
    return int(round(row["MSV"] * diff_factor * sf))

df["Status"] = df.apply(calc_status, axis=1)
df["타겟 등급"] = df.apply(calc_tier, axis=1)
df["우선순위"] = df.apply(calc_priority, axis=1)
df["굿_1P"] = df["굿웨어몰"].between(1, 10)  # 굿웨어몰 1페이지(1~10위) 여부

# ══════════════════════════════════════════════════════
# HEADER + KPI
# ══════════════════════════════════════════════════════
st.markdown("## 🎯 굿웨어몰 프로그래매틱 SEO 키워드 선점 분석")
st.markdown(
    "<div class='cap'>대상 <b>goodwearmall.com</b> · 경쟁사 "
    "<b>유니클로 · LF몰 · 이랜드몰</b> · 패션 카테고리 오가닉 키워드 "
    f"<b>{len(df)}</b>개 · 내부 SEO 툴 추정치</div>",
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
    st.caption("굿웨어몰 키워드 선점 분석")
    st.markdown("---")
    st.markdown("**Status 정의**")
    st.markdown(
        f"<span class='tag' style='background:{STATUS_COLOR['Strong']}'>Strong</span> 경쟁사보다 앞섬<br>"
        f"<span class='tag' style='background:{STATUS_COLOR['Weak']}'>Weak</span> 순위 있으나 열위<br>"
        f"<span class='tag' style='background:{STATUS_COLOR['Missing']}'>Missing</span> 순위 없음 → 선점 기회",
        unsafe_allow_html=True,
    )
    st.markdown("---")
    st.markdown("**선점 공식**")
    st.caption("MSV ↑ × Difficulty ↓ × Missing/Weak = 최우선. "
               "Strong이라도 10위 밖이면 최적화 필요.")
    st.markdown("---")
    st.caption("순위/보유량은 내부 SEO 툴 추정치. Semrush MCP "
               "(domain_organic / domain_domains)로 키워드 유니버스 확장 가능.")

# ══════════════════════════════════════════════════════
# 탭
# ══════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["🎯 기회 매트릭스", "🚀 선점 타겟", "🗂️ 카테고리 클러스터",
     "📊 Status 분석", "🔍 도메인 커버리지", "📖 가이드"]
)

# ──────────────────────────────────────────────────────
# TAB 1 — 기회 매트릭스 (Difficulty × MSV)
# ──────────────────────────────────────────────────────
with tab1:
    st.markdown("##### 키워드 기회 매트릭스 — 좌상단(고검색량·저난이도)이 황금존")
    fig = go.Figure()
    # 황금존 음영
    fig.add_shape(type="rect", x0=0, x1=25, y0=10000, y1=df["MSV"].max() * 1.05,
                  fillcolor="rgba(72,187,120,0.08)", line=dict(width=0), layer="below")
    fig.add_annotation(x=12.5, y=df["MSV"].max(), text="🥇 황금존 (선점 우선)",
                       showarrow=False, font=dict(color="#48bb78", size=12))
    for s in ["Missing", "Weak", "Strong"]:
        d = df[df["Status"] == s]
        fig.add_trace(go.Scatter(
            x=d["KD"], y=d["MSV"], mode="markers", name=s,
            marker=dict(size=(d["MSV"] / 1200 + 8), color=STATUS_COLOR[s],
                        opacity=0.78, line=dict(color="white", width=1)),
            customdata=d[["키워드", "타겟 등급", "굿웨어몰"]],
            hovertemplate=("<b>%{customdata[0]}</b><br>MSV %{y:,} · KD %{x}"
                           "<br>굿웨어몰 순위 %{customdata[2]}<br>%{customdata[1]}<extra></extra>"),
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
            ["키워드", "카테고리", "MSV", "KD", "유니클로", "LF몰", "이랜드몰"]],
        use_container_width=True, hide_index=True,
        column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")},
    )

# ──────────────────────────────────────────────────────
# TAB 2 — 선점 타겟 (우선순위 점수)
# ──────────────────────────────────────────────────────
with tab2:
    st.markdown("##### 프로그래매틱 SEO 선점 우선순위")
    st.caption("우선순위 = MSV × (난이도 할인) × 상태가중(Missing 1.0 / Weak 0.65 / Strong>10위 0.45). "
               "신규 페이지로 선점하거나 기존 페이지를 최적화할 순서.")
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
            ranked[["키워드", "타겟 등급", "MSV", "KD", "굿웨어몰", "Status", "우선순위"]],
            use_container_width=True, hide_index=True, height=max(420, topn * 17),
            column_config={
                "MSV": st.column_config.NumberColumn("MSV", format="%d"),
                "우선순위": st.column_config.ProgressColumn(
                    "우선순위", format="%d", min_value=0,
                    max_value=int(df["우선순위"].max())),
            },
        )
    # 등급 요약
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
# TAB 3 — 카테고리 클러스터 (허브-스포크 페이지 설계)
# ──────────────────────────────────────────────────────
with tab3:
    st.markdown("##### 카테고리 클러스터 — 허브·스포크 페이지 아키텍처")
    st.caption("프로그래매틱 SEO는 카테고리(허브) 아래 키워드별 페이지(스포크)를 묶는다. "
               "블록 크기=MSV, 색=Status. 빨간 블록이 많은 카테고리 = 선점 여지 큰 클러스터.")
    fig = px.treemap(
        df, path=[px.Constant("전체"), "카테고리", "키워드"], values="MSV",
        color="Status", color_discrete_map={**STATUS_COLOR, "(?)": "#cbd5e1"},
        custom_data=["KD", "굿웨어몰"],
    )
    fig.update_traces(
        marker=dict(line=dict(color="white", width=1)),
        hovertemplate="<b>%{label}</b><br>MSV %{value:,}<br>KD %{customdata[0]}<extra></extra>",
    )
    fig.update_layout(margin=dict(l=8, r=8, t=10, b=8), height=460)
    st.plotly_chart(fig, use_container_width=True)

    cat = (df.groupby("카테고리")
             .agg(키워드수=("키워드", "count"), 총MSV=("MSV", "sum"),
                  평균KD=("KD", "mean"),
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
# TAB 4 — Status 분석 (필터 가능 전체 테이블)
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

    # Status 분포 (개수 + 총 MSV)
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
        fdf[["키워드", "카테고리", "MSV", "KD", "굿웨어몰", "유니클로", "LF몰",
             "이랜드몰", "Status", "타겟 등급"]].sort_values("MSV", ascending=False),
        use_container_width=True, hide_index=True, height=460,
        column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")},
    )
    st.download_button("⬇️ 필터된 키워드 CSV 내려받기",
                       fdf.to_csv(index=False).encode("utf-8-sig"),
                       "goodwearmall_keywords.csv", "text/csv")

# ──────────────────────────────────────────────────────
# TAB 5 — 도메인 커버리지
# ──────────────────────────────────────────────────────
with tab5:
    st.markdown("##### 도메인별 커버리지 — 이 키워드셋에서 누가 얼마나 노출되나")
    rows = []
    for col in ["굿웨어몰", "유니클로", "LF몰", "이랜드몰"]:
        ranked_mask = df[col] > 0
        p1 = df[col].between(1, 10).sum()
        avg = df.loc[ranked_mask, col].mean()
        rows.append((col, int(ranked_mask.sum()), int(p1),
                     round(avg, 1) if ranked_mask.any() else 0))
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
        st.dataframe(cov, use_container_width=True, hide_index=True, height=200)
        st.caption("유니클로가 패션 일반 키워드를 가장 넓게 선점 → 핵심 경쟁자. "
                   "**LF몰·이랜드몰은 이 키워드셋에서 노출이 거의 없음** → "
                   "굿웨어몰이 유니클로만 넘으면 선점 가능한 영역이 넓음.")

    # 굿웨어몰 vs 유니클로 직접 비교 (둘 중 하나라도 순위 있는 키워드)
    st.markdown("##### 굿웨어몰 vs 유니클로 — 직접 맞대결 키워드")
    duel = df[(df["굿웨어몰"] > 0) | (df["유니클로"] > 0)].copy()
    duel["우위"] = duel.apply(
        lambda r: "굿웨어몰" if (r["굿웨어몰"] > 0 and (r["유니클로"] == 0 or r["굿웨어몰"] < r["유니클로"]))
        else ("유니클로" if r["유니클로"] > 0 else "굿웨어몰"), axis=1)
    win = (duel["우위"] == "굿웨어몰").sum()
    st.caption(f"맞대결 {len(duel)}개 중 굿웨어몰 우위 **{win}개** / 유니클로 우위 **{len(duel) - win}개**")
    st.dataframe(
        duel[["키워드", "MSV", "KD", "굿웨어몰", "유니클로", "Status"]]
        .sort_values("MSV", ascending=False),
        use_container_width=True, hide_index=True, height=320,
        column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")},
    )

# ──────────────────────────────────────────────────────
# TAB 6 — 가이드
# ──────────────────────────────────────────────────────
with tab6:
    st.markdown("##### 분석 대상")
    st.markdown(
        "<div class='card'>굿웨어몰(goodwearmall.com)과 경쟁 3사(유니클로·LF몰·이랜드몰)가 "
        "보유한 키워드 중 <b>패션 관련 키워드</b>만 선별해 정리했습니다.</div>",
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
        "굿웨어몰 순위가 경쟁 3사보다 높음 / "
        f"<span class='tag' style='background:{STATUS_COLOR['Weak']}'>Weak</span> 경쟁사보다 낮음 / "
        f"<span class='tag' style='background:{STATUS_COLOR['Missing']}'>Missing</span> 순위 없음<br>"
        "• <b>Strong이라도 순위가 10위 밖이면 SEO 최적화 필요</b>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.markdown("##### 프로그래매틱 SEO 실행 로드맵")
    st.markdown(
        "<div class='card'>"
        "1️⃣ <b>황금존 Missing 선점</b> — 청바지·바지·셔츠·반바지·후드집업·와이드팬츠 등 "
        "고MSV·저KD·미보유 키워드부터 카테고리 템플릿 페이지 생성<br>"
        "2️⃣ <b>허브-스포크 구조</b> — 카테고리(하의/아우터/상의…) 허브 아래 키워드별 페이지 연결, "
        "서브폴더 URL(`/category/`)로 도메인 권위 집중<br>"
        "3️⃣ <b>Weak 탈환</b> — 데님·니트·후드티 등 순위는 있으나 유니클로에 밀리는 키워드 콘텐츠 강화<br>"
        "4️⃣ <b>Strong 방어·순위↑</b> — 바람막이·브랜드 등 10위 밖 Strong 키워드 온페이지 최적화<br>"
        "5️⃣ <b>품질 관리</b> — 페이지마다 고유 가치(상품 큐레이션·코디·사이즈 가이드)로 thin content 회피"
        "</div>",
        unsafe_allow_html=True,
    )
    st.caption("*키워드 보유량/순위는 내부 SEO 툴 추정치로 실제 수치와 차이가 있을 수 있습니다.")

st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)
st.caption(f"패션 키워드 {len(df)}개 · 굿웨어몰 vs 유니클로·LF몰·이랜드몰 · 내부 SEO 툴 추정치 기반")
