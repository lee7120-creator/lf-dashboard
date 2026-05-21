import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="LF몰 세션 대시보드", layout="wide")
st.title("LF몰 세션 채널별 대시보드")

uploaded = st.file_uploader("xlsx 파일 업로드", type=["xlsx"])

if uploaded is None:
    st.info("좌측 상단에서 xlsx 파일을 업로드하세요.")
    st.stop()

@st.cache_data
def load_data(file):
    xl = pd.read_excel(file, sheet_name="세션", header=None)

    row0 = xl.iloc[0].tolist()
    col_to_category = {i: row0[i] for i in range(1, len(row0)) if pd.notna(row0[i])}

    data = xl.iloc[4:].copy()
    data = data.rename(columns={0: "date"})
    data["date"] = data["date"].apply(
        lambda x: str(int(x)) if pd.notna(x) and str(x).replace(".0","").isdigit() else None
    )
    data = data.dropna(subset=["date"])
    data = data[data["date"].str.match(r"^\d{8}$", na=False)]
    data["date"] = pd.to_datetime(data["date"], format="%Y%m%d")
    data["ym"] = data["date"].dt.to_period("M").astype(str)

    for col in data.columns[1:]:
        data[col] = pd.to_numeric(data[col], errors="coerce")

    categories = ["직접", "광고", "PUSH", "EP", "미디어커머스", "브랜드광고", "제휴"]
    result_rows = []
    for ym, grp in data.groupby("ym"):
        row = {"연월": ym}
        for cat in categories:
            cols = [i for i, c in col_to_category.items() if c == cat and i in grp.columns]
            row[cat] = grp[cols].sum().sum()
        row["합계"] = sum(row[c] for c in categories)
        result_rows.append(row)

    return pd.DataFrame(result_rows).sort_values("연월").reset_index(drop=True)

df = load_data(uploaded)
categories = ["직접", "광고", "PUSH", "EP", "미디어커머스", "브랜드광고", "제휴"]

# 기간 필터
months = df["연월"].tolist()
col_f1, col_f2 = st.columns(2)
start = col_f1.selectbox("시작 월", months, index=0)
end = col_f2.selectbox("종료 월", months, index=len(months)-1)
mask = (df["연월"] >= start) & (df["연월"] <= end)
filtered = df[mask].copy()

# 요약 지표
total = filtered["합계"].sum()
c1, c2, c3, c4 = st.columns(4)
c1.metric("총 세션", f"{total/1e8:.2f}억")
c2.metric("직접 비중", f"{filtered['직접'].sum()/total*100:.1f}%")
c3.metric("PUSH 비중", f"{filtered['PUSH'].sum()/total*100:.1f}%")
c4.metric("광고 비중", f"{filtered['광고'].sum()/total*100:.1f}%")

st.divider()

tab1, tab2, tab3 = st.tabs(["📈 월별 추이", "📊 채널 비중", "🔍 직접·PUSH 상세"])

with tab1:
    fig = px.line(
        filtered, x="연월", y=categories,
        title="채널별 월별 세션 추이",
        labels={"value": "세션 수", "variable": "채널"},
    )
    fig.update_layout(height=420, hovermode="x unified")
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    filtered_pct = filtered.copy()
    for cat in categories:
        filtered_pct[cat] = filtered_pct[cat] / filtered_pct["합계"] * 100
    fig2 = px.bar(
        filtered_pct, x="연월", y=categories,
        title="채널별 세션 비중 (%)",
        labels={"value": "비중(%)", "variable": "채널"},
    )
    fig2.update_layout(height=420, barmode="stack", hovermode="x unified")
    st.plotly_chart(fig2, use_container_width=True)

with tab3:
    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(
        x=filtered["연월"], y=filtered["직접"],
        name="직접", fill="tozeroy", mode="lines"
    ))
    fig3.add_trace(go.Scatter(
        x=filtered["연월"], y=filtered["PUSH"],
        name="PUSH", fill="tozeroy", mode="lines",
        line=dict(dash="dash")
    ))
    fig3.update_layout(
        title="직접 vs PUSH 세션 추이",
        height=420, hovermode="x unified"
    )
    st.plotly_chart(fig3, use_container_width=True)

st.divider()
st.subheader("월별 원본 데이터")
st.dataframe(
    filtered[["연월"] + categories + ["합계"]].style.format("{:,.0f}", subset=categories + ["합계"]),
    use_container_width=True, height=300
)
