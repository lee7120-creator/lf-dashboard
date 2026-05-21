import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="LF몰 세션 대시보드", layout="wide")
st.title("LF몰 세션 채널별 대시보드")

uploaded = st.file_uploader("xlsx 파일 업로드", type=["xlsx"])

if uploaded is None:
    st.info("xlsx 파일을 업로드하세요.")
    st.stop()

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

    for col in data.columns:
        if col not in ["date", "ym"]:
            data[col] = pd.to_numeric(data[col], errors="coerce")

    categories = ["직접", "광고", "PUSH", "EP", "미디어커머스", "브랜드광고", "제휴"]
    result_rows = []
    for ym, grp in data.groupby("ym"):
        row = {"ym": ym}
        for cat in categories:
            cols = [i for i, c in col_to_category.items() if c == cat and i in grp.columns]
            row[cat] = float(grp[cols].sum().sum())
        row["합계"] = sum(row[c] for c in categories)
        result_rows.append(row)

    df = pd.DataFrame(result_rows).sort_values("ym").reset_index(drop=True)
    return df

df = load_data(uploaded)
categories = ["직접", "광고", "PUSH", "EP", "미디어커머스", "브랜드광고", "제휴"]
colors = ["#185FA5", "#639922", "#BA7517", "#3C3489", "#888780", "#D85A30", "#0F6E56"]

months = df["ym"].tolist()
col_f1, col_f2 = st.columns(2)
start = col_f1.selectbox("시작 월", months, index=0)
end = col_f2.selectbox("종료 월", months, index=len(months)-1)
filtered = df[(df["ym"] >= start) & (df["ym"] <= end)].copy().reset_index(drop=True)

total = filtered["합계"].sum()
c1, c2, c3, c4 = st.columns(4)
c1.metric("총 세션", f"{total/1e8:.2f}억")
c2.metric("직접 비중", f"{filtered['직접'].sum()/total*100:.1f}%")
c3.metric("PUSH 비중", f"{filtered['PUSH'].sum()/total*100:.1f}%")
c4.metric("광고 비중", f"{filtered['광고'].sum()/total*100:.1f}%")

st.divider()

tab1, tab2, tab3 = st.tabs(["📈 월별 추이", "📊 채널 비중", "🔍 직접·PUSH 상세"])

with tab1:
    fig1 = go.Figure()
    for i, cat in enumerate(categories):
        fig1.add_trace(go.Scatter(
            x=filtered["ym"], y=filtered[cat],
            name=cat, mode="lines",
            line=dict(color=colors[i], width=2)
        ))
    fig1.update_layout(
        title="채널별 월별 세션 추이",
        height=420, hovermode="x unified",
        xaxis_title="월", yaxis_title="세션 수"
    )
    st.plotly_chart(fig1, use_container_width=True)

with tab2:
    fig2 = go.Figure()
    for i, cat in enumerate(categories):
        pct = (filtered[cat] / filtered["합계"] * 100).round(1)
        fig2.add_trace(go.Bar(
            x=filtered["ym"], y=pct,
            name=cat, marker_color=colors[i]
        ))
    fig2.update_layout(
        title="채널별 세션 비중 (%)",
        barmode="stack", height=420,
        hovermode="x unified",
        xaxis_title="월", yaxis_title="비중(%)"
    )
    st.plotly_chart(fig2, use_container_width=True)

with tab3:
    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(
        x=filtered["ym"], y=filtered["직접"],
        name="직접", fill="tozeroy", mode="lines",
        line=dict(color=colors[0], width=2)
    ))
    fig3.add_trace(go.Scatter(
        x=filtered["ym"], y=filtered["PUSH"],
        name="PUSH", fill="tozeroy", mode="lines",
        line=dict(color=colors[2], width=2, dash="dash")
    ))
    fig3.update_layout(
        title="직접 vs PUSH 세션 추이",
        height=420, hovermode="x unified",
        xaxis_title="월", yaxis_title="세션 수"
    )
    st.plotly_chart(fig3, use_container_width=True)

st.divider()
st.subheader("월별 원본 데이터")
st.dataframe(
    filtered[["ym"] + categories + ["합계"]],
    use_container_width=True,
    height=300
)
