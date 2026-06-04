"""첫구매 채널별 분석 대시보드"""

import io
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import seaborn as sns
import streamlit as st
from scipy import stats

# ── 페이지 설정 ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="첫구매 채널 분석",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 한국어 폰트 설정 ──────────────────────────────────────────────────────────
def _set_korean_font():
    import matplotlib.font_manager as fm
    candidates = [
        "NanumGothic", "NanumBarunGothic", "Malgun Gothic",
        "Apple SD Gothic Neo", "Noto Sans KR", "UnDotum", "Gulim",
    ]
    available = {f.name for f in fm.fontManager.ttflist}
    for c in candidates:
        if c in available:
            plt.rcParams["font.family"] = c
            break
    plt.rcParams["axes.unicode_minus"] = False
    sns.set_theme(style="whitegrid", font=plt.rcParams["font.family"])

_set_korean_font()

# ── 글로벌 CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .kpi-card{background:#f8fafc;border-radius:12px;padding:18px 20px;margin:4px 2px}
    .kpi-label{font-size:12px;color:#64748b;font-weight:600;margin-bottom:4px}
    .kpi-value{font-size:22px;font-weight:700;color:#1e293b;line-height:1.2}
    .kpi-delta-up{font-size:11px;color:#16a34a}
    .kpi-delta-dn{font-size:11px;color:#dc2626}
    section[data-testid="stSidebar"] .block-container{padding-top:1rem}
</style>
""", unsafe_allow_html=True)

# ── 채널 색상 ──────────────────────────────────────────────────────────────────
CHANNEL_COLORS = {
    "*TOTAL":    "#334155",
    "직접":      "#3b82f6",
    "광고":      "#f59e0b",
    "EP":        "#10b981",
    "PUSH":      "#8b5cf6",
    "제휴":      "#ef4444",
    "브랜드광고": "#06b6d4",
    "미디어커머스": "#f97316",
}
CHANNELS = ["직접", "광고", "EP", "PUSH", "제휴", "브랜드광고", "미디어커머스"]
PALETTE  = [CHANNEL_COLORS[c] for c in CHANNELS]

METRIC_KO = {
    "일평균거래액":  "일평균거래액",
    "거래액비중":    "거래액 비중",
    "일평균고객수":  "일평균고객수",
    "고객비중":      "고객 비중",
    "일평균객단가":  "일평균객단가",
    "유효회원수":    "유효회원수",
    "DAU":           "DAU",
    "유입율":        "유입율",
    "CR":            "전환율(CR)",
}

# ── 데이터 파싱 ───────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_data(file_bytes: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")

    # 연도 감지 (row 0)
    year = 2026
    for v in df_raw.iloc[0, :]:
        try:
            iv = int(v)
            if 2020 <= iv <= 2030:
                year = iv
                break
        except Exception:
            pass

    # 날짜 헤더 (row 2, col 6~)
    raw_dates = df_raw.iloc[2, 6:].tolist()
    dates = []
    for ds in raw_dates:
        if pd.isnull(ds):
            dates.append(pd.NaT)
        else:
            try:
                dates.append(pd.Timestamp(f"{year}/{str(ds).strip()}"))
            except Exception:
                dates.append(pd.NaT)

    n = len(dates)
    records = []
    current_metric = None
    for ri in range(3, df_raw.shape[0]):
        m   = df_raw.iloc[ri, 0]
        seg = df_raw.iloc[ri, 1]
        if not pd.isnull(m) and str(m).strip() not in ("-", ""):
            current_metric = str(m).strip()
        if pd.isnull(seg) or str(seg).strip() in ("-", "") or current_metric is None:
            continue
        seg  = str(seg).strip()
        vals = pd.to_numeric(df_raw.iloc[ri, 6:6 + n].values, errors="coerce")
        for dt, val in zip(dates, vals):
            if pd.isnull(dt):
                continue
            records.append({"date": dt, "metric": current_metric, "segment": seg, "value": val})

    df = pd.DataFrame(records)
    df["date"]  = pd.to_datetime(df["date"])
    df["month"] = df["date"].dt.to_period("M").astype(str)
    df["dow"]   = df["date"].dt.dayofweek
    return df

# ── 인사이트 저장/로드 ────────────────────────────────────────────────────────
STORE_FILE = "fp_insights.json"

def load_store() -> dict:
    if os.path.exists(STORE_FILE):
        try:
            with open(STORE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_store(store: dict):
    with open(STORE_FILE, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)

if "fp_store" not in st.session_state:
    st.session_state.fp_store = load_store()

def memo_block(key: str):
    store = st.session_state.fp_store
    val   = store.get(key, "")
    with st.expander("📝 인사이트 메모", expanded=bool(val)):
        new_val = st.text_area(
            "", value=val, key=f"_memo_{key}",
            placeholder="분석 인사이트를 입력하세요…", height=100,
        )
        if st.button("저장", key=f"_save_{key}"):
            store[key] = new_val
            save_store(store)
            st.success("저장되었습니다.")

# ── KPI 카드 ──────────────────────────────────────────────────────────────────
def kpi_card(label: str, value: str, delta_pct: float | None = None):
    delta_html = ""
    if delta_pct is not None:
        cls  = "kpi-delta-up" if delta_pct >= 0 else "kpi-delta-dn"
        sign = "▲" if delta_pct >= 0 else "▼"
        delta_html = f'<div class="{cls}">{sign} {abs(delta_pct):.1f}% (전주 대비)</div>'
    st.markdown(
        f'<div class="kpi-card">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'{delta_html}</div>',
        unsafe_allow_html=True,
    )

# ── 포맷 헬퍼 ─────────────────────────────────────────────────────────────────
def fmt_num(v, suffix=""):
    if pd.isnull(v): return "–"
    av = abs(v)
    if av >= 1e8: return f"{v/1e8:.1f}억{suffix}"
    if av >= 1e4: return f"{v/1e4:.0f}만{suffix}"
    return f"{v:,.0f}{suffix}"

def fmt_pct(v):
    if pd.isnull(v): return "–"
    return f"{v*100:.2f}%"

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("🛒 첫구매 분석")
    uploaded = st.file_uploader(
        "Excel 파일 업로드", type=["xlsx", "xls"], key="fp_up",
        help="첫구매 채널별 지표가 담긴 엑셀 파일을 업로드하세요.",
    )
    st.markdown("---")

    PAGE_LIST = [
        "01. 개요",
        "02. 거래액 분석",
        "03. 고객수 분석",
        "04. 채널 효율",
        "05. 채널 비중 추이",
        "06. 월별 요약",
    ]
    if "fp_page" not in st.session_state:
        st.session_state.fp_page = PAGE_LIST[0]
    page = st.radio("페이지", PAGE_LIST, key="fp_page")

# ── 데이터 로드 ───────────────────────────────────────────────────────────────
if uploaded is None:
    st.info("👈 사이드바에서 Excel 파일을 업로드해주세요.")
    st.stop()

with st.spinner("데이터 파싱 중…"):
    df = parse_data(uploaded.read())

# ── 공통 헬퍼 ─────────────────────────────────────────────────────────────────
def get_pivot(metric: str, segs: list | None = None) -> pd.DataFrame:
    mask = df["metric"] == metric
    if segs:
        mask &= df["segment"].isin(segs)
    return (
        df[mask]
        .pivot_table(index="date", columns="segment", values="value", aggfunc="first")
        .sort_index()
    )

def get_series(metric: str, seg: str) -> pd.Series:
    return (
        df[(df["metric"] == metric) & (df["segment"] == seg)]
        .sort_values("date")
        .set_index("date")["value"]
    )

def week_avg(metric: str, seg: str, last_date: pd.Timestamp):
    cur = df[
        (df["metric"] == metric) & (df["segment"] == seg) &
        (df["date"] >= last_date - pd.Timedelta(days=6))
    ]["value"].mean()
    prev = df[
        (df["metric"] == metric) & (df["segment"] == seg) &
        (df["date"] >= last_date - pd.Timedelta(days=13)) &
        (df["date"] <  last_date - pd.Timedelta(days=6))
    ]["value"].mean()
    delta = (cur - prev) / prev if prev and not np.isnan(prev) else None
    return cur, delta

def _fig(w=14, h=4):
    fig, ax = plt.subplots(figsize=(w, h))
    return fig, ax

def _show(fig):
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

def _monthly_avg(metric: str, segs: list) -> pd.DataFrame:
    return (
        df[(df["metric"] == metric) & (df["segment"].isin(segs))]
        .groupby(["month", "segment"])["value"]
        .mean()
        .unstack("segment")
        .reindex(columns=segs)
    )

# ══════════════════════════════════════════════════════════════════════════════
# 01. 개요
# ══════════════════════════════════════════════════════════════════════════════
if page == "01. 개요":
    st.header("📊 첫구매 채널 분석 — 개요")

    last_date = df["date"].max()
    date_range = f"{df['date'].min().strftime('%Y-%m-%d')} ~ {last_date.strftime('%Y-%m-%d')}"
    st.caption(f"데이터 기간: {date_range}  |  총 {df['date'].nunique()}일")

    # KPI 행
    kpi_specs = [
        ("일평균거래액", "*TOTAL", "일평균거래액",   lambda v: fmt_num(v, "원")),
        ("일평균고객수", "*TOTAL", "일평균고객수",   lambda v: f"{v:,.0f}명"),
        ("일평균객단가", "*TOTAL", "일평균객단가",   lambda v: fmt_num(v, "원")),
        ("CR",          "*TOTAL", "전환율(CR)",      fmt_pct),
        ("유입율",       "*TOTAL", "유입율",         fmt_pct),
    ]
    cols = st.columns(5)
    for col, (metric, seg, label, fmtr) in zip(cols, kpi_specs):
        cur, delta = week_avg(metric, seg, last_date)
        with col:
            kpi_card(label, fmtr(cur), delta_pct=delta * 100 if delta is not None else None)

    st.markdown("---")
    col_l, col_r = st.columns(2)

    # 최근 30일 채널 거래액 비중 바차트
    with col_l:
        st.subheader("채널별 거래액 비중 (최근 30일 평균)")
        share_30 = (
            df[
                (df["metric"] == "거래액비중") &
                (df["segment"].isin(CHANNELS)) &
                (df["date"] >= last_date - pd.Timedelta(days=29))
            ]
            .groupby("segment")["value"].mean()
            .reindex(CHANNELS).fillna(0)
        )
        fig, ax = _fig(6, 4)
        colors = [CHANNEL_COLORS[c] for c in share_30.index]
        sns.barplot(x=share_30.values * 100, y=share_30.index,
                    palette=colors, ax=ax, orient="h")
        for patch, val in zip(ax.patches, share_30.values):
            ax.text(val * 100 + 0.3,
                    patch.get_y() + patch.get_height() / 2,
                    f"{val*100:.1f}%", va="center", fontsize=10)
        ax.set_xlabel("거래액 비중 (%)")
        ax.set_ylabel("")
        ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))
        sns.despine(ax=ax)
        _show(fig)

    # 일평균거래액 전체 추이
    with col_r:
        st.subheader("일평균거래액 전체 추이 (7일 이동평균)")
        total_tx = get_series("일평균거래액", "*TOTAL").rolling(7).mean()
        fig, ax = _fig(6, 4)
        sns.lineplot(x=total_tx.index, y=total_tx.values / 1e8,
                     ax=ax, color="#3b82f6", linewidth=2.2)
        ax.fill_between(total_tx.index, total_tx.values / 1e8, alpha=0.15, color="#3b82f6")
        ax.set_ylabel("거래액 (억원)")
        ax.set_xlabel("")
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.1f}억"))
        plt.xticks(rotation=30, fontsize=8)
        sns.despine(ax=ax)
        _show(fig)

    # 채널별 최근 30일 KPI 요약표
    st.subheader("채널별 주요 지표 요약 (전체 기간 평균)")
    summary_rows = []
    for ch in CHANNELS:
        row = {"채널": ch}
        for m, fmtr in [
            ("일평균거래액", lambda v: fmt_num(v, "원")),
            ("일평균고객수", lambda v: f"{v:,.0f}명"),
            ("일평균객단가", lambda v: fmt_num(v, "원")),
            ("CR",           fmt_pct),
            ("유입율",        fmt_pct),
        ]:
            s = df[(df["metric"] == m) & (df["segment"] == ch)]["value"]
            row[METRIC_KO.get(m, m)] = fmtr(s.mean()) if len(s) else "–"
        summary_rows.append(row)
    st.dataframe(pd.DataFrame(summary_rows).set_index("채널"), use_container_width=True)

    memo_block("overview")

# ══════════════════════════════════════════════════════════════════════════════
# 02. 거래액 분석
# ══════════════════════════════════════════════════════════════════════════════
elif page == "02. 거래액 분석":
    st.header("💰 거래액 분석")

    # 채널별 일평균거래액 추이
    st.subheader("채널별 일평균거래액 추이 (7일 이동평균)")
    pivot_tx = get_pivot("일평균거래액", CHANNELS).rolling(7).mean()
    fig, ax = _fig(14, 5)
    for ch in CHANNELS:
        if ch in pivot_tx.columns:
            ax.plot(pivot_tx.index, pivot_tx[ch] / 1e8,
                    label=ch, color=CHANNEL_COLORS[ch], linewidth=1.8)
    ax.set_ylabel("거래액 (억원)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.1f}억"))
    ax.legend(loc="upper right", fontsize=9, ncol=3)
    plt.xticks(rotation=30, fontsize=8)
    sns.despine(ax=ax)
    _show(fig)

    # 월별 채널별 일평균거래액
    st.subheader("월별 채널별 일평균거래액")
    monthly_tx = _monthly_avg("일평균거래액", CHANNELS).reset_index().melt(
        id_vars="month", var_name="채널", value_name="거래액"
    )
    fig, ax = _fig(13, 5)
    sns.barplot(
        data=monthly_tx, x="month", y="거래액", hue="채널",
        palette=PALETTE, ax=ax,
    )
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x/1e8:.1f}억"))
    ax.set_xlabel("")
    ax.set_ylabel("일평균거래액 (억원)")
    ax.legend(loc="upper right", fontsize=9, ncol=2)
    plt.xticks(rotation=20, fontsize=9)
    sns.despine(ax=ax)
    _show(fig)

    # 거래액비중 히트맵 (월별)
    st.subheader("월별 채널 거래액 비중 히트맵 (%)")
    hm_tx = _monthly_avg("거래액비중", CHANNELS).fillna(0)
    fig, ax = _fig(12, 4)
    sns.heatmap(
        hm_tx.T * 100, annot=True, fmt=".1f", cmap="YlOrRd",
        linewidths=0.5, ax=ax, cbar_kws={"label": "비중 (%)"},
    )
    ax.set_xlabel("")
    ax.set_ylabel("")
    plt.xticks(rotation=30, fontsize=9)
    _show(fig)

    # 채널별 거래액 박스플롯
    st.subheader("채널별 일 거래액 분포 (Box Plot)")
    box_data = df[(df["metric"] == "일평균거래액") & (df["segment"].isin(CHANNELS))].copy()
    box_data["거래액(억원)"] = box_data["value"] / 1e8
    fig, ax = _fig(10, 4)
    sns.boxplot(
        data=box_data, x="segment", y="거래액(억원)",
        order=CHANNELS, palette=PALETTE, ax=ax,
        flierprops={"marker": ".", "markersize": 4},
    )
    ax.set_xlabel("")
    ax.set_ylabel("일평균거래액 (억원)")
    sns.despine(ax=ax)
    _show(fig)

    memo_block("tx_analysis")

# ══════════════════════════════════════════════════════════════════════════════
# 03. 고객수 분석
# ══════════════════════════════════════════════════════════════════════════════
elif page == "03. 고객수 분석":
    st.header("👥 고객수 분석")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("일평균고객수 추이 (7일 이동평균)")
        pivot = get_pivot("일평균고객수", CHANNELS).rolling(7).mean()
        fig, ax = _fig(7, 4)
        for ch in CHANNELS:
            if ch in pivot.columns:
                ax.plot(pivot.index, pivot[ch], label=ch,
                        color=CHANNEL_COLORS[ch], linewidth=1.6)
        ax.set_ylabel("고객수 (명)")
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{int(x):,}"))
        ax.legend(fontsize=8, ncol=2)
        plt.xticks(rotation=30, fontsize=8)
        sns.despine(ax=ax)
        _show(fig)

    with col2:
        st.subheader("DAU 추이 (7일 이동평균)")
        pivot = get_pivot("DAU", CHANNELS).rolling(7).mean()
        fig, ax = _fig(7, 4)
        for ch in CHANNELS:
            if ch in pivot.columns:
                ax.plot(pivot.index, pivot[ch], label=ch,
                        color=CHANNEL_COLORS[ch], linewidth=1.6)
        ax.set_ylabel("DAU")
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{int(x):,}"))
        ax.legend(fontsize=8, ncol=2)
        plt.xticks(rotation=30, fontsize=8)
        sns.despine(ax=ax)
        _show(fig)

    st.markdown("---")

    st.subheader("유효회원수 추이 (7일 이동평균)")
    pivot = get_pivot("유효회원수", CHANNELS).rolling(7).mean()
    fig, ax = _fig(14, 4)
    for ch in CHANNELS:
        if ch in pivot.columns:
            ax.plot(pivot.index, pivot[ch] / 1e4, label=ch,
                    color=CHANNEL_COLORS[ch], linewidth=1.6)
    ax.set_ylabel("유효회원수 (만명)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.1f}만"))
    ax.legend(fontsize=9, ncol=4)
    plt.xticks(rotation=30, fontsize=8)
    sns.despine(ax=ax)
    _show(fig)

    # 고객 비중 히트맵
    st.subheader("월별 채널 고객 비중 히트맵 (%)")
    hm_c = _monthly_avg("고객비중", CHANNELS).fillna(0)
    fig, ax = _fig(12, 4)
    sns.heatmap(
        hm_c.T * 100, annot=True, fmt=".1f", cmap="Blues",
        linewidths=0.5, ax=ax, cbar_kws={"label": "비중 (%)"},
    )
    ax.set_xlabel("")
    ax.set_ylabel("")
    plt.xticks(rotation=30, fontsize=9)
    _show(fig)

    # 고객수 vs 거래액 버블차트
    st.subheader("채널별 고객수 vs 거래액 (월 평균)")
    bubble_rows = []
    for ch in CHANNELS:
        c_m = df[(df["metric"] == "일평균고객수") & (df["segment"] == ch)].groupby("month")["value"].mean()
        t_m = df[(df["metric"] == "일평균거래액") & (df["segment"] == ch)].groupby("month")["value"].mean()
        for m in c_m.index.intersection(t_m.index):
            bubble_rows.append({"채널": ch, "month": m,
                                 "고객수": c_m[m], "거래액": t_m[m]})
    bdf = pd.DataFrame(bubble_rows)
    fig, ax = _fig(9, 5)
    for ch in CHANNELS:
        sub = bdf[bdf["채널"] == ch]
        if len(sub):
            ax.scatter(sub["고객수"], sub["거래액"] / 1e8,
                       label=ch, color=CHANNEL_COLORS[ch], s=70, alpha=0.85)
    ax.set_xlabel("일평균 고객수 (명)")
    ax.set_ylabel("일평균 거래액 (억원)")
    ax.legend(fontsize=9)
    sns.despine(ax=ax)
    _show(fig)

    memo_block("customer_analysis")

# ══════════════════════════════════════════════════════════════════════════════
# 04. 채널 효율
# ══════════════════════════════════════════════════════════════════════════════
elif page == "04. 채널 효율":
    st.header("⚡ 채널 효율 분석")

    # CR 추이
    st.subheader("전환율(CR) 추이 (7일 이동평균)")
    pivot_cr = get_pivot("CR", CHANNELS).rolling(7).mean()
    fig, ax = _fig(14, 4)
    for ch in CHANNELS:
        if ch in pivot_cr.columns:
            ax.plot(pivot_cr.index, pivot_cr[ch] * 100, label=ch,
                    color=CHANNEL_COLORS[ch], linewidth=1.8)
    ax.set_ylabel("CR (%)")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f%%"))
    ax.legend(fontsize=9, ncol=4)
    plt.xticks(rotation=30, fontsize=8)
    sns.despine(ax=ax)
    _show(fig)

    col1, col2 = st.columns(2)

    # 평균 CR 바차트
    with col1:
        st.subheader("채널별 평균 CR (전체 기간)")
        avg_cr = (
            df[(df["metric"] == "CR") & (df["segment"].isin(CHANNELS))]
            .groupby("segment")["value"].mean()
            .reindex(CHANNELS).fillna(0)
        )
        fig, ax = _fig(6, 4)
        sns.barplot(
            x=avg_cr.index, y=avg_cr.values * 100,
            palette=[CHANNEL_COLORS[c] for c in avg_cr.index], ax=ax,
        )
        for patch, val in zip(ax.patches, avg_cr.values):
            ax.text(
                patch.get_x() + patch.get_width() / 2,
                patch.get_height() + 0.002,
                f"{val*100:.2f}%", ha="center", va="bottom", fontsize=9,
            )
        ax.set_ylabel("CR (%)")
        ax.set_xlabel("")
        plt.xticks(rotation=20, fontsize=9)
        sns.despine(ax=ax)
        _show(fig)

    # 평균 유입율 바차트
    with col2:
        st.subheader("채널별 평균 유입율 (전체 기간)")
        avg_ir = (
            df[(df["metric"] == "유입율") & (df["segment"].isin(CHANNELS))]
            .groupby("segment")["value"].mean()
            .reindex(CHANNELS).fillna(0)
        )
        fig, ax = _fig(6, 4)
        sns.barplot(
            x=avg_ir.index, y=avg_ir.values * 100,
            palette=[CHANNEL_COLORS[c] for c in avg_ir.index], ax=ax,
        )
        for patch, val in zip(ax.patches, avg_ir.values):
            ax.text(
                patch.get_x() + patch.get_width() / 2,
                patch.get_height() + 0.002,
                f"{val*100:.2f}%", ha="center", va="bottom", fontsize=9,
            )
        ax.set_ylabel("유입율 (%)")
        ax.set_xlabel("")
        plt.xticks(rotation=20, fontsize=9)
        sns.despine(ax=ax)
        _show(fig)

    st.markdown("---")

    # 객단가 추이
    st.subheader("일평균객단가 추이 (7일 이동평균)")
    pivot_aov = get_pivot("일평균객단가", CHANNELS).rolling(7).mean()
    fig, ax = _fig(14, 4)
    for ch in CHANNELS:
        if ch in pivot_aov.columns:
            ax.plot(pivot_aov.index, pivot_aov[ch] / 1e4, label=ch,
                    color=CHANNEL_COLORS[ch], linewidth=1.8)
    ax.set_ylabel("객단가 (만원)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.1f}만"))
    ax.legend(fontsize=9, ncol=4)
    plt.xticks(rotation=30, fontsize=8)
    sns.despine(ax=ax)
    _show(fig)

    # 효율 지표 상관 히트맵
    st.subheader("효율 지표 간 상관관계 (전체 채널 통합)")
    eff_metrics = ["일평균거래액", "일평균고객수", "일평균객단가", "CR", "유입율", "DAU"]
    corr_rows = []
    for m in eff_metrics:
        s = (
            df[(df["metric"] == m) & (df["segment"] == "*TOTAL")]
            .set_index("date")["value"]
            .sort_index()
        )
        corr_rows.append(s.rename(METRIC_KO.get(m, m)))
    corr_df = pd.concat(corr_rows, axis=1).dropna()
    if len(corr_df) > 5:
        fig, ax = _fig(8, 6)
        mask = np.triu(np.ones_like(corr_df.corr(), dtype=bool))
        sns.heatmap(
            corr_df.corr(), annot=True, fmt=".2f", cmap="coolwarm",
            center=0, mask=mask, linewidths=0.5, ax=ax,
            cbar_kws={"label": "상관계수"},
        )
        ax.set_title("지표 간 상관관계 (Pearson r)")
        _show(fig)

    # 선형회귀: 유입율 → CR
    st.subheader("유입율 → CR 선형회귀 (채널별)")
    reg_rows = []
    for ch in CHANNELS:
        ir = df[(df["metric"] == "유입율") & (df["segment"] == ch)].set_index("date")["value"]
        cr = df[(df["metric"] == "CR")     & (df["segment"] == ch)].set_index("date")["value"]
        common = ir.index.intersection(cr.index)
        if len(common) < 10:
            continue
        slope, intercept, r, p, _ = stats.linregress(ir[common], cr[common])
        reg_rows.append({
            "채널": ch, "기울기": f"{slope:.4f}", "R²": f"{r**2:.3f}",
            "p-value": "< 0.001" if p < 0.001 else f"{p:.3f}",
            "해석": ("유의함" if p < 0.05 else "유의하지 않음"),
        })
    if reg_rows:
        st.dataframe(pd.DataFrame(reg_rows).set_index("채널"), use_container_width=True)
        st.caption("R² : 유입율이 CR 변동을 얼마나 설명하는지 (1에 가까울수록 강한 관계)")

    memo_block("efficiency")

# ══════════════════════════════════════════════════════════════════════════════
# 05. 채널 비중 추이
# ══════════════════════════════════════════════════════════════════════════════
elif page == "05. 채널 비중 추이":
    st.header("📈 채널 비중 추이")

    def stacked_area(metric: str, title: str, ylabel: str):
        pivot = get_pivot(metric, CHANNELS).fillna(method="ffill").fillna(0)
        fig, ax = _fig(14, 5)
        arrays = [
            pivot[c].values * 100 if c in pivot.columns else np.zeros(len(pivot))
            for c in CHANNELS
        ]
        ax.stackplot(pivot.index, arrays, labels=CHANNELS, colors=PALETTE, alpha=0.85)
        ax.set_ylabel(ylabel)
        ax.set_ylim(0, 100)
        ax.legend(loc="upper left", fontsize=9, ncol=2)
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))
        plt.xticks(rotation=30, fontsize=8)
        sns.despine(ax=ax)
        st.subheader(title)
        _show(fig)

    stacked_area("거래액비중", "거래액 비중 추이 (누적)", "거래액 비중 (%)")
    stacked_area("고객비중",   "고객 비중 추이 (누적)",   "고객 비중 (%)")

    # 요일별 CR 히트맵
    st.subheader("요일별 × 채널별 전환율(CR) 패턴")
    dow_map = {0: "월", 1: "화", 2: "수", 3: "목", 4: "금", 5: "토", 6: "일"}
    cr_df = df[(df["metric"] == "CR") & (df["segment"].isin(CHANNELS))].copy()
    cr_df["요일"] = cr_df["dow"].map(dow_map)
    cr_hm = (
        cr_df.groupby(["요일", "segment"])["value"].mean()
        .unstack("segment")
        .reindex(index=["월", "화", "수", "목", "금", "토", "일"], columns=CHANNELS)
        .fillna(0)
    )
    fig, ax = _fig(12, 4)
    sns.heatmap(
        cr_hm.T * 100, annot=True, fmt=".2f", cmap="RdYlGn",
        linewidths=0.5, ax=ax, cbar_kws={"label": "CR (%)"},
    )
    ax.set_xlabel("")
    ax.set_ylabel("")
    _show(fig)

    # 요일별 거래액 히트맵
    st.subheader("요일별 × 채널별 일평균거래액 패턴")
    tx_df = df[(df["metric"] == "일평균거래액") & (df["segment"].isin(CHANNELS))].copy()
    tx_df["요일"] = tx_df["dow"].map(dow_map)
    tx_hm = (
        tx_df.groupby(["요일", "segment"])["value"].mean()
        .unstack("segment")
        .reindex(index=["월", "화", "수", "목", "금", "토", "일"], columns=CHANNELS)
        .fillna(0)
    )
    fig, ax = _fig(12, 4)
    sns.heatmap(
        tx_hm.T / 1e8, annot=True, fmt=".2f", cmap="YlOrRd",
        linewidths=0.5, ax=ax, cbar_kws={"label": "거래액 (억원)"},
    )
    ax.set_xlabel("")
    ax.set_ylabel("")
    _show(fig)

    memo_block("share_trend")

# ══════════════════════════════════════════════════════════════════════════════
# 06. 월별 요약
# ══════════════════════════════════════════════════════════════════════════════
elif page == "06. 월별 요약":
    st.header("📅 월별 요약")

    SUMMARY_METRICS = [
        ("일평균거래액", lambda v: fmt_num(v, "원"),
         lambda x, _: f"{x/1e8:.1f}억"),
        ("일평균고객수", lambda v: f"{v:,.0f}명",
         lambda x, _: f"{x:,.0f}"),
        ("일평균객단가", lambda v: fmt_num(v, "원"),
         lambda x, _: f"{x/1e4:.1f}만"),
        ("CR",          fmt_pct,
         lambda x, _: f"{x*100:.2f}%"),
    ]

    for metric, cell_fmt, ax_fmt in SUMMARY_METRICS:
        st.subheader(f"월별 {METRIC_KO.get(metric, metric)}")
        monthly = _monthly_avg(metric, ["*TOTAL"] + CHANNELS)
        if monthly.empty:
            st.info("데이터 없음")
            continue

        # 테이블
        def safe_fmt(v):
            try:
                return cell_fmt(v) if not pd.isnull(v) else "–"
            except Exception:
                return "–"

        st.dataframe(
            monthly.apply(lambda col: col.map(safe_fmt)),
            use_container_width=True,
        )

        # 채널별 바차트 (TOTAL 제외)
        monthly_ch = monthly.drop(columns=["*TOTAL"], errors="ignore")
        long = monthly_ch.reset_index().melt(id_vars="month", var_name="채널", value_name="value")
        long = long[long["채널"].isin(CHANNELS)].dropna(subset=["value"])
        if len(long):
            fig, ax = _fig(12, 4)
            sns.barplot(
                data=long, x="month", y="value", hue="채널",
                hue_order=CHANNELS,
                palette=PALETTE, ax=ax,
            )
            ax.yaxis.set_major_formatter(mticker.FuncFormatter(ax_fmt))
            ax.set_xlabel("")
            ax.set_ylabel(METRIC_KO.get(metric, metric))
            ax.legend(fontsize=9, ncol=3)
            plt.xticks(rotation=20, fontsize=9)
            sns.despine(ax=ax)
            _show(fig)
        st.markdown("---")

    # 데이터 다운로드
    st.subheader("📥 전체 데이터 다운로드")
    all_pivot = df.pivot_table(
        index=["metric", "segment"], columns="date", values="value", aggfunc="first",
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        all_pivot.to_excel(writer, sheet_name="전체데이터")
    st.download_button(
        "⬇️ Excel 다운로드",
        data=buf.getvalue(),
        file_name="첫구매_채널분석.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    memo_block("monthly_summary")
