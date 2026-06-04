"""첫구매 채널별 분석 대시보드"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from scipy import stats
import io, json, os

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="첫구매 채널 분석", page_icon="🛒",
                   layout="wide", initial_sidebar_state="expanded")

INSIGHT_FILE = "fp_insights.json"

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;
  border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.vg{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;
  margin:8px 0;line-height:1.65;background:#ffffff}
.vr{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;
  margin:8px 0;line-height:1.65;background:#fff5f5}
.va{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;
  margin:8px 0;line-height:1.65;background:#fffbeb}
.vb{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;
  margin:8px 0;line-height:1.65;background:#f0f9ff}
.sdiv{border-top:1px solid #e2e8f0;margin:22px 0}
.appendix{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;
  padding:14px 18px;margin-top:12px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 색상 팔레트
# ══════════════════════════════════════════════════════
_PAL = {
    "blue":   ("rgba(79,143,255,1)",   "rgba(79,143,255,0.15)"),
    "red":    ("rgba(245,101,101,1)",  "rgba(245,101,101,0.15)"),
    "green":  ("rgba(72,187,120,1)",   "rgba(72,187,120,0.15)"),
    "amber":  ("rgba(237,137,54,1)",   "rgba(237,137,54,0.15)"),
    "purple": ("rgba(159,122,234,1)",  "rgba(159,122,234,0.15)"),
    "teal":   ("rgba(56,178,172,1)",   "rgba(56,178,172,0.15)"),
    "pink":   ("rgba(236,72,153,1)",   "rgba(236,72,153,0.15)"),
    "slate":  ("rgba(51,65,85,1)",     "rgba(51,65,85,0.15)"),
    "orange": ("rgba(249,115,22,1)",   "rgba(249,115,22,0.15)"),
}
def clr(n): return _PAL[n][0]
def cbg(n): return _PAL[n][1]

CHANNEL_PAL = {
    "직접":        "blue",
    "광고":        "amber",
    "EP":          "green",
    "PUSH":        "purple",
    "제휴":        "red",
    "브랜드광고":  "teal",
    "미디어커머스": "orange",
    "*TOTAL":      "slate",
}
CHANNELS = ["직접", "광고", "EP", "PUSH", "제휴", "브랜드광고", "미디어커머스"]

def ch_clr(ch): return clr(CHANNEL_PAL.get(ch, "blue"))
def ch_cbg(ch): return cbg(CHANNEL_PAL.get(ch, "blue"))

def base_layout(h=280, ysuffix="", title=""):
    return dict(
        paper_bgcolor="rgba(248,249,252,0)", plot_bgcolor="rgba(248,249,252,0)",
        font=dict(color="#475569", size=11),
        margin=dict(l=10, r=10, t=36, b=10),
        height=h, showlegend=False,
        title=dict(text=title, font=dict(color="#c0c8d8", size=13)),
        xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11)),
        yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11), ticksuffix=ysuffix),
    )

# ══════════════════════════════════════════════════════
# 인사이트 저장소
# ══════════════════════════════════════════════════════
def load_insights():
    if os.path.exists(INSIGHT_FILE):
        try:
            with open(INSIGHT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_insights(data):
    with open(INSIGHT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if "fp_ins" not in st.session_state:
    st.session_state.fp_ins = load_insights()

# ══════════════════════════════════════════════════════
# 통계 헬퍼
# ══════════════════════════════════════════════════════
def linreg(x, y):
    mask = ~np.isnan(x.astype(float)) & ~np.isnan(y.astype(float))
    if mask.sum() < 5:
        return dict(slope=np.nan, r2=np.nan, p=np.nan)
    sl, ic, r, p, _ = stats.linregress(x[mask], y[mask])
    return dict(slope=sl, intercept=ic, r2=r**2, p=p)

def sig_label(p):
    if np.isnan(p): return "–"
    if p < 0.001: return "p<0.001 (우연일 확률 0.1% 미만 — 매우 확실)"
    if p < 0.01:  return "p<0.01 (우연일 확률 1% 미만 — 신뢰할 수 있음)"
    if p < 0.05:  return "p<0.05 (우연일 확률 5% 미만 — 유의함)"
    return "유의하지 않음 (우연일 수 있음)"

def r2_label(r2):
    if np.isnan(r2): return "–"
    if r2 >= 0.3:  return f"R²={r2:.3f} → 뚜렷한 경향 (변동의 {r2*100:.0f}% 설명)"
    if r2 >= 0.1:  return f"R²={r2:.3f} → 약한 경향 (변동의 {r2*100:.0f}% 설명)"
    return f"R²={r2:.3f} → 경향 거의 없음 ({r2*100:.0f}% 미만 설명)"

def corr_label(r):
    if np.isnan(r): return "–"
    ar = abs(r)
    sign = "반대 방향" if r < 0 else "같은 방향"
    if ar >= 0.7:   strength = "매우 강하게"
    elif ar >= 0.5: strength = "강하게"
    elif ar >= 0.3: strength = "보통으로"
    elif ar >= 0.1: strength = "약하게"
    else:           strength = "거의 관계 없이"
    return f"{strength} {sign}으로 움직임"

def pct(a, b):
    return (b - a) / a * 100 if a and not np.isnan(float(a)) else 0

# ══════════════════════════════════════════════════════
# 파싱
# ══════════════════════════════════════════════════════
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
METRICS_LIST  = list(METRIC_KO.keys())
METRIC_LABELS = list(METRIC_KO.values())

def metric_color(m):
    return {"일평균거래액":"blue","거래액비중":"teal","일평균고객수":"purple",
            "고객비중":"purple","일평균객단가":"amber","유효회원수":"green",
            "DAU":"green","유입율":"red","CR":"pink"}.get(m,"blue")

def fmt_val(metric, v):
    if pd.isnull(v): return "–"
    if metric in ("일평균거래액","일평균객단가"):
        if abs(v) >= 1e8: return f"{v/1e8:.2f}억원"
        if abs(v) >= 1e4: return f"{v/1e4:.0f}만원"
        return f"{v:,.0f}원"
    if metric in ("일평균고객수","유효회원수","DAU"):
        if abs(v) >= 1e4: return f"{v/1e4:.1f}만명"
        return f"{v:,.0f}명"
    if metric in ("거래액비중","고객비중","유입율","CR"):
        return f"{v*100:.2f}%"
    return f"{v:,.3f}"

@st.cache_data(show_spinner=False)
def parse_xlsx(file_bytes: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")

    # 연도 감지 (row 0)
    year = 2026
    for v in df_raw.iloc[0, :]:
        try:
            iv = int(v)
            if 2020 <= iv <= 2030:
                year = iv; break
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
            if pd.isnull(dt): continue
            records.append({"date": dt, "metric": current_metric, "segment": seg, "value": val})

    df = pd.DataFrame(records)
    df["date"]    = pd.to_datetime(df["date"])
    df["month"]   = df["date"].dt.to_period("M").astype(str)
    df["quarter"] = df["date"].dt.to_period("Q").astype(str)
    df["year"]    = df["date"].dt.year
    df["dow"]     = df["date"].dt.dayofweek
    return df

# ══════════════════════════════════════════════════════
# 텍스트 포맷 헬퍼
# ══════════════════════════════════════════════════════
FMT_DEFAULTS = {"bold": False, "color": "#334155", "size": 13}
SIZE_OPTS    = [11, 12, 13, 14, 15, 16, 18, 20, 24]

def _norm(val, base=None):
    d = (base or FMT_DEFAULTS).copy()
    if isinstance(val, str):
        d["text"] = val; return d
    if isinstance(val, dict):
        merged = d.copy()
        merged.update({k: v for k, v in val.items() if k != "editing"})
        merged.setdefault("text", ""); return merged
    d["text"] = ""; return d

def _style(fmt):
    w = "700" if fmt.get("bold") else "400"
    c = fmt.get("color", "#334155")
    s = fmt.get("size", 13)
    return f"font-weight:{w};color:{c};font-size:{s}px;line-height:1.7"

def _fmt_ui(key, fmt):
    fc1, fc2, fc3 = st.columns([1, 2, 2])
    bold  = fc1.checkbox("굵게", value=bool(fmt.get("bold", False)), key=f"_ck_b_{key}")
    raw_c = fmt.get("color", "#334155")
    safe_c = raw_c if (isinstance(raw_c, str) and raw_c.startswith("#") and len(raw_c) == 7) else "#334155"
    color = fc2.color_picker("색상", value=safe_c, key=f"_ck_c_{key}")
    sz  = fmt.get("size", 13)
    idx = SIZE_OPTS.index(sz) if sz in SIZE_OPTS else SIZE_OPTS.index(13)
    size = fc3.selectbox("크기(px)", SIZE_OPTS, index=idx, key=f"_ck_s_{key}")
    return bold, color, size

# ══════════════════════════════════════════════════════
# 인사이트 에디터 (동적 추가/삭제)
# ══════════════════════════════════════════════════════
def insight_editor(page_key, _default=None):
    store = st.session_state.fp_ins
    if page_key not in store:
        store[page_key] = []
    items    = store[page_key]
    hide_key = f"__ihide_{page_key}__"
    if hide_key not in st.session_state:
        st.session_state[hide_key] = False

    h1, h2, h3 = st.columns([8, 1, 1])
    h1.markdown("**메모**", unsafe_allow_html=True)
    if h2.button("숨기기" if not st.session_state[hide_key] else "펼치기",
                 key=f"ihide_{page_key}", use_container_width=True):
        st.session_state[hide_key] = not st.session_state[hide_key]; st.rerun()
    if h3.button("+ 추가", key=f"iadd_{page_key}", use_container_width=True):
        items.append({"text": "", "editing": True, "bold": False, "color": "#334155", "size": 13})
        store[page_key] = items; st.session_state.fp_ins = store; save_insights(store); st.rerun()

    if not st.session_state[hide_key]:
        if not items:
            st.markdown("<p style='color:#94a3b8;font-size:12px;margin:4px 0'>+ 추가 버튼으로 메모를 입력하세요.</p>",
                        unsafe_allow_html=True)
        to_delete = []
        for i, raw_item in enumerate(items):
            item = _norm(raw_item) if not isinstance(raw_item, dict) else raw_item
            item.setdefault("bold", False); item.setdefault("color", "#334155"); item.setdefault("size", 13)
            c1, c2, c3 = st.columns([10, 1, 1])
            if item.get("editing", False):
                new_text = c1.text_area("", item.get("text", ""), key=f"ita_{page_key}_{i}",
                                         height=72, label_visibility="collapsed", placeholder="내용을 입력하세요.")
                bold, color, size = _fmt_ui(f"i_{page_key}_{i}", item)
                if c2.button("저장", key=f"isave_{page_key}_{i}", use_container_width=True):
                    items[i] = {"text": new_text, "editing": False, "bold": bold, "color": color, "size": size}
                    store[page_key] = items; st.session_state.fp_ins = store; save_insights(store); st.rerun()
                if c3.button("삭제", key=f"idel_{page_key}_{i}", use_container_width=True):
                    to_delete.append(i)
            else:
                c1.markdown(f"<p style='{_style(item)};margin:2px 0'>{item.get('text','')}</p>",
                            unsafe_allow_html=True)
                if c2.button("편집", key=f"iedit_{page_key}_{i}", use_container_width=True):
                    items[i]["editing"] = True
                    store[page_key] = items; st.session_state.fp_ins = store; save_insights(store); st.rerun()
                if c3.button("삭제", key=f"idel2_{page_key}_{i}", use_container_width=True):
                    to_delete.append(i)
        for idx in sorted(to_delete, reverse=True):
            items.pop(idx)
        if to_delete:
            store[page_key] = items; st.session_state.fp_ins = store; save_insights(store); st.rerun()

    store[page_key] = items; st.session_state.fp_ins = store; save_insights(store)

# ══════════════════════════════════════════════════════
# 편집 가능 텍스트
# ══════════════════════════════════════════════════════
if "fp_texts" not in st.session_state:
    st.session_state.fp_texts = load_insights().get("__fp_texts__", {})

TAG_FMT = {
    "h2": {"bold": True,  "color": "#1e293b", "size": 24},
    "h3": {"bold": True,  "color": "#1e293b", "size": 20},
    "p":  {"bold": False, "color": "#334155", "size": 13},
}

def editable_text(key, default, tag="p", style=""):
    base  = TAG_FMT.get(tag, FMT_DEFAULTS).copy()
    texts = st.session_state.fp_texts
    if key not in texts:
        texts[key] = {"text": default, **base}
    item = _norm(texts[key], base)
    ekey = f"__fpedit_{key}__"
    if ekey not in st.session_state:
        st.session_state[ekey] = False
    if st.session_state[ekey]:
        col1, col2 = st.columns([10, 1])
        new_val = col1.text_area("", item["text"], key=f"fpeta_{key}", height=80, label_visibility="collapsed")
        bold, color, size = _fmt_ui(key, item)
        if col2.button("확인", key=f"fpesave_{key}"):
            texts[key] = {"text": new_val, "bold": bold, "color": color, "size": size}
            st.session_state[ekey] = False
            all_data = load_insights(); all_data["__fp_texts__"] = texts
            save_insights(all_data); st.session_state.fp_texts = texts; st.rerun()
    else:
        item = _norm(texts[key], base)
        st.markdown(f'<{tag} style="{_style(item)};cursor:pointer;{style}">{item["text"]}</{tag}>',
                    unsafe_allow_html=True)
        if st.button("편집", key=f"fpeedit_{key}", help="편집"):
            st.session_state[ekey] = True; st.rerun()

def verdict_box(key, default_text, box_cls="vg"):
    store = st.session_state.fp_ins
    vkey  = f"__fpv_{key}__"
    ekey  = f"__fpvedit_{key}__"
    if vkey not in store:
        store[vkey] = {"text": default_text, "bold": False, "color": "#334155", "size": 13}
    item = _norm(store[vkey])
    if ekey not in st.session_state:
        st.session_state[ekey] = False
    if st.session_state[ekey]:
        c1, c2 = st.columns([11, 1])
        new_text = c1.text_area("", item["text"], key=f"fpvta_{key}", height=100, label_visibility="collapsed")
        bold, color, size = _fmt_ui(f"fpv_{key}", item)
        if c2.button("확인", key=f"fpvsave_{key}"):
            store[vkey] = {"text": new_text, "bold": bold, "color": color, "size": size}
            st.session_state[ekey] = False
            all_data = load_insights(); all_data.update(store)
            save_insights(all_data); st.session_state.fp_ins = store; st.rerun()
    else:
        item = _norm(store[vkey])
        txt = item["text"].replace("\n", "<br>")
        col_a, col_b = st.columns([20, 1])
        col_a.markdown(f'<div class="{box_cls}"><span style="{_style(item)}">{txt}</span></div>',
                       unsafe_allow_html=True)
        if col_b.button("편집", key=f"fpvedit_{key}"):
            st.session_state[ekey] = True; st.rerun()

# ══════════════════════════════════════════════════════
# 메모 블록 (기본 텍스트 제공)
# ══════════════════════════════════════════════════════
FP_DEFAULTS = {
    "top_overview":
        "채널별 거래액·고객수 분포를 파악합니다. 직접 채널이 거래액의 가장 큰 비중을 차지하며, "
        "채널별 객단가 차이에서 고객 품질의 차이를 확인할 수 있습니다.",
    "top_tx":
        "채널별 일평균거래액 추이와 월별 변화를 분석합니다. "
        "특정 채널의 거래액이 급등/급락하는 시점을 파악하고 원인을 탐색하세요.",
    "top_customer":
        "일평균고객수·DAU·유효회원수 추이를 채널별로 비교합니다. "
        "고객수 대비 DAU 비율이 높을수록 재방문이 활발한 채널입니다.",
    "top_efficiency":
        "CR이 높은 채널은 구매 의향이 높은 고객을 유입시킵니다. "
        "유입율이 높지만 CR이 낮은 채널은 브라우징 위주일 가능성이 있습니다.",
    "top_share":
        "채널별 비중의 시계열 변화로 포트폴리오 변화를 파악합니다. "
        "특정 채널 의존도가 높아진다면 리스크 분산 관점에서 검토가 필요합니다.",
    "top_monthly":
        "월별 집계로 계절성과 트렌드를 파악합니다. "
        "같은 월끼리 전년비(YoY) 비교를 활용하면 더 정확한 추세를 볼 수 있습니다.",
    "btm_overview":  "KPI는 필터 기간의 일평균입니다. 기간을 바꾸면 수치가 달라집니다.",
    "btm_tx":        "이동평균 일수를 늘리면 노이즈가 줄고 추세가 선명해집니다.",
    "btm_customer":  "고객수와 DAU가 함께 증가하는 채널은 신규·재방문 모두 건강한 상태입니다.",
    "btm_efficiency":"CR 추세선이 하락 중이라면 랜딩 품질 또는 상품 관련성을 점검하세요.",
    "btm_share":     "누적 영역 차트에서 특정 채널 면적이 줄어드는 것은 상대적 점유율 하락입니다.",
    "btm_monthly":   "기간 비교에서 증감율이 붉게 표시된 지표를 중심으로 원인을 분석하세요.",
    "btm_scatter":   "산점도의 점이 추세선 주변에 모일수록 R²가 높아 관계가 강합니다.",
    "btm_corr":      "절대값 0.5 이상인 상관계수를 중심으로 관계를 해석하세요.",
    "btm_dow":       "요일별 편차가 크다면 발송·프로모션 스케줄 조정이 효과적인 개선 방법이 될 수 있습니다.",
}

def memo_block(key):
    store = st.session_state.fp_ins
    mkey  = f"__fpm_{key}__"
    ekey  = f"__fpme_{key}__"
    hkey  = f"__fpmh_{key}__"
    default_text = FP_DEFAULTS.get(key, "")
    base = {"bold": False, "color": "#475569", "size": 13}
    if mkey not in store:
        store[mkey] = {"text": default_text, **base}
    item = _norm(store[mkey], base)
    if ekey not in st.session_state: st.session_state[ekey] = False
    if hkey not in st.session_state: st.session_state[hkey] = False
    if st.session_state[hkey]:
        if st.button("펼치기", key=f"fpmshow_{key}", help="메모 표시"):
            st.session_state[hkey] = False; st.rerun()
        return
    if st.session_state[ekey]:
        c1, c2, c3 = st.columns([10, 1, 1])
        new_val = c1.text_area("", item["text"], key=f"fpmta_{key}",
                                height=80, label_visibility="collapsed")
        bold, color, size = _fmt_ui(f"fpm_{key}", item)
        if c2.button("저장", key=f"fpmsave_{key}", use_container_width=True):
            store[mkey] = {"text": new_val, "bold": bold, "color": color, "size": size}
            st.session_state[ekey] = False
            all_data = load_insights(); all_data.update(store)
            save_insights(all_data); st.session_state.fp_ins = store; st.rerun()
        if c3.button("숨기기", key=f"fpmhide_{key}", use_container_width=True):
            st.session_state[hkey] = True; st.session_state[ekey] = False; st.rerun()
    else:
        item = _norm(store[mkey], base)
        txt = item["text"]
        c1, c2 = st.columns([20, 1])
        if txt:
            c1.markdown(f'<p style="{_style(item)};margin:6px 0">{txt}</p>', unsafe_allow_html=True)
        else:
            c1.markdown('<p style="font-size:12px;color:#cbd5e1;margin:4px 0">메모를 입력하세요.</p>',
                        unsafe_allow_html=True)
        if c2.button("편집", key=f"fpmedit_{key}", use_container_width=True):
            st.session_state[ekey] = True; st.rerun()
    store[mkey] = store.get(mkey, {"text": "", **base})
    st.session_state.fp_ins = store

# ══════════════════════════════════════════════════════
# 부속 도구
# ══════════════════════════════════════════════════════
def show_appendix(df_raw, label="근거 데이터"):
    with st.expander(f"근거 데이터 — {label}", expanded=False):
        st.markdown("<div class='appendix'>", unsafe_allow_html=True)
        st.dataframe(df_raw.reset_index(drop=True), use_container_width=True)
        csv = df_raw.to_csv(index=False).encode("utf-8-sig")
        st.download_button("CSV 다운로드", csv, file_name=f"{label}.csv", mime="text/csv")
        st.markdown("</div>", unsafe_allow_html=True)

def stat_explainer():
    with st.expander("통계 용어 해설 (클릭하여 펼치기)", expanded=False):
        verdict_box("fp_stat_r2",
            "R² (결정계수) — \"이 지표가 얼마나 일관된 패턴을 보이는가\"\n"
            "0~1 사이. 0.3 이상=뚜렷한 경향, 0.1~0.3=약한 경향, 0.1 미만=패턴 거의 없음.\n"
            "예: R²=0.28 → 시간 흐름에 따른 변화 패턴이 28% 설명됨", "vb")
        verdict_box("fp_stat_p",
            "p값 (유의확률) — \"이 패턴이 우연일 확률\"\n"
            "p<0.001 → 우연일 확률 0.1% 미만 = 거의 확실한 경향\n"
            "p<0.05  → 우연일 확률 5% 미만 = 통계적으로 유의함\n"
            "p≥0.05  → 우연일 수 있음 = 근거 부족", "vg")
        verdict_box("fp_stat_corr",
            "상관계수 — \"두 지표가 얼마나 함께 움직이는가\"\n"
            "-1~+1 사이. 음수=한쪽 오를 때 다른 쪽 내려감, 양수=같이 움직임.\n"
            "절대값 0.5 이상이면 강한 관계", "va")

# ══════════════════════════════════════════════════════
# 기간 비교 헬퍼
# ══════════════════════════════════════════════════════
COMPARE_OPTS = ["전월비 (MoM)", "전분기비 (QoQ)", "전년비 (YoY)", "전체 기간 처음↔끝"]

def get_compare_periods(df_src, mode):
    df_src = df_src.sort_values("date")
    last_d = df_src["date"].max()
    if mode == "전년비 (YoY)":
        cur  = df_src[df_src["date"].dt.year == last_d.year]
        prev = df_src[df_src["date"].dt.year == last_d.year - 1]
        return cur, prev, f"{last_d.year}년", f"{last_d.year-1}년"
    elif mode == "전분기비 (QoQ)":
        cq = pd.Period(last_d, "Q"); pq = cq - 1
        cur  = df_src[df_src["date"].dt.to_period("Q") == cq]
        prev = df_src[df_src["date"].dt.to_period("Q") == pq]
        return cur, prev, str(cq), str(pq)
    elif mode == "전월비 (MoM)":
        cm = pd.Period(last_d, "M"); pm = cm - 1
        cur  = df_src[df_src["date"].dt.to_period("M") == cm]
        prev = df_src[df_src["date"].dt.to_period("M") == pm]
        return cur, prev, str(cm), str(pm)
    else:
        mid  = df_src["date"].median()
        cur  = df_src[df_src["date"] >= mid]
        prev = df_src[df_src["date"] <  mid]
        return cur, prev, "후반기", "전반기"

def compare_table(metric, df_cur, df_prev, lbl_cur, lbl_prev, segs=None):
    segs = segs or (sel_channels + ["*TOTAL"])
    rows = []
    for seg in segs:
        vc = df_cur[(df_cur["metric"]==metric)&(df_cur["segment"]==seg)]["value"].mean()
        vp = df_prev[(df_prev["metric"]==metric)&(df_prev["segment"]==seg)]["value"].mean()
        chg = pct(vp, vc) if not (pd.isnull(vp) or vp == 0) else np.nan
        rows.append({"채널": seg, lbl_prev: fmt_val(metric, vp),
                     lbl_cur: fmt_val(metric, vc),
                     "증감율": f"{chg:+.1f}%" if not np.isnan(chg) else "–",
                     "_chg": chg})
    df_out = pd.DataFrame(rows)
    return df_out

def styled_compare(df_tbl):
    positive_good = {"일평균거래액","일평균고객수","일평균객단가","CR","유입율","DAU","유효회원수","거래액비중","고객비중"}
    def color_row(row):
        chg = row.get("_chg", np.nan)
        if pd.isnull(chg): return [""] * len(row)
        good = chg > 0
        color = "color:#16a34a;font-weight:600" if good else "color:#dc2626;font-weight:600"
        result = [""] * len(row)
        idx = list(row.index).index("증감율") if "증감율" in row.index else -1
        if idx >= 0: result[idx] = color
        return result
    display = df_tbl.drop(columns=["_chg"], errors="ignore")
    try:
        return display.style.apply(color_row, axis=1)
    except Exception:
        return display

# ══════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🛒 첫구매 분석")
    uploaded = st.file_uploader("엑셀 업로드", type=["xlsx", "xls"], label_visibility="collapsed")

    if uploaded:
        file_bytes = uploaded.read()
        df_all = parse_xlsx(file_bytes)
        date_min = df_all["date"].min()
        date_max = df_all["date"].max()
        st.success(f"✅ {df_all['date'].nunique()}일 데이터")
        st.caption(f"{date_min.date()} ~ {date_max.date()}")

        st.markdown("---")
        st.markdown("**📅 기간 필터**")
        d1 = st.date_input("시작일", date_min.date(), min_value=date_min.date(), max_value=date_max.date())
        d2 = st.date_input("종료일", date_max.date(), min_value=date_min.date(), max_value=date_max.date())

        st.markdown("**📆 요일 필터**")
        DOW_KR = ["월","화","수","목","금","토","일"]
        sel_dow = st.multiselect("요일 선택", DOW_KR, default=DOW_KR)
        dow_idx = [DOW_KR.index(d) for d in sel_dow]

        st.markdown("**📡 채널 필터**")
        avail_ch = [c for c in CHANNELS if c in df_all["segment"].unique()]
        sel_channels = st.multiselect("채널 선택", avail_ch, default=avail_ch)
        if not sel_channels:
            sel_channels = avail_ch

        # 필터 적용
        date_mask = (df_all["date"] >= pd.to_datetime(d1)) & (df_all["date"] <= pd.to_datetime(d2))
        dow_mask  = df_all["dow"].isin(dow_idx)
        df_full   = df_all[date_mask & dow_mask].copy()          # 채널 필터 없음 (TOTAL 포함)
        df_f      = df_full[df_full["segment"].isin(sel_channels + ["*TOTAL"])].copy()

        st.markdown("---")
        st.caption(f"필터 후: {df_f['date'].nunique()}일")
    else:
        df_all = None; df_f = None; df_full = None
        sel_channels = CHANNELS

    st.markdown("---")
    PAGE_LIST = [
        "01. 개요",
        "02. 거래액 분석",
        "03. 고객수 분석",
        "04. 채널 효율",
        "05. 채널 비중 추이",
        "06. 요일별 패턴",
        "07. 월별 요약",
    ]
    page = st.radio("분석 주제", PAGE_LIST, label_visibility="collapsed")

    if df_all is not None:
        st.markdown("---")
        if st.button("메모 초기화", use_container_width=True):
            st.session_state.fp_ins = {}; save_insights({}); st.rerun()
        lines = ["# 첫구매 채널 분석 인사이트"]
        for pk, items in st.session_state.fp_ins.items():
            if pk.startswith("__") or not isinstance(items, list): continue
            lines.append(f"\n## {pk}")
            for it in items:
                txt = it.get("text", str(it)) if isinstance(it, dict) else str(it)
                lines.append(f"- {txt}")
        st.download_button("메모 내보내기", "\n".join(lines),
                           file_name="fp_insights.txt", mime="text/plain",
                           use_container_width=True)

# ══════════════════════════════════════════════════════
# 업로드 전 안내
# ══════════════════════════════════════════════════════
if df_all is None:
    st.title("🛒 첫구매 채널 분석 대시보드")
    st.markdown("왼쪽에서 **첫구매 채널별 지표 엑셀 파일**을 업로드하면 분석이 시작됩니다.")
    c1, c2, c3, c4 = st.columns(4)
    for cw, title, items in [
        (c1, "거래액 지표",  ["일평균거래액", "거래액 비중", "일평균객단가"]),
        (c2, "고객 지표",    ["일평균고객수", "유효회원수", "DAU", "고객 비중"]),
        (c3, "효율 지표",    ["전환율(CR)", "유입율"]),
        (c4, "채널 구분",    ["직접", "광고", "EP", "PUSH", "제휴", "브랜드광고", "미디어커머스"]),
    ]:
        with cw:
            st.info(f"**{title}**\n\n" + "\n\n".join(f"• {i}" for i in items))
    st.stop()

# ══════════════════════════════════════════════════════
# 공통 헬퍼 (필터 적용 후)
# ══════════════════════════════════════════════════════
DOW_MAP = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}

def get_series(metric: str, seg: str, df_src=None) -> pd.Series:
    src = df_src if df_src is not None else df_f
    sub = src[(src["metric"] == metric) & (src["segment"] == seg)].sort_values("date")
    return sub.set_index("date")["value"]

def get_monthly(metric: str, segs: list, df_src=None) -> pd.DataFrame:
    src = df_src if df_src is not None else df_f
    avail = [s for s in segs if s in src["segment"].unique()]
    return (
        src[(src["metric"] == metric) & (src["segment"].isin(avail))]
        .groupby(["month", "segment"])["value"].mean()
        .unstack("segment")
        .reindex(columns=[s for s in segs if s in avail])
    )

def legend_layout(layout: dict) -> dict:
    layout["showlegend"] = True
    layout["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)",
                             font=dict(color="#64748b", size=10))
    return layout

meta = dict(start=str(df_all["date"].min().date()),
            end=str(df_all["date"].max().date()),
            days=df_all["date"].nunique())

# ══════════════════════════════════════════════════════
# 01. 개요
# ══════════════════════════════════════════════════════
if page == "01. 개요":
    editable_text("fp_h_overview", "첫구매 채널 분석 — 개요", "h2",
                  "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"전체 데이터: {meta['start']} ~ {meta['end']} · {meta['days']}일 "
               f"| 필터 적용: {df_f['date'].nunique()}일")
    memo_block("top_overview")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # KPI 카드
    kpi_specs = [
        ("일평균거래액","일평균거래액"),("일평균고객수","일평균고객수"),
        ("일평균객단가","일평균객단가"),("CR","전환율(CR)"),("유입율","유입율"),
    ]
    cols = st.columns(5)
    for col, (metric, label) in zip(cols, kpi_specs):
        cur_s = df_f[(df_f["metric"]==metric)&(df_f["segment"]=="*TOTAL")]["value"]
        all_s = df_all[(df_all["metric"]==metric)&(df_all["segment"]=="*TOTAL")]["value"]
        v_now = cur_s.mean() if len(cur_s) else np.nan
        v_all = all_s.mean() if len(all_s) else np.nan
        delta = pct(v_all, v_now) if not pd.isnull(v_all) and v_all != 0 else None
        col.metric(label, fmt_val(metric, v_now),
                   f"{delta:+.1f}% vs 전체 평균" if delta is not None else None)
    memo_block("btm_overview")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 채널 비중 + 전체 거래액 추이
    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("채널별 거래액 비중 (필터 기간 평균)")
        share = (
            df_f[(df_f["metric"]=="거래액비중")&(df_f["segment"].isin(sel_channels))]
            .groupby("segment")["value"].mean()
            .reindex([c for c in sel_channels if c in df_f["segment"].unique()])
            .dropna()
        )
        fig = go.Figure(go.Bar(
            x=(share.values * 100).tolist(), y=share.index.tolist(),
            orientation="h",
            marker_color=[ch_cbg(c) for c in share.index],
            marker_line_color=[ch_clr(c) for c in share.index], marker_line_width=1.5,
            text=[f"{v*100:.1f}%" for v in share.values], textposition="outside",
            textfont=dict(color="#64748b", size=11),
        ))
        ly = base_layout(300, "%", "거래액 비중 (%)")
        ly["xaxis"]["ticksuffix"] = "%"
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    with col_r:
        st.subheader("일평균거래액 전체 추이 (7일 이동평균)")
        s_total = get_series("일평균거래액", "*TOTAL", df_full).rolling(7).mean().dropna()
        fig2 = go.Figure(go.Scatter(
            x=s_total.index.strftime("%m/%d").tolist(),
            y=(s_total / 1e8).tolist(),
            mode="lines", line=dict(color=clr("blue"), width=2),
            fill="tozeroy", fillcolor=cbg("blue"),
        ))
        ly2 = base_layout(300, "억원", "일평균거래액 (억원, 7일 MA)")
        ly2["xaxis"]["nticks"] = 15; ly2["xaxis"]["tickangle"] = -30
        fig2.update_layout(**ly2)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 채널별 요약 테이블
    st.subheader("채널별 주요 지표 요약 (필터 기간 평균)")
    rows = []
    for ch in sel_channels:
        row = {"채널": ch}
        for m in ["일평균거래액","일평균고객수","일평균객단가","CR","유입율"]:
            s = df_f[(df_f["metric"]==m)&(df_f["segment"]==ch)]["value"]
            row[METRIC_KO.get(m, m)] = fmt_val(m, s.mean()) if len(s) else "–"
        rows.append(row)
    st.dataframe(pd.DataFrame(rows).set_index("채널"), use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 지표 추이 선택
    st.subheader("지표 추이")
    sa, sb = st.columns([2, 1])
    sel_m_lbl = sa.selectbox("지표", METRIC_LABELS, key="ov_m")
    agg_ov    = sb.selectbox("집계 단위", ["월별","분기별"], key="ov_agg")
    sel_m     = METRICS_LIST[METRIC_LABELS.index(sel_m_lbl)]
    grp_col   = "month" if agg_ov == "월별" else "quarter"
    total_agg = (
        df_full[(df_full["metric"]==sel_m)&(df_full["segment"]=="*TOTAL")]
        .groupby(grp_col)["value"].mean().reset_index()
    )
    if len(total_agg):
        cn = metric_color(sel_m)
        fig3 = go.Figure(go.Scatter(
            x=total_agg[grp_col].astype(str).tolist(), y=total_agg["value"].tolist(),
            mode="lines+markers", line=dict(color=clr(cn), width=2),
            marker=dict(size=5, color=clr(cn)),
            fill="tozeroy", fillcolor=cbg(cn),
        ))
        ly3 = base_layout(260, title=f"{sel_m_lbl} 추이 ({agg_ov})")
        ly3["xaxis"]["tickangle"] = -30
        fig3.update_layout(**ly3)
        st.plotly_chart(fig3, use_container_width=True)
        show_appendix(total_agg.rename(columns={grp_col:"기간","value":sel_m_lbl}), f"개요_{sel_m_lbl}")

    # 기간 비교
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간 비교 — 전체 지표")
    cmp_ov = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_ov")
    dc, dp, lc, lp = get_compare_periods(df_full, cmp_ov)
    if len(dc) and len(dp):
        for m in ["일평균거래액","일평균고객수","일평균객단가","CR","유입율"]:
            tbl = compare_table(m, dc, dp, lc, lp)
            st.caption(f"**{METRIC_KO.get(m,m)}**")
            st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)

    insight_editor("fp_overview")

# ══════════════════════════════════════════════════════
# 02. 거래액 분석
# ══════════════════════════════════════════════════════
elif page == "02. 거래액 분석":
    editable_text("fp_h_tx", "거래액 분석", "h2",
                  "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"필터 기간: {df_f['date'].min().date()} ~ {df_f['date'].max().date()}")
    memo_block("top_tx")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    verdict_box("fp_tx_v1",
        "채널별 거래액 추이를 비교합니다. 직접 채널이 비중이 크지만, "
        "광고·EP 채널의 추이 변화가 전체 거래액에 미치는 영향도 함께 살펴보세요.", "vg")

    sel_tx_m  = st.selectbox("거래액 지표", ["일평균거래액","거래액비중","일평균객단가"], key="tx_m")
    roll_tx   = st.slider("이동평균 (일)", 1, 30, 7, key="tx_roll")

    st.subheader(f"채널별 {METRIC_KO.get(sel_tx_m, sel_tx_m)} 추이")
    fig = go.Figure()
    for ch in sel_channels:
        s = get_series(sel_tx_m, ch)
        if len(s) < 2: continue
        sm = s.rolling(roll_tx).mean().dropna()
        if sel_tx_m == "일평균거래액":
            yv = (sm / 1e8).tolist(); ysuf = "억원"
        elif sel_tx_m == "거래액비중":
            yv = (sm * 100).tolist(); ysuf = "%"
        else:
            yv = (sm / 1e4).tolist(); ysuf = "만원"
        fig.add_trace(go.Scatter(
            x=sm.index.strftime("%m/%d").tolist(), y=yv,
            mode="lines", name=ch, line=dict(color=ch_clr(ch), width=1.8),
        ))
    ly = base_layout(320, ysuf, f"채널별 {METRIC_KO.get(sel_tx_m,sel_tx_m)} ({roll_tx}일 MA)")
    legend_layout(ly); ly["xaxis"]["nticks"] = 15; ly["xaxis"]["tickangle"] = -30
    fig.update_layout(**ly)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 월별 grouped bar
    st.subheader(f"월별 채널별 {METRIC_KO.get(sel_tx_m, sel_tx_m)}")
    monthly_tx = get_monthly(sel_tx_m, sel_channels)
    if len(monthly_tx):
        fig2 = go.Figure()
        for ch in [c for c in sel_channels if c in monthly_tx.columns]:
            yraw = monthly_tx[ch].tolist()
            if sel_tx_m == "일평균거래액":
                yv2 = [v/1e8 if not pd.isnull(v) else None for v in yraw]
            elif sel_tx_m == "거래액비중":
                yv2 = [v*100 if not pd.isnull(v) else None for v in yraw]
            else:
                yv2 = [v/1e4 if not pd.isnull(v) else None for v in yraw]
            fig2.add_trace(go.Bar(
                x=monthly_tx.index.tolist(), y=yv2, name=ch,
                marker_color=ch_cbg(ch), marker_line_color=ch_clr(ch), marker_line_width=1,
            ))
        ly2 = base_layout(300, title=f"월별 채널별 {METRIC_KO.get(sel_tx_m,sel_tx_m)}")
        ly2["barmode"] = "group"; legend_layout(ly2)
        ly2["xaxis"]["tickangle"] = -30
        fig2.update_layout(**ly2)
        st.plotly_chart(fig2, use_container_width=True)
        show_appendix(monthly_tx.reset_index().rename(columns={"month":"월"}), f"월별_{sel_tx_m}")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 거래액비중 히트맵
    st.subheader("월별 채널 거래액 비중 히트맵")
    hm = get_monthly("거래액비중", sel_channels).fillna(0)
    if len(hm):
        fig3 = go.Figure(go.Heatmap(
            z=(hm.T * 100).values.tolist(),
            x=hm.index.tolist(), y=hm.columns.tolist(),
            colorscale=[[0,"rgba(248,249,252,1)"],[1,"rgba(79,143,255,0.9)"]],
            text=(hm.T * 100).round(1).values.tolist(),
            texttemplate="%{text:.1f}%",
            showscale=True, colorbar=dict(title="비중%"),
        ))
        ly3 = base_layout(280, title="월별 채널 거래액 비중 (%)")
        ly3["xaxis"]["tickangle"] = -30
        fig3.update_layout(**ly3)
        st.plotly_chart(fig3, use_container_width=True)
        show_appendix(hm.reset_index().rename(columns={"month":"월"}), "거래액비중히트맵")

    memo_block("btm_tx")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 기간 비교
    st.subheader("기간 비교 — 거래액")
    cmp_tx = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_tx")
    dc, dp, lc, lp = get_compare_periods(df_full, cmp_tx)
    if len(dc) and len(dp):
        for m in ["일평균거래액","거래액비중","일평균객단가"]:
            tbl = compare_table(m, dc, dp, lc, lp)
            st.caption(f"**{METRIC_KO.get(m,m)}**")
            st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)
            show_appendix(tbl.drop(columns=["_chg"], errors="ignore"), f"거래액비교_{m}")

    insight_editor("fp_tx")

# ══════════════════════════════════════════════════════
# 03. 고객수 분석
# ══════════════════════════════════════════════════════
elif page == "03. 고객수 분석":
    editable_text("fp_h_cust", "고객수 분석", "h2",
                  "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"필터 기간: {df_f['date'].min().date()} ~ {df_f['date'].max().date()}")
    memo_block("top_customer")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    cust_metrics = ["일평균고객수","DAU","유효회원수","고객비중"]
    sel_cm  = st.selectbox("고객 지표", cust_metrics, key="cust_m")
    roll_c  = st.slider("이동평균 (일)", 1, 30, 7, key="cust_roll")

    st.subheader(f"채널별 {METRIC_KO.get(sel_cm, sel_cm)} 추이")
    fig = go.Figure()
    for ch in sel_channels:
        s = get_series(sel_cm, ch)
        if len(s) < 2: continue
        sm = s.rolling(roll_c).mean().dropna()
        yv = (sm * 100).tolist() if sel_cm == "고객비중" else sm.tolist()
        fig.add_trace(go.Scatter(
            x=sm.index.strftime("%m/%d").tolist(), y=yv,
            mode="lines", name=ch, line=dict(color=ch_clr(ch), width=1.8),
        ))
    suf = "%" if sel_cm == "고객비중" else "명"
    ly = base_layout(300, suf, f"채널별 {METRIC_KO.get(sel_cm,sel_cm)} ({roll_c}일 MA)")
    legend_layout(ly); ly["xaxis"]["nticks"] = 15; ly["xaxis"]["tickangle"] = -30
    fig.update_layout(**ly)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader(f"월별 채널별 {METRIC_KO.get(sel_cm, sel_cm)}")
        mc = get_monthly(sel_cm, sel_channels)
        if len(mc):
            fig2 = go.Figure()
            for ch in [c for c in sel_channels if c in mc.columns]:
                yv2 = [(v*100 if sel_cm=="고객비중" else v) if not pd.isnull(v) else None
                       for v in mc[ch].tolist()]
                fig2.add_trace(go.Bar(
                    x=mc.index.tolist(), y=yv2, name=ch,
                    marker_color=ch_cbg(ch), marker_line_color=ch_clr(ch), marker_line_width=1,
                ))
            ly2 = base_layout(280, title=f"월별 {METRIC_KO.get(sel_cm,sel_cm)}")
            ly2["barmode"] = "group"; legend_layout(ly2); ly2["xaxis"]["tickangle"] = -30
            fig2.update_layout(**ly2)
            st.plotly_chart(fig2, use_container_width=True)

    with col_r:
        st.subheader("고객 비중 히트맵 (월별)")
        hm_c = get_monthly("고객비중", sel_channels).fillna(0)
        if len(hm_c):
            fig3 = go.Figure(go.Heatmap(
                z=(hm_c.T * 100).values.tolist(),
                x=hm_c.index.tolist(), y=hm_c.columns.tolist(),
                colorscale=[[0,"rgba(248,249,252,1)"],[1,"rgba(159,122,234,0.9)"]],
                text=(hm_c.T * 100).round(1).values.tolist(),
                texttemplate="%{text:.1f}%",
                showscale=True, colorbar=dict(title="비중%"),
            ))
            ly3 = base_layout(280, title="월별 고객 비중 (%)")
            ly3["xaxis"]["tickangle"] = -30
            fig3.update_layout(**ly3)
            st.plotly_chart(fig3, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 고객수 vs 거래액 산점도
    st.subheader("채널별 고객수 vs 거래액 (월별 평균)")
    bub = []
    for ch in sel_channels:
        cm = df_f[(df_f["metric"]=="일평균고객수")&(df_f["segment"]==ch)].groupby("month")["value"].mean()
        tm = df_f[(df_f["metric"]=="일평균거래액")&(df_f["segment"]==ch)].groupby("month")["value"].mean()
        for mo in cm.index.intersection(tm.index):
            bub.append({"채널":ch,"고객수":cm[mo],"거래액억":tm[mo]/1e8,"월":mo})
    bdf = pd.DataFrame(bub)
    if len(bdf):
        fig4 = go.Figure()
        for ch in sel_channels:
            sub = bdf[bdf["채널"]==ch]
            if not len(sub): continue
            fig4.add_trace(go.Scatter(
                x=sub["고객수"].tolist(), y=sub["거래액억"].tolist(),
                mode="markers+text", name=ch,
                marker=dict(color=ch_clr(ch), size=10, opacity=0.85),
                text=sub["월"].tolist(), textposition="top center",
                textfont=dict(size=9, color="#94a3b8"),
            ))
        ly4 = base_layout(340, title="채널별 고객수 vs 거래액 (월별)")
        ly4["xaxis"]["title"] = "일평균 고객수 (명)"
        ly4["yaxis"]["title"] = "일평균 거래액 (억원)"
        legend_layout(ly4)
        fig4.update_layout(**ly4)
        st.plotly_chart(fig4, use_container_width=True)
        show_appendix(bdf, "고객수_거래액_산점도")

    memo_block("btm_customer")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    st.subheader("기간 비교 — 고객 지표")
    cmp_c = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_c")
    dc, dp, lc, lp = get_compare_periods(df_full, cmp_c)
    if len(dc) and len(dp):
        for m in ["일평균고객수","DAU","유효회원수"]:
            tbl = compare_table(m, dc, dp, lc, lp)
            st.caption(f"**{METRIC_KO.get(m,m)}**")
            st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)

    insight_editor("fp_customer")

# ══════════════════════════════════════════════════════
# 04. 채널 효율
# ══════════════════════════════════════════════════════
elif page == "04. 채널 효율":
    editable_text("fp_h_eff", "채널 효율 분석", "h2",
                  "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"필터 기간: {df_f['date'].min().date()} ~ {df_f['date'].max().date()}")
    memo_block("top_efficiency")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    verdict_box("fp_eff_v1",
        "CR이 높은 채널은 구매 의향이 높은 고객을 유입시킵니다. "
        "유입율이 높지만 CR이 낮은 채널은 브라우징 위주이거나 "
        "랜딩 페이지 최적화가 필요한 경우일 수 있습니다.", "vg")

    roll_e = st.slider("이동평균 (일)", 1, 30, 7, key="eff_roll")

    # CR 추이
    st.subheader("전환율(CR) 추이")
    fig = go.Figure()
    for ch in sel_channels:
        s = get_series("CR", ch)
        if len(s) < 2: continue
        sm = s.rolling(roll_e).mean().dropna()
        fig.add_trace(go.Scatter(
            x=sm.index.strftime("%m/%d").tolist(), y=(sm * 100).tolist(),
            mode="lines", name=ch, line=dict(color=ch_clr(ch), width=1.8),
        ))
    ly = base_layout(300, "%", f"채널별 CR (%, {roll_e}일 MA)")
    legend_layout(ly); ly["xaxis"]["nticks"] = 15; ly["xaxis"]["tickangle"] = -30
    fig.update_layout(**ly)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("채널별 평균 CR")
        avg_cr = (
            df_f[(df_f["metric"]=="CR")&(df_f["segment"].isin(sel_channels))]
            .groupby("segment")["value"].mean()
            .reindex([c for c in sel_channels if c in df_f["segment"].unique()]).dropna()
        )
        fig2 = go.Figure(go.Bar(
            x=avg_cr.index.tolist(), y=(avg_cr.values * 100).tolist(),
            marker_color=[ch_cbg(c) for c in avg_cr.index],
            marker_line_color=[ch_clr(c) for c in avg_cr.index], marker_line_width=1.5,
            text=[f"{v*100:.2f}%" for v in avg_cr.values],
            textposition="outside", textfont=dict(color="#64748b", size=11),
        ))
        ly2 = base_layout(280, "%", "채널별 평균 CR")
        fig2.update_layout(**ly2)
        st.plotly_chart(fig2, use_container_width=True)

    with col2:
        st.subheader("채널별 평균 유입율")
        avg_ir = (
            df_f[(df_f["metric"]=="유입율")&(df_f["segment"].isin(sel_channels))]
            .groupby("segment")["value"].mean()
            .reindex([c for c in sel_channels if c in df_f["segment"].unique()]).dropna()
        )
        fig3 = go.Figure(go.Bar(
            x=avg_ir.index.tolist(), y=(avg_ir.values * 100).tolist(),
            marker_color=[ch_cbg(c) for c in avg_ir.index],
            marker_line_color=[ch_clr(c) for c in avg_ir.index], marker_line_width=1.5,
            text=[f"{v*100:.2f}%" for v in avg_ir.values],
            textposition="outside", textfont=dict(color="#64748b", size=11),
        ))
        ly3 = base_layout(280, "%", "채널별 평균 유입율")
        fig3.update_layout(**ly3)
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 객단가 추이
    st.subheader("일평균객단가 추이")
    fig4 = go.Figure()
    for ch in sel_channels:
        s = get_series("일평균객단가", ch)
        if len(s) < 2: continue
        sm = s.rolling(roll_e).mean().dropna()
        fig4.add_trace(go.Scatter(
            x=sm.index.strftime("%m/%d").tolist(), y=(sm / 1e4).tolist(),
            mode="lines", name=ch, line=dict(color=ch_clr(ch), width=1.8),
        ))
    ly4 = base_layout(300, "만원", f"채널별 객단가 (만원, {roll_e}일 MA)")
    legend_layout(ly4); ly4["xaxis"]["nticks"] = 15; ly4["xaxis"]["tickangle"] = -30
    fig4.update_layout(**ly4)
    st.plotly_chart(fig4, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 산점도 + 상관 분석
    st.subheader("지표 간 상관 분석")
    st.caption("채널을 선택하면 해당 채널의 일별 데이터로 두 지표의 관계를 분석합니다.")
    ea, eb, ec = st.columns(3)
    xm_lbl  = ea.selectbox("X축 지표", METRIC_LABELS, index=METRIC_LABELS.index("전환율(CR)"), key="eff_x")
    ym_lbl  = eb.selectbox("Y축 지표", METRIC_LABELS, index=METRIC_LABELS.index("유입율"),      key="eff_y")
    ch_scat = ec.selectbox("채널",     sel_channels + ["*TOTAL"], key="eff_ch")
    xm = METRICS_LIST[METRIC_LABELS.index(xm_lbl)]
    ym = METRICS_LIST[METRIC_LABELS.index(ym_lbl)]

    xs = df_f[(df_f["metric"]==xm)&(df_f["segment"]==ch_scat)].set_index("date")["value"]
    ys = df_f[(df_f["metric"]==ym)&(df_f["segment"]==ch_scat)].set_index("date")["value"]
    common = xs.index.intersection(ys.index)
    if len(common) > 5:
        xv = xs[common].values; yv = ys[common].values
        r_val = np.corrcoef(xv, yv)[0, 1]
        reg_r = linreg(xv, yv)
        c1, c2, c3 = st.columns(3)
        c1.metric("상관계수", f"{r_val:.3f}", corr_label(r_val),
                   help="두 지표가 얼마나 함께 움직이는지 (-1~+1)")
        c2.metric("R²", f"{reg_r['r2']:.3f}", sig_label(reg_r["p"]),
                   help="추세선이 데이터를 얼마나 잘 설명하는지 (0~1)")
        c3.metric("표본 수", f"{len(common)}일",
                   help="30일 이상이어야 통계 결과를 신뢰할 수 있습니다.")
        st.caption(f"해석: {corr_label(r_val)} · {r2_label(reg_r['r2'])} · {sig_label(reg_r['p'])}")
        stat_explainer()

        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(
            x=xv.tolist(), y=yv.tolist(), mode="markers",
            marker=dict(color=ch_clr(ch_scat), size=5, opacity=0.65),
            text=xs[common].index.strftime("%Y-%m-%d").tolist(), name=ch_scat,
        ))
        xfit = np.linspace(xv.min(), xv.max(), 100)
        yfit = reg_r["slope"] * xfit + reg_r["intercept"]
        fig5.add_trace(go.Scatter(
            x=xfit.tolist(), y=yfit.tolist(), mode="lines",
            line=dict(color=clr("red"), width=1.5, dash="dot"), name="추세선",
        ))
        ly5 = base_layout(300, title=f"{xm_lbl} vs {ym_lbl} (r={r_val:.3f})")
        ly5["xaxis"]["title"] = xm_lbl; ly5["yaxis"]["title"] = ym_lbl
        legend_layout(ly5)
        fig5.update_layout(**ly5)
        st.plotly_chart(fig5, use_container_width=True)
        memo_block("btm_scatter")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 지표 상관 매트릭스 (TOTAL 기준)
    st.subheader("지표 간 상관 매트릭스 (전체 채널 합산 기준)")
    mat_keys   = ["일평균거래액","일평균고객수","일평균객단가","CR","유입율","DAU"]
    mat_labels = [METRIC_KO.get(m, m) for m in mat_keys]
    mat_df = pd.DataFrame({
        METRIC_KO.get(m, m): df_full[(df_full["metric"]==m)&(df_full["segment"]=="*TOTAL")]
                             .set_index("date")["value"]
        for m in mat_keys
    }).dropna()
    if len(mat_df) > 10:
        cm = mat_df.corr().round(3)
        fig6 = go.Figure(go.Heatmap(
            z=cm.values.tolist(), x=mat_labels, y=mat_labels,
            colorscale=[[0,"rgba(245,101,101,0.8)"],[0.5,"rgba(248,249,252,1)"],
                        [1,"rgba(79,143,255,0.8)"]],
            zmid=0, text=cm.round(2).values.tolist(), texttemplate="%{text}",
            showscale=True,
        ))
        ly6 = base_layout(380, title="지표 간 상관 매트릭스")
        ly6["xaxis"]["tickangle"] = -30
        fig6.update_layout(**ly6)
        st.plotly_chart(fig6, use_container_width=True)
        show_appendix(cm.reset_index().rename(columns={"index":"지표"}), "상관매트릭스")
        memo_block("btm_corr")

    memo_block("btm_efficiency")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    st.subheader("기간 비교 — 효율 지표")
    cmp_e = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_e")
    dc, dp, lc, lp = get_compare_periods(df_full, cmp_e)
    if len(dc) and len(dp):
        for m in ["CR","유입율","일평균객단가"]:
            tbl = compare_table(m, dc, dp, lc, lp)
            st.caption(f"**{METRIC_KO.get(m,m)}**")
            st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)

    insight_editor("fp_efficiency")

# ══════════════════════════════════════════════════════
# 05. 채널 비중 추이
# ══════════════════════════════════════════════════════
elif page == "05. 채널 비중 추이":
    editable_text("fp_h_share", "채널 비중 추이", "h2",
                  "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"필터 기간: {df_f['date'].min().date()} ~ {df_f['date'].max().date()}")
    memo_block("top_share")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    def stacked_area(metric: str, title: str, unit: str = "%", scale: float = 100):
        pivot = (
            df_f[(df_f["metric"]==metric)&(df_f["segment"].isin(sel_channels))]
            .pivot_table(index="date", columns="segment", values="value", aggfunc="first")
            .reindex(columns=[c for c in sel_channels if c in df_f["segment"].unique()])
            .ffill().fillna(0)
        )
        fig = go.Figure()
        for ch in [c for c in sel_channels if c in pivot.columns]:
            fig.add_trace(go.Scatter(
                x=pivot.index.strftime("%m/%d").tolist(),
                y=(pivot[ch] * scale).tolist(),
                mode="lines", name=ch,
                line=dict(width=0.5, color=ch_clr(ch)),
                fill="tonexty", fillcolor=ch_cbg(ch),
                stackgroup="one",
            ))
        ly = base_layout(340, unit, title)
        legend_layout(ly); ly["xaxis"]["nticks"] = 15; ly["xaxis"]["tickangle"] = -30
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("거래액 비중 추이 (채널별 누적)")
    stacked_area("거래액비중", "거래액 비중 추이 (%)", "%", 100)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("고객 비중 추이 (채널별 누적)")
    stacked_area("고객비중", "고객 비중 추이 (%)", "%", 100)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 이중축: 거래액비중 vs CR 특정 채널
    st.subheader("채널별 거래액비중 vs CR 이중축 비교")
    ch_dual = st.selectbox("채널 선택", sel_channels, key="share_dual")
    ms, mc_s = get_series("거래액비중", ch_dual), get_series("CR", ch_dual)
    if len(ms) > 5 and len(mc_s) > 5:
        ms_sm = ms.rolling(7).mean().dropna()
        mc_sm = mc_s.rolling(7).mean().dropna()
        common_d = ms_sm.index.intersection(mc_sm.index)
        if len(common_d) > 3:
            fig_dual = go.Figure()
            fig_dual.add_trace(go.Scatter(
                x=[d.strftime("%m/%d") for d in common_d],
                y=(ms_sm[common_d] * 100).tolist(),
                name="거래액 비중(%)", yaxis="y",
                mode="lines", line=dict(color=ch_clr(ch_dual), width=2),
            ))
            fig_dual.add_trace(go.Scatter(
                x=[d.strftime("%m/%d") for d in common_d],
                y=(mc_sm[common_d] * 100).tolist(),
                name="CR(%)", yaxis="y2",
                mode="lines", line=dict(color=clr("red"), width=2, dash="dot"),
            ))
            ly_d = base_layout(300, title=f"{ch_dual} — 거래액비중 vs CR (이중축)")
            ly_d["showlegend"] = True
            ly_d["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)",
                                   font=dict(color="#64748b"))
            ly_d["yaxis"]["title"]  = "거래액 비중 (%)"
            ly_d["yaxis2"] = dict(overlaying="y", side="right", title="CR (%)",
                                   tickfont=dict(color="#64748b", size=11),
                                   gridcolor="rgba(0,0,0,0)")
            ly_d["xaxis"]["nticks"] = 15; ly_d["xaxis"]["tickangle"] = -30
            fig_dual.update_layout(**ly_d)
            st.plotly_chart(fig_dual, use_container_width=True)

    memo_block("btm_share")
    insight_editor("fp_share")

# ══════════════════════════════════════════════════════
# 06. 요일별 패턴
# ══════════════════════════════════════════════════════
elif page == "06. 요일별 패턴":
    editable_text("fp_h_dow", "요일별 패턴 분석", "h2",
                  "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"필터 기간: {df_f['date'].min().date()} ~ {df_f['date'].max().date()}")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    sel_dow_m_lbl = st.selectbox("분석 지표", METRIC_LABELS, key="dow_m")
    sel_dow_m     = METRICS_LIST[METRIC_LABELS.index(sel_dow_m_lbl)]

    # 요일별 평균 바차트 (채널별)
    st.subheader(f"요일별 {sel_dow_m_lbl} 평균 (채널별)")
    dow_agg = (
        df_f[(df_f["metric"]==sel_dow_m)&(df_f["segment"].isin(sel_channels))]
        .assign(요일=lambda x: x["dow"].map(DOW_MAP))
        .groupby(["요일","segment"])["value"].mean()
        .unstack("segment")
        .reindex(index=[v for v in ["월","화","수","목","금","토","일"] if v in
                        df_f[df_f["metric"]==sel_dow_m]["dow"].map(DOW_MAP).unique()])
        .reindex(columns=[c for c in sel_channels if c in df_f["segment"].unique()])
    )
    if len(dow_agg):
        fig = go.Figure()
        for ch in [c for c in sel_channels if c in dow_agg.columns]:
            scale = 100 if sel_dow_m in ("거래액비중","고객비중","CR","유입율") else (1e8 if sel_dow_m=="일평균거래액" else 1)
            yv = [(v*scale if sel_dow_m in ("거래액비중","고객비중","CR","유입율") else v/scale if sel_dow_m in ("일평균거래액","일평균객단가") else v)
                  if not pd.isnull(v) else None for v in dow_agg[ch].tolist()]
            fig.add_trace(go.Bar(
                x=dow_agg.index.tolist(), y=yv, name=ch,
                marker_color=ch_cbg(ch), marker_line_color=ch_clr(ch), marker_line_width=1,
            ))
        ly = base_layout(300, title=f"요일별 {sel_dow_m_lbl}")
        ly["barmode"] = "group"; legend_layout(ly)
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)
        show_appendix(dow_agg.reset_index().rename(columns={"요일":"요일"}), f"요일별_{sel_dow_m}")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 요일 × 채널 히트맵 (CR)
    st.subheader("요일 × 채널 전환율(CR) 히트맵")
    cr_dow = (
        df_f[(df_f["metric"]=="CR")&(df_f["segment"].isin(sel_channels))]
        .assign(요일=lambda x: x["dow"].map(DOW_MAP))
        .groupby(["요일","segment"])["value"].mean()
        .unstack("segment")
        .reindex(index=[v for v in ["월","화","수","목","금","토","일"] if v in
                        df_f[df_f["metric"]=="CR"]["dow"].map(DOW_MAP).unique()],
                 columns=[c for c in sel_channels if c in df_f["segment"].unique()])
        .fillna(0)
    )
    if len(cr_dow):
        fig2 = go.Figure(go.Heatmap(
            z=(cr_dow.T * 100).values.tolist(),
            x=cr_dow.index.tolist(), y=cr_dow.columns.tolist(),
            colorscale=[[0,"rgba(245,101,101,0.6)"],[0.5,"rgba(248,249,252,1)"],
                        [1,"rgba(72,187,120,0.9)"]],
            zmid=(cr_dow.T * 100).values.mean(),
            text=(cr_dow.T * 100).round(2).values.tolist(),
            texttemplate="%{text:.2f}%",
            showscale=True, colorbar=dict(title="CR%"),
        ))
        ly2 = base_layout(280, title="요일 × 채널 CR (%)")
        fig2.update_layout(**ly2)
        st.plotly_chart(fig2, use_container_width=True)
        show_appendix(cr_dow.T.reset_index().rename(columns={"segment":"채널"}), "요일별CR히트맵")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 요일 × 채널 거래액 히트맵
    st.subheader("요일 × 채널 거래액 히트맵")
    tx_dow = (
        df_f[(df_f["metric"]=="일평균거래액")&(df_f["segment"].isin(sel_channels))]
        .assign(요일=lambda x: x["dow"].map(DOW_MAP))
        .groupby(["요일","segment"])["value"].mean()
        .unstack("segment")
        .reindex(index=[v for v in ["월","화","수","목","금","토","일"] if v in
                        df_f[df_f["metric"]=="일평균거래액"]["dow"].map(DOW_MAP).unique()],
                 columns=[c for c in sel_channels if c in df_f["segment"].unique()])
        .fillna(0)
    )
    if len(tx_dow):
        fig3 = go.Figure(go.Heatmap(
            z=(tx_dow.T / 1e8).values.tolist(),
            x=tx_dow.index.tolist(), y=tx_dow.columns.tolist(),
            colorscale=[[0,"rgba(248,249,252,1)"],[1,"rgba(237,137,54,0.9)"]],
            text=(tx_dow.T / 1e8).round(2).values.tolist(),
            texttemplate="%{text:.2f}억",
            showscale=True, colorbar=dict(title="거래액(억)"),
        ))
        ly3 = base_layout(280, title="요일 × 채널 거래액 (억원)")
        fig3.update_layout(**ly3)
        st.plotly_chart(fig3, use_container_width=True)

    memo_block("btm_dow")
    insight_editor("fp_dow")

# ══════════════════════════════════════════════════════
# 07. 월별 요약
# ══════════════════════════════════════════════════════
elif page == "07. 월별 요약":
    editable_text("fp_h_monthly", "월별 요약", "h2",
                  "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"필터 기간: {df_f['date'].min().date()} ~ {df_f['date'].max().date()}")
    memo_block("top_monthly")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    SUMM = [
        ("일평균거래액", 1e8, "억원"),
        ("일평균고객수", 1,   "명"),
        ("일평균객단가", 1e4, "만원"),
        ("CR",           100, "%"),
        ("유입율",        100, "%"),
    ]

    for metric, scale, unit in SUMM:
        st.subheader(f"월별 {METRIC_KO.get(metric, metric)}")
        monthly_m = get_monthly(metric, ["*TOTAL"] + sel_channels, df_full)
        if monthly_m.empty:
            st.info("데이터 없음"); continue

        def safe_fmt(v, _m=metric):
            try: return fmt_val(_m, v) if not pd.isnull(v) else "–"
            except Exception: return "–"

        # 테이블
        display_tbl = monthly_m.copy()
        for col in display_tbl.columns:
            display_tbl[col] = display_tbl[col].apply(lambda v: safe_fmt(v))
        st.dataframe(display_tbl, use_container_width=True)

        # 채널별 바차트
        mc2 = monthly_m.drop(columns=["*TOTAL"], errors="ignore")
        mc2 = mc2.reindex(columns=[c for c in sel_channels if c in mc2.columns])
        if len(mc2.columns):
            fig = go.Figure()
            for ch in mc2.columns:
                yv = [(v/scale if scale != 100 else v*100) if not pd.isnull(v) else None
                      for v in mc2[ch].tolist()]
                fig.add_trace(go.Bar(
                    x=mc2.index.tolist(), y=yv, name=ch,
                    marker_color=ch_cbg(ch), marker_line_color=ch_clr(ch), marker_line_width=1,
                ))
            ly = base_layout(280, unit, f"월별 채널별 {METRIC_KO.get(metric,metric)} ({unit})")
            ly["barmode"] = "group"; legend_layout(ly); ly["xaxis"]["tickangle"] = -20
            fig.update_layout(**ly)
            st.plotly_chart(fig, use_container_width=True)
            show_appendix(mc2.reset_index().rename(columns={"month":"월"}), f"월별_{metric}")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    memo_block("btm_monthly")

    # 기간 비교
    st.subheader("기간 비교 — 전체 지표")
    cmp_mn = st.selectbox("비교 기준", COMPARE_OPTS, key="cmp_mn")
    dc, dp, lc, lp = get_compare_periods(df_full, cmp_mn)
    if len(dc) and len(dp):
        for m in METRICS_LIST:
            tbl = compare_table(m, dc, dp, lc, lp)
            if len(tbl):
                st.caption(f"**{METRIC_KO.get(m,m)}**")
                st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)
        show_appendix(pd.concat([compare_table(m, dc, dp, lc, lp).assign(지표=m)
                                  for m in METRICS_LIST], ignore_index=True), "전체지표기간비교")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 전체 데이터 다운로드
    st.subheader("📥 전체 데이터 다운로드")
    all_pivot = df_full.pivot_table(
        index=["metric","segment"], columns="date", values="value", aggfunc="first"
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        all_pivot.to_excel(writer, sheet_name="전체데이터")
    st.download_button(
        "⬇️ Excel 다운로드", data=buf.getvalue(),
        file_name="첫구매_채널분석.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    insight_editor("fp_monthly")
