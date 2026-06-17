# -*- coding: utf-8 -*-
"""
LF몰 CRM 발송성과 대시보드
─────────────────────────────────────────────────────────────
발송기획(문구포함) 시트 + 발송성과(실적) 시트를 (발송일 + AF코드)로 머지하여
"어떤 문구·오퍼·타이밍 패턴이 성과를 만드는가"를 도출하는 대시보드.

· 조인 방향: 실적(성과) 기준 — 기획에만 있고 실제 발송 안 된 건은 제외
· 문구 자동 태깅(규칙 기반) + Claude AI 인사이트/처방
· 디자인: 기존 발송 피로도 / 첫구매 주간보고 대시보드 슬레이트 팔레트 계승

데이터 로직(parse_perf_bytes / parse_plan_bytes / merge_perf_plan / tag_copy)은
Streamlit 의존이 없는 순수 함수이며 모듈 import 만으로 테스트 가능하다.
앱 UI 는 main() 안에 있고 `python -m streamlit run` 시에만 실행된다.
"""
import io, os, re, json, hashlib, datetime
import numpy as np
import pandas as pd

# ══════════════════════════════════════════════════════════════════════
# 1. 순수 데이터 로직 (Streamlit 비의존)
# ══════════════════════════════════════════════════════════════════════

AF_RE   = re.compile(r'^(AP|PB)\d+$', re.I)          # PUSH AF코드 형식
WEEK_RE = re.compile(r'\d{1,2}\s*월\s*\d\s*주차')    # 기획 주차 시트명

# 실적 '소재별 실적(당주)' 시트 컬럼 → 표준 키
PERF_COLMAP = {
    "날짜": "date", "요일": "dow_k", "시간대": "hour", "타겟 구분": "target",
    "발송유형": "stype", "BPU": "bpu", "우선순위": "prio", "카테고리": "cat",
    "속성": "attr", "담당자": "owner", "브랜드": "brand", "AF코드": "af",
    "기획전": "promo", "발송": "send", "UV": "uv", "VISIT": "visit",
    "고객수": "cust", "주문건수": "oc", "주문금액": "amt",
    "유입전환율": "infl_cr", "주문전환율": "ord_cr", "효율": "eff",
}
NUM_COLS = ["hour", "send", "uv", "visit", "cust", "oc", "amt", "infl_cr", "ord_cr", "eff"]


def _norm_date(v):
    """엑셀 셀 값 → 'YYYYMMDD' 문자열 또는 None."""
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y%m%d")
    s = re.sub(r'\.0$', '', str(v).strip())
    return s if re.match(r'^\d{8}$', s) else None


def parse_perf_bytes(file_bytes):
    """실적 엑셀 → 캠페인 단위 DataFrame. '소재별 실적(당주)' 시트(또는 AF코드 헤더 시트)."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    target_sheet = None
    if "소재별 실적(당주)" in wb.sheetnames:
        target_sheet = "소재별 실적(당주)"
    else:                                   # 헤더에 AF코드+주문금액 있는 첫 시트
        for s in wb.sheetnames:
            ws = wb[s]
            hdr = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
            hdr = [str(x).strip() if x is not None else "" for x in hdr]
            if "AF코드" in hdr and ("주문금액" in hdr or "발송" in hdr):
                target_sheet = s
                break
    if target_sheet is None:
        wb.close()
        raise ValueError("실적 시트(소재별 실적/AF코드 헤더)를 찾지 못했습니다.")

    ws = wb[target_sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return pd.DataFrame()
    hdr = [str(x).strip() if x is not None else "" for x in rows[0]]
    idx = {h: i for i, h in enumerate(hdr)}
    recs = []
    for r in rows[1:]:
        afi = idx.get("AF코드")
        if afi is None or afi >= len(r) or r[afi] is None:
            continue
        af = str(r[afi]).strip()
        if not AF_RE.match(af):
            continue
        rec = {}
        for kcol, key in PERF_COLMAP.items():
            i = idx.get(kcol)
            rec[key] = r[i] if (i is not None and i < len(r)) else None
        rec["af"] = af
        rec["date"] = _norm_date(rec.get("date"))
        recs.append(rec)
    return _finalize(pd.DataFrame(recs))


def _finalize(df):
    """숫자 변환 + 파생지표(rps/aov/dt/dow) 계산. 신규 파싱·저장소 로드 양쪽에서 재사용."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    df = df.copy()
    df["date"] = df["date"].astype(str).str.replace(r'\.0$', '', regex=True)
    for c in NUM_COLS:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["rps"] = np.where(df["send"].fillna(0) > 0, df["amt"] / df["send"], 0.0)   # 발송건당 거래액
    df["aov"] = np.where(df["oc"].fillna(0) > 0, df["amt"] / df["oc"], 0.0)       # 객단가
    df["dt"] = pd.to_datetime(df["date"], format="%Y%m%d", errors="coerce")
    df["dow"] = df["dt"].dt.dayofweek
    return df


def parse_plan_bytes(file_bytes):
    """기획 엑셀(통합본) → {(date, af): (title, body)} 사전.

    165개 주차 시트를 순회하며, 'AF코드' 가 들어간 헤더 행을 만날 때마다
    컬럼 매핑(일자/AF코드/제목/내용)을 갱신해 하위 표가 여러 개여도 견고하게 파싱한다.
    """
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    lookup = {}
    for sname in wb.sheetnames:
        if not WEEK_RE.search(sname):
            continue
        ws = wb[sname]
        c_af = c_date = c_title = c_body = None
        # 헤더가 2줄로 쪼개져 있으므로(45행=AF코드, 46행=제목/내용) 컬럼 매핑을 행마다 '누적' 갱신
        for row in ws.iter_rows(values_only=True, max_col=45):
            cells = [("" if v is None else str(v)) for v in row]
            for j, c in enumerate(cells):
                cs = c.strip()
                if cs == "AF코드":
                    c_af = j
                elif cs == "제목":
                    c_title = j
                elif cs == "내용":
                    c_body = j
                elif (("일자" in cs or cs == "날짜") and len(cs) < 16):
                    c_date = j
            if c_af is None or c_af >= len(cells):
                continue
            af = cells[c_af].strip()
            if not AF_RE.match(af):
                continue
            d = _norm_date(row[c_date]) if (c_date is not None and c_date < len(row)) else None
            if d is None:                                # 날짜열이 비면 보통 B열(index 1)에 있음
                d = _norm_date(row[1]) if len(row) > 1 else None
            title = cells[c_title].strip() if (c_title is not None and c_title < len(cells)) else ""
            body = cells[c_body].strip() if (c_body is not None and c_body < len(cells)) else ""
            if title or body:
                lookup[(d, af)] = (title, body)
    wb.close()
    return lookup


def merge_perf_plan(perf_df, plan_lookup, keep_unmatched=False):
    """실적 기준 조인 — 문구(제목/내용) 부착. keep_unmatched=False면 문구 없는 행 제외."""
    if perf_df.empty:
        return perf_df.copy()
    df = perf_df.copy()
    titles, bodies, matched = [], [], []
    for _, r in df.iterrows():
        key = (r["date"], r["af"])
        tb = plan_lookup.get(key)
        if tb is None:                                   # 날짜 없이 AF코드만으로 폴백
            cand = [v for (d, a), v in plan_lookup.items() if a == r["af"] and d == r["date"]]
            tb = cand[0] if cand else None
        if tb:
            titles.append(tb[0]); bodies.append(tb[1]); matched.append(True)
        else:
            titles.append(""); bodies.append(""); matched.append(False)
    df["title"] = titles
    df["body"] = bodies
    df["matched"] = matched
    if not keep_unmatched:
        df = df[df["matched"]].reset_index(drop=True)
    return df


# ── 누적 저장소 — 업로드할 때마다 (발송일+AF코드) 키로 병합·영속 ──────────
DATA_STORE = "send_perf_store.csv"
STORE_COLS = list(PERF_COLMAP.values()) + ["title", "body", "matched"]


def _to_bool(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return (not (isinstance(v, float) and np.isnan(v))) and bool(v)
    return str(v).strip().lower() in ("true", "1", "1.0", "yes", "y")


def load_store():
    """누적 CSV → DataFrame. 없으면 빈 프레임."""
    if os.path.exists(DATA_STORE):
        try:
            d = pd.read_csv(DATA_STORE, encoding="utf-8-sig", dtype={"date": str, "af": str})
            return d
        except Exception:
            pass
    return pd.DataFrame(columns=STORE_COLS)


def save_store(df):
    df[[c for c in STORE_COLS if c in df]].to_csv(DATA_STORE, index=False, encoding="utf-8-sig")


def expand_uploads(files):
    """업로드 목록 → [(표시이름, xlsx바이트)]. zip 은 내부 .xlsx 를 모두 풀어낸다(연도별 폴더 OK)."""
    import zipfile
    out = []
    for f in files:
        data = f.getvalue()
        nm = f.name
        if nm.lower().endswith(".zip"):
            try:
                zf = zipfile.ZipFile(io.BytesIO(data))
                for zi in zf.infolist():
                    if zi.is_dir():
                        continue
                    inner = zi.filename
                    base = inner.split("/")[-1]
                    if "__MACOSX" in inner or base.startswith(".") or not base.lower().endswith(".xlsx"):
                        continue
                    try:                                      # 한글 파일명 복원(cp437→cp949)
                        disp = base.encode("cp437").decode("cp949")
                    except Exception:
                        disp = base
                    out.append((disp, zf.read(zi)))
            except Exception as e:
                out.append((f"{nm} (zip 오류: {e})", None))
        else:
            out.append((nm, data))
    return out


def merge_store(old, new):
    """기존 누적 + 신규 업로드 병합 — 같은 (발송일, AF코드)는 신규 우선."""
    def _pick(d):
        if d is None or len(d) == 0:
            return pd.DataFrame(columns=STORE_COLS)
        return d[[c for c in STORE_COLS if c in d]].copy()
    both = pd.concat([_pick(old), _pick(new)], ignore_index=True)
    if both.empty:
        return both
    both["date"] = both["date"].astype(str).str.replace(r'\.0$', '', regex=True)
    return both.drop_duplicates(subset=["date", "af"], keep="last").reset_index(drop=True)


# ── 문구 자동 태깅 ─────────────────────────────────────────────────────
EMOJI_RE = re.compile(
    "[" "\U0001F300-\U0001FAFF" "\U00002600-\U000027BF" "\U0001F000-\U0001F0FF"
    "\U0000FE00-\U0000FE0F" "\U00002190-\U000021FF" "\U00002B00-\U00002BFF"
    "\U000023E0-\U000023FF" "✅❌✨⚠⏰⏳⌚❤❗❓✌☝"
    "]", flags=re.UNICODE)

KW = {
    "혜택강조": ["할인", "세일", "쿠폰", "적립", "증정", "특가", "사은품", "무료배송", "무배",
              "%", "최대", "반값", "균일가", "得", "혜택", "기프트", "페이백", "캐시백"],
    "긴급성":  ["오늘", "마지막", "마감", "종료", "임박", "한정", "지금", "단독", "오픈", "D-",
              "단 ", "단하루", "오직", "초읽기", "마지막날", "곧 ", "예정", "남았", "놓치"],
    "개인화":  ["#{", "고객명", "님,", "님!", "회원님", "등급", "관심", "찜한", "장바구니"],
    "호기심":  ["쉿", "확인", "알림", "[?]", "❓", "비밀", "공개", "대기", "미확인", "두근", "설마",
              "이것", "단 한", "혹시"],
}


def _has(s, words):
    return any(w in s for w in words)


def _s(v):
    """NaN/None/숫자 → 안전한 문자열."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v)


def tag_copy(title, body=""):
    """제목(+본문) → 문구 속성 dict. 분석의 핵심 축."""
    t = _s(title)
    full = (t + " " + _s(body)).strip()
    return {
        "이모지":   bool(EMOJI_RE.search(t)),
        "숫자노출": bool(re.search(r'\d', t)),
        "혜택강조": _has(full, KW["혜택강조"]),
        "긴급성":   _has(full, KW["긴급성"]),
        "개인화":   _has(full, KW["개인화"]),
        "호기심":   _has(full, KW["호기심"]),
        "질문형":   ("?" in t or "？" in t),
        "대괄호":   bool(re.search(r'[\[\]【】（）()]', t)),
        "제목길이": len(t),
        "본문길이": len(body or ""),
    }


TAG_BOOLS = ["이모지", "숫자노출", "혜택강조", "긴급성", "개인화", "호기심", "질문형", "대괄호"]


def add_tags(df):
    """머지 DataFrame 에 문구 속성 컬럼 추가."""
    if df.empty:
        return df
    tags = df.apply(lambda r: tag_copy(r.get("title", ""), r.get("body", "")), axis=1)
    tdf = pd.DataFrame(list(tags), index=df.index)
    return pd.concat([df, tdf], axis=1)


# ══════════════════════════════════════════════════════════════════════
# 2. Streamlit 앱
# ══════════════════════════════════════════════════════════════════════
def main():
    import streamlit as st
    import plotly.graph_objects as go
    from scipy import stats

    st.set_page_config(page_title="LF몰 발송성과 대시보드", layout="wide",
                       initial_sidebar_state="expanded")
    st.markdown("""
    <style>
    [data-testid="stAppViewContainer"]{background:#f8f9fc}
    [data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
    [data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
    [data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
    [data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
    .vg{border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin:8px 0;line-height:1.7;background:#ffffff}
    .sdiv{border-top:1px solid #e2e8f0;margin:22px 0}
    .stat-label{font-size:11px;color:#545c6a;margin-bottom:3px;font-weight:500;letter-spacing:.04em;text-transform:uppercase}
    .appendix{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin-top:12px;font-size:13px;color:#475569}
    .tag{display:inline-block;font-size:11px;padding:1px 7px;border-radius:10px;margin:1px 2px;border:1px solid}
    h1,h2,h3{color:#1e293b}
    </style>""", unsafe_allow_html=True)

    PALETTE = {
        "blue": "rgba(79,143,255,1)", "red": "rgba(245,101,101,1)", "green": "rgba(72,187,120,1)",
        "amber": "rgba(237,137,54,1)", "purple": "rgba(159,122,234,1)", "teal": "rgba(56,178,172,1)",
        "slate": "rgba(100,116,139,1)",
    }

    def base_layout(h=300, ysuffix="", title=""):
        return dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#475569", size=11), margin=dict(l=10, r=10, t=40, b=10),
                    height=h, showlegend=False,
                    title=dict(text=title, font=dict(color="#94a3b8", size=13)),
                    xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#e2e8f0",
                               tickfont=dict(color="#64748b", size=11)),
                    yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                               tickfont=dict(color="#64748b", size=11), ticksuffix=ysuffix))

    def sig_label(p):
        if p is None or (isinstance(p, float) and np.isnan(p)):
            return "–"
        if p < 0.001: return "p<0.001 · 매우 확실"
        if p < 0.01:  return "p<0.01 · 신뢰 가능"
        if p < 0.05:  return "p<0.05 · 유의함"
        return f"p={p:.2f} · 유의하지 않음(우연일 수 있음)"

    def welch(a, b):
        a = np.asarray(a, float); b = np.asarray(b, float)
        a = a[~np.isnan(a)]; b = b[~np.isnan(b)]
        if len(a) < 3 or len(b) < 3:
            return np.nan
        try:
            return stats.ttest_ind(a, b, equal_var=False).pvalue
        except Exception:
            return np.nan

    # ── 캐시 래퍼 (무거운 파싱 1회만) ──
    @st.cache_data(show_spinner=False)
    def cached_perf(b): return parse_perf_bytes(b)

    @st.cache_data(show_spinner=False)
    def cached_plan(b): return parse_plan_bytes(b)

    AI_MODELS = {
        "Claude Opus 4.8 (최고 품질)": "claude-opus-4-8",
        "Claude Sonnet 4.6 (균형)": "claude-sonnet-4-6",
        "Claude Haiku 4.5 (빠름·저렴)": "claude-haiku-4-5",
    }

    def anthropic_key():
        try:
            if "ANTHROPIC_API_KEY" in st.secrets:
                return st.secrets["ANTHROPIC_API_KEY"]
        except Exception:
            pass
        return os.environ.get("ANTHROPIC_API_KEY")

    def ai_generate(system, user, model):
        key = anthropic_key()
        if not key:
            return None, "ANTHROPIC_API_KEY 미설정 — Streamlit Secrets 또는 환경변수에 추가하세요."
        try:
            import anthropic
        except ImportError:
            return None, "anthropic 패키지가 없습니다. requirements.txt 반영 후 재배포하세요."
        try:
            client = anthropic.Anthropic(api_key=key)
            resp = client.messages.create(model=model, max_tokens=2200, system=system,
                                           messages=[{"role": "user", "content": user}])
            text = "".join(b.text for b in resp.content if b.type == "text").strip()
            return (text or None), (None if text else "빈 응답")
        except Exception as e:
            return None, f"생성 오류: {e}"

    # 숫자 포맷 헬퍼
    def won(v):
        if v is None or (isinstance(v, float) and np.isnan(v)): return "–"
        if abs(v) >= 1e8: return f"{v/1e8:.2f}억"
        if abs(v) >= 1e4: return f"{v/1e4:,.0f}만"
        return f"{v:,.0f}"

    # ══════════════════════════════════════════════════════════════
    # 사이드바 — 업로드 / 필터 / 페이지
    # ══════════════════════════════════════════════════════════════
    st.sidebar.markdown("### 📨 발송성과 대시보드")
    st.sidebar.caption("기획(문구) × 실적(성과) 머지 분석")

    perf_files = st.sidebar.file_uploader(
        "① 발송실적 파일 (주차별 xlsx · 연도별 ZIP 가능)",
        type=["xlsx", "zip"], accept_multiple_files=True,
        help="주차별 실적 xlsx 를 여러 개, 또는 연도 폴더를 ZIP 으로 묶어 올리면 전부 풀어 머지합니다.")
    plan_file = st.sidebar.file_uploader(
        "② 발송기획 파일 (통합본)", type=["xlsx"],
        help="새 실적 주차를 추가할 때만 필요합니다. 이미 누적된 데이터만 볼 땐 생략 가능.")

    stored = load_store()
    parse_log = []
    new_raw = None

    if perf_files:
        if not plan_file:
            st.sidebar.warning("새 실적을 머지하려면 기획 파일(②)도 올려주세요.\n(없으면 기존 누적 데이터만 표시)")
        else:
            try:
                with st.spinner("기획 파일(통합본) 파싱 중… 최초 1회만 수십 초 소요됩니다."):
                    plan_lookup = cached_plan(plan_file.getvalue())
            except Exception as e:
                st.error(f"기획 파일 파싱 실패: {e}"); st.stop()
            frames = []
            expanded = expand_uploads(perf_files)
            prog = st.sidebar.progress(0.0, text="실적 파일 머지 중…")
            for k, (nm, b) in enumerate(expanded):
                prog.progress((k + 1) / max(len(expanded), 1), text=f"머지 중… {nm[:24]}")
                if b is None:
                    parse_log.append(f"· {nm}"); continue
                try:
                    pdf = cached_perf(b)
                    mdf = merge_perf_plan(pdf, plan_lookup, keep_unmatched=True)
                    frames.append(mdf[[c for c in STORE_COLS if c in mdf]])
                    mr = mdf["matched"].mean() * 100 if len(mdf) else 0
                    parse_log.append(f"· {nm[:26]} — {len(mdf)}건 (매칭 {mr:.0f}%)")
                except Exception as e:
                    parse_log.append(f"· {nm[:26]} — 실패: {e}")
            prog.empty()
            if frames:
                new_raw = pd.concat(frames, ignore_index=True)

    # 누적 병합 — 저장 전까지는 미리보기(영구 반영 X)
    if new_raw is not None and len(new_raw):
        work = merge_store(stored, new_raw)
        ko = set(zip(stored["date"].astype(str), stored["af"])) if len(stored) else set()
        kn = set(zip(new_raw["date"].astype(str), new_raw["af"]))
        st.sidebar.success(f"신규 {len(new_raw)}건 → 누적 {len(work)}건 "
                           f"(추가 {len(kn - ko)} · 갱신 {len(kn & ko)})")
        if st.sidebar.button("💾 저장 (누적 반영)", use_container_width=True):
            save_store(work); st.cache_data.clear()
            st.sidebar.success("저장됨 ✓ — 다음 세션에도 유지됩니다.")
        st.sidebar.caption("※ 저장을 눌러야 영구 반영됩니다. (미저장 시 이번 세션만 분석)")
    else:
        work = stored

    if work is None or len(work) == 0:
        st.title("LF몰 CRM 발송성과 대시보드")
        st.markdown("""
        <div class="vg">
        <b>발송기획(문구) 시트</b>와 <b>발송실적(성과) 시트</b>를 <code>(발송일 + AF코드)</code>로 머지해
        <b>어떤 문구·오퍼·타이밍 패턴이 성과를 만드는지</b> 도출합니다.<br><br>
        왼쪽에서 <b>① 실적 파일(주차별)</b> 과 <b>② 기획 파일(통합본)</b> 을 올린 뒤
        <b>「저장」</b>을 누르면 누적됩니다.<br>
        · 조인은 <b>실적 기준</b> — 기획에만 있고 실제 발송 안 된 건은 자동 제외됩니다.<br>
        · 다음부터는 <b>새 주차 실적만</b> 올려 누적하면 되고, 누적된 데이터는 기획 파일 없이도 분석됩니다.
        </div>""", unsafe_allow_html=True)
        st.stop()

    # 작업 데이터 확정: 파생 재계산 + 타입 정리 + 문구 재태깅
    raw = _finalize(work.copy())
    raw["title"] = raw["title"].fillna("").astype(str) if "title" in raw else ""
    raw["body"] = raw["body"].fillna("").astype(str) if "body" in raw else ""
    raw["matched"] = raw["matched"].map(_to_bool) if "matched" in raw else False
    raw = add_tags(raw)

    # ── 필터 ──
    st.sidebar.markdown("---")
    st.sidebar.markdown("#### 필터")
    only_matched = st.sidebar.checkbox("문구 매칭된 캠페인만", value=True,
                                       help="실적엔 있으나 기획 문구를 못 찾은 건 제외")
    df = raw[raw["matched"]] if only_matched else raw

    cats = sorted([c for c in df["cat"].dropna().unique()])
    stypes = sorted([c for c in df["stype"].dropna().unique()])
    sel_cat = st.sidebar.multiselect("카테고리", cats)
    sel_st = st.sidebar.multiselect("발송유형", stypes)
    min_send = st.sidebar.number_input("최소 발송수 (분석 표본)", value=5000, step=1000, min_value=0)
    if sel_cat: df = df[df["cat"].isin(sel_cat)]
    if sel_st:  df = df[df["stype"].isin(sel_st)]
    fdf = df[df["send"].fillna(0) >= min_send].reset_index(drop=True)

    st.sidebar.markdown("---")
    page = st.sidebar.radio("페이지", [
        "01. 종합 요약", "08. 전체 효율·추이", "02. 문구 속성별 성과", "03. 캠페인 리더보드",
        "04. 카테고리·시간대 매트릭스", "05. 타이밍·피로도", "06. AI 처방", "07. 데이터·다운로드",
    ])
    model_name = st.sidebar.selectbox("AI 모델", list(AI_MODELS.keys()))
    model = AI_MODELS[model_name]

    st.sidebar.markdown("---")
    st.sidebar.markdown("#### 처리 현황")
    drange = "—"
    if df["date"].notna().any():
        ds = sorted(df["date"].dropna().unique())
        drange = f"{ds[0]} ~ {ds[-1]}"
    st.sidebar.caption(f"전체 {len(raw)}건 · 매칭 {raw['matched'].mean()*100:.0f}% · 분석표본 {len(fdf)}건\n\n기간 {drange}")
    if parse_log:
        with st.sidebar.expander("파싱 로그"):
            st.text("\n".join(parse_log))

    with st.sidebar.expander("누적 데이터 관리"):
        st.caption(f"현재 누적(미저장 포함) {len(work)}건")
        if len(work):
            st.download_button(
                "⬇ 누적 백업 (CSV)",
                work[[c for c in STORE_COLS if c in work]].to_csv(index=False).encode("utf-8-sig"),
                file_name="send_perf_store_backup.csv", mime="text/csv", use_container_width=True)
            st.caption("앱 재배포 시 누적이 초기화될 수 있으니 주기적으로 백업하세요. 이 CSV를 다시 올리면 복원됩니다.")
        rest = st.file_uploader("복원/추가 (백업 CSV)", type=["csv"], key="restore_store")
        if rest is not None:
            try:
                d = pd.read_csv(rest, encoding="utf-8-sig", dtype={"date": str, "af": str})
                save_store(merge_store(load_store(), d)); st.cache_data.clear()
                st.success("복원·병합 완료 ✓ 새로고침하세요")
            except Exception as e:
                st.error(f"복원 실패: {e}")
        if st.button("🗑 누적 초기화", use_container_width=True, key="clear_store"):
            if os.path.exists(DATA_STORE):
                os.remove(DATA_STORE)
            st.cache_data.clear()
            st.success("초기화됨 — 새로고침하세요")

    # 지표 메타
    METRIC_OPTS = {
        "주문전환율": ("ord_cr", "%", PALETTE["purple"]),
        "유입전환율": ("infl_cr", "%", PALETTE["blue"]),
        "발송건당거래액(RPS)": ("rps", "원", PALETTE["green"]),
        "객단가(AOV)": ("aov", "원", PALETTE["amber"]),
        "거래액": ("amt", "원", PALETTE["teal"]),
    }

    # 전환율은 캠페인별 단순 평균으로 비교(표본은 사이드바 '최소 발송수'로 통제)

    # ══════════════════════════════════════════════════════════════
    # PAGE 01 — 종합 요약
    # ══════════════════════════════════════════════════════════════
    if page.startswith("01"):
        st.title("종합 요약")
        st.caption(f"분석 표본: 발송 {min_send:,}건 이상 · {len(fdf)}개 캠페인 · {drange}")
        base = fdf if len(fdf) else df
        c = st.columns(4)
        c[0].metric("발송 캠페인 수", f"{len(base):,}")
        c[1].metric("총 발송 건수", won(base["send"].sum()))
        c[2].metric("총 거래액", won(base["amt"].sum()))
        c[3].metric("문구 매칭률", f"{raw['matched'].mean()*100:.0f}%")
        c = st.columns(4)
        c[0].metric("평균 유입전환율", f"{base['infl_cr'].mean()*100:.2f}%")
        c[1].metric("평균 주문전환율", f"{base['ord_cr'].mean()*100:.2f}%")
        c[2].metric("평균 RPS(발송건당)", won(base["rps"].mean()))
        c[3].metric("평균 객단가", won(base["aov"].mean()))

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        cc = st.columns(2)
        win = base.sort_values("ord_cr", ascending=False).head(8)
        los = base[base["send"] >= min_send].sort_values("ord_cr").head(8)
        with cc[0]:
            st.markdown("##### 🏆 주문전환율 TOP")
            st.dataframe(win[["title", "cat", "send", "ord_cr", "rps"]].rename(
                columns={"title": "제목", "cat": "카테고리", "send": "발송", "ord_cr": "주문CR", "rps": "RPS"}
            ).style.format({"발송": "{:,.0f}", "주문CR": "{:.2%}", "RPS": "{:,.0f}"}),
                hide_index=True, use_container_width=True)
        with cc[1]:
            st.markdown("##### 🧊 주문전환율 BOTTOM")
            st.dataframe(los[["title", "cat", "send", "ord_cr", "rps"]].rename(
                columns={"title": "제목", "cat": "카테고리", "send": "발송", "ord_cr": "주문CR", "rps": "RPS"}
            ).style.format({"발송": "{:,.0f}", "주문CR": "{:.2%}", "RPS": "{:,.0f}"}),
                hide_index=True, use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.markdown("##### 🤖 AI 핵심 인사이트")
        if st.button("AI 인사이트 생성", key="ai_sum"):
            facts = build_facts(base)
            system = ("당신은 LF몰 CRM PUSH 발송 분석가입니다. 주어진 데이터만 근거로 "
                      "핵심 인사이트 3~5개를 한국어 불릿으로 작성하세요. 수치를 지어내지 말고 "
                      "문구 패턴과 성과의 관계에 집중하세요. 출력은 HTML, 불릿은 <br>로 구분, "
                      "긍정은 <span style='color:#16a34a;font-weight:700'>…</span>, "
                      "부정/주의는 <span style='color:#dc2626;font-weight:700'>…</span>로 감싸세요.")
            with st.spinner("생성 중…"):
                txt, err = ai_generate(system, facts, model)
            if err: st.warning(err)
            else: st.session_state["ai_sum_txt"] = txt
        if st.session_state.get("ai_sum_txt"):
            st.markdown(f'<div class="vg">{st.session_state["ai_sum_txt"]}</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════
    # PAGE 02 — 문구 속성별 성과 (핵심)
    # ══════════════════════════════════════════════════════════════
    elif page.startswith("02"):
        st.title("문구 속성별 성과")
        st.caption("각 문구 속성 보유/미보유 그룹의 평균 성과 차이 + 통계 유의성(Welch t-검정)")
        mlabel = st.selectbox("성과 지표", list(METRIC_OPTS.keys()))
        mcol, msuf, mclr = METRIC_OPTS[mlabel]
        base = fdf
        if len(base) < 6:
            st.info("표본이 부족합니다. 사이드바 '최소 발송수'를 낮춰 보세요."); st.stop()

        rows = []
        for tag in TAG_BOOLS:
            yes = base[base[tag]][mcol].dropna().values
            no = base[~base[tag]][mcol].dropna().values
            if len(yes) == 0 or len(no) == 0:
                continue
            my, mn = float(np.mean(yes)), float(np.mean(no))
            rows.append(dict(속성=tag, 보유평균=my, 보유n=len(yes), 미보유평균=mn, 미보유n=len(no),
                             차이=my - mn, p=welch(yes, no)))
        adf = pd.DataFrame(rows).sort_values("차이", ascending=False)

        # Δ 막대
        is_pct = mcol in ("ord_cr", "infl_cr")
        delta_disp = adf["차이"] * (100 if is_pct else 1)
        fig = go.Figure(go.Bar(
            x=delta_disp, y=adf["속성"], orientation="h",
            marker_color=[PALETTE["green"] if v >= 0 else PALETTE["red"] for v in adf["차이"]],
            text=[f"{v:+.2f}{'%p' if is_pct else ''}" for v in delta_disp], textposition="outside"))
        fig.update_layout(**base_layout(h=360, title=f"{mlabel} — 속성 보유 시 평균 차이 (보유 − 미보유)"))
        st.plotly_chart(fig, use_container_width=True)

        def fmtv(v):
            return f"{v*100:.2f}%" if is_pct else (won(v) if mcol in ("rps", "aov", "amt") else f"{v:,.1f}")
        show = adf.copy()
        show["보유평균"] = show["보유평균"].map(fmtv)
        show["미보유평균"] = show["미보유평균"].map(fmtv)
        show["차이"] = [f"{v:+.2f}%p" if is_pct else f"{v:+,.0f}" for v in delta_disp]
        show["유의성"] = show["p"].map(sig_label)
        st.dataframe(show[["속성", "보유평균", "보유n", "미보유평균", "미보유n", "차이", "유의성"]],
                     hide_index=True, use_container_width=True)
        st.markdown('<div class="appendix">속성은 제목·본문 문구를 규칙 기반으로 자동 태깅한 결과입니다. '
                    'n(표본수)이 작으면 차이가 우연일 수 있으니 유의성을 함께 보세요. '
                    '전환율은 캠페인별 단순 평균이며, 최소 발송수 필터로 소표본을 통제합니다.</div>',
                    unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════
    # PAGE 03 — 캠페인 리더보드
    # ══════════════════════════════════════════════════════════════
    elif page.startswith("03"):
        st.title("캠페인 리더보드")
        mlabel = st.selectbox("정렬 지표", list(METRIC_OPTS.keys()))
        mcol = METRIC_OPTS[mlabel][0]
        asc = st.radio("정렬", ["높은순", "낮은순"], horizontal=True) == "낮은순"
        base = fdf.sort_values(mcol, ascending=asc)
        tagcols = [t for t in TAG_BOOLS]
        view = base.copy()
        view["속성"] = view[tagcols].apply(lambda r: " ".join(t for t in tagcols if r[t]), axis=1)
        cols = ["date", "cat", "brand", "title", "send", "infl_cr", "ord_cr", "rps", "amt", "속성"]
        ren = {"date": "날짜", "cat": "카테고리", "brand": "브랜드", "title": "제목", "send": "발송",
               "infl_cr": "유입CR", "ord_cr": "주문CR", "rps": "RPS", "amt": "거래액"}
        st.dataframe(view[cols].rename(columns=ren).style.format(
            {"발송": "{:,.0f}", "유입CR": "{:.2%}", "주문CR": "{:.2%}", "RPS": "{:,.0f}", "거래액": "{:,.0f}"}),
            hide_index=True, use_container_width=True, height=560)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.markdown("##### 🔍 문구 원문 확인")
        opts = {f"[{r['date']}] {r['title'][:40]} (CR {r['ord_cr']*100:.2f}%)": i
                for i, r in base.head(40).reset_index(drop=True).iterrows()}
        bb = base.head(40).reset_index(drop=True)
        sel = st.selectbox("캠페인 선택", list(opts.keys()))
        if sel is not None:
            r = bb.loc[opts[sel]]
            st.markdown(f'<div class="vg"><b>제목</b><br>{r["title"]}<br><br>'
                        f'<b>내용</b><br>{(r["body"] or "—").replace(chr(10),"<br>")}</div>',
                        unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════
    # PAGE 04 — 카테고리·시간대 매트릭스
    # ══════════════════════════════════════════════════════════════
    elif page.startswith("04"):
        st.title("카테고리·시간대 매트릭스")
        mlabel = st.selectbox("지표", list(METRIC_OPTS.keys()))
        mcol = METRIC_OPTS[mlabel][0]
        is_pct = mcol in ("ord_cr", "infl_cr")
        base = fdf

        def heat(idx, col, title):
            pv = base.pivot_table(index=idx, columns=col, values=mcol, aggfunc="mean")
            if pv.empty:
                st.info("데이터 부족"); return
            z = pv.values * (100 if is_pct else 1)
            fig = go.Figure(go.Heatmap(
                z=z, x=[str(c) for c in pv.columns], y=[str(i) for i in pv.index],
                colorscale="Blues", text=np.round(z, 2), texttemplate="%{text}",
                textfont=dict(size=10), colorbar=dict(thickness=10)))
            fig.update_layout(**base_layout(h=420, title=title))
            st.plotly_chart(fig, use_container_width=True)

        heat("cat", "stype", f"카테고리 × 발송유형 — 평균 {mlabel}")
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        heat("hour", "dow_k", f"시간대 × 요일 — 평균 {mlabel}")

    # ══════════════════════════════════════════════════════════════
    # PAGE 05 — 타이밍·피로도
    # ══════════════════════════════════════════════════════════════
    elif page.startswith("05"):
        st.title("타이밍 · 피로도")
        mlabel = st.selectbox("지표", list(METRIC_OPTS.keys()))
        mcol = METRIC_OPTS[mlabel][0]
        is_pct = mcol in ("ord_cr", "infl_cr")
        base = fdf

        def barby(key, title, order=None):
            g = base.groupby(key)[mcol].mean()
            if order: g = g.reindex([o for o in order if o in g.index])
            y = g.values * (100 if is_pct else 1)
            fig = go.Figure(go.Bar(x=[str(i) for i in g.index], y=y,
                                   marker_color=METRIC_OPTS[mlabel][2],
                                   text=[f"{v:.2f}" for v in y], textposition="outside"))
            fig.update_layout(**base_layout(h=320, ysuffix=("%" if is_pct else ""), title=title))
            st.plotly_chart(fig, use_container_width=True)

        cc = st.columns(2)
        with cc[0]: barby("hour", f"시간대별 평균 {mlabel}")
        with cc[1]: barby("dow_k", f"요일별 평균 {mlabel}", order=["월", "화", "수", "목", "금", "토", "일"])

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.markdown("##### 발송량 구간별 성과 (피로도 — 한 캠페인 발송 규모 vs 효율)")
        b = base[base["send"] > 0].copy()
        if len(b) >= 10:
            b["bin"] = pd.qcut(b["send"], q=5, labels=False, duplicates="drop")
            g = b.groupby("bin", observed=True)[mcol].mean().sort_index()
            g.index = [f"Q{i+1}" for i in range(len(g))]
            y = g.values * (100 if is_pct else 1)
            fig = go.Figure(go.Bar(x=[str(i) for i in g.index], y=y, marker_color=PALETTE["amber"],
                                   text=[f"{v:.2f}" for v in y], textposition="outside"))
            fig.update_layout(**base_layout(h=320, ysuffix=("%" if is_pct else ""),
                                            title=f"발송 규모 5분위(Q1 소량→Q5 대량)별 평균 {mlabel}"))
            st.plotly_chart(fig, use_container_width=True)
            st.markdown('<div class="appendix">대량 발송 구간(Q5)에서 효율이 떨어진다면 모수 확대의 한계수익이 '
                        '낮다는 신호입니다. 단, 발송유형/카테고리 구성 차이가 섞일 수 있습니다.</div>',
                        unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════
    # PAGE 06 — AI 처방
    # ══════════════════════════════════════════════════════════════
    elif page.startswith("06"):
        st.title("AI 처방 — 다음 캠페인 카피 가이드")
        st.caption("머지 데이터(문구 속성 × 성과)를 근거로 Claude가 패턴 종합 + 실행 가이드를 작성합니다.")
        base = fdf
        if st.button("AI 처방 생성", key="ai_rx"):
            facts = build_facts(base, with_attr=True)
            system = ("당신은 LF몰 CRM PUSH 카피라이팅 전략가입니다. 주어진 '문구 속성별 성과'와 "
                      "'상·하위 캠페인 문구'를 근거로 다음을 한국어로 작성하세요: "
                      "1) 성과를 가른 핵심 패턴 3가지, 2) 다음 캠페인에 권장하는 카피 공식(Do) 5개, "
                      "3) 피해야 할 패턴(Don't) 3개, 4) 바로 쓸 수 있는 제목 예시 3개. "
                      "수치를 지어내지 말고 데이터 근거를 함께 제시하세요. 출력은 HTML, 소제목은 <b>, "
                      "항목은 <br>로 구분.")
            with st.spinner("생성 중…"):
                txt, err = ai_generate(system, facts, model)
            if err: st.warning(err)
            else: st.session_state["ai_rx_txt"] = txt
        if st.session_state.get("ai_rx_txt"):
            st.markdown(f'<div class="vg">{st.session_state["ai_rx_txt"]}</div>', unsafe_allow_html=True)
        with st.expander("AI에 전달되는 데이터 미리보기"):
            st.text(build_facts(base, with_attr=True))

    # ══════════════════════════════════════════════════════════════
    # PAGE 08 — 전체 효율·추이 (send_dashboard 피로도 관점 계승)
    # ══════════════════════════════════════════════════════════════
    elif page.startswith("08"):
        st.title("전체 효율 · 추이")
        st.caption("누적된 전 발송(문구 매칭 여부 무관)을 주차 단위로 집계한 전체 효율·피로도 관점. "
                   "필터(최소 발송수)와 무관하게 전체 모수를 사용합니다.")
        g = raw.dropna(subset=["dt"]).copy()
        if sel_cat: g = g[g["cat"].isin(sel_cat)]
        if sel_st:  g = g[g["stype"].isin(sel_st)]
        g = g[g["send"].fillna(0) > 0]
        if len(g) < 3:
            st.info("데이터가 부족합니다. 더 많은 주차를 업로드하세요."); st.stop()

        g["주"] = g["dt"].dt.to_period("W").apply(lambda p: p.start_time)
        rows = []
        for wkstart, d in g.groupby("주"):
            s, u, o, a = d["send"].sum(), d["uv"].sum(), d["oc"].sum(), d["amt"].sum()
            rows.append(dict(주=wkstart, 발송=s, 거래액=a, 캠페인수=len(d),
                             유입전환율=(u / s if s else np.nan),
                             주문전환율=(o / u if u else np.nan),
                             RPS=(a / s if s else np.nan)))
        wk = pd.DataFrame(rows).sort_values("주").reset_index(drop=True)

        c = st.columns(4)
        c[0].metric("누적 발송", won(g["send"].sum()))
        c[1].metric("누적 거래액", won(g["amt"].sum()))
        c[2].metric("전체 가중 주문전환율",
                    f"{(g['oc'].sum()/g['uv'].sum())*100:.2f}%" if g["uv"].sum() else "–")
        c[3].metric("전체 RPS", won(g["amt"].sum() / g["send"].sum() if g["send"].sum() else np.nan))

        # 주차별 총 발송량(막대) vs RPS(선) — 이중축
        fig = go.Figure()
        fig.add_bar(x=wk["주"], y=wk["발송"], name="발송량", marker_color=PALETTE["slate"], opacity=0.45)
        fig.add_trace(go.Scatter(x=wk["주"], y=wk["RPS"], name="RPS(발송건당 거래액)",
                                 mode="lines+markers", line=dict(color=PALETTE["green"], width=2), yaxis="y2"))
        lay = base_layout(h=380, title="주차별 총 발송량 vs RPS — 규모를 키울수록 효율이 떨어지나(피로도)")
        lay["showlegend"] = True
        lay["legend"] = dict(orientation="h", y=1.12, bgcolor="rgba(0,0,0,0)")
        lay["yaxis2"] = dict(overlaying="y", side="right", showgrid=False,
                             tickfont=dict(color="#64748b", size=11))
        fig.update_layout(**lay)
        st.plotly_chart(fig, use_container_width=True)

        # 발송량 ↔ 가중 주문전환율 회귀 (피로도 통계 검증)
        x = wk["발송"].values.astype(float)
        y = (wk["주문전환율"].values.astype(float)) * 100
        m = ~np.isnan(x) & ~np.isnan(y)
        if m.sum() >= 5:
            sl, ic, r, p, _ = stats.linregress(x[m], y[m])
            arrow = "효율 하락(피로도 신호)" if r < 0 else "효율 동반 상승"
            st.markdown(
                f'<div class="appendix"><b>피로도 검증</b> — 주차 발송량 ↔ 가중 주문전환율: '
                f'상관 r={r:.2f} ({arrow}), {sig_label(p)}. '
                f'기울기 {sl*1e5:+.4f}%p / 발송 10만건 증가당.</div>', unsafe_allow_html=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.markdown("##### 주차별 전체 실적")
        show = wk.copy()
        show["주"] = show["주"].dt.strftime("%Y-%m-%d")
        st.dataframe(show.style.format({
            "발송": "{:,.0f}", "거래액": "{:,.0f}", "캠페인수": "{:,.0f}",
            "유입전환율": "{:.2%}", "주문전환율": "{:.2%}", "RPS": "{:,.0f}"}),
            hide_index=True, use_container_width=True, height=360)
        st.markdown('<div class="appendix">‘인당 발송 건수’ 기반 피로도(고객 중복 제거)는 이 데이터만으론 계산되지 않습니다 '
                    '— 전사 MTD 발송상세가 필요합니다. 여기서는 캠페인 합산 기준 전체 효율을 봅니다.</div>',
                    unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════
    # PAGE 07 — 데이터·다운로드
    # ══════════════════════════════════════════════════════════════
    else:
        st.title("데이터 · 다운로드")
        st.markdown(f"**머지 결과** — 전체 {len(raw)}건 · 문구 매칭 {raw['matched'].sum()}건 "
                    f"({raw['matched'].mean()*100:.0f}%)")
        st.dataframe(df.drop(columns=["dt"], errors="ignore"), hide_index=True,
                     use_container_width=True, height=420)
        st.download_button("📥 머지 데이터 CSV 다운로드",
                           df.drop(columns=["dt"], errors="ignore").to_csv(index=False).encode("utf-8-sig"),
                           file_name="발송성과_머지.csv", mime="text/csv")
        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.markdown("##### ⚠️ 매칭 진단 — 실적은 있으나 기획 문구를 못 찾은 건")
        miss = raw[~raw["matched"]][["date", "af", "cat", "brand", "send", "amt"]]
        if len(miss):
            st.dataframe(miss.rename(columns={"date": "날짜", "af": "AF코드", "cat": "카테고리",
                                              "brand": "브랜드", "send": "발송", "amt": "거래액"}),
                         hide_index=True, use_container_width=True)
            st.caption("AF코드 오타·기획시트 미등록·날짜 불일치 가능성. 기획 파일의 해당 주차 시트를 확인하세요.")
        else:
            st.success("모든 실적 캠페인에 문구가 매칭되었습니다.")


def build_facts(df, with_attr=False, metric_col="ord_cr"):
    """AI에 넘길 사실 요약 텍스트 생성 (Streamlit 비의존)."""
    if df is None or len(df) == 0:
        return "데이터 없음"
    lines = []
    lines.append(f"[기간] {df['date'].min()} ~ {df['date'].max()} · 캠페인 {len(df)}건")
    lines.append(f"[평균] 유입전환율 {df['infl_cr'].mean()*100:.2f}% · "
                 f"주문전환율 {df['ord_cr'].mean()*100:.2f}% · "
                 f"RPS {df['rps'].mean():,.0f}원 · 객단가 {df['aov'].mean():,.0f}원")
    win = df.sort_values(metric_col, ascending=False).head(6)
    los = df.sort_values(metric_col).head(6)
    lines.append("\n[주문전환율 상위 문구]")
    for _, r in win.iterrows():
        lines.append(f" - CR {r['ord_cr']*100:.2f}% / 발송 {int(r['send']):,} / {r['cat']} / “{r['title']}”")
    lines.append("\n[주문전환율 하위 문구]")
    for _, r in los.iterrows():
        lines.append(f" - CR {r['ord_cr']*100:.2f}% / 발송 {int(r['send']):,} / {r['cat']} / “{r['title']}”")
    if with_attr:
        lines.append("\n[문구 속성별 평균 주문전환율 (보유 vs 미보유)]")
        for tag in TAG_BOOLS:
            if tag not in df: continue
            yes = df[df[tag]]["ord_cr"]; no = df[~df[tag]]["ord_cr"]
            if len(yes) and len(no):
                lines.append(f" - {tag}: 보유 {yes.mean()*100:.2f}%(n={len(yes)}) vs "
                             f"미보유 {no.mean()*100:.2f}%(n={len(no)}) "
                             f"→ {(yes.mean()-no.mean())*100:+.2f}%p")
    return "\n".join(lines)


if __name__ == "__main__":
    main()
