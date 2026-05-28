import streamlit as st
import plotly.graph_objects as go
import seaborn as sns

# ══════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════
st.set_page_config(
    page_title="LF Mall CRM 고객 여정 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&family=DM+Mono:wght@400;500&display=swap');
* { font-family: 'Noto Sans KR', sans-serif; }
[data-testid="stAppViewContainer"] { background: #F4F5F7; }
[data-testid="stMain"] { background: #F4F5F7; }
.block-container { padding: 1.2rem 2rem 3rem !important; max-width: 100% !important; }
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stDecoration"] { display: none; }

/* Topbar strip */
.crm-topbar {
  background: #1A1917; border-radius: 10px; padding: 14px 22px;
  display: flex; align-items: center; gap: 14px; margin-bottom: 22px;
}
.crm-topbar-title { color: #A8A49A; font-size: 13px; }
.crm-topbar-badge {
  font-family: 'DM Mono', monospace; font-size: 9px; letter-spacing:.14em;
  color: #C8381A; text-transform: uppercase;
}

/* Flowchart click buttons */
.stage-click-row { display: flex; gap: 0; margin-top: -6px; margin-bottom: 18px; }

/* Campaign card */
.ccard {
  background: #fff; border: 1px solid #E4E8EE;
  border-radius: 10px; padding: 14px 16px 12px;
  margin-bottom: 10px; border-left: 4px solid #ccc;
  transition: box-shadow .15s, transform .1s;
  cursor: pointer;
}
.ccard:hover { box-shadow: 0 4px 18px rgba(0,0,0,.09); transform: translateY(-1px); }
.ccard-head { display: flex; align-items: center; gap: 8px; margin-bottom: 4px; }
.ccard-code { font-family: 'DM Mono', monospace; font-size: 10px; color: #8A8778; }
.ccard-label { font-size: 13px; font-weight: 700; flex: 1; }
.ccard-timing { font-family: 'DM Mono', monospace; font-size: 10.5px; color: #94A0B2; }

/* Status badges */
.badge { padding: 2px 8px; border-radius: 12px; font-size: 10.5px; font-weight: 600; }
.badge-new  { background:#EBF7EF; color:#1A6E3C; border:1px solid rgba(26,110,60,.18); }
.badge-mod  { background:#EBF1FA; color:#1A5A8A; border:1px solid rgba(26,90,138,.18); }
.badge-keep { background:#F5F4F1; color:#5A5650; border:1px solid #D8D6CE; }
.badge-del  { background:#FDEFED; color:#C8381A; border:1px solid rgba(200,56,26,.18); }

/* Detail section */
.dbox {
  background: #F9FAFB; border: 1px solid #E4E8EE; border-radius: 8px;
  padding: 10px 14px; margin-top: 4px; font-size: 12.5px; color: #3A4455;
  line-height: 1.7; white-space: pre-wrap; word-break: keep-all;
}
.dbox-h { border-left: 3px solid #1A5A8A; background: #EEF4FB; }
.dlabel {
  font-family: 'DM Mono', monospace; font-size: 9px; color: #94A0B2;
  text-transform: uppercase; letter-spacing:.1em; margin-bottom: 3px; margin-top: 14px;
}

/* Gap banner */
.gap-banner {
  background: #FFFBF0; border: 1px solid #F0D580; border-radius: 8px;
  padding: 10px 14px; font-size: 12px; color: #92600A;
  margin-bottom: 16px; line-height: 1.55;
}

/* Summary bar */
.sumbar { display: flex; gap: 1px; border-radius: 8px; overflow: hidden; margin-bottom: 20px; border: 1px solid #E4E8EE; }
.sumcell { flex: 1; background: white; padding: 12px 16px; }
.sumnum { font-family: 'DM Mono', monospace; font-size: 22px; font-weight: 700; line-height: 1; margin-bottom: 2px; }
.sumlbl { font-size: 10px; color: #94A0B2; }

/* Section divider */
.sdiv { border-top: 1px solid #E4E8EE; margin: 18px 0; }

/* Tabs override for dark pill style */
.stTabs [data-baseweb="tab-list"] {
  background: #ECEEF2; padding: 4px; border-radius: 8px; gap: 3px;
}
.stTabs [data-baseweb="tab"] {
  border-radius: 6px; padding: 6px 18px; font-size: 12px; font-weight: 500;
}
.stTabs [aria-selected="true"] {
  background: white !important; box-shadow: 0 1px 4px rgba(0,0,0,.12) !important;
}

/* Back button */
.stButton > button[kind="secondary"] {
  background: white; border: 1px solid #D8DCE4; border-radius: 8px;
  font-size: 12px; padding: 6px 14px;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# DATA
# ══════════════════════════════════════════════════
STAGES = [
    {"id": 0, "label": "가입 / 미구매",     "short": "가입·미구매",  "color": "#2D2B28", "mosu": "33,167명/월",
     "gap": "⚠ 골든타임(0~2h) 트리거 공백 — 첫 접점이 D+1 08:00 SMS"},
    {"id": 1, "label": "탐색 / 구매결정",   "short": "탐색·구매결정","color": "#C8381A", "mosu": "상시",
     "gap": "⚠ 주문서(결제창) 이탈 대응 전무 — PUSH 동의율 순증 –300명/월"},
    {"id": 2, "label": "1회구매 → 재구매",  "short": "1회→재구매",   "color": "#1A5A8A", "mosu": "29,647명/월",
     "gap": "⚠ 배송완료 → D+7 사이 7일 공백 — 리뷰 유도·알림톡 채널 미활용"},
    {"id": 3, "label": "2회구매 → VIP 이관","short": "2회→VIP이관",  "color": "#8A5A00", "mosu": "17,791명/월",
     "gap": "⚠ 3회 구매 후 VIP 이관 자동화 없음 — 등급 넛지(인앱) 전무"},
    {"id": 4, "label": "VIP / Win-back",    "short": "VIP·Win-back", "color": "#1A6E3C", "mosu": "24,195명/월",
     "gap": "⚠ 일반 고객 Win-back 시퀀스 전무 — EV00 VIP 30일 미방문만 대응"},
]

CH = {
    "sms":     {"label": "SMS/LMS", "color": "#C8381A", "icon": "📱"},
    "push":    {"label": "PUSH",    "color": "#1A5A8A", "icon": "🔔"},
    "alimtok": {"label": "알림톡",  "color": "#B07A00", "icon": "💬"},
    "inapp":   {"label": "인앱",    "color": "#1A6E3C", "icon": "📲"},
}

STATUS_MAP = {
    "new":  {"label": "➕ 신규", "cls": "badge-new",  "color": "#1A6E3C"},
    "mod":  {"label": "🔧 수정", "cls": "badge-mod",  "color": "#1A5A8A"},
    "keep": {"label": "✅ 유지", "cls": "badge-keep", "color": "#5A5650"},
    "del":  {"label": "❌ 삭제", "cls": "badge-del",  "color": "#C8381A"},
}

# [stage, ch, code, label, timing, trigger, offer, message]
ASIS = [
  [0,"sms","LL01","웰컴 쿠폰 5종 안내","D+1 08:00","가입 후 1일",
   "플러스쿠폰 5종 / 앱 다운로드 3,000M",
   "(광고)[LFmall] #{고객명}님, LFmall 가입을 축하드립니다.\n신규가입 웰컴 선물이 도착했어요! [플러스쿠폰 5종] 지급 완료♥\n▼ 내 보유쿠폰 보기\n잠깐! 지금 LFmall 앱 다운로드 시 3,000 마일리지가 쏙!"],
  [0,"sms","LL02","첫구매 5% 장바구니쿠폰","D+3 08:00","가입 후 3일 미구매",
   "첫구매 시크릿 5% 장바구니쿠폰 (max 3만원, 7일)",
   "(광고)[LFmall] #{고객명} 고객님, 시크릿 쿠폰이 도착했어요.\n단 7일의 혜택! 장바구니 통째로 할인♥"],
  [0,"sms","LL07","첫구매 5% 쿠폰 (재발급)","D+14 08:00","가입 후 14일 미구매",
   "첫구매 시크릿 5% 장바구니쿠폰 (max 3만원, 7일)",
   "(광고)[LFmall] #{고객명} 고객님, 담아두신 그 상품 시크릿 쿠폰으로 최대할인♥\n쿠폰은 이미 발급 완료 결제 전 꼭 적용하세요!"],
  [0,"sms","LL05","100일 기념 쿠폰","가입 100일 11:00","가입 100일차",
   "5,000원 장바구니쿠폰 (30일)",
   "(광고)[LFmall] 우리.. 오늘 100일♥\n오늘은 #{고객명}님과 LFmall이 만난 지 100일 되는 날이에요!\n소중한 인연을 기념하며 시크릿 [5,000원 장바구니 쿠폰]을 선물 드려요."],
  [0,"alimtok","K0400","통합회원 전환 완료 안내","즉시 발송","통합회원 전환 완료 시",
   "LF Members 멤버십 혜택 안내",
   "LF Members 통합회원 전환이 완료되었습니다.\n멤버십 혜택을 확인하고 즐거운 쇼핑하세요!"],

  [1,"push","O001","장바구니 혜택가 알림","실시간","장바구니 상품 5%↑ 할인 / 30일내 구매자 제외",
   "할인율 직접 노출",
   "(광고) ⭐ {ID}님 장바구니 상품! {DC_RATE}% 할인! ☘️ 놓치면 후회!"],
  [1,"push","O002","장바구니 품절 임박 알림","실시간","재고 10개 미만 / 7일 내 동일 제외",
   "희소성 소구",
   "(광고) {ID}님 상품 품절 주의보⏳ 현재 100명이 장바구니에 담았어요❗\n재고 소량 품절임박⚠️ 품절 전 선점하세요 ☛"],
  [1,"push","P003","장바구니 상품 가격 인하","실시간","장바구니 담은 상품 5%↑ 할인 / 기구매자 제외",
   "가격 인하 직접 알림",
   "(광고) ⚡ {DC_RATE}% 할인! ⚡ {ID}님 담아둔 이 상품 지금 사야 해요!"],
  [1,"push","P004","찜 상품 가격 인하","실시간","최근 1주 내 찜 상품 5%↑ / 재고 5↑ / 20만↑",
   "가격 인하 + 희소성",
   "(광고) 내가 찜❤️했던 상품이 {DC_RATE}% 할인!? 이건 사야해! ❤️"],
  [1,"push","P005","최근 본 상품 가격 인하","실시간","최근 1주 내 조회 상품 5%↑ / 재고 5↑ / 20만↑",
   "관심 상품 가격 인하",
   "(광고) ⭐ {ID}님 관심 상품! {DC_RATE}% 할인! ☘️ 놓치면 후회!"],
  [1,"push","P006","예약판매 입고 안내","실시간","최근 1주 진입 예약상품 + 당일 예약판매 종료",
   "즉시 배송 가능 알림",
   "(광고) ☕ 예약판매 상품이 입고 됐어요! 지금 주문하면 바로 배송 ✨"],
  [1,"push","P007","입점 상품 가격 인하","실시간","최근 본 입점 상품 추가 할인 시",
   "추가 할인 기회",
   "(광고) ⭐추가 할인 시작! 최근에 보신 상품! 득템 할 기회예요 ☘️"],
  [1,"push","P_BR","기획전-브랜드 매칭","실시간","조회·담은 브랜드 기획전 매칭",
   "관심 브랜드 기획전 연결",
   "(광고) 혜택 모아모아 [브랜드] 기획전에서 바로 득템 찬스!\n취향 저격 준비 완료!"],
  [1,"push","P_CA","기획전-카테고리 매칭","실시간","조회 카테고리 기획전 매칭",
   "관심 카테고리 기획전 연결",
   "(광고) 혜택 모아모아 [카테고리] 기획전에서 바로 득템 찬스!"],
  [1,"alimtok","K0279","LIVE 방송 예정 알림","방송 20분 전","알림 신청자 대상",
   "LIVE 방송 사전 알림",
   "알림 신청하신 LF LIVE가 곧 시작됩니다.\n▶ 방송명/혜택/방송시간 안내"],
  [1,"alimtok","K0280","LIVE 방송 시작 알림","방송 시작 즉시","알림 신청자 대상",
   "LIVE 방송 시작 알림",
   "알림 신청하신 LF LIVE가 지금 시작합니다.\n▶ 지금 바로 입장하기"],

  [2,"sms","LL45","배송완료 감사 쿠폰","배송완료 16:00","첫 구매 배송 완료 후",
   "첫구매 감사 10,000원 플러스쿠폰 (30일)",
   "(광고)[LFmall] 첫 구매 감사 쿠폰\n#{고객명} 고객님, 상품은 잘 받아보셨나요?\n또 만나고 싶은 마음을 담아, 첫 구매 감사 쿠폰 드려요! ♥"],
  [2,"sms","LL11","재구매 6% 쿠폰","D+7 11:00","1회 구매 후 7일",
   "재구매 시크릿 6% 장바구니쿠폰 (7일)",
   "(광고)[LFmall] 첫 구매 한 #{고객명}님, 시크릿 혜택 팡!♥\n첫 만남의 설렘을 담아, [장바구니 6% 쿠폰]이 발급되었습니다."],
  [2,"sms","LL12","쿠폰 소멸 리마인드","D+12 11:00","기발급 쿠폰 소멸 D-3 (재발급 X)",
   "기발급 쿠폰 소멸 전 리마인드",
   "(광고)[LFmall] #{고객명} 고객님, 딱! 한번 발급되는 구매 감사 쿠폰이 소멸 예정입니다.\n최대 인기 상품들에 즉시 적용됩니다."],
  [2,"sms","LL15","시크릿 6% 쿠폰 재발급","D+15 15:30","1회 구매 후 15일 미구매",
   "시크릿 6% 장바구니쿠폰 8만원↑ (7일)",
   "(광고)[LFmall] 이 문자는 놓치면 후회해요!\n#{고객명}님만을 위한 쿠폰이 도착했습니다."],

  [3,"sms","LL13","2회 감사 6% 쿠폰","D+7 15:30","2회 구매 후 7일",
   "재구매 시크릿 6% 장바구니쿠폰 (7일)",
   "(광고)[LFmall] #{고객명}님, 두 번째 구매 감사합니다.\n다음에 또 만나고 싶은 마음으로 시크릿 쿠폰 팡!♥"],
  [3,"sms","LL17","쿠폰 소멸 리마인드","D+12 08:00","기발급 쿠폰 소멸 D-3",
   "기발급 쿠폰 소멸 리마인드",
   "(광고)[LFmall] 앗, #{고객명}님 고객님의 구매 감사 쿠폰이 소멸 예정입니다.\n고객님이 봤던 상품들에 즉시 적용 가능하니, 지금 확인하세요 ♥"],
  [3,"sms","LL14","시크릿 6% 쿠폰 재발급","D+15 17:00","2회 구매 후 15일 미구매",
   "시크릿 6% 장바구니쿠폰 8만원↑ (7일)",
   "(광고)[LFmall] 쉿, 이 문자는 놓치면 후회해요!\n#{고객명}님 만을 위한 시크릿 쿠폰이 도착했습니다."],
  [3,"sms","LL18","쿠폰 D+21 리마인드","D+21 08:00","2회 구매 후 21일",
   "기발급 쿠폰 소멸 리마인드 (최후)",
   "(광고)[LFmall] 앗, #{고객명} 고객님의 구매 감사 쿠폰이 소멸 예정입니다.\n최대 인기 상품들에 즉시 적용!"],

  [4,"sms","EV21","승급 유도 안내","매월 23일 14:00","BK/SV 등급 가망 고객",
   "등급 상승 잔여 금액 안내",
   "(문구 미등록 — 담당자 확인 필요)\n핵심 소구: 등급 기회를 놓치지 마세요 / 등급 상승금액만큼만 구매하시면 다음 달 등급 업"],
  [4,"sms","EV25","멤버십 쿠폰 미사용 알림","매월 28일 16:00","VIP 멤버십 쿠폰 미사용 고객",
   "VIP 멤버십 쿠폰 소멸 전 사용 유도",
   "(광고)[LFmall] #{고객명}님이 쇼핑백에 넣어두신 상품에 적용 가능한 멤버십 쿠폰이 아직 남아있어요!\n혜택 소멸 3일전, 지금 바로 확인해보세요."],
  [4,"sms","EV00","30일 미방문 복귀 유도","매일 13:00","최근 30일↑ 미방문 + 6개월내 2회↑ 구매",
   "개인화 추천 업데이트 + 보유 혜택 확인",
   "(광고)[LFmall] #{고객명}님, 방문이 뜸하신 동안 취향에 맞는 상품들이 새로 업데이트 되었어요!\n▶ 오늘 기준 내 취향 보기\n▶ 내 보유 혜택 보기"],
]

# [stage, ch, code, label, timing, trigger, offer, message, status]
TOBE = [
  [0,"alimtok","신규-001","가입 즉시 웰컴 알림톡","가입 즉시","가입 완료 즉시 (0분)",
   "웰컴쿠폰팩(최대 128,000원) + 앱 설치 시 추가 20% 쿠폰",
   "#{이름}님, LFmall 가족이 되신 것을 환영합니다! 🎉\n웰컴쿠폰팩(최대 128,000원)이 발급됐어요.\n앱 설치 시 추가 20% 쿠폰(1일 한정)!","new"],
  [0,"push","신규-002","2h 골든타임 PUSH (장바구니)","가입 후 2시간","장바구니 이탈 감지 시",
   "웰컴쿠폰 + 장바구니 상품 직링크",
   "⚠ 아직 마치지 못한 #{상품명} 주문이 있어요!\n보유 쿠폰 혜택이 사라질 수도 있어요!","new"],
  [0,"push","신규-003","2h 골든타임 PUSH (탐색 이탈)","가입 후 2시간","기획전·상품상세·검색 이탈 시",
   "웰컴쿠폰 + 퍼널별 개인화 랜딩",
   "📌 방금 보셨던 #{기획전명} 아직 진행 중이에요!\n웰컴쿠폰 적용하고 지금 바로 쇼핑해보세요 >","new"],
  [0,"alimtok","K0400","통합회원 전환 완료 안내","즉시 발송","통합회원 전환 완료 시",
   "LF Members 멤버십 혜택 안내",
   "LF Members 통합회원 전환이 완료되었습니다. 멤버십 혜택을 확인하고 즐거운 쇼핑하세요!","keep"],
  [0,"sms","LL01","웰컴 쿠폰 안내 + 알림톡 병행","D+1 08:00","가입 D+1 미구매",
   "웰컴혜택 재안내 + 앱 설치 마일리지",
   "알림톡 채널 병행 추가 → 버튼형으로 쿠폰함·앱설치 딥링크 연결","mod"],
  [0,"sms","LL02","첫구매 5% 장바구니쿠폰","D+3 08:00","가입 D+3 미구매",
   "첫구매 시크릿 5% 쿠폰 (7일)",
   "(유지) 시크릿 쿠폰이 도착했어요 / 단 7일의 혜택 / 장바구니 득템 찬스","keep"],
  [0,"sms","LL07","첫구매 5% 쿠폰 (재발급)","D+14 08:00","가입 D+14 미구매",
   "첫구매 시크릿 5% 쿠폰 (7일)",
   "(유지) 담아두신 그 상품 / 시크릿 쿠폰으로 최대 할인 / 쿠폰은 이미 발급 완료","keep"],
  [0,"sms","LL05","100일 기념 쿠폰","가입 100일 11:00","가입 100일차",
   "5,000원 장바구니쿠폰 (30일)",
   "(유지) LFmall과 만난 지 100일 / 시크릿 5,000원 쿠폰 선물","keep"],
  [0,"sms","신규-004","앱 미설치 구매자 앱설치 유도","D+1","앱 미설치 구매 완료 후",
   "앱 설치 유도 + 앱 전용 20% 쿠폰",
   "앱에서 배송 현황 실시간 확인 + 앱 전용 20% 쿠폰! 지금 앱 설치하고 알림 받기 →","new"],

  [1,"push","신규-101","주문서 이탈 PUSH 1h","이탈 후 1시간","결제창 진입 후 미결제 1h 경과",
   "장바구니 직링크 + 쿠폰 잔여 알림",
   "⚠ 아직 마치지 못한 #{상품명} 주문이 있어요!\n보유 쿠폰 혜택이 사라질 수도 있어요!","new"],
  [1,"sms","신규-102","주문서 이탈 D+1 LMS","이탈 후 D+1","결제창 진입 후 미결제 24h 경과",
   "할인쿠폰 적용 안내 + 주문 완료 유도",
   "놓치고 계신 할인쿠폰이 있어요!\n주문 중이던 #{상품명}… 지금 쿠폰 적용하고 완료하세요!","new"],
  [1,"push","O001","장바구니 혜택가 알림","실시간","장바구니 상품 5%↑ 할인",
   "할인율 직접 노출",
   "(유지) ⭐ {ID}님 장바구니 상품! {DC_RATE}% 할인!","keep"],
  [1,"push","O002","장바구니 품절 임박 알림","실시간","재고 10개 미만",
   "희소성 소구",
   "(유지) {ID}님 상품 품절 주의보⏳ 재고 소량 품절임박⚠️","keep"],
  [1,"push","P003~P007","가격 인하 PUSH 4종","실시간","장바구니·찜·최근본·입점 가격 인하",
   "관심 상품 가격 인하 알림",
   "(유지) 관심 상품 {DC_RATE}% 할인!","keep"],
  [1,"push","P_BR","기획전-브랜드 (운영중)","실시간","관심 브랜드 기획전 매칭",
   "관심 브랜드 기획전 연결",
   "(유지) 혜택 모아모아 [브랜드] 기획전에서 바로 득템 찬스!","keep"],
  [1,"push","P_CA","기획전-카테고리 (운영중)","실시간","관심 카테고리 기획전 매칭",
   "관심 카테고리 기획전 연결",
   "(유지) [카테고리] 기획전에서 바로 득템 찬스!","keep"],
  [1,"alimtok","K027X","LIVE 방송 알림 (운영전환)","방송 전","알림 신청자 대상",
   "LIVE 방송 사전·시작 알림",
   "대기 → 운영 전환. K0279/K0280 알림톡 버튼형으로 개선.","mod"],
  [1,"push","예정","검색 이탈 리타겟팅","이탈 후 2~4h","검색 후 미구매 이탈",
   "검색어 기반 추천 상품",
   "🔎 방금 찾으신 \"#{검색어}\" 관련 상품이에요!\n할인 적용된 지금 바로 확인해보세요 >","new"],

  [2,"sms","LL45","배송완료 감사 + 알림톡 병행","배송완료","첫 구매 배송 완료",
   "첫구매 감사 10,000원 쿠폰 (30일)",
   "알림톡 채널 병행 추가 → 이미지+버튼형으로 쿠폰함 딥링크 연결","mod"],
  [2,"alimtok","신규-201","배송완료 D+1~3 리뷰+쿠폰","배송완료 D+1~3","배송완료 후 리뷰 미작성",
   "리뷰 작성 시 500P + 재구매 쿠폰 10% (14일)",
   "#{상품명} 잘 받아보셨나요? 😊\n리뷰 작성 시 500P + 재구매 쿠폰 10%(14일) 발급!","new"],
  [2,"alimtok","신규-202","배송완료 D+7 스타일링 콘텐츠","배송완료 D+7","배송완료 후 7일",
   "활용팁·스타일링 콘텐츠 (세일즈 소구 없음)",
   "구매하신 #{상품명} 잘 활용하고 계신가요?\n소재에 맞는 세탁·보관법 안내드려요 → 앱 내 스타일 콘텐츠","new"],
  [2,"sms","LL11","재구매 6% 쿠폰 + 알림톡 병행","D+7 11:00","1회 구매 후 7일",
   "재구매 시크릿 6% 장바구니쿠폰 (7일)",
   "알림톡 채널 추가 + 개인화 추천 버튼 → 쿠폰함 바로가기 / 추천 상품 딥링크","mod"],
  [2,"sms","LL12","쿠폰 리마인드 (삭제 검토)","D+12","기발급 쿠폰 소멸 리마인드",
   "❌ CTR·CR 실측 후 효과 낮으면 폐지",
   "CTR·CR 실측 결과에 따라 폐지 또는 알림톡 전환 검토.\n현재 7일 내 복수 SMS 수신 피로도 우려.","del"],
  [2,"sms","LL15","시크릿 6% 쿠폰 재발급","D+15 15:30","1회 구매 후 15일 미구매",
   "시크릿 6% 장바구니쿠폰 8만원↑ (7일)",
   "(유지) 이 문자는 놓치면 후회해요 / 고객님만을 위한 쿠폰","keep"],

  [3,"inapp","신규-301","2회 구매 직후 등급 넛지","즉시 (앱 오픈 시)","2회 구매 완료 직후",
   "다음 등급까지 잔여 금액 실시간 안내",
   "벌써 2번째 구매 감사합니다 🎉\nPurple까지 OO만원 남았어요 👉","new"],
  [3,"sms","LL13","2회 감사 6% 쿠폰","D+7 15:30","2회 구매 후 7일",
   "재구매 시크릿 6% 쿠폰 (7일)",
   "(유지) 두 번째 구매 감사합니다 / 시크릿 쿠폰 팡","keep"],
  [3,"sms","LL17","쿠폰 리마인드 (삭제 검토)","D+12","기발급 쿠폰 소멸 리마인드",
   "❌ CTR·CR 실측 후 효과 낮으면 폐지",
   "CTR·CR 실측 결과에 따라 폐지 검토.","del"],
  [3,"sms","LL14","시크릿 6% 쿠폰 재발급","D+15 17:00","2회 구매 후 15일 미구매",
   "시크릿 6% 장바구니쿠폰 8만원↑ (7일)",
   "(유지) 이 문자는 놓치면 후회해요 / 고객님만을 위한 쿠폰","keep"],
  [3,"sms","LL18","D+21 리마인드 (삭제 검토)","D+21","2회 구매 후 21일",
   "❌ 월평균 948명, 실효성 점검 필요",
   "CTR·CR 실측 결과에 따라 폐지 검토.","del"],
  [3,"push","신규-302","D+14 개인화 추천 PUSH","D+14","2회 구매 후 14일 미구매",
   "구매 이력 기반 1:1 개인화 추천",
   "#{이름}님이 좋아하는 스타일의 신상이 도착했어요 👗","new"],

  [4,"alimtok","신규-401","3회 구매 → VIP 이관 알림톡","즉시","3회 구매 달성 시",
   "LF Members Black/Purple 혜택 안내",
   "#{이름}님이 드디어 VIP 등급에 도달했어요! 🌟\nLF Members Black/Purple 혜택 확인하세요.","new"],
  [4,"sms","EV21","승급 유도 안내","매월 23일 14:00","승급 가망 고객",
   "등급 상승 잔여 금액 안내",
   "(유지) 등급 기회를 놓치지 마세요 / 등급 상승금액만큼만 구매하시면 다음 달 등급 업","keep"],
  [4,"sms","EV25","멤버십 쿠폰 미사용 알림","매월 28일 16:00","VIP 멤버십 쿠폰 미사용",
   "VIP 멤버십 쿠폰 소멸 전 사용 유도",
   "(유지) 쇼핑에 넣어두신 상품에 적용 가능한 멤버십 쿠폰이 아직 남아있어요!","keep"],
  [4,"sms","EV00","30일 미방문 복귀 유도","매일 13:00","VIP 30일↑ 미방문",
   "개인화 추천 + 보유 혜택 확인",
   "(유지) 취향에 맞는 상품들이 새로 업데이트 되었어요!","keep"],
  [4,"sms","신규-402","Win-back 90일 (1차)","90일 미방문","일반 고객 90일↑ 미방문",
   "무료배송 쿠폰 7일",
   "#{이름}님, 오랜만이에요 😢\n오늘 돌아오시면 7일 무료배송 쿠폰을 드려요.\n새로운 스타일 업데이트됐어요!","new"],
  [4,"sms","신규-403","Win-back 180일 (2차)","180일 미방문","일반 고객 180일↑ 미방문",
   "10% 할인쿠폰 7일",
   "#{이름}님이 보고 싶어요!\n특별히 10% 쿠폰 준비했어요 → 지금 쇼핑하기","new"],
  [4,"sms","신규-404","Win-back 270일 (3차)","270일 미방문","일반 고객 270일↑ 미방문",
   "15% 할인쿠폰 7일 (최후 수단)",
   "마지막으로 한 번만 더 만나고 싶어요 💌\n15% 쿠폰 드릴게요. 기한은 7일!","new"],
]

# ══════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════
def get_offer_tags(offer: str) -> list[str]:
    tags = []
    if any(k in offer for k in ["마일리지", "M ", "M)", "3,000M", "포인트", "500P", "마일"]):
        tags.append("마일리지")
    if any(k in offer for k in ["쿠폰", "%", "할인", "무료배송"]):
        tags.append("쿠폰")
    return tags or ["기타"]

def filter_items(data, stage_id, ch_filter="전체", offer_filter="전체"):
    items = [d for d in data if d[0] == stage_id]
    if ch_filter != "전체":
        items = [d for d in items if CH.get(d[1], {}).get("label") == ch_filter]
    if offer_filter != "전체":
        items = [d for d in items if offer_filter in get_offer_tags(d[6])]
    return items

def badge_html(status_key):
    if not status_key or status_key not in STATUS_MAP:
        return ""
    s = STATUS_MAP[status_key]
    return f'<span class="badge {s["cls"]}">{s["label"]}</span>'

def ch_dot(ch_key, size=9):
    color = CH.get(ch_key, {}).get("color", "#888")
    return f'<span style="display:inline-block;width:{size}px;height:{size}px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle"></span>'

# ══════════════════════════════════════════════════
# FLOWCHART (Plotly)
# ══════════════════════════════════════════════════
def make_journey_chart(data, selected_stage=None):
    # Seaborn palette for chart accents
    pal = sns.color_palette("muted", 5)

    BOX_W, BOX_H = 2.0, 1.4
    GAP = 0.55
    Y0 = 2.0
    TOTAL_W = len(STAGES) * BOX_W + (len(STAGES) - 1) * GAP

    fig = go.Figure()

    scatter_x, scatter_y, scatter_hover, scatter_custom = [], [], [], []

    for i, stage in enumerate(STAGES):
        x = i * (BOX_W + GAP)
        cx = x + BOX_W / 2

        items = [d for d in data if d[0] == i]
        n = len(items)
        ch_counts = {}
        for it in items:
            ch_counts[it[1]] = ch_counts.get(it[1], 0) + 1

        is_sel = (selected_stage == i)
        fill_color = stage["color"]
        border_w = 3 if is_sel else 0
        border_color = "#FFFFFF" if is_sel else fill_color
        opacity = 1.0 if is_sel else 0.82

        # Stage box
        fig.add_shape(
            type="rect", x0=x, y0=Y0, x1=x + BOX_W, y1=Y0 + BOX_H,
            fillcolor=fill_color, opacity=opacity,
            line=dict(color=border_color, width=border_w),
            layer="below",
        )
        if is_sel:
            fig.add_shape(
                type="rect", x0=x - 0.06, y0=Y0 - 0.06,
                x1=x + BOX_W + 0.06, y1=Y0 + BOX_H + 0.06,
                fillcolor="rgba(0,0,0,0)",
                line=dict(color="#FFFFFF", width=2.5, dash="dot"),
            )

        # Stage name
        fig.add_annotation(
            x=cx, y=Y0 + BOX_H - 0.28,
            text=f"<b>{stage['short']}</b>",
            font=dict(color="white", size=11, family="Noto Sans KR"),
            showarrow=False, xanchor="center",
        )
        # Campaign count
        fig.add_annotation(
            x=cx, y=Y0 + BOX_H / 2 - 0.05,
            text=f"<b>{n}</b>",
            font=dict(color="white", size=22, family="DM Mono"),
            showarrow=False, xanchor="center",
        )
        fig.add_annotation(
            x=cx, y=Y0 + 0.15,
            text="건",
            font=dict(color="rgba(255,255,255,0.7)", size=10),
            showarrow=False, xanchor="center",
        )

        # Mosu below box
        fig.add_annotation(
            x=cx, y=Y0 - 0.2,
            text=f"<span style='color:#64748B;font-size:9px'>{stage['mosu']}</span>",
            font=dict(color="#64748B", size=9),
            showarrow=False, xanchor="center",
        )

        # Channel dots below
        ch_order = ["sms", "push", "alimtok", "inapp"]
        dot_x_start = cx - 0.36
        for ci, ch_key in enumerate(ch_order):
            if ch_key in ch_counts:
                dot_color = CH[ch_key]["color"]
                fig.add_shape(
                    type="circle",
                    x0=dot_x_start + ci * 0.2 - 0.07,
                    y0=Y0 - 0.55,
                    x1=dot_x_start + ci * 0.2 + 0.07,
                    y1=Y0 - 0.41,
                    fillcolor=dot_color, opacity=0.85,
                    line=dict(color=dot_color, width=0),
                )

        # Connector arrow
        if i < len(STAGES) - 1:
            fig.add_annotation(
                x=x + BOX_W + GAP * 0.05, y=Y0 + BOX_H / 2,
                ax=x + BOX_W + GAP * 0.95, ay=Y0 + BOX_H / 2,
                xref="x", yref="y", axref="x", ayref="y",
                showarrow=True, arrowhead=2, arrowsize=1.1,
                arrowwidth=2, arrowcolor="#CBD5E1",
            )

        # Invisible scatter for click detection
        scatter_x.append(cx)
        scatter_y.append(Y0 + BOX_H / 2)
        scatter_hover.append(
            f"<b>{stage['label']}</b><br>캠페인 {n}건<br><i>클릭해서 상세보기</i>"
        )
        scatter_custom.append(i)

    # Scatter trace for clicks
    fig.add_trace(go.Scatter(
        x=scatter_x, y=scatter_y,
        mode="markers",
        marker=dict(size=55, color="rgba(0,0,0,0)", symbol="square"),
        hovertemplate="%{text}<extra></extra>",
        text=scatter_hover,
        customdata=scatter_custom,
        showlegend=False,
    ))

    fig.update_layout(
        height=240,
        margin=dict(l=10, r=10, t=10, b=30),
        xaxis=dict(visible=False, range=[-0.3, TOTAL_W + 0.3]),
        yaxis=dict(visible=False, range=[Y0 - 0.9, Y0 + BOX_H + 0.4]),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        dragmode=False,
        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Noto Sans KR"),
    )
    return fig

# ══════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════
def ss(k, default):
    if k not in st.session_state:
        st.session_state[k] = default

ss("page", "main")
ss("stage", 0)
ss("view", "tobe")

# ══════════════════════════════════════════════════
# TOPBAR (persistent)
# ══════════════════════════════════════════════════
col_logo, col_title, col_view = st.columns([0.12, 0.6, 0.28])
with col_logo:
    st.markdown('<div style="font-family:DM Mono,monospace;font-size:10px;color:#C8381A;letter-spacing:.14em;padding-top:6px">● CRM</div>', unsafe_allow_html=True)
with col_title:
    st.markdown('<div style="font-size:13px;color:#3A4455;padding-top:5px;font-weight:500">LF Mall 자동화 메시지 · 고객 여정 대시보드</div>', unsafe_allow_html=True)
with col_view:
    view_choice = st.radio(
        "뷰", ["AS-IS", "TO-BE"], horizontal=True,
        index=0 if st.session_state.view == "asis" else 1,
        label_visibility="collapsed",
        key="view_radio",
    )
    new_view = "asis" if view_choice == "AS-IS" else "tobe"
    if new_view != st.session_state.view:
        st.session_state.view = new_view
        st.rerun()

st.divider()

DATA = ASIS if st.session_state.view == "asis" else TOBE

# ══════════════════════════════════════════════════
# PAGE: MAIN — full journey flowchart
# ══════════════════════════════════════════════════
if st.session_state.page == "main":
    st.markdown("#### 전체 고객 여정")
    st.caption("단계 블록을 클릭하거나 아래 버튼을 눌러 캠페인 상세를 확인하세요.")

    # ── Plotly flowchart ──────────────────────────
    fig = make_journey_chart(DATA)
    event = st.plotly_chart(
        fig,
        use_container_width=True,
        on_select="rerun",
        key="journey_chart",
        config={"displayModeBar": False},
    )
    if event and event.selection and event.selection.points:
        stage_idx = int(event.selection.points[0].customdata)
        st.session_state.stage = stage_idx
        st.session_state.page = "detail"
        st.rerun()

    # ── Stage button row (always visible fallback) ─
    st.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)
    cols = st.columns(5)
    for i, (stage, col) in enumerate(zip(STAGES, cols)):
        items = [d for d in DATA if d[0] == i]
        with col:
            btn_label = f"{stage['short']}\n▶ {len(items)}건"
            if st.button(btn_label, key=f"sb_{i}", use_container_width=True):
                st.session_state.stage = i
                st.session_state.page = "detail"
                st.rerun()

    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    st.divider()

    # ── Overview stats ─────────────────────────────
    st.markdown("#### 캠페인 현황 요약")
    total = len(DATA)
    by_ch = {k: sum(1 for d in DATA if d[1] == k) for k in CH}

    c0, c1, c2, c3, c4 = st.columns(5)
    cells = [
        (c0, str(total), "전체 캠페인", "#1A1917"),
        (c1, str(by_ch["sms"]), "SMS / LMS", "#C8381A"),
        (c2, str(by_ch["push"]), "PUSH", "#1A5A8A"),
        (c3, str(by_ch["alimtok"]), "알림톡", "#B07A00"),
        (c4, str(by_ch["inapp"]), "인앱", "#1A6E3C"),
    ]
    for col, num, lbl, color in cells:
        with col:
            st.markdown(
                f'<div class="sumcell" style="border-radius:10px;border:1px solid #E4E8EE">'
                f'<div class="sumnum" style="color:{color}">{num}</div>'
                f'<div class="sumlbl">{lbl}</div></div>',
                unsafe_allow_html=True,
            )

    if st.session_state.view == "tobe":
        st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
        cnt = {k: 0 for k in STATUS_MAP}
        for d in TOBE:
            if len(d) > 8 and d[8] in cnt:
                cnt[d[8]] += 1
        cc0, cc1, cc2, cc3 = st.columns(4)
        for col, key, label in zip(
            [cc0, cc1, cc2, cc3],
            ["new", "mod", "keep", "del"],
            ["신규 추가", "수정·개선", "유지", "삭제검토"],
        ):
            s = STATUS_MAP[key]
            with col:
                st.markdown(
                    f'<div class="sumcell" style="border-radius:10px;border:1px solid #E4E8EE">'
                    f'<div class="sumnum" style="color:{s["color"]}">{cnt[key]}</div>'
                    f'<div class="sumlbl">{label}</div></div>',
                    unsafe_allow_html=True,
                )

# ══════════════════════════════════════════════════
# PAGE: DETAIL — stage campaigns
# ══════════════════════════════════════════════════
else:
    stage = STAGES[st.session_state.stage]

    # ── Back + stage selector ──────────────────────
    col_back, col_stages = st.columns([0.12, 0.88])
    with col_back:
        if st.button("← 전체 여정", key="back_btn"):
            st.session_state.page = "main"
            st.rerun()
    with col_stages:
        stage_labels = [s["short"] for s in STAGES]
        sel = st.radio(
            "단계 선택", stage_labels, horizontal=True,
            index=st.session_state.stage,
            label_visibility="collapsed",
            key="stage_radio",
        )
        new_stage = stage_labels.index(sel)
        if new_stage != st.session_state.stage:
            st.session_state.stage = new_stage
            st.rerun()

    # ── Mini flowchart ─────────────────────────────
    fig2 = make_journey_chart(DATA, selected_stage=st.session_state.stage)
    event2 = st.plotly_chart(
        fig2,
        use_container_width=True,
        on_select="rerun",
        key="detail_chart",
        config={"displayModeBar": False},
    )
    if event2 and event2.selection and event2.selection.points:
        stage_idx = int(event2.selection.points[0].customdata)
        if stage_idx != st.session_state.stage:
            st.session_state.stage = stage_idx
            st.rerun()

    # ── Stage header ───────────────────────────────
    st.markdown(
        f'<div style="background:{stage["color"]};border-radius:10px;padding:16px 22px;'
        f'color:white;margin-bottom:16px;">'
        f'<div style="font-size:17px;font-weight:900;letter-spacing:-.02em">{stage["label"]}</div>'
        f'<div style="font-size:11px;opacity:.75;margin-top:4px">{stage["mosu"]}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # GAP banner
    st.markdown(
        f'<div class="gap-banner">{stage["gap"]}</div>',
        unsafe_allow_html=True,
    )

    # ── Filters ────────────────────────────────────
    f_col1, f_col2 = st.columns(2)

    with f_col1:
        ch_labels = ["전체", "SMS/LMS", "PUSH", "알림톡", "인앱"]
        ch_tabs = st.tabs([f"📡 발송채널  ▸  {l}" if l == "전체" else l for l in ch_labels])

    with f_col2:
        offer_labels = ["전체", "쿠폰", "마일리지"]
        offer_tabs = st.tabs([f"🎁 오퍼  ▸  {l}" if l == "전체" else l for l in offer_labels])

    # Determine active tab via session state or default
    ss("ch_tab_idx", 0)
    ss("offer_tab_idx", 0)

    for i, tab in enumerate(ch_tabs):
        with tab:
            selected_ch = ch_labels[i]
            for j, otab in enumerate(offer_tabs):
                with otab:
                    selected_offer = offer_labels[j]
                    items = filter_items(DATA, st.session_state.stage, selected_ch, selected_offer)

                    if not items:
                        st.info("해당 조건의 캠페인이 없습니다.")
                        continue

                    # TOBE 변경 요약 (tobe 뷰에서만)
                    if st.session_state.view == "tobe":
                        cnt = {k: 0 for k in STATUS_MAP}
                        for d in items:
                            if len(d) > 8 and d[8] in cnt:
                                cnt[d[8]] += 1
                        sc0, sc1, sc2, sc3 = st.columns(4)
                        for col, key in zip([sc0, sc1, sc2, sc3], ["new", "mod", "keep", "del"]):
                            s = STATUS_MAP[key]
                            with col:
                                st.markdown(
                                    f'<div class="sumcell" style="border-radius:8px;border:1px solid #E4E8EE;padding:8px 12px">'
                                    f'<div class="sumnum" style="color:{s["color"]};font-size:18px">{cnt[key]}</div>'
                                    f'<div class="sumlbl">{s["label"]}</div></div>',
                                    unsafe_allow_html=True,
                                )
                        st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)

                    # Group by channel
                    ch_order = ["sms", "alimtok", "push", "inapp"]
                    for ch_key in ch_order:
                        ch_items = [d for d in items if d[1] == ch_key]
                        if not ch_items:
                            continue

                        ch_info = CH[ch_key]
                        st.markdown(
                            f'<div style="display:flex;align-items:center;gap:8px;margin:16px 0 8px">'
                            f'<div style="width:18px;height:3px;border-radius:2px;background:{ch_info["color"]}"></div>'
                            f'<span style="font-family:DM Mono,monospace;font-size:10px;color:{ch_info["color"]};'
                            f'font-weight:600;text-transform:uppercase;letter-spacing:.1em">'
                            f'{ch_info["icon"]} {ch_info["label"]}</span>'
                            f'<span style="font-family:DM Mono,monospace;font-size:9px;color:{ch_info["color"]};'
                            f'background:{ch_info["color"]}18;border:1px solid {ch_info["color"]}30;'
                            f'padding:2px 8px;border-radius:10px">{len(ch_items)}건</span></div>',
                            unsafe_allow_html=True,
                        )

                        card_cols = st.columns(2)
                        for ci, item in enumerate(ch_items):
                            code    = item[2]
                            label   = item[3]
                            timing  = item[4]
                            trigger = item[5]
                            offer   = item[6]
                            message = item[7]
                            status  = item[8] if len(item) > 8 else None
                            s_info  = STATUS_MAP.get(status) if status else None

                            offer_tags = get_offer_tags(offer)
                            tag_html = " ".join(
                                f'<span style="font-size:9px;background:{"#FEF9E8" if t=="마일리지" else "#EEF8FF"};'
                                f'color:{"#92600A" if t=="마일리지" else "#1A5A8A"};'
                                f'border:1px solid {"#F0D580" if t=="마일리지" else "#A8CEE8"};'
                                f'padding:1px 6px;border-radius:8px">{"🏅" if t=="마일리지" else "🎫"} {t}</span>'
                                for t in offer_tags if t != "기타"
                            )

                            status_badge = (
                                f'<span class="badge {s_info["cls"]}">{s_info["label"]}</span>'
                                if s_info else
                                '<span style="font-size:10px;color:#94A0B2">AS-IS 운영중</span>'
                            )

                            with card_cols[ci % 2]:
                                with st.expander(
                                    f"**{label}**  \n"
                                    f"⏱ {timing}",
                                    expanded=False,
                                ):
                                    st.markdown(
                                        f'<div style="display:flex;gap:6px;align-items:center;margin-bottom:10px">'
                                        f'<span style="font-family:DM Mono,monospace;font-size:10px;color:#94A0B2">{code}</span>'
                                        f'{status_badge}'
                                        f'{tag_html}</div>',
                                        unsafe_allow_html=True,
                                    )
                                    st.markdown('<div class="dlabel">🎯 트리거 조건</div>', unsafe_allow_html=True)
                                    st.markdown(f'<div class="dbox">{trigger}</div>', unsafe_allow_html=True)
                                    st.markdown('<div class="dlabel">🎁 오퍼 / 혜택</div>', unsafe_allow_html=True)
                                    st.markdown(f'<div class="dbox dbox-h">{offer}</div>', unsafe_allow_html=True)
                                    st.markdown('<div class="dlabel">💬 메시지 / 문구</div>', unsafe_allow_html=True)
                                    st.markdown(f'<div class="dbox">{message}</div>', unsafe_allow_html=True)
