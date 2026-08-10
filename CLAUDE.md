# LF Mall CRM 고객 여정 대시보드 — 개발 가이드

## 프로젝트 개요

저장소에 여러 Streamlit 앱이 있다. **지금 활발히 개발하는 건 아래 둘**이고,
`crm_journey.*`는 초기 산출물이라 손댈 일이 거의 없다.

| 파일 | 앱 | 규모 |
|------|-----|------|
| `send_perf_dashboard.py` | 발송성과 대시보드 (11페이지 + 하위탭) | ~7,300줄 |
| `weekly_report.py` | 주간보고 (7페이지) | ~2,500줄 |
| `crm_journey.html` / `.py` | 고객 여정 (초기 산출물) | — |

- **배포**: Streamlit Share (main 머지 시 자동 배포)
- **저장소**: `lee7120-creator/lf-dashboard`
- **브랜치**: 세션마다 지정된 `claude/*` 브랜치 → main squash 머지 후 배포 반영

---

## ⚠️ 테스트 · CI — 머지 전 반드시 통과

**이걸 안 돌리고 머지해서 앱 전체가 다운된 적이 두 번 있다.** 문법 오류도 린트 경고도
아니고, 해당 페이지를 열어야만 재현되는 종류라 코드 리뷰로는 안 잡힌다.

```bash
python tests/check_shadowing.py           # 몇 초 — 전역 헬퍼 섀도잉 정적 검사
python tests/test_plan_merge.py           # 몇 초 — 실적↔기획 문구 조인 (날짜·AF 키)
python tests/test_brand_classify.py       # 몇 초 — 브랜드 자동 분류 (오탐 방어)
python tests/test_store_layer.py          # 몇 초 — push dtype 방어 (문자열 라운드트립)
python tests/test_table_export.py         # 몇 초 — 표 엑셀 내보내기 (숫자 서식·글자색)
python tests/test_hour_norm.py            # 몇 초 — 발송 시간대 자동 보정 (2400·8000)
python tests/smoke_pages.py               # 몇 분 — 발송성과 전 페이지·하위탭 렌더
python tests/test_filter_follows_upload.py # 몇 분 — 업로드 후 기간 필터·차트 레이블
python tests/smoke_weekly_report.py       # 몇 분 — 주간보고 전 페이지·라디오 렌더
```

`.github/workflows/dashboard-ci.yml`이 PR·푸시에서 **문법 → 섀도잉 → 문구조인 → 브랜드 →
저장소 → 엑셀 → 스모크 → 필터 → 주간보고** 순으로 자동 실행한다. 두 대시보드나 `tests/`를 건드렸으면 CI가 초록인 걸
보고 머지할 것.

> 문구 조인이 깨지면 **화면은 멀쩡히 뜨고 문구 칸만 빈다** — 스모크로는 절대 안 잡히니
> `parse_perf_bytes` · `_parse_plan_sheet` · `merge_perf_plan` · `parse_plan_gsheet`를
> 건드렸으면 `test_plan_merge.py`를 꼭 볼 것.

**스모크는 합성 데이터로 실제 렌더**한다. 페이지 목록은 앱에서 직접 읽으므로 페이지를
추가·개명해도 자동으로 따라간다. 새 하위탭·라디오를 추가하면 커버리지에 자동 포함된다.

### 실제로 났던 사고 (같은 실수 반복 금지)

| 사고 | 원인 | 막는 방법 |
|------|------|-----------|
| 11개 페이지 전부 NameError | `main()` 안에서 지역변수 이름으로 전역 헬퍼 `_s` 사용 | `check_shadowing.py` |
| 주간보고 7페이지 전부 AttributeError | `re.search(...).group()` None 미가드 | `smoke_weekly_report.py` |

---

## 전역 헬퍼 섀도잉 금지 (중요)

`main()`은 수천 줄짜리 단일 함수다. **그 안 어디서든 모듈 전역 함수와 같은 이름으로
대입하면, `main()` 전체에서 그 이름이 지역변수로 승격**된다. 중첩 함수
(`render_messages` 등)에서 부르던 전역 헬퍼는 클로저 자유변수로 묶여, 대입 줄이
실행되기 전에는 미바인딩이라 `NameError`로 죽는다.

```python
# ❌ 절대 금지 — _s 는 모듈 전역 헬퍼(NaN→빈문자열)
for _i, _y in enumerate(years):
    _s = g[g["_yr"] == _y]      # 이 한 줄이 main() 안의 _s() 호출 전부를 죽인다

# ✅ 다른 이름을 쓴다
    _sy = g[g["_yr"] == _y]
```

지역변수는 `_sy`, `_gg`, `_ser`처럼 **전역 헬퍼와 겹치지 않는 이름**으로. 헷갈리면
`python tests/check_shadowing.py`를 돌려 보면 된다.

---

## 파일 구조

```
crm_journey.html   ← 대시보드 본체 (전부 여기)
crm_journey.py     ← Streamlit 래퍼 (건드릴 일 거의 없음)
requirements.txt
```

`crm_journey.py` 내용 (참고용, 수정 불필요):
```python
import streamlit as st
import streamlit.components.v1 as components
import pathlib

st.set_page_config(
    page_title="LF Mall CRM 자동화 메시지 — 고객 여정 대시보드",
    layout="wide",
    initial_sidebar_state="collapsed"
)
st.markdown("""<style>#MainMenu, header, footer { visibility: hidden; }</style>""", unsafe_allow_html=True)

html_path = pathlib.Path(__file__).parent / "crm_journey.html"
components.html(html_path.read_text(encoding="utf-8"), height=920, scrolling=True)
```

---

## 데이터 모델

카드(메시지) 한 건의 구조:

```js
{
  id: "uuid",           // crypto.randomUUID()
  code: "A-001",        // AF코드
  label: "가입 환영",   // 캠페인명
  ch: "sms",            // 발송채널 (아래 CH 참고)
  push: false,          // PUSH 여부
  timing: "가입 D+1",   // 발송시점 (사람이 읽는 형태)
  trigger: "...",       // 트리거 조건
  offer: "...",         // 오퍼/혜택
  msg: "...",           // 메시지 내용
  status: "new",        // 상태 (아래 ST 참고)
  stageId: "stage-1",   // 소속 스테이지 ID
  view: "asis",         // "asis" | "tobe"
  ord: 0,               // 드래그앤드롭 정렬 순서 (timingNum 기반으로 seed)
}
```

**채널 (CH)**:
```js
const CH = {
  sms:     { label:'SMS',   c:'#B83A3A', bg:'#FEF2F2', icon:'📱' },
  alimtok: { label:'알림톡', c:'#A07010', bg:'#FEF8EA', icon:'💬' },
  friends: { label:'플친',  c:'#367A4C', bg:'#EEF8F1', icon:'💚' },
  email:   { label:'이메일', c:'#2E68B0', bg:'#EEF3FA', icon:'✉️' },
  inapp:   { label:'인앱',  c:'#7B5BC0', bg:'#F3EFFB', icon:'📲' },
  none:    { label:'없음',  c:'#96938C', bg:'#F2F1EE', icon:'∅'  },
};
```

**상태 (ST)**:
```js
const ST = {
  new: { label:'신규추가', c:'#367A4C', bg:'#EBF8F2', b:'#B4DEC8' },
  mod: { label:'수정개선', c:'#2E68B0', bg:'#EEF3FA', b:'#B4CCE8' },
  kep: { label:'유지',    c:'#706E68', bg:'#F2F1EE', b:'#D8D6CE' },
  del: { label:'삭제검토', c:'#B03030', bg:'#FEF1F0', b:'#E8C0BC' },
};
```

---

## 핵심 아키텍처 패턴

### 1. localStorage 영속성
```js
const STORE_KEY = 'lfmall_crm_v3';  // 데이터 구조 바뀔 때 버전 올려야 함

function saveData() { localStorage.setItem(STORE_KEY, JSON.stringify({stages, cards})); }
function loadData() {
  const raw = localStorage.getItem(STORE_KEY);
  if (raw) { const d = JSON.parse(raw); stages = d.stages; cards = d.cards; }
  else { stages = DEFAULT_STAGES; cards = DEFAULT_CARDS; }
}
```
> 데이터 구조(타이밍 형식, 채널 enum 등) 바뀌면 반드시 STORE_KEY 버전 올릴 것.
> 안 올리면 구버전 캐시가 그대로 로드됨.

### 2. XSS 방지 esc() 헬퍼
사용자 입력값을 innerHTML에 넣을 때 반드시 `esc()` 적용:
```js
function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, m =>
    ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[m])
  );
}
// 줄바꿈이 필요한 경우: esc(s).replace(/\n/g, '<br>')  ← 반드시 esc 먼저!
```

### 3. 렌더 함수 호출 순서
```js
loadData();
ensureOrder();   // ord 필드가 없는 카드에 timingNum 기반으로 seed
renderSidebar(); // 필터 사이드바
render();        // 메인 콘텐츠
```

### 4. 필터 상태 (FILTERS)
```js
let FILTERS = {
  search: '',
  channels: new Set(),
  statuses: new Set(),
  views: new Set(),
  stages: new Set(),
  pushOnly: false,
};

function cardMatches(c) {
  if (FILTERS.channels.size && !FILTERS.channels.has(c.ch)) return false;
  if (FILTERS.pushOnly && !c.push) return false;
  if (FILTERS.statuses.size && (!c.status || !FILTERS.statuses.has(c.status))) return false;
  if (FILTERS.search) {
    const q = FILTERS.search.toLowerCase();
    const hay = [c.code, c.label, c.timing, c.trigger, c.offer, c.msg].join(' ').toLowerCase();
    if (!hay.includes(q)) return false;
  }
  return true;
}
```

> **주의**: `onSearch()`는 `renderSidebar()`를 호출하지 않는다. 검색창 포커스 유지를 위해서임.
> 필터 토글 함수들은 반드시 `renderSidebar(); render();` 둘 다 호출.

### 5. timingNum — 정렬용 숫자 변환
```js
function timingNum(t) {
  if (!t) return 999;
  const s = t.toLowerCase();
  if (s.includes('실시간')) return -1;
  if (s.includes('방송') && s.includes('전')) return -0.5;
  if (s.includes('즉시')) return 0;
  const hm = s.match(/(\d+)\s*(?:시간|h)/); if (hm) return parseInt(hm[1]) / 24;
  const dm = s.match(/d\s*\+\s*(\d+)/i);    if (dm) return parseFloat(dm[1]);
  if (s.includes('배송완료')) return 0.1;
  const dom = s.match(/(\d+)\s*일/);         if (dom) return parseInt(dom[1]);
  if (s.includes('매일')) return 30;
  return 999;
}
```

### 6. 드래그앤드롭
- 카드에 `draggable="true"` + `ondragstart` / 드롭존에 `ondragover` / `ondrop`
- `ord` 필드로 순서 저장. 드롭 후 `saveData()` 호출.
- 같은 row 내 이동 및 다른 stage/row 간 이동 모두 지원.

### 7. 사이드바 토글
```js
function toggleSidebar() { document.body.classList.toggle('sidebar-collapsed'); }
```
```css
.sidebar { position: fixed; left: 0; top: 48px; width: 216px; transition: transform .22s; }
.content { margin-left: 216px; transition: margin-left .22s; }
body.sidebar-collapsed .sidebar { transform: translateX(-100%); }
body.sidebar-collapsed .content { margin-left: 0; }
```

---

## 레이아웃 구조

```
[Topbar 48px — sticky]
  ☰버튼 | LF 로고 | [sep] | 저장됨 | 내보내기 | 불러오기

[Sidebar 216px — fixed left, top:48px]
  필터 | 초기화
  검색창
  발송채널 chips
  상태 chips
  PUSH chips
  AS-IS/TO-BE chips
  스테이지 chips
  [카드 수 카운터]

[Content — margin-left:216px]
  [Stage Section × N]
    [Stage Header — 색상 편집 가능]
    [Gap Row — 골든타임 등 메모]
    [AS-IS Row]
      [card] [card] ... [+ 버튼]
    [divider]
    [TO-BE Row]
      [card] [card] ... [+ 버튼]
```

---

## 카드 뱃지 순서 (중요)

왼쪽부터: **PUSH 뱃지 → 채널 뱃지 → 상태 뱃지**

```html
<div class="card-badges">
  ${c.push ? '<span class="push-badge">🔔 PUSH</span>' : ''}
  <span class="ch-badge" style="background:${ch.bg};color:${ch.c};border-color:${ch.c}30">
    ${ch.icon} ${esc(ch.label)}
  </span>
  ${stBadge}
</div>
```

뱃지 3개 동시에 떠도 안 깨지려면:
- `flex-wrap: wrap; row-gap: 4px` — card-badges
- `white-space: nowrap; flex-shrink: 0` — 각 뱃지
- `status-badge`에 `margin-left: auto` **절대 쓰지 말 것** (두 번째 줄로 밀림)
- 카드 너비는 최소 210px 필요

---

## 타이밍 문자열 형식

사람이 읽기 쉬운 앵커 형식으로 작성:

| 형식 예시 | timingNum 결과 |
|-----------|---------------|
| 즉시 | 0 |
| 가입 D+1 | 1 |
| 구매 D+7 | 7 |
| 미방문 D+90 | 90 |
| 2시간 후 | 0.083 |
| 실시간 | -1 |
| 방송 1시간 전 | -0.5 |

---

## 흔히 생기는 버그 & 해결법

| 버그 | 원인 | 해결 |
|------|------|------|
| 카드 뱃지 줄바꿈 깨짐 | `margin-left:auto` on status badge | 제거, `flex-shrink:0` 추가 |
| HTML 특수문자 깨짐/주입 | innerHTML에 raw 삽입 | `esc()` 적용 |
| 구버전 데이터 로드 | STORE_KEY 그대로 | 버전 올리기 (v3→v4) |
| 검색 중 포커스 빠짐 | onSearch가 renderSidebar 호출 | onSearch에서 renderSidebar 제거 |
| 새 스테이지 칩 미갱신 | 색상/라벨 변경 후 사이드바 미갱신 | color picker change + label save에 `renderSidebar()` 추가 |
| import 후 칩 미갱신 | importJSON이 renderSidebar 미호출 | `importJSON` 마지막에 `renderSidebar()` 추가 |

---

## 새 기능 추가 시 체크리스트

- [ ] 사용자 입력 → innerHTML 넣는 곳에 `esc()` 적용했는가?
- [ ] 데이터 구조 변경 시 `STORE_KEY` 버전 올렸는가?
- [ ] 필터 관련 변경 시 `renderSidebar()` + `render()` 둘 다 호출하는가?
- [ ] 카드 너비 210px 유지하는가?
- [ ] 타이밍 문자열 형식이 `timingNum()`으로 파싱되는가?

---

## 효과적이었던 프롬프트 패턴

### 채널/상태 추가
```
발송채널에 [채널명]도 추가해줘. 색상은 [색상 설명]으로.
```
→ CH 객체에 항목 추가 + 사이드바 chip 자동 반영됨

### 타이밍 표시 변경
```
발송시점 표시를 [기존 형식] 대신 [새 형식]으로 변경해줘.
예시: 가입 D+1, 구매 D+7 이렇게.
```
→ 기존 카드 데이터 일괄 마이그레이션 + STORE_KEY 버전 업 필요

### 채널 일괄 변경
```
[조건]에 해당하는 카드들의 채널을 [채널]로 일괄 변경해줘.
```
→ DEFAULT_CARDS 배열 직접 수정 + STORE_KEY 버전 업

### 필터 사이드바 확장
```
사이드바 필터에 [항목] 기준으로도 필터링 추가해줘.
```
→ FILTERS 객체에 필드 추가 + renderSidebar() + cardMatches() 수정

### 드래그앤드롭
```
카드를 드래그앤드롭으로 같은 row 안에서, 그리고 다른 stage/row 간에도 이동 가능하게 해줘.
```
→ ord 필드 + draggable 속성 + 이벤트 핸들러

### 뱃지 순서/위치
```
[뱃지명] 뱃지가 [위치]에 오도록 변경해줘.
```
→ cardHTML 내 뱃지 HTML 순서 변경 + flex 속성 확인

---

## Git 워크플로

```bash
# 개발 브랜치 확인 (세션마다 지정된 claude/* 브랜치)
git branch --show-current

# 머지 전 검사 — CI와 같은 순서
python tests/check_shadowing.py && python tests/smoke_pages.py && python tests/smoke_weekly_report.py

# 변경 커밋
git add send_perf_dashboard.py
git commit -m "기능 설명"
git push -u origin "$(git branch --show-current)"

# PR 생성 (GitHub MCP)
# mcp__github__create_pull_request 또는 mcp__github__list_pull_requests로 기존 PR 확인 후
# mcp__github__update_pull_request 로 업데이트

# 머지 전 draft → ready 전환 필수
# mcp__github__update_pull_request { draft: false }
# mcp__github__merge_pull_request { merge_method: "squash" }

# 머지 후 — squash는 새 커밋을 만들므로 브랜치가 머지 전 커밋을 계속 가리킨다.
# main에 맞춰 두지 않으면 다음 작업에서 non-fast-forward로 푸시가 막힌다.
git fetch origin main && git checkout -B "$(git branch --show-current)" origin/main
git push --force-with-lease origin "$(git branch --show-current)"
```

> 다른 세션이 동시에 main에 머지하는 일이 잦다. 푸시가 거절되면 원격 브랜치 tip이
> **이미 머지된 내 커밋**인지 먼저 확인할 것(내용 diff가 비면 그렇다). 그 경우엔
> 최신 main 위로 rebase 후 `--force-with-lease`가 맞다.

---

## 구글시트 연동 (발송성과 대시보드)

### 개요
`send_perf_dashboard.py`는 구글시트를 영속 저장소로 사용한다.
`gspread` + `google-auth` 서비스 계정 방식이며, 키는 Streamlit Cloud Secrets에 TOML로 저장한다.

### Streamlit Secrets 설정 위치
Streamlit Cloud → 앱 **Settings** → **Secrets** 탭

### Secrets TOML 형식
```toml
[gcp_service_account]
type = "service_account"
project_id = "quick-doodad-397006"
private_key_id = "키ID"
private_key = "-----BEGIN PRIVATE KEY-----\nMIIEv...base64...==\n-----END PRIVATE KEY-----\n"
client_email = "googlesheet@quick-doodad-397006.iam.gserviceaccount.com"
client_id = "116048104131558945028"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/googlesheet%40quick-doodad-397006.iam.gserviceaccount.com"
universe_domain = "googleapis.com"
```

### private_key 주의사항
- Streamlit TOML이 `\n`을 깨뜨리거나 `-----BEGIN/END-----` 마커를 누락시킬 수 있음
- `_fix_pem()` 함수가 자동 복구: 마커 없이 base64만 있어도 표준 PEM으로 재구성
- `_pem_diag()` 함수가 진단: `BEGIN:있음/없음 END:있음/없음 본문:N자` 형태로 에러 원인 표시
- **키는 절대 코드나 깃에 넣지 말 것** — `.gitignore`에 `.streamlit/secrets.toml` 등록됨

### 코드 구조
```python
# 스프레드시트 제목 매핑
GS_TITLES = {"campaign": "campaign_store", "mtd": "mtd_store", "promo": "promo_store"}

# 키 로드 → PEM 복구 → 인증 → 시트 열기
def gs_open(creds_dict, spreadsheet):
    info["private_key"] = _fix_pem(info.get("private_key"))
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open(spreadsheet)  # URL/키/제목 모두 허용

# 저장/로드 — 구글시트 실패 시 로컬 CSV 폴백
_save_gs(kind, df)   # 시트에 덮어쓰기
_load_gs(kind)       # 시트에서 DataFrame 로드
```

### 구글시트 공유 설정
스프레드시트를 서비스 계정 이메일(`client_email`)에 **편집자** 권한으로 공유해야 함.

### 흔한 에러
| 에러 | 원인 | 해결 |
|------|------|------|
| Unable to load PEM file | `private_key`의 마커/줄바꿈 깨짐 | `_fix_pem()`이 자동 복구 (PR #108) |
| BEGIN:없음 END:없음 | Secrets TOML이 마커를 날림 | 코드가 자동으로 마커 추가 |
| 본문:0자 | `private_key`가 비어있음 | Secrets에 키 값 재입력 |
| 403 Forbidden | 시트 공유 안 됨 | `client_email`에 편집자 권한 공유 |

---

## Streamlit 배포 주의사항

- `height=920` — 대시보드가 잘리면 이 값 키워야 함
- `scrolling=True` — 내부 스크롤 허용
- Streamlit iframe 내부라 `window.localStorage`는 정상 작동함
- 외부 폰트(Google Fonts)는 iframe 내부에서도 로드됨

---

## 발송성과 대시보드 (send_perf_dashboard.py) 개발 규칙

### 차트 공통 규칙
- **모든 차트 레이아웃은 `base_layout()`** 을 쓴다 — 폰트(Pretendard)·호버라벨·막대 라운딩 포함.
  시계열이면 `base_layout(..., hover="x")` 로 통합 툴팁+크로스헤어를 켠다.
- 스케일이 다른 두 지표는 기본적으로 `stacked_panels()` (X축 공유 상막대/하선 패널).
  **이중축이 필요하면 `overlay_dual()` 헬퍼로만** (사용자 명시 요청으로 도입 — 축 색=시리즈 색
  매칭이 강제됨). 그 외 임의의 이중축 구현 금지, 기존 overlay_dual 사용처를 규칙 위반으로
  '수정'하지 말 것.
- **소구 속성 시리즈 색은 `tag_color(속성명)`** — TAG_BOOLS 순서에 고정. 필터가 바뀌어도 같은 속성=같은 색.
- 시간대(HHMM) 표시는 반드시 `fmt_hhmm()`(차트/표) 또는 `_hm_label()`(셀렉트박스 format_func).
  `int(hour)`를 그대로 문자열에 붙이면 1050 → "1050시" 버그가 된다.
- **시간대 해석의 단일 창구는 `norm_hhmm()`** — `hour_of_day`·`fmt_hhmm`·`hhmm_to_minutes`가
  전부 이걸 쓴다. 보정은 **`_finalize`에서 한 번만**(파싱·저장소 로드 공통). 화면마다
  따로 고치면 슬롯 집계는 원값(`8000`), 차트는 보정값(`800`)으로 갈려 같은 발송이
  두 줄로 쪼개진다.
- `use_container_width` 금지(Streamlit 제거 예정) — **`width="stretch"`** 사용.
- `st.components.v1.html` 금지(2026-06-01 제거 예정, 이미 지남) — **`st.iframe(html, height=)`** 사용.
- **선이 3개 넘는 라인차트에 모든 시리즈 데이터 레이블을 켜지 말 것.** 값이 비슷한 구간에서
  숫자가 통째로 포개져 아무것도 안 읽힌다. 기본은 **주 시리즈만** 찍고, 전체 표시는 옵션으로
  빼되 그때도 시리즈마다 `textposition`을 어긋나게 준다
  (`top center` / `bottom center` / `top right` / `bottom left` 순환 + `cliponaxis=False`).
  나머지 값은 통합 툴팁(`hover="x"`)과 아래 표가 이미 담당한다.

### 페이지 작성 규칙
- 모든 페이지 하단에 **`glossary()`** 호출 (용어 주석 접이식). 새 용어를 쓰면 glossary()에 항목 추가.
- 문구 표+원문 보기는 `render_messages(df, mcol, key)` 재사용 — 행클릭(on_select) 연동 포함.
  시간대가 중요한 화면이면 `show_hour=True`로 발송시간 칼럼을 켠다(기본은 꺼짐).
- 행클릭·셀렉트박스를 직접 만들 땐 **`guard_select(key, opts)`** 를 selectbox 직전에 호출
  (지표/필터 변경으로 세션 선택값이 옵션 밖이 되는 문제의 공용 가드).
- **`key`가 붙은 위젯은 `value=`/`default=`가 최초 1회만 반영된다** — 그 뒤론 세션값이 이긴다.
  그래서 데이터 범위에서 파생되는 위젯(기간 필터 등)은 업로드로 범위가 넓어져도 **옛 범위를
  계속 써서 새 날짜가 화면에서 통째로 사라진다**(F5로 세션을 날려야 보임). 범위가 바뀌면
  직전 전체 범위와 세션값을 비교해서, **손 안 댄 필터는 새 범위로 갱신하고 직접 좁힌 필터는
  새 경계 안으로 clamp**할 것. 세션값이 이미 있으면 `value=`를 같이 넘기지 말 것(경고).
- 캠페인/슬롯 **순위·추천은 raw 평균 정렬 금지** — `rank_adjusted(df, col, ascending)` 사용
  (비율=Jeffreys 경계, 금액=수축 평균 — 소표본 요행의 순위 점령 방지. 표시 값은 원값 유지).
- 수평 범례는 `legend_h()`, 막대 텍스트 라벨은 `bar_label(v, col, is_pct)` 재사용.
- 히트맵은 n<3 셀 마스킹 + hover에 표본수(n) 표기 관행 유지.
- **표는 `st.dataframe` 대신 `table()`** (주간보고는 `wtable()`) — 검은 헤더 스타일
  엑셀 다운로드 버튼이 자동으로 붙는다. 시트명·파일명은 `dl_name=`으로 준다.
  컬럼 컨테이너 안이면 `with col:` 로 감싸고 `table()`을 부를 것(`col.dataframe`은 안 됨).
  내보내기 본체는 `table_export.py`(두 앱 공용) — 값은 **숫자 + 엑셀 표시형식**으로
  넣는다. 문자열로 넣으면 받는 쪽에서 정렬·합계가 안 된다.
  Styler로 입힌 **글자색·굵기·셀 배경도 그대로 따라간다**(`△` 빨강 / `+` 초록).
  단 `Styler._translate()`는 **끝나면서 `ctx`를 비우므로**, 색은 `_compute()` 직후에
  먼저 떠 놓고 표시값을 나중에 읽어야 한다. 순서를 바꾸면 색만 조용히 사라진다
  (`test_table_export.py`의 `t_delta_font_color_carries_to_excel`이 잡는다).
- **업로드 문구(title/body)를 unsafe_allow_html로 렌더할 땐 반드시 `esc()`** (실데이터에
  'F&C' 같은 &·< 포함 문구가 있어 태그로 오해석됨). **AI 응답 렌더는 `safe_ai_html()`**
  (프롬프트에 사용자 문구가 들어가므로 script/이벤트핸들러 제거 — span 색상은 보존됨).

### 저장 계층 규칙
- `storage_save()`는 **성공 여부(bool)를 반환** — 반드시 `if storage_save(...):` 로 받아서
  성공 시에만 세션 갱신·성공 메시지. (읽기 실패 상태에선 덮어쓰기 방지를 위해 저장이 차단됨)
- push 데이터는 로드 직후 **`finalize_push()`** 통과 필수 (gsheets 라운드트립이 전값을
  문자열로 되돌림 — campaign의 `_finalize`와 대칭). `get_push_stats`는 **사이드바 렌더
  경로**라 여기서 죽으면 11개 페이지가 전부 다운된다 — 함수 안에도 dtype 방어가 있지만
  호출부에서 finalize를 빠뜨리지 말 것 (`test_store_layer.py`가 잡는다).
- '오늘' 계산은 `date.today()` 금지 — **`today_kst()`** 사용 (서버 UTC라 KST 새벽에 주 경계가 밀림).
  **다운로드 파일명의 날짜도 포함** — UTC면 KST 새벽 0~9시에 하루 전 날짜로 찍힌다.

### 기간 비교 규칙 (전주 · 전월 동주 · 전년 동주 · 전년 동요일)
주차·요일을 맞추는 방식이 화면마다 다르면 숫자가 어긋나 보인다. 아래로 통일한다.

| 비교 | 기준 | 왜 |
|------|------|-----|
| 전월 동주 | 기준주 **목요일**의 한 달 전이 속한 주 | 월요일 앵커를 쓰면 한 달 전 날짜가 주 꼬리에 걸려 한 주 이른 주가 잡힌다 |
| 전년 동주 | 같은 **ISO 주차** | `datetime.date.fromisocalendar(y-1, w, 1)` |
| 전년 동요일 | 같은 **ISO 주차 + 같은 요일** | 날짜를 그대로 1년 빼면 요일이 어긋나 발송 패턴 자체가 달라진다 |
| 대응 주차 없음 | 그 달 **마지막 주**로 대체하고 화면에 명시 | 주차가 달마다 4~5개라 5주차는 전월에 없는 경우가 많다 (`week_like`) |

- 전년에 그 ISO 주가 없으면(53주차) **364일 전**(=정확히 52주)으로 폴백 — 요일이 보존된다.
- 대체했으면 **어느 주와 비교했는지 화면에 반드시 표기**한다(뱃지·캡션).
- 기준주가 **부분 주**(진행 중이거나 실적 미완결)면 비교 대상도 **동요일 누계로 잘라서**
  집계한다(`_elapsed`). 안 그러면 2일치가 7일치와 맞붙어 △70%대 가짜 급락이 뜬다.
  KPI 카드뿐 아니라 **분해 차트 등 같은 화면의 모든 비교**에 같은 창을 적용할 것.

### UX 라이팅 (토스 라이팅 원칙)
- **해요체로 통일.** 합쇼체(`분석합니다`·`권장합니다`·`검토하십시오`)와 섞으면 여러 사람이
  이어붙인 인상을 준다. 두 대시보드 모두 해요체다.
- 한 문장에 한 메시지. 줄표(`—`)로 절을 이어붙이지 말고 문장을 끊는다.
- 없어도 되는 말은 뺀다: `※`, 중복 괄호 부연, "참고하세요" 같은 상투구.
- 과공손 정리: "계산해 드려요" → "계산해요". 지시조 대신 제안형: "재점검하십시오" → "점검해 보세요".
- **AI 프롬프트 문자열(`system`/`user`)은 건드리지 말 것** — 출력 서식(계층형 불릿 등)을
  지시하는 내용이라 말투를 바꾸면 생성 결과가 깨진다.
- 문구를 대량 교체할 땐 원문이 파일에 **정확히 1번** 나올 때만 치환하고 아니면 중단하게 짤 것.
  여러 줄로 쪼개진 리터럴(`"a" "b"`)은 조각 단위로 바꾸면 들여쓰기를 안 건드린다.
  (AST 스팬으로 자를 거면 `col_offset`이 **UTF-8 바이트** 오프셋이라 한글에서 어긋난다 — 바이트로 처리)

### 폰트
- Pretendard 전면 적용: `.streamlit/config.toml`의 `theme.font`+`fontFaces`(표는 캔버스 렌더링이라
  CSS로는 안 바뀌고 테마로만 적용됨) + 앱 CSS @import(이중 안전망) + `base_layout` font family.
- CSS로 폰트를 만질 때 Material 아이콘(`[data-testid="stIconMaterial"]`) 폰트는 보호할 것.

### 데이터 파싱 주의
- 실적 날짜 헤더는 변형 다양(`일자`, `일자(8자리)`, `발송일자`…) — `parse_perf_bytes`의 퍼지 매칭 유지.
- 기획 시트명은 무슬래시 형식(`1월 1주차(1229~14)`) — `_week_sheet_end_date`가 'N월' 힌트로 분해.
- 6자리 날짜 오타(`250111`)는 `_norm_date`가 `20`을 붙여 복구.
- 매칭 키는 `(date, af)` 정확 일치 — AF코드만으로 폴백하면 매주 재사용되는 코드라 엉뚱한 문구가 붙는다.
- 시간대 칸도 수기 입력이라 흔들린다. `norm_hhmm`이 **근거가 있는 두 가지만** 복구한다:
  `2400`→자정(오타가 아니라 24시 표기 — 실백업에서 `야간푸시` 카테고리에 `0`과 공존),
  `8000`→`800`(뒤 0 오타 — 같은 AF코드·같은 문구가 전부 `800` 발송).
  **`HH00` 꼴일 때만** 뒤 0을 떼서, `2460`을 `02:46` 같은 없는 시각으로 지어내지 않는다.
  복구 못 하면 NaN으로 두고 화면에 몇 건인지 띄운다 — 조용히 지어내는 것보다 낫다.
  규칙을 손대면 `test_hour_norm.py`가 본검사다 (실백업 근거를 주석에 남겨 둘 것).

### 브랜드 자동 분류 (brand2 · sales_org · brand_kind)
실적 엑셀의 `브랜드` 칸은 **담당자 수기 입력이라 캠페인명·행사명이 섞여 있다**
('위켄드세일 추가 적립', 'L+DAY 쇼핑핫타임 1차', 'DD', 'HZ'). 그대로 차원으로 쓰면
브랜드 비교가 안 되므로 **제목·내용·브랜드칸을 같이 훑어 다시 태깅**한다.

- **정본은 `data/brand_map.csv`** (brand=base 브랜드명, org=영업)와
  **`data/brand_code_map.csv`** (ADMIN브랜드코드 → base 브랜드). 둘 다 영업별 운영브랜드
  시트에서 `tools/build_brand_map.py`로 생성한다. 라인 접미어(남성·여성·세트·골프…)를
  떼어 base 브랜드로 굴린 결과다.
- **브랜드코드(`HZ`·`DM`·`JN`)는 브랜드칸 정확일치일 때만** 쓴다. 'AE'·'SD'처럼 짧고
  의미 없는 문자열이라 문구 안에서 부분 일치시키면 오탐이 쏟아진다.
- 한 코드가 서로 다른 base 브랜드를 가리키면 신뢰할 수 없어 빌드 단계에서 뺀다.
- 시트에 없는 표기(영문·사내 약어)는 **`BRAND_ALIAS`** 에 한 줄 추가. 값은 반드시
  `brand_map.csv`의 `brand`와 정확히 같아야 영업까지 붙는다.
- 매칭은 **공백·구분자 제거 + 대문자** 정규화 후 **가장 긴 이름 우선**
  ('질스튜어트 뉴욕'이 '질스튜어트'보다 먼저). 표기만 다른 같은 브랜드
  ('일꼬르소'/'일 꼬르소')는 정규화 키로 한 이름에 모은다 — 안 그러면 화면에서 두 줄로 갈린다.
- **단어 경계를 본다.** 정규화하면 공백이 사라져 원문 위치를 잃으므로 인덱스 표를 같이
  만들어, 매치 **앞에 한글이 붙어 있으면 버린다**(업'데이트'·설'프라이'즈). 뒤는 안 본다 —
  '헤지스아이템'·'헤지스품절임박'처럼 명사를 붙여 쓰는 문구가 많아 꼬리말을 막으면
  멀쩡한 매칭을 대량으로 잃는다. 뒤쪽 오탐('클리어런스'⊃클리어)은 `BRAND_STOP`이 맡는다.
- **짧은 이름이 최대 위험**이다. 9천 개 사전엔 '객'·'갭'·'고요'처럼 일반 문장에 그대로
  박히는 이름이 많다. 그래서 3글자 미만은 **자사(e-영업1·2)만 부분 일치**를 허용하고
  나머지는 브랜드칸 정확일치로 강등한다. 일반어와 겹치는 건 **`BRAND_STOP`**.
- **트리거 키워드(`TRIGGER_KW`)는 브랜드칸에서만** 본다. 문구에 적용하면 '무료배송'·
  '장바구니에 담아두신'이 죄다 트리거로 잡힌다.
- 브랜드를 못 집어도 미분류로 버리지 않는다:
  전관행사 → 제휴몰 → 트리거발송 → `{카테고리} 외` → 기타 순으로 묶는다(`brand_kind`).
- **검증은 실백업으로**: `python tools/audit_brand_classify.py <백업.zip>` — 구분 분포·
  브랜드별 실제 문구·미분류 상위를 뽑는다. 1.2만 건 기준 브랜드 특정 42%, 고유 브랜드 333개.
  일반어가 브랜드 목록에 보이면 `BRAND_STOP` 후보다.
- 사전을 고치면 **`BRANDSET_VER`가 `TAGSET_VER`에 물려** `prepare_raw` 캐시가 자동
  무효화된다. 새 사전 요소를 추가하면 **`brandset_ver()` 인자에도 넣을 것**
  (빠뜨리면 `test_brand_classify.py`의 `t_brandset_ver_changes_with_dict`가 잡는다).
- **브랜드로 분석·표시하는 곳은 전부 `brand2`를 쓴다** — 사이드바 브랜드·영업·구분
  필터, 거래액 드라이버 차트, 브랜드 랭킹, 문구 표, 담당자 표, 검색 대상.
  원본 `brand`는 문구 미매칭 표에서 '브랜드(원본)'으로만 남긴다(디버깅 단서).

### 문구 소스 — 기획 구글시트 자동 연결
평소 흐름은 **실적 엑셀만 올리고 문구는 기획 구글시트에서 가져오는 것**이다.
`_PLAN_SHEET_URL`(사이드바 코드 상단)에 시트가 고정돼 있고, 실적 파일을 올렸는데 세션에
lookup이 없으면 `_fetch_plan_gs(quiet=True)`가 **한 번 자동으로 가져온다.**

- 수동 버튼(`📥 기획 문구 다시 가져오기`)은 **시트에서 문구를 고친 뒤 갱신**하는 용도다.
  세션 캐시를 무시하고 항상 새로 읽는다.
- `parse_plan_gsheet`는 **금주 시트까지 포함하고 차주 이후는 제외**한다. `recent=N`(기본 12)로
  읽을 주차 수를 제한하니, 옛 실적을 재업로드할 땐 이 값을 늘려야 문구가 붙는다.
- 자동 가져오기가 실패해도 업로드 흐름을 막지 않는다. 실패 사유는 `plan_lookup_err`에
  담아 사이드바 경고에 함께 띄운다. **문구를 못 가져오면 그 업로드는 합치지 않는다**
  (빈 문구로 기존 누적본을 덮어쓰지 않기 위해서다).
