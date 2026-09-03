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
python tests/test_copy_effect.py          # 몇 초 — 기획시트 W·X열(문구변경·잔여모수)·발송유형 정규화
python tests/smoke_pages.py               # 몇 분 — 발송성과 전 페이지·하위탭 렌더
python tests/test_filter_follows_upload.py # 몇 분 — 업로드 후 기간 필터·차트 레이블
python tests/test_send_volume_band.py     # 몇 분 — 발송량 최적 구간(MTD·앱푸시 결합 회귀)
python tests/test_prio_reduction.py       # 몇 분 — 우선순위 정규화·발송 감축 효과(구성 효과 방어)
python tests/test_site_metrics.py         # 몇 분 — 사이트 회원UV·거래액 파싱·주간보고 MTD 행
python tests/test_target_ctr.py           # 몇 분 — 목표 CTR 역산(달성 가능성 판정)
python tests/test_inflow_funnel.py        # 몇 분 — 유입 퍼널(유니크유입 파생·단계 비율)
python tests/test_orgcat.py               # 몇 분 — 조직×카테고리 파싱·저장·화면(구분06×구분07)
python tests/smoke_weekly_report.py       # 몇 분 — 주간보고 전 페이지·라디오 렌더
```

`.github/workflows/dashboard-ci.yml`이 PR·푸시에서 **문법 → 섀도잉 → 문구조인 → 브랜드 →
저장소 → 엑셀 → 스모크 → 필터 → 최적구간 → 감축효과 → 사이트지표 → 목표CTR → 유입퍼널 → 조직카테고리 → 주간보고** 순으로 자동 실행한다. 두 대시보드나 `tests/`를 건드렸으면 CI가 초록인 걸
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
- **월 누계(MTD)의 마감일도 `ref_ws + _elapsed`다** — `ref_we`(기준주 일요일)를 쓰면
  아직 오지 않은 날을 마감으로 잡는다. 월말이 월요일이면(2025-03-31·2026-08-31 등)
  그 일요일이 다음 달이라 **'당월'이 한 주 일찍 넘어가고 표 전체가 '–'로 빈다**.
  완결 주에선 `ref_ws+6 == ref_we`라 동작이 같아서 평소엔 티가 안 난다
  (`test_site_metrics.py`의 `t_mtd_window_clamps_to_partial_week` — 재현 날짜를 고정해 뒀다).

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
  **실백업 근거**: AF코드 1개당 평균 36개 날짜에 재사용되고(AP80은 84개 날짜), 89%가 2개 이상
  날짜에 쓰인다. 날짜를 빼면 조인이 통째로 망가진다.
- **lookup 값은 `(제목, 내용, 문구변경, 잔여모수)` 4-튜플이다.** `(title, body)`로 고정해서
  풀면 터진다 — 주간보고 「금주 집행 내용 요약」 자동 생성이 W·X열이 붙은 뒤부터 운영에서
  `ValueError`로 죽었다. 소비처는 `merge_perf_plan`처럼 **길이를 보고 관용적으로** 읽을 것.
  세션에 lookup이 있어야 밟는 경로라 스모크로는 안 잡힌다
  (`test_plan_merge.py`의 `t_exec_note_reads_four_tuple_lookup`).
- 기획시트 **오른쪽 별도 블록**에 `문구 변경 여부`(W)·`잔여모수 추가 여부`(X)가 있다.
  헤더 표기가 흔들려 이름 **부분 일치**로 찾는다(`문구`+`변경` / `잔여모수`).
  값도 수기라 `문구변경`·`문구 변경`이 섞인다 — `_plan_flag()`가 정규화한다.
  **빈칸은 '검토했지만 그대로 감'이라 대조군**이다. True로 잘못 읽으면 비교가 통째로 무너진다.
  같은 `(date, af)`가 두 번 나오면(원발송 + 잔여모수) 체크는 **OR로 합친다** — 덮어쓰면
  뒤 행의 빈칸이 앞 행의 체크를 지운다.
- 발송유형(`stype`)은 실적 파일에서 오지만 표기가 흔들린다(2026-08부터 `컨틴전시 A/B` →
  `컨틴`, `시그니처`/`시그니쳐`). **`norm_stype()`을 거칠 것** — 안 그러면 같은 유형이
  차트에서 두세 줄로 갈린다. `우수발송 1~8`은 기본으로 묶고, 번호가 필요하면 `group=False`.

### 「11. 개선 효과 검증」 — 세 탭은 질문이 서로 다르다
| 탭 | 답하는 질문 | 데이터 |
|----|------------|--------|
| 문구 개선 효과 | 우리측이 다듬은 문구가 나았나 | 기획시트 W열 |
| 잔여모수 추가 효과 | 잔여모수를 더 태운 게 나았나 | 기획시트 X열 |
| 컨틴 구좌 효율 | **이번 주 컨틴 구좌에 뭘 넣을까** | 발송유형(`컨틴`) + 평일 16시 이력 |

- **대조군은 `POLICY_CHANGE_DATE`(2026-08-01) 이후로 한정하는 게 기본이다.** 그 전에는
  구좌가 10개였고 타겟팅 없는 남은발송이 20% 가까이 나갔다. 운영 자체가 달라서 섞으면
  '문구를 고쳐서'가 아니라 '운영이 달라서' 생긴 차이를 문구 효과로 읽게 된다.
  「전체 기간」으로 넓혀 볼 수는 있게 두되 기본값은 정책 변경 이후다.
- 앞의 둘은 `_eff_block()`(단순 비교 + 층화 보정 + Welch)을 공유한다. 헬퍼는 **페이지 분기
  앞**에 둔다 — `welch`·`sig_label`·`table`이 `main()` 지역이라 모듈로 못 뺀다.
- **총량(거래액·발송)을 비교 카드에 넣지 말 것.** 두 군의 표본 크기 차이가 그대로 찍혀
  '변경군이 45% 적다'로 읽힌다. 전부 비율 지표(CTR·주문CR·RPS·객단가)로 본다.
- **컨틴 구좌는 회고가 아니라 판단 지원 화면이다.** 2026-08 신설 16시 구좌라 자체 실적이
  거의 없어서, 기본 기준 표본을 **평일 16시 슬롯 전체 이력**으로 두고 컨틴 건이 쌓이면
  좁혀 보게 했다. 2025년 `컨틴전시 A/B`(주말 17시 308건)는 **성격이 다른 구좌**라
  **화면 전체를 `POLICY_CHANGE_DATE` 이후로 잘라** 아예 들어오지 않게 한다(기간 옵션도
  그 안에서만 좁힌다). 뺀 건수는 캡션에 밝힌다.
- **판단 기준은 영업 세일즈 푸시와의 비교다.** 카테고리 순위는 '컨틴에 뭘 넣을까'엔 답해도
  '컨틴이 나은가'엔 답하지 못한다. 영업 세일즈 푸시 = **BPU 칸이 `1BPU`~`4BPU`** 인 발송
  (컨틴 제외) — 마케팅·편성·브랜드컨텐츠 요청분은 성격이 달라 뺀다. 컨틴은 16시 고정이라
  시간대를 안 맞추면 '컨틴이 나은지'가 아니라 '16시가 나은지'를 보게 되므로
  **「같은 슬롯(평일 16시)으로 맞춰 비교」** 옵션을 둔다.
- 순위는 `rank_adjusted`로 정렬한다(표시 값은 원값). 소표본 요행이 1등을 먹는 걸 막는다.
- **추이 차트는 소표본 주·부분 월을 뺀다.** 1~2건짜리 주는 CTR이 0%·5%로 튀어 추세를
  가리고, 진행 중인 달은 비중이 흔들린다. 양쪽 군 `min_n`건 이상인 주만 그리고, 뺀
  개수를 캡션에 밝힌다. 빠진 주는 NaN으로 남겨 **`connectgaps=False`로 선을 끊을 것** —
  안 그러면 1년치 공백을 직선 하나로 이어 그 사이에 데이터가 있는 것처럼 보인다.
  **남은 주가 3개 미만이면 선차트를 아예 그리지 말고 표로만 보여줄 것.** 점이 1~2개면
  plotly가 x축을 밀리초(`23:59:59.999`)까지 쪼개 눈금을 만들어 더 못 읽는다.
  3개 이상일 때도 `xaxis`에 `type="date"`·`tickformat`·`dtick`을 명시한다.
- 컨틴 탭의 **판단 지표는 선택식**(`METRIC_OPTS` 기반)이다. 단일 지표로는 판단이 안 된다 —
  실제로 CTR 1위는 `제휴·스포츠`, RPS 1위는 `마케팅·통합`으로 갈린다. 카테고리 랭킹·소구
  리프트·추천 문장이 전부 선택 지표를 따른다.
- **총 거래액은 이 탭의 지표에서 뺐다.** 발송량이 큰 카테고리가 무조건 위로 와서 구좌 배정
  판단에 못 쓴다. 대신 `거래액(캠페인 건당)`을 넣었다(발송 1건당은 RPS가 이미 담당).
  이 지표는 `RANK_ND`에 없어 `rank_adjusted`가 원값을 돌려주므로, 정렬만 같은 수축 공식을
  직접 적용한다.
- 시간대 칸도 수기 입력이라 흔들린다. `norm_hhmm`이 **근거가 있는 두 가지만** 복구한다:
  `2400`→자정(오타가 아니라 24시 표기 — 실백업에서 `야간푸시` 카테고리에 `0`과 공존),
  `8000`→`800`(뒤 0 오타 — 같은 AF코드·같은 문구가 전부 `800` 발송).
  **`HH00` 꼴일 때만** 뒤 0을 떼서, `2460`을 `02:46` 같은 없는 시각으로 지어내지 않는다.
  복구 못 하면 NaN으로 두고 화면에 몇 건인지 띄운다 — 조용히 지어내는 것보다 낫다.
  규칙을 손대면 `test_hour_norm.py`가 본검사다 (실백업 근거를 주석에 남겨 둘 것).

### 「6. 효율·피로도 › 발송량 최적 구간」 — 늘려도 되나에 답하는 화면
2026-08에 영업 구좌를 10개→5개로 줄였다가, 남은모수 추가·컨틴 배정으로 **발송량을 다시
올려보는 중**이다. 그 판단을 지원하는 화면이라 회고가 아니라 결정 지원이다.

- 세 블록의 질문이 다르다: **① 얼마나 올릴까**(총량↔효율 트레이드오프) · **② 올려도 되나**
  (이탈 손익분기) · **③ 어디에 태울까**(발송유형별 건당 효율). 순서를 섞지 말 것.
- **회귀는 전부 요일 잔차(`_dow_residual`)로 한다.** 원값으로 돌리면 '주말엔 적게 보내고
  매출도 적다'가 발송 효과로 잡혀 상관이 부풀려진다.
- **이탈 기울기가 음수거나 p≥0.05면 손익분기를 계산하지 말 것.** 분모가 0에 가까우면
  '이탈 1명당 3억원'처럼 아무 의미 없는 큰 수가 나온다. 그 경우엔 '신호가 안 보인다'고
  말하고 끝낸다 (`test_send_volume_band.py`의 `t_flat_churn_reports_no_signal`).
- 손익분기는 **하루치 거래액 증분 ÷ 하루치 이탈 증분**이라 LTV와 단위가 다르다. 화면에
  **상한선으로만 쓰라고 명시**한다.
- ③의 발송유형은 반드시 **`norm_stype()`** 을 거친다. 안 거치면 `컨틴전시 A`/`컨틴전시 B`가
  따로 놀아 컨틴 구좌 효율이 반토막으로 보인다. 대조군 기본값은 **`POLICY_CHANGE_DATE` 이후**
  (11번 페이지와 같은 이유).
- 관측 데이터라 **인과가 아니라 상관**이라는 경고를 화면 맨 위에 둔다. 발송을 늘린 날은
  대개 행사일이라 행사 효과가 섞인다.
- 이 탭은 MTD가 있어야 열리고 ②는 앱푸시까지 있어야 계산된다 — **스모크는 캠페인 데이터만
  넣으므로 이 경로를 안 밟는다.** 손대면 `test_send_volume_band.py`가 본검사다.

### 「6. 효율·피로도 › 유입 퍼널」 — 어느 단계에서 새는지
전사 MTD 파일엔 유니크유입·총유입·구매고객수가 다 들어 있는데 화면엔 CTR·RPS만 나왔다.
발송 고객 → 유입 고객 → 구매 고객으로 끊어 보면 **발송량을 늘릴 일인지 문구를 고칠
일인지**가 갈린다.

- **단계 비율은 일별 비율의 평균이 아니라 기간 합계끼리 나눈 값이다.** 발송량이 날마다
  크게 흔들려서 일별 비율을 평균 내면 소량 발송일이 과대 대표된다 (실제로 5만↔50만이
  섞이면 2.7% vs 6.0%로 갈린다 — `test_inflow_funnel.py`의 `t_step_rates_use_period_sums`).
- 비교 창은 **선택 창 길이만큼 뒤로 민 직전 같은 기간**이다. 기본값이 전체 기간이라 앞이
  비는 경우가 대부분이라, 그때는 **선택 구간을 반으로 잘라 전반↔후반**으로 대체하고
  **무엇과 비교했는지 캡션에 밝힌다**(컬럼명도 `직전/선택 기간` ↔ `전반/후반`으로 바뀐다).
- 파생 3종은 전부 **유니크유입이 분모**다: `uniq_cr`(구매고객÷유니크유입) ·
  `rev_per_uniq`(거래액÷유니크유입) · `inflow_dup`(총유입÷유니크유입).
  기존 `purchaseRate`는 분모가 **발송 고객수**라 답하는 질문이 다르다 — 섞어 쓰지 말 것.
  0 분모는 NaN으로 두고 화면엔 `–`로 찍는다(inf를 그대로 뿌리면 카드가 깨진다).
- **`compute_mtd`에 파생을 추가하면 `MTDSET_VER`도 같이 올릴 것.** `cached_compute_mtd`는
  `@st.cache_data`라 자기 소스만 캐시 키로 본다 — `prepare_raw`/`TAGSET_VER`와 같은 함정이고,
  빠뜨리면 옛 프레임이 남아 화면이 통째로 빈다 (`t_cache_marker_mentions_uniq`).
- 이 탭은 MTD가 있어야 열린다 — **스모크는 캠페인 데이터만 넣으므로 본문을 안 밟는다.**
  손대면 `test_inflow_funnel.py`가 본검사다.

### 우선순위(prio) — 0순위는 1순위에 합친다
실적 엑셀의 `우선순위`는 같은 시간대에 몇 번째로 나갔는지다. **`0`은 실백업 1.2만 건에서
2건뿐**이라(1순위 3,539건) 따로 두면 표본 2짜리 칸이 생겨 순번↔효율 관계를 왜곡한다.
가장 먼저 나간다는 뜻이 같으니 합친다.

- 정규화는 **`norm_prio()`** 한 곳에서. `_finalize`가 **`prio_g`** 를 파생으로 붙이고,
  **분석·필터·표시는 전부 `prio_g`** 를 쓴다(사이드바 우선순위 필터·9번 페이지·감축 효과 탭).
  원본 `prio`는 그대로 남긴다 — 다운로드·디버깅용.
- 선택지에서 값이 사라지면 **세션에 남은 옛 선택('0')이 multiselect를 예외로 죽인다** —
  `guard_multi(key, opts)`를 위젯 직전에 부를 것(selectbox의 `guard_select`와 짝).
- **`_finalize`에 파생을 추가하면 `TAGSET_VER` 표식도 같이 올릴 것.** `prepare_raw`는
  `@st.cache_data`라 `_finalize` 변경을 모른다 — 옛 프레임이 남으면 그 파생을 쓰는 화면이
  통째로 빈다. 증상이 '데이터가 없다'로만 보여서 원인이 안 드러난다(`hour:norm1`·`prio:g1`이
  그래서 붙어 있다). 표식을 잊어도 죽지 않게 **캐시 밖에서 한 번 복구**하고
  (`prio_series()`처럼) 소비처는 헬퍼를 거칠 것.

### 「11. 개선 효과 검증 › 발송 감축 효과」 — 구성 변화를 효과로 읽지 않기
'남은발송을 빼서 피로도를 낮췄으니 앞 순위 지표가 올랐을 것'이라는 가설을 전후로 본다.
여기서 조용히 틀리는 방식이 정해져 있다.

- **남은발송은 순위별로 몰려 있다.** 실백업(2026)에서 1순위 발송의 0.9%뿐인데 2·3·4순위는
  50~56%다. 그래서 그냥 전후를 맞대면 **빼기만 해도 그 순위 평균이 저절로 올라간다**.
  기본값은 **양쪽 구간에서 남은발송을 다 뺀** 같은 성격끼리의 비교다(`남은발송 빼고 비교`).
  한쪽만 빼면 비교가 통째로 기울어진다 (`test_prio_reduction.py`).
- **전후 창은 같은 길이로 자른다.** 긴 쪽이 계절·행사를 더 많이 먹는다.
- **8월 정책 변경(구좌 10→5)이 '후' 구간에 섞이면** 남은발송 제거 효과와 구분이 안 된다 —
  기본은 `POLICY_CHANGE_DATE` 전까지만 본다.
- **모수 크기도 같이 본다.** 구좌를 줄이면 한 번에 더 많이 보내게 되고(실백업에서 1~3순위
  캠페인당 발송이 15~17만 → 18~20만), 리스트 아래쪽까지 내려가 CTR이 떨어지는 게 정상이다.
- 판정 문장은 **유의성까지 봐야 '지지'** 라고 말한다. 방향만 보고 결론 내리면 우연을
  성과로 읽는다.

### 조직×카테고리 실적 (MICRO 대시보드 export) — 주간보고 「08」
`구분06(BPU) × 구분07(카테고리)` 두 축짜리 export다. 기존 store는 `segment` 한 칸뿐이라
여기 밀어 넣으면 두 축이 뭉개진다 — **별도 store**(`wr_orgcat_store.csv`,
`org`·`cat` 두 칸)에 쌓는다.

- **헤더 구성이 단위마다 다르다.** 월·주는 `연도 / LFMS / 기간 / (구분+마감)` 4행이고,
  일은 마감 행이 없어 3행이다. 행 번호를 박으면 한 단위만 조용히 깨지니 **내용으로 찾을 것**
  (구분06이 있는 행·연도 행·Y/N만 있는 행·기간 라벨 행).
- **LFMS 포함여부(Y/N)는 모집단이 다른 축이라 키에 넣는다.** 빼면 같은 기간·같은 조직 값이
  서로를 조용히 덮어쓴다. 화면 선택지는 **고른 집계 단위 안에서** 뽑을 것 — 단위마다 받아 온
  export가 달라(일별만 Y) 전역 목록을 쓰면 '일'로 바꿨을 때 이전 선택 N이 남아 데이터가
  있는데도 빈 화면이 된다 (`test_orgcat.py`의 `t_lfms_options_follow_granularity`).
- **연도·조직·LFMS는 전부 병합셀**이라 오른쪽/아래로 이어받아야 한다. 한 파일에 2개년이 온다.
- 카테고리 `-`는 카테고리 구분이 없는 조직(SPACE-R 등)의 자리표시라 `*TOTAL`과 겹친다 — **버린다**.
- 지표는 마스터와 같은 이름으로 맞춘다(`일평균거래액`→`첫구매 거래액`). 그래야 `fmt_value`·
  `PCT_METRICS`가 그대로 통한다. `상품CR`은 비율이라 `PCT_METRICS`에 넣었다.
- **마스터 파서가 집어삼키지 않게 라우팅을 먼저 가른다**(`route_push`가 조직×카테고리면
  둘 다 None). 인식 목록에도 따로 찍어 '왜 안 올라가지'를 없앤다.
- 「08」은 사이드바 기준 기간에 기대지 않고 **자체 기간 선택**을 쓴다 — 일 단위까지 오고
  커버리지가 마스터와 달라서다. 덕분에 마스터가 아직 없어도 이 화면만은 열린다.
- `*TOTAL`은 하위 항목의 합이 아니라 **파일이 준 값** 그대로다(사이트 지표와 같은 이유).

**진단 화면으로서의 규칙 — 표만 늘어놓지 않는다**

- **가산 지표는 거래액 하나뿐이다.** 실파일에서 조직 합이 전체와 오차 0.00%로 맞는다.
  고객수(+2.4%)·상품UV(+33%)는 유니크 값이라 같은 사람이 여러 조직에 잡혀 합이 전체를
  넘는다 — **기여도 분해를 붙이면 거짓말이 된다**(`ORGCAT_ADDITIVE`).
  비가산 지표를 고르면 기여도 칼럼을 빼고 왜 없는지 화면에 밝힌다.
- **기여도(%p)의 분모는 전년 전체**다. 그래야 다 더해서 전체 전년비와 같아진다
  (`t_contributions_sum_to_total_yoy`). 개별 증감을 전체 증감으로 나누면 상쇄 구간에서
  300%·800%가 튀어나온다 — 요약 문장에선 3배를 넘으면 아예 생략한다.
- **`거래액 = 상품UV × 상품CR × 객단가`가 원본에서 정확히 성립한다**(모든 조직·카테고리,
  오차 0.000%). 이 항등식이 '왜'의 축이다: 증감을 **LMDI(로그평균 디비지아)** 로 쪼개면
  세 기여액의 합이 실제 증감과 원 단위까지 같고, 대입 순서에 따라 답이 달라지지 않는다
  (`factor_split`). **0·음수·결측이 하나라도 있으면 로그가 정의되지 않으니 None을 돌려
  숫자를 지어내지 말 것** (`t_factor_split_refuses_nonpositive`).
- 화면은 **전체 → 조직 → 카테고리**로 파고든다. 레벨이 내려가면 ①의 표가 그 아래 단계로
  바뀌고, ②의 요인 분해는 늘 **지금 보고 있는 대상**을 설명한다.
- **조사는 `josa()`로 붙인다.** '객단가이 가장'·'조직는'이 실제로 났다. 발송성과에서
  겪은 '시그니처이 가장 나아요'와 같은 버그다. 검사할 땐 마크다운 원문이 `**객단가**이`라
  평문으로 만든 뒤 봐야 걸린다 (`_plain()`).
- 워터폴은 **상위 8개만 세우고 나머지는 '기타 N곳'** 으로 합친다. 0원짜리 조직까지 다
  세우면 눈금이 뭉개져 정작 큰 항목이 안 읽힌다.

### 사이트 회원UV · 거래액 (전사 채널×디바이스 리포트)
CRM 발송 실적이 아니라 **사이트 전체 유입·매출**이다. 태블로 export 두 벌(회원UV·거래액)이
같은 모양으로 오고, 「12. 회원UV·거래액」과 주간보고 MTD 표에서 쓴다.

- 구조: 0행=연 · 1행=월 · 2행=일(연·월은 병합셀이라 **ffill 필수**), A열=블록, B열=채널,
  C열=디바이스, D열부터 날짜. **`기준` 블록만 읽는다** — 나머지(전주비·전년비…)는 증감률이라
  실측치로 읽으면 값이 ±0.1 근처로 깔린다.
- **`Total` 행은 하위 항목의 합이 아니다.** 유니크 방문자라 채널·디바이스를 더하면 중복이
  섞인다. 반드시 파일이 준 `Total` 행을 그대로 쓸 것.
- **두 리포트는 구조가 완전히 같다.** 회원UV는 사람 수라 정수만, 거래액은 환산값이라 소수가
  섞인다(실파일 0% vs 85%) — `site_metric_kind()`가 이걸로 가른다. 파일명에 '거래액·매출'이
  있으면 그게 우선. 어느 파일을 무엇으로 읽었는지 **사이드바에 찍어** 오판을 눈으로 잡는다.
- 원본 단위는 **회원UV=명 · 거래액=천원**이고 화면은 **천명 · 백만원**이라 둘 다 `SITE_DIV`
  (=1000)로 나눈다. 근거: Total 일평균 UV 25.2만명 · 거래액 9.8억원 → 방문 1명당 3,880원.
- 저장은 `(date, ch, dev)` 키에 **칼럼별 coalesce**(`merge_site_store`). 행 교체로 하면
  회원UV만 다시 올렸을 때 이미 쌓인 거래액이 NaN으로 날아간다.
- 주·월 집계의 회원UV는 **일별 값의 평균**이지 기간 순방문자가 아니다. 화면에 명시할 것.
  진행 중인 주·달은 합계로 보면 뚝 떨어져 보이므로 기본값은 일평균이다.
- **기간 자르기는 집계 *전* 일별 단계에서** 한다. 집계 후에 자르면 라벨이 창 밖으로 나가는
  마지막 부분 주가 통째로 사라져 '최근 데이터가 없다'로 보인다.
- **최근 N개로 잘라 그리지 말 것.** 2년치를 올려도 화면엔 120일만 나와 '데이터가 잘렸다'로
  읽힌다. 기간은 `date_input`으로 노출한다(사이드바 기간 필터와 같은 세션 갱신 처리 필요 —
  새 파일로 범위가 넓어지면 손 안 댄 필터는 따라가게. 비교 기준은 **직전 기본값**이다).
  기본 시작일은 **`SITE_DEFAULT_FROM`(2025-01-01)**, 고를 수 있는 하한은 데이터 전체라
  2024년도 시작일만 당기면 보인다 — 기본을 좁히는 것과 데이터를 자르는 건 다르다.
- 디바이스 비중 파이는 **연 단위(올해 vs 전년)** 로 본다. 기간을 반년으로 잡아도 궁금한 건
  연 비교라서다. 전년은 **같은 날짜까지 잘라(YTD)** 계절성이 섞이지 않게 한다.
- 전년 비교선은 **시차 창을 따로 집계해 x를 앞으로 되돌린다**. 현재 축에 `reindex`로 끌어오면
  부분 주에서 한 칸씩 어긋난다. 시차는 일·주 **364일**(52주, 요일 보존), 월 **12개월**.
- 주간보고 MTD 표에 넣는 건 **앱푸시 = `PUSH` 채널 × `App` 디바이스 교집합** 한 줄이다.
  둘을 따로 넣으면 '앱으로 들어온 광고 유입'과 'PC로 받은 푸시'가 섞여 앱푸시 성과가 아니다.
  **데이터가 없으면 통째로 뺀다**(빈 행이 남으면 0인지 없는 건지 구분이 안 된다).
  값은 누계가 아니라 **그 기간의 일평균**이다 — 발송 실적과 분모가 달라 합계로 섞을 수 없다.

### 「4. 성과 진단 › 전환·AOV 진단 ③ 목표 CTR 역산」
`거래액 = 발송 × CTR × 주문CR × 객단가`라, 나머지를 지금 수준으로 붙들면 목표 거래액에
필요한 CTR이 나눗셈 한 번으로 나온다. 위험한 건 **그 숫자가 달성 가능한 수준인지 말하지
않는 것**이다 — 그대로 실행 목표가 된다.

- 목표 CTR을 **최근 52주 주별 CTR 분포**와 대조해 판정한다: 최고치 초과면 「CTR만으로는 못
  메워요」, 상위 10% 안이면 「도달한 적은 있지만 유지가 어렵다」, 중앙값 부근이면 「노려볼
  만하다」. 분포가 8주 미만이면 판정하지 않는다.
- **한 레버만 움직일 때의 필요 변화율은 넷이 모두 같다**(곱이라서). 표에 그렇게 적어 두고,
  '어디를 움직이는 게 현실적인가'로 고르게 한다.
- CTR 단독이 무리일 때를 위해 **다른 레버를 기준 수준으로 되돌렸을 때 필요한 CTR**을 같이
  보여 준다. 실백업에서는 발송량만 전년 수준으로 되돌리면 CTR·주문CR·객단가가 지금
  그대로여도 거래액이 전년의 100%가 된다 — 격차의 사실상 전부가 발송량이었다.
- 이 블록만 **사이드바 기간 필터 대신 자체 구간 선택**을 쓴다(전년 같은 기간을 같이 봐야
  해서). 속성·매칭·최소발송 필터는 그대로 적용한다.
- **비교 구간 시차는 창 길이만큼 민다.** 28일로 고정해 두면 8주·13주 창에서 비교 구간이
  현재 창을 파고들어 같은 발송을 양쪽이 나눠 갖는다 (`test_target_ctr.py`의
  `t_prior_window_does_not_overlap_current`). 전년만 364일(52주라 요일 보존) 고정이다.
- 네 레버가 서로 독립이라는 전제라, CTR을 올리면 유입 품질이 묽어져 주문CR이 같이 내려갈 수
  있다. 화면에 **목표치는 상한이 아니라 출발점**이라고 밝힌다.

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
- **시트명에서 기간을 못 읽는 주차(`undated`)도 반드시 읽을 것.** 예전엔 `recent`가
  설정된 기본 경로에서 통째로 빠져, 제목에 `(810~816)`이 아직 안 붙은 **이번 주 시트가
  조용히 유실**됐다. 화면엔 '매칭률 0%'로만 떠서 원인이 안 보인다
  (`test_plan_merge.py`의 `t_undated_sheet_survives_recent_cap`).
- 읽은 주차 목록은 `plan_lookup_sheets`에 남기고 **최신 주차명을 화면에 띄운다** —
  개수만 보여주면 '이번 주 시트를 읽긴 했나'를 확인할 방법이 없다.
- **`문구 매칭된 것만`(기본 켜짐)이 발송일을 통째로 지우면 본문 상단에 알리고
  「필터 끄기」 버튼을 같이 준다.** 안 그러면 증상이 '기준일 목록에 날짜가 없다'로
  나타나 업로드 실패로 읽힌다. 일부만 빠진 건 알리지 않는다(상시 17% 남짓이라 소음).
  위젯 세션값은 **`on_click` 콜백에서만** 바꿀 것 — 위젯 생성 뒤 직접 대입하면 예외다
  (`test_filter_follows_upload.py`의 `t_unmatched_days_are_announced`).
