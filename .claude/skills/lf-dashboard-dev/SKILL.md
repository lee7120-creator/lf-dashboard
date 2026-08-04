---
name: lf-dashboard-dev
description: 발송성과(send_perf_dashboard.py)·주간보고(weekly_report.py) 대시보드 작업 절차. 페이지·하위탭·차트 추가, 화면 문구 대량 수정, 버그 재현·수정, 검사 통과 후 PR·머지·브랜치 동기화까지의 순서를 다룬다. 규칙 자체는 CLAUDE.md에 있고 이 스킬은 "어떤 순서로 하는지"다.
---

## 언제 사용하나
이 저장소의 두 Streamlit 대시보드를 손댈 때. 새 페이지·하위탭·차트를 넣거나, 화면
문구를 여러 개 한꺼번에 바꾸거나, 버그 리포트를 받아 고치거나, 머지 직전 점검할 때.

## 먼저 읽을 것
`CLAUDE.md`의 **테스트·CI** / **전역 헬퍼 섀도잉 금지** / **기간 비교 규칙** /
**UX 라이팅** 절. 규칙은 거기 있고, 여기는 그 규칙을 적용하는 **순서**만 적는다.

## 구조상 먼저 알아야 할 3가지
1. **`main()`이 수천 줄 단일 함수다.** 그 안 어디서든 모듈 전역 함수와 같은 이름으로
   대입하면 `main()` 전체에서 그 이름이 지역변수로 승격돼, 중첩 함수에서 부르던 전역
   헬퍼가 `NameError`로 죽는다. 실제로 이것 때문에 11개 페이지가 전부 다운된 적 있다.
2. **페이지 분기가 `if "키" in page:` 문자열 매칭이다.** 새 페이지명이 기존 키를
   포함하면 엉뚱한 분기로 샌다 (예: `"타이밍"`이 이미 키라 `"타이밍 비교"`는 위험).
3. **스토어 주입 방식이 앱마다 다르다.** 발송성과는 `st.session_state["camp_store"]`에
   DataFrame, 주간보고는 작업 디렉터리의 `wr_data_store.csv`.

---

## 워크플로 A — 새 페이지·하위탭 추가

### A1. 이름 충돌 확인
```bash
grep -nE '^\s+(el)?if ".+" in page:' send_perf_dashboard.py
```
새 이름이 기존 키의 부분 문자열이 아닌지 본다.
**Done when:** 기존 어느 분기 키에도 걸리지 않는 이름을 정함.

### A2. 그룹에 등록
`CAMPAIGN_GROUPS`의 해당 그룹 리스트에 이름을 추가한다. 값 리스트의 각 문자열은
아래 분기의 매칭 키를 포함해야 한다.

### A3. 분기 추가
기존 페이지 블록 사이에 `elif "새키" in page:` 로 넣는다. 페이지 맨 끝에 **`glossary()`**
를 부른다(전 페이지 공통 관행). 새 용어를 썼으면 `glossary()` 본문에도 항목 추가.

재사용할 것: `base_layout()` · `stacked_panels()` · `legend_h()` · `bar_label()` ·
`rank_adjusted()` · `guard_select()` · `render_messages(..., show_hour=)` · `fmt_hhmm()` ·
`won()` · `esc()`. 직접 만들지 말고 CLAUDE.md의 차트 규칙을 따를 것.

### A4. 렌더 확인 (검사 전에 눈으로)
```python
import sys; sys.path.insert(0,'.'); sys.path.insert(0,'tests')
from streamlit.testing.v1 import AppTest
from smoke_pages import synth_store
at = AppTest.from_file('send_perf_dashboard.py', default_timeout=300)
at.session_state['camp_store'] = synth_store(); at.run()
at.sidebar.radio[0].set_value("5. 맥락·타이밍"); at.run()
[r for r in at.radio if r.label != "페이지"][0].set_value("새 하위탭"); at.run()
print(at.exception[0].value if at.exception else "OK")
```
**엣지 케이스도 같이 태울 것** — 데이터 하루치만, 전년 데이터 없음, 값 전부 NaN.
"데이터 없음"은 크래시가 아니라 **안내 문구**로 끝나야 한다.

### A5. 검사
```bash
python tests/check_shadowing.py && python tests/smoke_pages.py
```
스모크는 페이지 목록을 앱에서 직접 읽고 하위탭까지 순회하므로, 새 탭은 **자동으로**
커버된다. 따로 테스트를 추가할 필요 없다.
**Done when:** 섀도잉 0건, 스모크 전부 OK.

---

## 워크플로 B — 화면 문구 대량 수정

손으로 하나씩 고치지 말 것. 아래 순서가 안전하다.

### B1. 대상 추출 (AST)
`st.caption/info/warning/success/error/title/help`뿐 아니라
**`st.markdown(f'<div class="appendix">…')` 안의 설명 블록도 놓치지 말 것.**
한 번 이걸 빠뜨려서 18곳이 남았다.
```python
import ast, pathlib, re
src = pathlib.Path("send_perf_dashboard.py").read_text(encoding="utf-8")
t = ast.parse(src)
docs = {ast.get_docstring(n, clean=False) for n in ast.walk(t)
        if isinstance(n, (ast.FunctionDef, ast.Module, ast.ClassDef))}
for n in ast.walk(t):
    if isinstance(n, ast.Constant) and isinstance(n.value, str) and n.value not in docs:
        ...  # 패턴 매칭
```

### B2. 치환 스크립트 (가드 필수)
- 원문이 파일에 **정확히 1번** 나올 때만 치환하고, 아니면 **즉시 중단**한다
  (부분 적용 방지). 0회면 소스에서 줄바꿈으로 쪼개진 것이니 **조각 단위**로 나눈다.
- 여러 줄 리터럴(`"a" "b"`)은 조각 단위로 바꾸면 들여쓰기를 안 건드린다.
- **AI 프롬프트(`system`/`user`)는 제외.** 출력 서식을 지시하는 내용이라 말투를 바꾸면
  생성 결과가 깨진다. 치환 후 프롬프트가 그대로인지 마커로 검증할 것.

> AST 소스 스팬으로 자를 거면 `col_offset`이 **UTF-8 바이트** 오프셋이다. 한글 소스에서
> 문자 인덱스로 자르면 리터럴이 깨진다 — `read_bytes()`로 받아 바이트로 splice할 것.

### B3. 검증
```bash
python -m compileall -q send_perf_dashboard.py weekly_report.py
python tests/check_shadowing.py && python tests/smoke_pages.py && python tests/smoke_weekly_report.py
```
바꾼 문구가 **로직에 쓰이지 않는지** 교차 검증한다(비교문·딕셔너리 키·인덱싱).
페이지 분기 문자열·컬럼명·세션 키를 건드리면 앱이 조용히 망가진다.
**Done when:** 문법·섀도잉·스모크 전부 통과 + 로직 사용 문자열 0건.

---

## 워크플로 C — 버그 수정

### C1. 재현부터 (고치기 전에)
리포트를 받으면 **먼저 실패하는 테스트/스크립트를 만든다.** 못 재현하면 못 고친 것이다.
앱 레벨 재현이 가장 확실하다(AppTest로 해당 페이지를 실제로 띄워 예외를 받는다).
**Done when:** 리포트와 같은 예외/오답을 손에 넣음.

### C2. 최소 수정
동작이 바뀌는 범위를 좁힌다. 완결 주·정상 입력에서 **기존과 완전히 동일**한지
동치성으로 확인할 것(예: `_elapsed=6`이면 기존 슬라이스와 같음).

### C3. 테스트가 실제로 잡는지 확인
고친 뒤 **버그를 다시 주입해서** 테스트가 실패(exit 1)하는지 보고 원복한다.
이 왕복을 안 하면 "통과했다"가 아무 의미 없다.
**Done when:** 주입 시 FAIL → 원복 시 PASS 확인.

### C4. 커버리지 공백이 원인이면 같이 메운다
그 버그가 왜 안 걸렸는지 본다. 테스트가 안 도는 영역이었다면 **테스트도 같이 추가**한다.
(주간보고 렌더 테스트가 아예 없어서 앱 전면 다운을 놓친 적 있다.)

---

## 머지 절차

```bash
python tests/check_shadowing.py && python tests/smoke_pages.py && python tests/smoke_weekly_report.py
git add -A && git commit -m "..."
git fetch origin main && git merge-base --is-ancestor origin/main HEAD && echo "main 포함 OK"
git push -u origin "$(git branch --show-current)"
```
1. **draft PR로 생성** (`mcp__github__create_pull_request` `draft: true`)
2. **CI가 초록인 걸 확인** — `mcp__github__actions_get` `method: get_workflow_run`,
   `minimal_output: true`. `conclusion: success` 확인 전엔 머지하지 말 것.
3. draft 해제 → squash 머지
4. **브랜치를 main에 동기화** (안 하면 다음 작업에서 푸시가 막힌다)
   ```bash
   git fetch origin main && git checkout -B "$(git branch --show-current)" origin/main
   git push --force-with-lease origin "$(git branch --show-current)"
   ```

---

## 자주 밟는 지뢰

| 증상 | 원인 | 확인 |
|------|------|------|
| 전 페이지 `NameError: free variable` | 지역변수가 전역 헬퍼 이름을 가림 | `python tests/check_shadowing.py` |
| `AttributeError: 'NoneType' … 'group'` | `re.search(...).group()` 무가드 | 매치 결과를 항상 `if m:`로 가드 |
| 푸시가 non-fast-forward로 거절 | squash 머지라 원격 브랜치가 머지 전 커밋을 가리킴 | 내용 diff가 비면 `--force-with-lease`가 맞음 |
| 다른 세션이 main을 먼저 옮김 | 동시 작업 | 최신 main 위로 rebase 후 재검증 |
| 문구를 바꿨는데 앱이 이상 | 페이지 분기 문자열·컬럼명·세션 키를 건드림 | B3의 로직 사용 교차 검증 |
| `st.iframe` height 에러 | `height=0`을 거부함 | 스크립트 주입용이면 `height=1` |
| 다운로드 파일명이 하루 전 | 서버가 UTC | `today_kst()` 사용 |
| 부분 주에서 △70% 가짜 급락 | 기준주 2일치를 전주 7일치와 비교 | 같은 화면의 **모든** 비교에 `_elapsed` 적용 |
