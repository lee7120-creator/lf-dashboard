"""리런 비용 측정 — 두 대시보드 공용.

**추측으로 최적화하지 말라는 CLAUDE.md 규칙의 실행 도구다.** 예전 측정
(발송성과 19.2초→7.6초 · 주간보고 14.60초→6.20초)은 일회성 스크립트로 해서
방법만 문서에 남고 도구는 사라졌다. 다음 사람이 같은 함정을 다시 밟지 않게
여기 둔다.

함정 둘 — 둘 다 실제로 밟았다.

1. **`ScriptCache`를 공유시키지 않으면 측정이 통째로 덮인다.** `AppTest`는
   `at.run()`마다 스크립트를 새로 컴파일하는데(운영은 1회) 발송성과가 1.2만 줄이라
   그것만 14초다. 페이지별 차이가 그 밑에 묻혀 안 보인다.
2. **프로파일은 «스크립트 스레드»를 떠야 한다.** 메인 스레드에 `cProfile`을 걸면
   스크립트는 다른 스레드에서 도는 탓에 `time.sleep`만 보인다.

쓰는 법:

    python tools/perf_rerun.py                      # 발송성과 전 페이지
    python tools/perf_rerun.py --app weekly         # 주간보고 전 페이지
    python tools/perf_rerun.py --rows 40            # 데이터 규모를 키워서
    python tools/perf_rerun.py --page "6. 발송량·피로도" --profile 25

고친 뒤엔 **같은 명령으로 다시 재서 전후 값을 커밋·CLAUDE.md에 남길 것.**
숫자가 없으면 다음 사람이 이유를 몰라 되돌린다.
"""
import argparse
import cProfile
import os
import pathlib
import pstats
import shutil
import sys
import tempfile
import threading
import time

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

# ── 함정 ① 스크립트 캐시 공유 ──────────────────────────────────────────
# `LocalScriptRunner.__init__`이 `ScriptCache()`를 «새로» 만든다. 모듈 네임스페이스의
# 이름을 갈아 끼우면 호출 시점에 우리 것이 잡힌다(인스턴스 생성 전이기만 하면 된다).
import streamlit.testing.v1.local_script_runner as _lsr          # noqa: E402
from streamlit.runtime.scriptrunner.script_cache import ScriptCache  # noqa: E402

_SHARED_CACHE = ScriptCache()
_lsr.ScriptCache = lambda *a, **k: _SHARED_CACHE

from streamlit.testing.v1 import AppTest                          # noqa: E402

TIMEOUT = 900

# ── 함정 ② 스크립트 스레드 프로파일 ────────────────────────────────────
_PROF = {"on": False, "stats": None}
_orig_thread_run = threading.Thread.run


def _thread_run(self, *a, **k):
    if not _PROF["on"]:
        return _orig_thread_run(self, *a, **k)
    pr = cProfile.Profile()
    pr.enable()
    try:
        return _orig_thread_run(self, *a, **k)
    finally:
        pr.disable()
        if _PROF["stats"] is None:
            _PROF["stats"] = pstats.Stats(pr)
        else:
            _PROF["stats"].add(pr)


threading.Thread.run = _thread_run


def _page_radio(at):
    """페이지 라디오 — 발송성과는 사이드바 첫 라디오, 주간보고는 label='페이지'."""
    rs = [r for r in at.sidebar.radio if r.label == "페이지"] or \
         [r for r in at.radio if r.label == "페이지"] or list(at.sidebar.radio)
    return rs[0] if rs else None


class SendPerf:
    """발송성과 — 실적을 «세션»에 들고 있다."""
    name = "발송성과"
    app = str(ROOT / "send_perf_dashboard.py")

    def __init__(self, rows):
        from smoke_pages import synth_store
        self.store = synth_store(weeks=70, per_day=rows, seed=7)

    def describe(self):
        return f"캠페인 {len(self.store):,}건"

    def setup(self):
        return None                      # cwd를 안 옮겨도 된다

    def fresh(self):
        at = AppTest.from_file(self.app, default_timeout=TIMEOUT)
        at.session_state["camp_store"] = self.store
        at.run()
        return at


class Weekly:
    """주간보고 — 저장소를 «디스크»에서 매번 읽는다(서명 캐시가 걸린 자리다)."""
    name = "주간보고"
    app = str(ROOT / "weekly_report.py")

    def __init__(self, rows):
        from smoke_weekly_report import synth_store
        self.store = synth_store()
        self._tmp = None

    def describe(self):
        return f"마스터 {len(self.store):,}행"

    def setup(self):
        self._tmp = tempfile.mkdtemp(prefix="perf_wr_")
        shutil.copy(self.app, os.path.join(self._tmp, "weekly_report.py"))
        te = ROOT / "table_export.py"
        if te.exists():
            shutil.copy(te, os.path.join(self._tmp, "table_export.py"))
        self.store.to_csv(os.path.join(self._tmp, "wr_data_store.csv"),
                          index=False, encoding="utf-8-sig")
        self.app = os.path.join(self._tmp, "weekly_report.py")
        os.chdir(self._tmp)
        return self._tmp

    def fresh(self):
        at = AppTest.from_file(self.app, default_timeout=TIMEOUT)
        at.run()
        return at


APPS = {"send": SendPerf, "weekly": Weekly}


def measure(adapter, page, runs):
    """그 페이지를 연 뒤의 리런 비용. 워밍업 1회는 버리고 min·avg를 낸다.

    min을 같이 보는 건 GC·스케줄링 잡음이 avg만 밀어 올리기 때문이다."""
    at = adapter.fresh()
    if at.exception:
        raise RuntimeError(at.exception[0].value)
    r = _page_radio(at)
    if r is not None and page is not None:
        r.set_value(page)
        at.run()
        if at.exception:
            raise RuntimeError(at.exception[0].value)
    at.run()                                           # 워밍업
    ts = []
    for _ in range(runs):
        t0 = time.perf_counter()
        at.run()
        ts.append(time.perf_counter() - t0)
    return min(ts), sum(ts) / len(ts)


def main():
    ap = argparse.ArgumentParser(description="대시보드 리런 비용 측정")
    ap.add_argument("--app", choices=list(APPS), default="send")
    ap.add_argument("--rows", type=int, default=25,
                    help="발송성과: 하루당 캠페인 수 (25면 실백업 규모인 1.2만 건)")
    ap.add_argument("--runs", type=int, default=3, help="페이지당 측정 횟수")
    ap.add_argument("--page", help="이 페이지만 잰다 (생략하면 전부)")
    ap.add_argument("--profile", type=int, metavar="N",
                    help="측정과 함께 상위 N개 함수를 뜬다 (--page와 같이 쓰는 게 좋다)")
    a = ap.parse_args()

    adapter = APPS[a.app](a.rows)
    adapter.setup()
    print(f"{adapter.name} — {adapter.describe()}\n")

    t0 = time.perf_counter()
    at = adapter.fresh()
    if at.exception:
        print(f"초기 렌더 실패: {at.exception[0].value}")
        return 1
    print(f"첫 렌더(스크립트 컴파일 포함) {time.perf_counter() - t0:.2f}초")
    print("  두 번째부터는 ScriptCache를 공유해 컴파일이 빠집니다.\n")

    r = _page_radio(at)
    pages = [a.page] if a.page else (list(r.options) if r is not None else [None])

    if a.profile:
        _PROF["on"] = True
    rows, total = [], 0.0
    for p in pages:
        try:
            mn, av = measure(adapter, p, a.runs)
        except Exception as e:                          # noqa: BLE001
            print(f"  {p}: 실패 — {str(e)[:150]}")
            continue
        rows.append((p, mn, av))
        total += mn
    _PROF["on"] = False

    print(f"{'페이지':<34}{'min':>9}{'avg':>9}")
    for p, mn, av in sorted(rows, key=lambda x: -x[1]):
        print(f"  {str(p):<32}{mn:>8.2f}s{av:>8.2f}s")
    print(f"\n  합계(min) {total:.2f}초 · {len(rows)}개 페이지")

    if a.profile and _PROF["stats"] is not None:
        print(f"\n상위 {a.profile}개 (누적 시간)")
        _PROF["stats"].sort_stats("cumulative").print_stats(a.profile)
    return 0


if __name__ == "__main__":
    sys.exit(main())
