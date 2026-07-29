"""발송성과 대시보드 전 페이지 스모크 — 합성 데이터로 모든 페이지를 실제 렌더해 본다.

머지 전에 '어느 페이지든 열자마자 죽는' 류의 회귀를 잡는 게 목적이다.
실제로 2026-07 전역 헬퍼 섀도잉(`_s`) 사고 때 11개 페이지가 전부 다운됐는데,
이 스모크를 돌렸다면 머지 전에 걸렸다.

로컬 실행:
    python tests/smoke_pages.py
"""
import sys
import pathlib

import numpy as np
import pandas as pd
from streamlit.testing.v1 import AppTest

ROOT = pathlib.Path(__file__).resolve().parent.parent
APP = ROOT / "send_perf_dashboard.py"
TIMEOUT = 300


def synth_store(weeks=18, per_day=2, seed=1):
    """업로드 실적과 같은 모양의 합성 캠페인 데이터.

    날짜는 '오늘' 기준 상대값이라 시간이 지나도 계속 유효하다(부분 주 경로도 자연히 탄다).
    """
    rng = np.random.default_rng(seed)
    today = pd.Timestamp.utcnow().tz_localize(None).normalize()
    start = today - pd.Timedelta(weeks=weeks)
    rows = []
    for i, d in enumerate(pd.date_range(start, today, freq="D")):
        for j in range(per_day):
            rows.append(dict(
                date=d.strftime("%Y%m%d"), af=f"AP{i:04d}{j}",
                hour=rng.choice(["1000", "1430", "2000"]),
                target=rng.choice(["신규", "휴면", "전체"]),
                stype=rng.choice(["정기", "이벤트"]), bpu=rng.choice(["1BPU", "2BPU"]),
                prio=str(rng.integers(1, 4)), cat=rng.choice(["패션", "리빙", "뷰티"]),
                attr="정상", owner=rng.choice(["김", "이"]),
                brand=rng.choice(["닥스", "헤지스", "질스튜어트"]),
                send=int(rng.integers(5_000, 50_000)), uv=int(rng.integers(200, 3_000)),
                visit=int(rng.integers(300, 4_000)), cust=int(rng.integers(10, 200)),
                oc=int(rng.integers(5, 150)), amt=int(rng.integers(10**5, 5 * 10**6)),
                infl_cr=float(rng.uniform(0.01, 0.10)), ord_cr=float(rng.uniform(0.01, 0.08)),
                promo="", title=f"제목{i} 세일&특가 <단독>", body="본문 내용입니다",
                matched=True,
            ))
    return pd.DataFrame(rows)


def _fresh(store):
    at = AppTest.from_file(str(APP), default_timeout=TIMEOUT)
    at.session_state["camp_store"] = store
    at.run()
    return at


def main():
    if not APP.exists():
        print(f"앱 파일을 찾을 수 없어요: {APP}")
        return 1
    store = synth_store()
    print(f"합성 데이터 {len(store):,}건 ({store['date'].min()} ~ {store['date'].max()})\n")

    # 페이지 목록은 앱에서 직접 읽는다 — 페이지가 추가/개명돼도 자동으로 따라간다
    at = _fresh(store)
    if at.exception:
        print(f"  FAIL (초기 렌더): {at.exception[0].value}")
        return 1
    if not at.sidebar.radio:
        print("  FAIL: 사이드바 페이지 라디오를 찾지 못했어요")
        return 1
    pages = list(at.sidebar.radio[0].options)
    print(f"페이지 {len(pages)}개 발견\n")

    fails = []
    for p in pages:
        try:
            a = _fresh(store)
            if a.exception:
                raise RuntimeError(f"초기 렌더: {a.exception[0].value}")
            a.sidebar.radio[0].set_value(p)
            a.run()
            if a.exception:
                raise RuntimeError(a.exception[0].value)
            print(f"  OK   {p}")
        except Exception as e:                       # noqa: BLE001 — 어떤 예외든 실패로 기록
            print(f"  FAIL {p}: {str(e)[:200]}")
            fails.append(p)

    print()
    if fails:
        print(f"실패 {len(fails)}/{len(pages)}: {fails}")
        return 1
    print(f"전 페이지 스모크 통과 ✅ ({len(pages)}/{len(pages)})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
