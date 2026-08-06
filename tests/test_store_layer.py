"""저장소 계층 방어 테스트 — dtype이 문자열로 되돌아와도 앱이 죽지 않아야 한다.

구글시트 라운드트립은 모든 값을 문자열로 되돌린다. push 저장소는 `finalize_push()`로
복구하는 게 원칙인데, 한 경로라도 빠뜨리면 `get_push_stats`가 'str vs Timestamp'
비교로 터진다. 이 함수는 사이드바 렌더 경로라 **11개 페이지가 전부 다운**된다.

실백업으로 전 페이지를 돌리다 실제로 재현돼서 방어를 넣고 여기에 고정한다.

로컬 실행:
    python tests/test_store_layer.py
"""
import pathlib
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import send_perf_dashboard as S  # noqa: E402

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def raw_push():
    """구글시트에서 막 읽어온 모양 — 전부 문자열."""
    return pd.DataFrame([
        {"date": "2026-07-01", "group": "앱푸시_동의자수", "consent": "700000",
         "added": "1000", "removed": "400", "diff": "600", "is_outlier": "False"},
        {"date": "2026-07-02", "group": "앱푸시_동의자수", "consent": "700600",
         "added": "1200", "removed": "500", "diff": "700", "is_outlier": "False"},
        {"date": "2026-07-03", "group": "앱푸시_동의자수", "consent": "701300",
         "added": "900", "removed": "300", "diff": "600", "is_outlier": "True"},
    ])


@case
def t_get_push_stats_survives_string_dtypes():
    """문자열 dtype이 들어와도 죽지 않고, finalize된 것과 같은 값을 낸다."""
    got = S.get_push_stats(raw_push(), "2026-07-01", "2026-07-31", "앱푸시_동의자수")
    want = S.get_push_stats(S.finalize_push(raw_push()), "2026-07-01", "2026-07-31",
                            "앱푸시_동의자수")
    assert got is not None, "문자열 dtype에서 None이 나왔어요"
    assert got == want, f"{got} != {want}"


@case
def t_outlier_excluded():
    """is_outlier가 문자열 'True'여도 제외돼야 한다."""
    r = S.get_push_stats(raw_push(), "2026-07-01", "2026-07-31", "앱푸시_동의자수")
    assert r["tot_added"] == 2200, r          # 이상치 행(900) 제외
    assert r["last_consent"] == 700600, r     # 이상치 제외 후 마지막 날


@case
def t_empty_and_none_are_safe():
    """빈 프레임·None·컬럼 부재에서 예외 대신 None."""
    assert S.get_push_stats(None, "2026-07-01", "2026-07-31", "g") is None
    assert S.get_push_stats(pd.DataFrame(), "2026-07-01", "2026-07-31", "g") is None
    assert S.get_push_stats(pd.DataFrame({"x": [1]}), "2026-07-01", "2026-07-31", "g") is None


@case
def t_no_match_returns_none():
    """기간·그룹이 안 맞으면 None (0으로 위장하지 않는다)."""
    assert S.get_push_stats(raw_push(), "2020-01-01", "2020-12-31", "앱푸시_동의자수") is None
    assert S.get_push_stats(raw_push(), "2026-07-01", "2026-07-31", "없는그룹") is None


@case
def t_finalize_push_is_idempotent():
    """두 번 통과시켜도 같은 결과 — 여러 경로에서 중복 호출된다."""
    once = S.finalize_push(raw_push())
    twice = S.finalize_push(once)
    assert len(once) == len(twice)
    assert list(once["date"]) == list(twice["date"])
    assert list(once["is_outlier"]) == list(twice["is_outlier"])


def main():
    fails = []
    for fn in CASES:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {e}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"저장소 계층 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
