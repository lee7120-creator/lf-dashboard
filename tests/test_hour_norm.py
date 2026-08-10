"""발송 시간대 자동 보정 테스트 — 실적 엑셀의 수기 입력 흔들림을 표준 HHMM으로.

실백업(12,729건)에 실제로 들어 있던 두 가지를 근거로 복구 규칙을 정했다.

| 원값 | 복구 | 근거 |
|------|------|------|
| `2400` (3건) | 자정 `00시` | 전부 `야간푸시` 카테고리. 같은 카테고리에 `0`(4건)이 공존한다 |
| `8000` (4건) | `800`(08시) | 해당 AF코드(AP59·AP60)의 평소 발송이 `800`(129건), 같은 문구도 `800`(12건) |

**지어내지 않는 것도 같이 고정한다.** `2460`을 `02:46` 같은 없는 시각으로 복구하면
조용히 틀린 분석이 된다. 못 읽으면 None으로 두고 화면에서 몇 건인지 알리는 게 맞다.

보정은 `_finalize` 한 곳에서만 한다. 화면마다 따로 고치면 슬롯 집계는 원값(8000),
차트는 보정값(800)으로 갈려 같은 발송이 두 줄로 쪼개진다.

로컬 실행:
    python tests/test_hour_norm.py
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


@case
def t_valid_values_pass_through():
    """멀쩡한 값은 손대지 않는다."""
    for v in (0, 800, 830, 1050, 1200, 1915, 1920, 2200, 2359):
        assert S.norm_hhmm(v) == v, (v, S.norm_hhmm(v))
    assert S.norm_hhmm(23) == 2300                        # 한 자리·두 자리는 '시'로 읽음
    assert S.norm_hhmm(8) == 800
    assert S.norm_hhmm(30) == 30                          # 00:30


@case
def t_2400_is_midnight():
    """`2400`은 오타가 아니라 24시 표기 — 자정으로 읽는다 (실백업 야간푸시 3건)."""
    assert S.norm_hhmm(2400) == 0
    assert S.norm_hhmm(2430) == 30
    assert S.hour_of_day(2400) == 0
    assert S.fmt_hhmm(2400) == "00시"
    assert S.hhmm_to_minutes(2400) == 0


@case
def t_trailing_zero_typo():
    """`8000`은 `800`의 뒤 0 오타 (실백업 4건 — AF코드·문구 모두 08시 발송)."""
    assert S.norm_hhmm(8000) == 800
    assert S.hour_of_day(8000) == 8
    assert S.fmt_hhmm(8000) == "08시"
    assert S.hhmm_to_minutes(8000) == 8 * 60
    assert S.norm_hhmm(9000) == 900
    assert S.norm_hhmm(12000) == 1200


@case
def t_does_not_invent_times():
    """복구 근거가 없으면 지어내지 말고 None — 조용히 틀린 분석보다 낫다."""
    assert S.norm_hhmm(2460) is None                      # 02:46으로 만들면 안 됨
    assert S.norm_hhmm(2461) is None
    assert S.norm_hhmm(1260) is None                      # 12:60 — 분이 범위 밖
    assert S.norm_hhmm(99999) is None
    assert S.norm_hhmm(-1) is None
    assert S.norm_hhmm(None) is None
    assert S.norm_hhmm("") is None
    assert S.norm_hhmm("아침") is None
    assert S.fmt_hhmm(2460) == "–"
    assert S.hhmm_to_minutes(2460) == 0


@case
def t_string_input_from_gsheet():
    """구글시트 라운드트립은 전부 문자열로 되돌린다."""
    assert S.norm_hhmm("8000") == 800
    assert S.norm_hhmm("2400") == 0
    assert S.norm_hhmm("800.0") == 800


@case
def t_finalize_normalizes_column():
    """보정은 `_finalize` 한 곳에서 — 저장소 로드·신규 파싱 양쪽에 같이 걸린다."""
    df = pd.DataFrame({
        "date": ["20260702", "20251117", "20260702"],
        "af": ["AP59", "AP02", "AP10"],
        "hour": [8000, 2400, 1050],
        "send": [100, 200, 300], "uv": [1, 2, 3], "oc": [1, 1, 1],
        "amt": [10, 20, 30], "title": ["a", "b", "c"],
    })
    out = S._finalize(df)
    assert list(out["hour"]) == [800.0, 0.0, 1050.0], list(out["hour"])


@case
def t_finalize_keeps_numeric_dtype():
    """dtype이 object로 새면 시간대 필터의 astype(str) 비교가 조용히 깨진다.

    필터는 선택지도 비교값도 같은 칼럼의 `astype(str)`로 만든다. object가 되면
    '800'과 '800.0'이 섞여, 선택지는 뜨는데 아무것도 안 걸리는 상태가 된다.
    보정 전(`pd.to_numeric`만 하던 때)과 같은 dtype이어야 한다.
    """
    for raw, fixed in ([8000], [800]), ([8000, 2460], [800, None]):
        df = pd.DataFrame({"date": ["20260702"] * len(raw), "af": list("AB"[:len(raw)]),
                           "hour": raw, "send": [1] * len(raw), "uv": [1] * len(raw),
                           "oc": [1] * len(raw), "amt": [1] * len(raw),
                           "title": ["a"] * len(raw)})
        got = S._finalize(df)["hour"]
        want = pd.to_numeric(pd.Series(fixed), errors="coerce")
        assert got.dtype == want.dtype, (raw, got.dtype, want.dtype)
        assert got.dtype != object


@case
def t_finalize_survives_unreadable_hour():
    """못 읽는 값이 섞여도 죽지 않고 NaN으로 남긴다."""
    df = pd.DataFrame({"date": ["20260702", "20260702"], "af": ["A", "B"],
                       "hour": [2460, "아침"], "send": [1, 1], "uv": [1, 1],
                       "oc": [1, 1], "amt": [1, 1], "title": ["a", "b"]})
    out = S._finalize(df)
    assert out["hour"].isna().all()


@case
def t_no_hour_column_is_safe():
    """시간대 칸이 아예 없는 파일도 있다."""
    df = pd.DataFrame({"date": ["20260702"], "af": ["A"], "send": [1], "uv": [1],
                       "oc": [1], "amt": [1], "title": ["a"]})
    out = S._finalize(df)
    assert "hour" not in out.columns


@case
def t_repaired_rows_group_together():
    """같은 08시 발송이 원값(8000)·보정값(800)으로 갈리지 않는다.

    보정을 화면마다 따로 하면 슬롯 집계에서 두 줄로 쪼개져 표본이 반토막 난다.
    """
    df = pd.DataFrame({
        "date": ["20260702", "20260709"], "af": ["AP59", "AP59"],
        "hour": [800, 8000], "send": [100, 100], "uv": [1, 1],
        "oc": [1, 1], "amt": [1, 1], "title": ["a", "b"],
    })
    out = S._finalize(df)
    assert out.groupby("hour").ngroups == 1, out["hour"].tolist()


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
    print(f"시간대 보정 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
