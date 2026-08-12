"""업로드로 데이터 범위가 넓어졌을 때 사이드바 기간 필터가 따라가는지.

`key`가 붙은 Streamlit 위젯은 `value=`가 최초 1회만 반영되고 그 뒤론 세션값이 이긴다.
그래서 8/2까지 있는 상태에서 앱을 연 뒤 8/3~8/4를 올리면, 기간 필터가 옛 범위를
그대로 적용해 새 날짜가 화면에서 통째로 사라졌다(F5로 세션을 날려야 보였다).

렌더 자체는 성공하므로 스모크로는 안 잡힌다 — 그래서 따로 둔다.

로컬 실행:
    python tests/test_filter_follows_upload.py
"""
import pathlib
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest        # noqa: E402

from smoke_pages import synth_store             # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300


def _sv(at, key, default=None):
    """AppTest 세션값 조회 — 미설정 키에서 예외 대신 default."""
    try:
        return at.session_state[key]
    except Exception:                                     # noqa: BLE001
        return default


def _goto_yoy(at):
    """5. 맥락·타이밍 › 전년 동요일 비교 페이지로 이동."""
    at.sidebar.radio[0].set_value("5. 맥락·타이밍")
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    if subs:
        for opt in subs[0].options:
            if "전년 동요일" in opt:
                subs[0].set_value(opt)
                at.run()
                break
    return at


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_new_dates_selectable_without_refresh():
    """업로드로 늘어난 날짜가 새로고침 없이 기준일 옵션에 나와야 한다."""
    full = synth_store(weeks=8)
    dates = sorted(full["date"].unique())
    old, new = dates[:-2], dates[-2:]

    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = full[full["date"].isin(old)].reset_index(drop=True)
    at.run()
    assert not at.exception, at.exception[0].value
    assert _sv(at, "flt_date") is not None, "기간 필터가 초기화되지 않았어요"

    # 업로드+저장으로 새 날짜가 스토어에 들어온 상황 — 세션은 유지된다(= F5 안 누름)
    at.session_state["camp_store"] = full
    at.run()
    assert not at.exception, at.exception[0].value

    span = _sv(at, "flt_date")
    want_hi = pd.Timestamp(new[-1]).date()
    assert span[1] == want_hi, f"기간 필터 종료일이 {span[1]} — {want_hi}로 따라가야 해요"

    at = _goto_yoy(at)
    assert not at.exception, at.exception[0].value
    sels = [s for s in at.selectbox if s.label == "기준일"]
    assert sels, "기준일 셀렉트박스를 찾지 못했어요"
    opts = list(sels[0].options)
    for d in new:
        lab = pd.Timestamp(d).strftime("%Y-%m-%d")
        assert any(lab in o for o in opts), f"{lab}이 기준일 옵션에 없어요 — {opts[:4]}"


@case
def t_user_narrowed_range_is_kept():
    """사용자가 직접 좁힌 기간은 새 데이터가 들어와도 멋대로 넓히지 않는다."""
    full = synth_store(weeks=8)
    dates = sorted(full["date"].unique())
    old, new = dates[:-2], dates[-2:]

    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = full[full["date"].isin(old)].reset_index(drop=True)
    at.run()

    lo = pd.Timestamp(old[0]).date()
    narrow_hi = pd.Timestamp(old[len(old) // 2]).date()
    at.session_state["flt_date"] = (lo, narrow_hi)         # 사용자가 직접 좁힘
    at.run()
    assert not at.exception, at.exception[0].value

    at.session_state["camp_store"] = full
    at.run()
    assert not at.exception, at.exception[0].value
    span = _sv(at, "flt_date")
    assert span[1] == narrow_hi, f"직접 좁힌 종료일이 {span[1]}로 바뀌었어요 (원래 {narrow_hi})"
    assert pd.Timestamp(new[-1]).date() > span[1], "테스트 전제가 깨졌어요"


@case
def t_unmatched_days_are_announced():
    """문구가 안 붙은 날짜가 통째로 숨겨지면 알리고, 한 번에 풀 수 있어야 한다.

    「문구 매칭된 것만」은 기본 켜짐이라, 방금 올린 실적에 문구가 안 붙으면 그 날짜가
    모든 화면에서 사라진다. 증상이 '기준일 목록에 날짜가 없다'로 나타나 업로드 실패로
    읽힌다 — 실제로 그렇게 리포트를 받았다.
    """
    full = synth_store(weeks=8)
    dates = sorted(full["date"].unique())
    new = dates[-2:]
    full = full.copy()
    _new_rows = full["date"].isin(new)
    full.loc[_new_rows, ["title", "body"]] = ""            # 기획 시트 미적재 상황
    full.loc[_new_rows, "matched"] = False

    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = full
    at.run()
    assert not at.exception, at.exception[0].value

    warns = [str(w.value) for w in at.warning if "문구 매칭된 것만" in str(w.value)]
    assert warns, f"숨겨진 날짜 알림이 없어요 — warnings={[str(w.value)[:60] for w in at.warning]}"
    for d in new:
        lab = f"{pd.Timestamp(d):%-m/%d}"
        assert lab in warns[0], f"{lab}이 알림에 없어요 — {warns[0][:120]}"

    btn = [b for b in at.button if b.label == "필터 끄기"]
    assert btn, "「필터 끄기」 버튼이 없어요"
    btn[0].click()
    at.run()
    assert not at.exception, at.exception[0].value
    assert at.session_state["flt_matched"] is False, "버튼이 필터를 끄지 못했어요"
    assert not [w for w in at.warning if "문구 매칭된 것만" in str(w.value)], \
        "필터를 껐는데 알림이 그대로예요"

    at = _goto_yoy(at)
    assert not at.exception, at.exception[0].value
    opts = list([s for s in at.selectbox if s.label == "기준일"][0].options)
    for d in new:
        lab = pd.Timestamp(d).strftime("%Y-%m-%d")
        assert any(lab in o for o in opts), f"필터를 껐는데도 {lab}이 안 보여요 — {opts[:4]}"


@case
def t_no_notice_when_nothing_hidden():
    """문구가 다 붙어 있으면 알림을 띄우지 않는다 (상시 노출 방지)."""
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=8)
    at.run()
    assert not at.exception, at.exception[0].value
    assert not [w for w in at.warning if "문구 매칭된 것만" in str(w.value)], \
        "숨겨진 날짜가 없는데 알림이 떴어요"


@case
def t_label_mode_switches_render():
    """'값 표시' 모드를 바꿔도 시간대 차트가 정상 렌더된다."""
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=8)
    at.run()
    at = _goto_yoy(at)
    assert not at.exception, at.exception[0].value
    sels = [s for s in at.selectbox if s.label == "값 표시"]
    assert sels, f"'값 표시' 셀렉트박스가 없어요 — {[s.label for s in at.selectbox]}"
    for mode in sels[0].options:
        tgt = [s for s in at.selectbox if s.label == "값 표시"][0]
        tgt.set_value(mode)
        at.run()
        assert not at.exception, f"값 표시={mode}에서 실패: {at.exception[0].value}"


def main():
    fails = []
    for fn in CASES:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {str(e)[:300]}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"필터·레이블 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
