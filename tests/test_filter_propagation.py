"""상단 컨트롤이 **그 아래 블록을 다 끌고 가는지** — 한 화면이 두 가지를 말하지 않게.

이 저장소에서 세 번 난 사고가 같은 모양이다. 상단에 지표·기준을 고르는 컨트롤이
있는데 **아래 블록 하나가 자기 값을 따로 계산**해서, 컨트롤을 바꿔도 거기만 안
따라온다. 셋 다 «화면은 멀쩡히 뜨고 값만 틀리는» 꼴이라 눈으로도 코드 리뷰로도
안 잡혔고, 전부 사용자가 지적한 뒤에야 **그 자리에만** 검사가 붙었다.

| 사고 | 어디 |
|---|---|
| 「비교 지표」를 거래액으로 바꿔도 소구 블록만 CTR | 5. 타이밍·카테고리 › 전년 동요일 비교 |
| 「비교 기준」을 전년 동월로 바꿔도 카테고리 표가 그대로 | 1. 주간보고 |
| 「조회 기간」을 3주로 좁혀도 아래 추이는 연중 | 주간보고 02. 첫구매 퍼널 |

그래서 **자리마다 검사를 붙이는 대신 규칙을 하나 둔다** — 컨트롤을 바꾸면
«그 아래» 블록은 **전부** 값이 달라져야 하고, 안 달라져도 되는 블록만 이유와 함께
목록에 적는다. 블록을 새로 달면 자동으로 규칙에 걸리니, 다음 사고는 지적받기 전에
잡힌다. 반대 방향도 같이 본다 — 컨트롤은 **«위»를 건드리면 안 된다**
(`t_comparison_basis_does_not_touch_the_cards_above`가 그 자리다).

**제목만 보는 검사는 소용없다.** 제목만 바뀌는 게 이 저장소의 전형적인 조용한
실패라, 차트는 y값·표는 셀까지 대조한다.

로컬 실행:
    python tests/test_filter_propagation.py
    python tests/test_filter_propagation.py 전년      # 이름이 든 계약만
"""
import base64
import json
import pathlib
import sys

import numpy as np

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest            # noqa: E402

from smoke_pages import group_of, synth_store       # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300

# 값을 보여 주는 블록. **타입 이름을 넘겨짚지 말 것** — 처음에 표를
# `arrow_data_frame`으로 적었다가 표 4개를 통째로 못 보고도 초록이었다. 아래
# `_snap`이 「블록이 하나도 안 잡혔으면」 그 자리에서 멈춘다.
BLOCKS = ("plotly_chart", "dataframe", "metric")
# 컨트롤 — 값이 아니라 입력이라 대조에서 뺀다.
CONTROLS = ("selectbox", "radio", "button_group", "multiselect", "checkbox",
            "number_input", "slider", "select_slider", "date_input", "text_input",
            "pills", "segmented_control", "toggle")

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def _nums(y):
    """plotly JSON은 숫자 배열을 base64(`bdata`)로 싣기도 한다."""
    if isinstance(y, dict) and "bdata" in y:
        return np.frombuffer(base64.b64decode(y["bdata"]),
                             dtype=np.dtype(y.get("dtype", "f8"))).tolist()
    return list(y or [])


def _label_and_value(n):
    """블록 하나 → (사람이 읽을 이름, 값 지문). 값은 **숫자·셀까지** 담는다."""
    t = n.type
    if t == "plotly_chart":
        sp = json.loads(n.proto.spec)
        title = ((sp.get("layout") or {}).get("title") or {}).get("text") or "(무제 차트)"
        val = json.dumps([[_nums(tr.get(ax)) for ax in ("x", "y")]
                          for tr in sp.get("data", [])], default=str)
        return title, val
    if t == "dataframe":
        df = n.value
        cols = [str(c) for c in getattr(df, "columns", [])]
        return "|".join(cols[:3]) or "(무제 표)", df.to_csv()
    return str(getattr(n, "label", "")), str(getattr(n, "value", ""))


def _snap(at, ctrl_key):
    """화면을 «컨트롤 위 / 아래»로 갈라 (이름, 값) 목록을 만든다."""
    above, below, seen_ctrl = [], [], False
    for n in at.main:
        t = getattr(n, "type", "")
        if t in CONTROLS:
            if getattr(n, "key", None) == ctrl_key:
                seen_ctrl = True
            continue
        if t not in BLOCKS:
            continue
        (below if seen_ctrl else above).append(_label_and_value(n))
    assert seen_ctrl, f"컨트롤 {ctrl_key!r} 이(가) 화면에 없어요"
    assert below, (f"{ctrl_key!r} 아래에 값 블록이 하나도 없어요 — "
                   "이 픽스처론 전파를 못 봅니다")
    return above, below


def _open(tab, store, **sess):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = store
    for k, v in sess.items():
        at.session_state[k] = v
    at.run()
    at.sidebar.radio[0].set_value(group_of(tab))
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    if subs and tab in list(subs[0].options):
        subs[0].set_value(tab)
        at.run()
    assert not at.exception, at.exception[0].value
    return at


def _diff(a, b):
    """두 스냅샷을 **자리별로** 맞댄다. 블록 수가 달라지면 그것부터 말한다."""
    assert len(a) == len(b), (
        f"컨트롤을 바꿨더니 블록 개수가 달라졌어요 ({len(a)} → {len(b)}). "
        "블록이 생기거나 사라지면 자리 대조가 무의미해져요")
    same, moved = [], []
    for (la, va), (lb, vb) in zip(a, b):
        (same if va == vb else moved).append((la, lb))
    return same, moved


def check(tab, ctrl_key, values, stays_below=(), follows_above=(),
          store=None, **sess):
    """컨트롤 하나의 «전파 계약»을 확인한다.

    `stays_below` — 아래인데 **안 바뀌어도 되는** 블록(이름 조각 → 이유).
    `follows_above` — 위인데 **바뀌어도 되는** 블록. 둘 다 적은 게 실제로 화면에
    나와야 한다. 블록을 지우고 목록만 남으면 그 줄이 썩은 채로 규칙을 느슨하게
    만드는데, 증상이 없어서 영영 안 드러난다.
    """
    store = synth_store(weeks=60, seed=5) if store is None else store
    snaps = []
    for v in values:
        at = _open(tab, store, **{ctrl_key: v}, **sess)
        snaps.append(_snap(at, ctrl_key))
    (ab_a, be_a), (ab_b, be_b) = snaps

    def _hit(exc, la, lb):
        return next((e for e in exc if e in la or e in lb), None)

    be_same, be_moved = _diff(be_a, be_b)
    bad = [la for la, lb in be_same if not _hit(stays_below, la, lb)]
    assert not bad, (
        f"[{tab}] 「{ctrl_key}」를 {values[0]!r}→{values[1]!r}로 바꿨는데 "
        f"**아래 블록이 안 따라와요**: {bad}\n"
        "    같은 화면이 두 가지를 말하게 돼요. 일부러 안 따라가는 블록이면 "
        "`stays_below`에 이유와 함께 적어 주세요.")

    ab_same, ab_moved = _diff(ab_a, ab_b)
    bad = [la for la, lb in ab_moved if not _hit(follows_above, la, lb)]
    assert not bad, (
        f"[{tab}] 「{ctrl_key}」가 **위 블록까지 바꿨어요**: {bad}\n"
        "    위는 그 컨트롤이 없는 맥락이라 같이 바뀌면 어디서 바뀐 건지 못 되짚어요.")

    # 목록이 썩지 않게 — 적어 둔 예외가 실제로 화면에 있어야 한다
    seen = [la for la, _ in be_a + be_b + ab_a + ab_b]
    stale = [e for e in tuple(stays_below) + tuple(follows_above)
             if not any(e in s for s in seen)]
    assert not stale, (
        f"[{tab}] 예외 목록에 적어 둔 블록이 화면에 없어요: {stale}\n"
        "    블록을 지웠으면 목록에서도 빼세요 — 남겨 두면 규칙만 느슨해져요.")
    return len(be_moved), len(be_same)


# ══════════════════════════════════════════════════════════════════════
# 전파 계약
# ══════════════════════════════════════════════════════════════════════

@case
def t_yoy_metric_picker_drives_everything_below():
    """「비교 지표」(`p05b_met`) — 시간대 차트·표·일자 합산·소구 블록을 다 끈다.

    2026-10 사고 자리다. 소구 블록만 CTR에 박혀 있었고 제목에도 「소구 속성별 CTR」
    이라 적혀 있어 눈으로도 안 잡혔다.
    """
    moved, same = check(
        "전년 동요일 비교", "p05b_met", ("CTR", "거래액"),
        stays_below=("날짜|시간|카테고리",))   # 문구 표 — CTR 순 정렬이 공용 헬퍼 인자다
    assert moved >= 4, f"따라간 블록이 {moved}개뿐이에요 — 픽스처가 너무 얕아요"


@case
def t_yoy_ratio_basis_drives_everything_below():
    """「비율 기준」(`p05b_basis`) — 합산↔캠페인 평균. 소구 블록까지 같이 간다."""
    check("전년 동요일 비교", "p05b_basis", ("합산", "캠페인 평균"),
          stays_below=("날짜|시간|카테고리",))


def main():
    want = sys.argv[1:]
    fails = []
    for fn in [f for f in CASES if not want or any(w in f.__name__ for w in want)]:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {str(e)[:600]}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"필터 전파 계약 {len(CASES) if not want else '선택'}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
