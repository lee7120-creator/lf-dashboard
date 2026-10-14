"""플레이북 추천 되짚기 — **백테스트가 거짓말하지 않는지.**

「발송 플레이북」은 지난 성과로 (요일 × 시간) 슬롯을 추천하는데, 맞았는지 아무도
안 봤다. 되짚기가 없으면 추천을 믿을 근거도, 고칠 근거도 안 쌓인다.

**백테스트가 거짓말하는 길은 정해져 있다.**

| 길 | 막는 법 |
|---|---|
| 평가 주를 학습에 넣는다(미래 훔쳐보기) | 학습 창을 주 시작 **전**으로 자르고, 쓴 추천을 밖에서 되계산해 대조 |
| 전체 평균과 맞댄다 | 전체엔 추천분이 들어 있어 리프트가 저절로 깎인다 — **나머지**와 맞댄다 |
| 적은 주로 판정한다 | `BT_MIN_WEEKS` 미만이면 **판정하지 않는다** |
| 화면과 다른 랭킹을 되짚는다 | 추천 규칙은 `playbook_slots` **한 곳**이고 화면도 그걸 본다 |

로컬 실행:
    python tests/test_playbook_backtest.py
"""
import ast
import pathlib
import sys

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest            # noqa: E402

import send_perf_dashboard as A                     # noqa: E402
from smoke_pages import group_of, synth_store       # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300
TAB = "발송 플레이북"
GCOL = "ord_cr"

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def _store(weeks=40, seed=3):
    """파생 칼럼(`dow_k` 등)까지 붙인 프레임 — 화면이 보는 것과 같은 모양."""
    return A._finalize(synth_store(weeks=weeks, seed=seed))


# ══════════════════════════════════════════════════════════════════════
# 미래 훔쳐보기 — 이 검사가 이 파일의 존재 이유다
# ══════════════════════════════════════════════════════════════════════

@case
def t_recommendation_uses_only_past_data():
    """**그 주에 쓴 추천은 «그 주 시작 전» 데이터만으로 나와야 한다.**

    통계로 에둘러 보지 않는다 — 백테스트가 실어 준 추천을 밖에서 과거 데이터만으로
    되계산해 **그대로 같은지** 본다. 평가 주를 한 줄이라도 학습에 넣으면 여기서 갈린다.

    **픽스처가 규칙을 반증하는지 먼저 확인한다** — 과거만 쓴 추천과 전체를 쓴 추천이
    같아 버리면 훔쳐봐도 값이 안 변해서 아무것도 못 본다.
    """
    d = _store()
    bt = A.backtest_slots(d, GCOL)
    assert len(bt) >= 5, f"평가 주가 {len(bt)}개뿐이라 이 검사가 헛돌아요"

    dd = d.copy()
    dd["_dt"] = pd.to_datetime(dd["date"], format="%Y%m%d", errors="coerce")
    dd["_wk"] = dd["_dt"].dt.to_period("W-SUN").dt.start_time

    differ = 0
    for _, r in bt.iterrows():
        past = A.playbook_slots(dd[dd["_wk"] < r["주"]], GCOL)
        allin = A.playbook_slots(dd[dd["_wk"] <= r["주"]], GCOL)
        assert list(r["추천"]) == list(past), (
            f"{r['주']:%Y-%m-%d} 주의 추천이 «과거만»으로 만든 것과 달라요 — "
            f"미래를 훔쳐봤어요\n  썼다: {r['추천']}\n  과거만: {past}")
        if list(past) != list(allin):
            differ += 1
    assert differ, ("이 픽스처론 규칙을 못 깨요 — 과거만 쓴 추천과 평가 주까지 쓴 "
                    "추천이 모든 주에서 같아요")


@case
def t_the_baseline_is_the_rest_not_the_whole():
    """비교 기준은 **그 주의 나머지 발송**이다.

    전체 평균과 맞대면 분모에 추천분이 같이 들어가 리프트가 저절로 깎인다. 두 건수를
    더하면 그 주 전체가 되는지로 본다(겹치지도, 빠지지도 않아야 한다).
    """
    d = _store()
    bt = A.backtest_slots(d, GCOL)
    assert len(bt), "평가 주가 없어요"
    dd = d.copy()
    dd["_wk"] = (pd.to_datetime(dd["date"], format="%Y%m%d", errors="coerce")
                 .dt.to_period("W-SUN").dt.start_time)
    for _, r in bt.iterrows():
        n = int((dd["_wk"] == r["주"]).sum())
        assert r["추천 건수"] + r["나머지 건수"] == n, (
            f"{r['주']:%Y-%m-%d}: 추천 {r['추천 건수']} + 나머지 {r['나머지 건수']} "
            f"≠ 그 주 전체 {n} — 겹치거나 빠진 게 있어요")
        assert r["추천 건수"] > 0 and r["나머지 건수"] > 0


@case
def t_a_real_signal_is_found():
    """신호가 **있으면** 찾아내야 한다 — 「못 찾는 백테스트」는 안전한 게 아니다.

    한 슬롯만 꾸준히 좋게 심으면 추천이 그 슬롯을 집고 리프트가 양수여야 한다.
    이게 없으면 위 「신호 없음」 결과가 그냥 고장 난 건지 구분이 안 된다.
    """
    rows = []
    for w in range(14):
        for dow in ("월", "화", "수", "목", "금"):
            for h in (1000, 1400):
                good = (dow == "수" and h == 1400)
                for k in range(2):
                    rows.append({
                        "date": (pd.Timestamp("2026-01-05") + pd.Timedelta(weeks=w)
                                 + pd.Timedelta(days="월화수목금".index(dow))).strftime("%Y%m%d"),
                        "af": f"AP{w}{dow}{h}{k}", "dow_k": dow, "hour": h,
                        "send": 100_000.0, "uv": 5_000.0,
                        "oc": 500.0 * (2.0 if good else 1.0),
                        "amt": 5_000_000.0, "cat": "패션",
                        "ord_cr": 0.20 if good else 0.10})
    d = pd.DataFrame(rows)
    bt = A.backtest_slots(d, GCOL)
    vd = A.backtest_verdict(bt)
    assert vd["verdict"] is True, f"있는 신호를 못 찾았어요 — {vd}"
    assert ("수", 1400) in list(bt["추천"].iloc[-1]), bt["추천"].iloc[-1]


@case
def t_noise_is_not_reported_as_a_finding():
    """신호가 **없으면** 없다고 해야 한다 — 매주 뭔가를 찾아내면 그게 거짓 발견이다."""
    vd = A.backtest_verdict(A.backtest_slots(_store(), GCOL))
    assert vd["verdict"] is not True, f"무작위 데이터에서 신호를 찾았다고 해요 — {vd}"


@case
def t_too_few_weeks_is_not_judged():
    """주가 모자라면 **판정하지 않는다** — 적은 표본으로 판정하는 게 제일 쉬운 거짓말이다."""
    bt = A.backtest_slots(_store(weeks=10), GCOL)
    vd = A.backtest_verdict(bt)
    if len(bt) >= A.BT_MIN_WEEKS:
        bt = bt.head(A.BT_MIN_WEEKS - 1)
        vd = A.backtest_verdict(bt)
    assert vd["verdict"] is None, f"{len(bt)}주로 판정했어요 — {vd}"
    assert "최소" in str(vd["why"]), vd["why"]


@case
def t_an_empty_or_tiny_store_does_not_crash():
    """저장소가 비었거나 축 칼럼이 없어도 화면이 죽으면 안 된다."""
    for d in (None, pd.DataFrame(), _store().drop(columns=["dow_k"]),
              _store().head(1)):
        bt = A.backtest_slots(d, GCOL)
        assert isinstance(bt, pd.DataFrame), bt
        assert A.backtest_verdict(bt)["verdict"] is None
    assert A.playbook_slots(None, GCOL) == []
    assert A.playbook_slots(pd.DataFrame(), GCOL) == []


@case
def t_the_screen_and_the_backtest_share_one_ranking():
    """추천 규칙은 **`playbook_slots` 한 곳**이다.

    화면이 따로 정렬하면 되짚기가 «화면이 추천하지 않은 것»을 재게 되고, 숫자는
    멀쩡히 나오는데 뜻이 없다. 플레이북 분기 안에서 `rank_adjusted`를 직접 부르는
    슬롯 정렬이 **없어야** 한다.
    """
    src = pathlib.Path(APP).read_text(encoding="utf-8")
    fns = [n.name for n in ast.walk(ast.parse(src)) if isinstance(n, ast.FunctionDef)]
    assert "playbook_slots" in fns and "backtest_slots" in fns, fns

    # **플레이북 분기 «안»만 본다.** 파일 전체에서 찾으면 「5. 타이밍·카테고리 ›
    # 타이밍·발송슬롯」이 걸린다 — 거기는 사용자가 최소 표본을 직접 정하고 슬롯을
    # 전부 보는, **질문이 다른 화면**이라 공용 함수로 묶으면 그 화면이 바뀐다.
    # 처음에 파일 전체로 적었다가 그 오탐을 밟았다.
    i = src.find('elif "플레이북" in page:')
    assert i > 0, "플레이북 분기를 못 찾았어요 — 분기 조건이 바뀌었나요"
    j = src.find("\n    elif ", i + 1)
    blk = src[i:j if j > 0 else len(src)]
    assert "playbook_slots(" in blk, "플레이북 화면이 공용 추천 함수를 안 써요"
    assert 'rank_adjusted(_sf' not in blk and 'slot["_rk"] = rank_adjusted(' not in blk, \
        "플레이북이 슬롯 순위를 따로 매기고 있어요 — `playbook_slots`를 쓰세요"
    assert "backtest_slots(" in blk and "backtest_verdict(" in blk, \
        "플레이북 화면에 되짚기가 없어요"


# ══════════════════════════════════════════════════════════════════════
# 화면
# ══════════════════════════════════════════════════════════════════════

def _open(store, **sess):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = store
    for k, v in sess.items():
        at.session_state[k] = v
    at.run()
    at.sidebar.radio[0].set_value(group_of(TAB))
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    subs[0].set_value(TAB)
    at.run()
    assert not at.exception, at.exception[0].value
    return at


def _texts(at):
    out = []
    for coll in (at.caption, at.info, at.success, at.warning, at.markdown):
        out += [str(getattr(e, "value", "")) for e in coll]
    return out


@case
def t_the_screen_says_the_verdict_and_how_it_measured():
    """화면이 **판정과 방법**을 같이 말해야 한다 — 숫자만 보면 전진식인지 모른다."""
    at = _open(synth_store(weeks=40, seed=3))
    txt = " ".join(_texts(at))
    assert "주 중" in txt and "부호검정" in txt, f"판정을 안 말해요 — {txt[:200]!r}"
    assert "전진식" in txt, f"어떻게 쟀는지 안 밝혀요 — {txt[:200]!r}"
    assert "나머지" in txt, "무엇과 맞댔는지 안 밝혀요"


@case
def t_the_backtest_table_is_on_screen():
    """주별 되짚기 표가 떠야 한다 — 판정 한 줄만 있으면 왜 그런지 못 되짚는다."""
    at = _open(synth_store(weeks=40, seed=3))
    cols = [list(d.value.columns) for d in at.dataframe]
    hit = [c for c in cols if "추천 건수" in [str(x) for x in c]]
    assert hit, f"되짚기 표가 없어요 — {cols}"
    assert "리프트" in [str(x) for x in hit[0]], hit[0]


def main():
    want = sys.argv[1:]
    fails = []
    for fn in [f for f in CASES if not want or any(w in f.__name__ for w in want)]:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {str(e)[:400]}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"플레이북 되짚기 테스트 {len(CASES) if not want else '선택'}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
