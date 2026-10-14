"""주간 변화 다이제스트 — «달라진 칸»을 찾되 **거짓 발견을 안 만드는지**.

화면이 서른 개 넘는데 「지난 기간 대비 뭐가 달라졌나」를 사람이 돌며 눈으로 찾는다.
그래서 매주 같은 걸 놓친다. 그 자리를 메우는 화면인데, **여기서 틀리는 방식이 둘이다.**

1. **보정을 안 하면 아무 일 없어도 목록이 찬다.** 축 여섯 × 값 수십 × 지표 둘이면
   비교가 수백 개라, p<0.05만 쓰면 5%가 저절로 「유의미」가 된다.
2. **보정만 하면 반대로 거의 다 걸린다.** 칸 하나가 발송 수십만 건이면 0.1%p 차이도
   p값이 0으로 떨어진다. 실제로 합성 1.2만 건에서 **검정 40개 중 38개가 유의**로 떴다.
   38개를 들이미는 다이제스트는 0개짜리와 똑같이 안 읽힌다.

그래서 문턱이 둘이다 — **변화폭이 1차, q가 2차.** 정렬도 변화폭 순이다(q는 거의 다
0이라 정렬 키로 못 쓴다). 그리고 **「신호 없음」도 그대로 말한다** — 매주 뭔가를
찾아내려 드는 순간 그게 거짓 발견이다.

로컬 실행:
    python tests/test_change_digest.py
"""
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
TAB = "변화 다이제스트"

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# ══════════════════════════════════════════════════════════════════════
# 순수 함수
# ══════════════════════════════════════════════════════════════════════

@case
def t_bh_matches_the_hand_calculation():
    """BH 보정은 손계산과 맞아야 한다 — 라이브러리를 «맞게 부르는지»가 쟁점이다.

    q_i = min over j>=i of (m/j)·p_j. p=[.001,.01,.04,.5], m=4 →
    [.004, .02, .0533…, .5].
    """
    got = list(A.bh_qvalues([0.001, 0.01, 0.04, 0.5]))
    want = [0.004, 0.02, 4 / 3 * 0.04, 0.5]
    for g, w in zip(got, want):
        assert abs(g - w) < 1e-12, f"{got} != {want}"


@case
def t_bh_does_not_count_cells_it_could_not_measure():
    """못 잰 칸(NaN)을 검정 수에 넣으면 보정이 **느슨해진다.**

    [.001, nan, .04]에서 m은 3이 아니라 2다 → [.002, nan, .04].
    """
    got = list(A.bh_qvalues([0.001, float("nan"), 0.04]))
    assert abs(got[0] - 0.002) < 1e-12 and np.isnan(got[1]) \
        and abs(got[2] - 0.04) < 1e-12, got
    assert list(A.bh_qvalues([])) == [], "빈 목록에서 죽으면 안 돼요"


@case
def t_two_prop_z_signs_the_difference_and_is_symmetric():
    """차이는 **기준 − 직전**이고, 뒤집으면 부호만 바뀌고 p는 같아야 한다."""
    d1, p1 = A.two_prop_z(100, 1000, 80, 1000)
    d2, p2 = A.two_prop_z(80, 1000, 100, 1000)
    assert abs(d1 - 0.02) < 1e-12 and abs(d2 + 0.02) < 1e-12, (d1, d2)
    assert abs(p1 - p2) < 1e-12, (p1, p2)
    assert A.two_prop_z(100, 1000, 100, 1000)[1] == 1.0


@case
def t_a_cell_it_cannot_measure_is_not_a_finding():
    """분모 0·양쪽 0%는 **nan**이지 0이 아니다 — 0으로 두면 «변화 없음»으로 읽힌다."""
    assert all(np.isnan(v) for v in A.two_prop_z(1, 0, 1, 10)), A.two_prop_z(1, 0, 1, 10)
    assert np.isnan(A.two_prop_z(0, 100, 0, 100)[1]), "양쪽 0%면 잴 게 없어요"


def _cell(date, **kw):
    r = {"date": date, "af": "AP01", "hour": 1000, "target": "전체", "stype": "정기",
         "bpu": "1BPU", "prio": "1", "prio_g": "1", "cat": "패션", "attr": "통합",
         "owner": "A", "brand": "닥스", "brand2": "닥스", "dow_k": "월",
         "send": 100_000.0, "uv": 5_000.0, "oc": 500.0, "amt": 5_000_000.0}
    r.update(kw)
    return r


def _two_windows(n=4, **cur):
    """같은 칸을 두 기간에 n건씩. `cur`로 기준 기간만 흔든다."""
    rows = [_cell(f"2026100{i + 1}", **cur) for i in range(n)]
    rows += [_cell(f"2026090{i + 1}") for i in range(n)]
    d = pd.DataFrame(rows)
    return d[d["date"].str.startswith("202610")], d[d["date"].str.startswith("202609")]


@case
def t_small_cells_never_reach_the_test():
    """표본이 작은 칸은 **검정에 안 올린다** — 요행이 목록을 점령하면 안 된다.

    발송을 문턱 아래로 내리면 그 칸이 통째로 빠져야 한다(`rank_adjusted`와 같은 이유).
    """
    cur, prv = _two_windows(n=4, uv=9_000.0)
    assert len(A.change_digest(cur, prv)), "이 픽스처론 규칙을 못 깨요 — 원래 잡혀야 해요"
    small_c = cur.assign(send=A.DIGEST_MIN_SEND / 10)
    small_p = prv.assign(send=A.DIGEST_MIN_SEND / 10)
    assert not len(A.change_digest(small_c, small_p)), "발송이 적은 칸이 검정에 올라갔어요"
    few_c, few_p = cur.head(A.DIGEST_MIN_CAMP - 1), prv.head(A.DIGEST_MIN_CAMP - 1)
    assert not len(A.change_digest(few_c, few_p)), "캠페인이 적은 칸이 검정에 올라갔어요"


@case
def t_rows_are_ordered_by_effect_not_by_q():
    """**정렬은 변화폭 순이다.** 발송이 크면 q가 거의 다 0이라 정렬 키로 못 쓴다.

    변화폭이 큰 칸과 작은 칸을 같이 심고, 큰 쪽이 위에 오는지 본다. q로 정렬하면
    둘 다 0이라 순서가 입력 순서에 좌우된다.
    """
    # **두 정렬이 어긋나게 심어야 한다.** 처음엔 「살짝 움직인 칸 / 크게 움직인 칸」을
    # 같은 표본으로 심었는데, 크게 움직인 쪽이 q도 작아서 **q로 정렬해도 순서가 같았다**
    # — 그 주입이 안 잡혔다. 어긋나게 하려면 «표본이 거대하고 변화는 작은 칸»과
    # «표본이 작고 변화는 큰 칸»이 필요하다.
    rows = []
    for i in range(4):      # A — 발송 2천만, CTR 5.05% → 5.00% (변화율 −1%, p≈4e-13)
        rows.append(_cell(f"2026100{i + 1}", cat="A", send=5_000_000.0, uv=250_000.0))
        rows.append(_cell(f"2026090{i + 1}", cat="A", send=5_000_000.0, uv=252_500.0))
    for i in range(4):      # B — 발송 3.2만, CTR 0.10% → 0.22% (변화율 +120%, p≈1e-4)
        rows.append(_cell(f"2026100{i + 1}", cat="B", send=8_000.0, uv=17.5))
        rows.append(_cell(f"2026090{i + 1}", cat="B", send=8_000.0, uv=8.0))
    d = pd.DataFrame(rows)
    out = A.change_digest(d[d["date"].str.startswith("202610")],
                          d[d["date"].str.startswith("202609")],
                          dims=(("cat", "카테고리"),))
    ctr = out[out["지표"] == "CTR"].reset_index(drop=True)
    assert len(ctr) == 2, ctr
    # 픽스처가 규칙을 반증하는지 먼저 확인 — 두 정렬이 같은 답이면 아무것도 못 본다
    _qa = float(ctr.loc[ctr["값"] == "A", "q"].iloc[0])
    _qb = float(ctr.loc[ctr["값"] == "B", "q"].iloc[0])
    assert _qa < _qb, (f"이 픽스처론 규칙을 못 깨요 — q로 정렬해도 순서가 같아요 "
                       f"(A {_qa:.3g} / B {_qb:.3g})")
    assert ctr.loc[0, "값"] == "B", \
        f"변화폭이 큰 칸이 위에 와야 해요 — {ctr[['값', '변화율', 'q']].to_dict('records')}"
    assert abs(ctr.loc[0, "변화율"]) > abs(ctr.loc[1, "변화율"])


@case
def t_the_digest_actually_applies_the_correction():
    """`change_digest`가 **q를 BH로 만든다.** 위 `bh_qvalues` 검사만으론 «부르는지»를
    못 본다 — p를 그대로 q에 넣어도 그 검사는 통과한다.

    BH q는 정의상 **언제나 p 이상**이고, 검정이 여럿이면 적어도 하나는 **더 커야**
    한다(안 그러면 보정이 안 걸린 것이다).
    """
    out = A.change_digest(*_two_windows(n=4, uv=5_200.0))
    ok = out[out["p"].notna()]
    assert len(ok) >= 2, f"이 픽스처론 보정을 못 봐요 — 검정이 {len(ok)}개뿐이에요"
    assert (ok["q"] >= ok["p"] - 1e-12).all(), "q가 p보다 작아요 — BH가 아니에요"
    assert (ok["q"] > ok["p"] + 1e-12).any(),         f"q가 p와 똑같아요 — 보정을 안 거쳤어요\n{ok[['지표', 'p', 'q']]}"


@case
def t_empty_input_is_an_empty_result_not_a_crash():
    """저장소가 비었거나 한쪽 기간이 비어도 화면이 죽으면 안 된다."""
    cur, prv = _two_windows()
    for a, b in ((None, prv), (cur, None), (cur.head(0), prv), (cur, prv.head(0))):
        out = A.change_digest(a, b)
        assert isinstance(out, pd.DataFrame) and not len(out), out


@case
def t_a_one_row_window_does_not_crash():
    """창에 캠페인이 **딱 한 건**일 때 죽으면 안 된다.

    pandas 3.0.6의 `GroupBy.get_group`은 **프레임이 1행이면** 스칼라 키로 불러도
    `KeyError`로 죽는다. 실측으로 «행 1·그룹 1»만 깨지고 «행 2·그룹 1»·«행 4·그룹 1»
    은 멀쩡하다 — 그래서 **「축에 값이 하나뿐」으로 짚으면 안 잡힌다**(처음에 그렇게
    적었다가 주입을 놓쳤다). 운영에선 브랜드 하나로 좁히고 「최근 7일」을 보면
    바로 밟고, `min_camp` 가드는 그 줄보다 **뒤**라 못 막는다.
    """
    cur, prv = _two_windows(n=1)                    # ← 1행이어야 한다. 2행이면 안 깨진다
    assert len(cur) == 1 and len(prv) == 1, (len(cur), len(prv))
    out = A.change_digest(cur, prv)                 # 죽으면 여기서 끝난다
    assert list(out.columns) == list(A.DIGEST_COLS), out.columns


@case
def t_an_empty_result_keeps_its_columns():
    """결과가 비어도 **컬럼은 그대로** — 한쪽 경로만 컬럼이 없으면 소비처가 죽는다.

    「검정할 칸이 하나도 없을 때」만 밟는 경로라 평소엔 안 드러난다.
    """
    cur, prv = _two_windows(n=1)                    # 캠페인 1건 — 전부 걸러진다
    for out in (A.change_digest(None, None), A.change_digest(cur, prv)):
        assert not len(out), out
        assert list(out.columns) == list(A.DIGEST_COLS), \
            f"빈 결과의 컬럼이 달라요 — {list(out.columns)}"


@case
def t_a_missing_axis_is_skipped_not_fatal():
    """축 컬럼이 없는 저장소(옛 백업)에서도 나머지 축은 돌아야 한다."""
    cur, prv = _two_windows(n=4, uv=9_000.0)
    out = A.change_digest(cur.drop(columns=["brand2"]), prv.drop(columns=["brand2"]))
    assert len(out), "축 하나가 없다고 전부 멈췄어요"
    assert "브랜드" not in set(out["축"]), out["축"].unique()


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
def t_the_effect_gate_narrows_the_list():
    """「최소 변화폭」을 올리면 목록이 **줄어야** 한다 — 안 줄면 문턱이 안 걸린 것이다."""
    store = synth_store(weeks=30, seed=3)
    seen = {}
    for lab in ("±10% 이상", "±30% 이상"):
        at = _open(store, dg_rel=lab)
        seen[lab] = len(at.dataframe[0].value) if at.dataframe else 0
    assert seen["±30% 이상"] < seen["±10% 이상"], \
        f"문턱을 올려도 목록이 그대로예요 — {seen}"


@case
def t_no_signal_is_said_plainly():
    """두 기간이 똑같으면 **「변화 없음」이라고 말해야** 한다.

    매주 뭔가를 찾아내려 드는 순간 그게 거짓 발견이다. 두 창에 **같은 값**을 심어
    「찾아낼 게 없는」 상황을 만든다.
    """
    base = synth_store(weeks=30, seed=3).copy()
    dt = pd.to_datetime(base["date"], format="%Y%m%d", errors="coerce")
    keep = base[dt >= dt.max() - pd.Timedelta(days=27)].copy()
    for c, v in (("send", 200_000.0), ("uv", 10_000.0), ("oc", 1_000.0),
                 ("amt", 10_000_000.0)):
        keep[c] = v
    at = _open(keep.reset_index(drop=True))
    txt = " ".join(_texts(at))
    assert "변화가 없어요" in txt, f"신호가 없는데 그렇게 말하지 않아요 — {txt[:300]!r}"
    assert not (at.dataframe and len(at.dataframe[0].value)), \
        "변화가 없는데 목록이 떴어요"


@case
def t_the_two_windows_do_not_overlap():
    """두 창이 겹치면 같은 발송을 양쪽이 나눠 갖는다 — 날짜를 화면에서 읽어 본다."""
    at = _open(synth_store(weeks=30, seed=3), dg_win="최근 7일")
    cap = next((t for t in _texts(at) if "기준 **" in t and "직전 **" in t), None)
    assert cap, f"기간 캡션이 없어요 — {_texts(at)[:3]}"
    import re
    ds = re.findall(r"(\d{4}-\d{2}-\d{2})", cap)
    assert len(ds) == 4, f"날짜 네 개를 못 읽었어요 — {cap!r}"
    c_lo, c_hi, p_lo, p_hi = ds
    assert p_hi < c_lo, f"두 창이 겹쳐요 — 직전 {p_lo}~{p_hi} / 기준 {c_lo}~{c_hi}"
    assert pd.Timestamp(c_hi) - pd.Timestamp(c_lo) == pd.Timestamp(p_hi) - pd.Timestamp(p_lo), \
        f"두 창 길이가 달라요 — {cap!r}"


@case
def t_the_screen_says_how_many_it_tested():
    """**몇 개를 검정했는지 밝힌다.** 안 밝히면 목록 길이만 보고 믿게 된다."""
    at = _open(synth_store(weeks=30, seed=3))
    txt = " ".join(_texts(at))
    assert "검정 **" in txt and "변화폭 통과" in txt and "q<" in txt, \
        f"검정 수·문턱 통과 수를 안 밝혀요 — {txt[:300]!r}"


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
    print(f"변화 다이제스트 테스트 {len(CASES) if not want else '선택'}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
