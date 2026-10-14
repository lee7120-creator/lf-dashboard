"""소구 사전 후보 발굴 — **거짓 발견을 사전에 집어넣지 않는지.**

`KW` 사전은 손으로 채운 목록이라 문구가 바뀌어도 안 따라간다. 담당자가 새로 쓰기
시작한 말은 어느 속성에도 안 잡혀 분석에서 통째로 빠지는데, 증상은 '미분류가 좀
많네'로만 보인다. 그래서 문구에서 후보 말을 캐어 성과가 갈리는 걸 추린다.

**여기서 거짓말하는 길이 다이제스트보다 하나 더 많다.**

| 길 | 막는 법 |
|---|---|
| 후보가 수천 개라 5%가 저절로 「유의」로 뜬다 | 최소 표본 → 변화폭 문턱 → **BH FDR** |
| 찾은 데이터에서 확인까지 한다 | **홀드아웃** — 앞에서 찾고 뒤에서 같은 방향인지 본다 |
| 거의 모든 건에 있는 말을 비교한다 | `PHRASE_MAX_SHARE` — 비교군이 안 남으면 뺀다 |
| 접두어 조각이 목록을 채운다 | 같은 칸을 집는 말은 **가장 긴 것만** 남긴다 |
| 화면이 사전을 자동으로 고친다 | **안 고친다** — 후보 목록까지만 낸다 |

**사전은 한 번 넣으면 전 화면의 축이 되고 과거 분석까지 다시 계산된다.** 다이제스트는
그 주만 틀리고 끝나지만 사전은 안 그렇다 — 홀드아웃이 한 겹 더 있는 이유다.

로컬 실행:
    python tests/test_phrase_mining.py
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
TAB = "소구 사전 후보"

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# ──────────────────────────────────────────────────────────────────────
# 픽스처 — 말 하나에 «진짜 효과»를 심는다
# ──────────────────────────────────────────────────────────────────────

def _frame(days=200, per_day=6, seed=7, word="타임딜",
           lift_train=1.5, lift_hold=1.5, every=3, base_ctr=0.020,
           send=50_000):
    """`word`가 든 캠페인의 CTR만 `lift`배로 심은 프레임.

    `lift_hold`를 따로 받는 건 **홀드아웃에서 뒤집히는 말**을 만들기 위해서다.
    앞·뒤를 같은 비율로 심으면 홀드아웃 규칙을 깨도 값이 안 변해 안 잡힌다.
    """
    rng = np.random.default_rng(seed)
    dates = pd.date_range("2026-01-01", periods=days, freq="D")
    cut = dates[int(days * (1 - A.PHRASE_HOLDOUT)) - 1]
    rows = []
    for i, d in enumerate(dates):
        for j in range(per_day):
            hit = (i * per_day + j) % every == 0
            lift = lift_train if d <= cut else lift_hold
            ctr = base_ctr * (lift if hit else 1.0)
            rows.append({
                "date": d, "send": send, "uv": int(rng.binomial(send, ctr)),
                "title": f"오늘 {word} 시작" if hit else "오늘 일반 안내",
                "body": "본문입니다",
            })
    return pd.DataFrame(rows)


def _flat(days=200, per_day=6, seed=11):
    """효과가 «없는» 프레임 — 말은 섞여 있지만 CTR이 말과 무관하다."""
    rng = np.random.default_rng(seed)
    words = ["가가가", "나나나", "다다다", "라라라", "마마마"]
    dates = pd.date_range("2026-01-01", periods=days, freq="D")
    rows = []
    for i, d in enumerate(dates):
        for j in range(per_day):
            rows.append({
                "date": d, "send": 50_000,
                "uv": int(rng.binomial(50_000, 0.020)),
                "title": f"오늘 {words[(i * per_day + j) % len(words)]} 안내",
                "body": "본문입니다",
            })
    return pd.DataFrame(rows)


def _words(out):
    return [str(w) for w in out["단어"]] if len(out) else []


# ══════════════════════════════════════════════════════════════════════
# 홀드아웃 — 이 검사가 이 파일의 존재 이유다
# ══════════════════════════════════════════════════════════════════════

@case
def t_a_word_that_flips_in_the_holdout_is_dropped():
    """**앞에서 좋고 뒤에서 나쁜 말은 후보가 아니다.**

    선택과 확인을 같은 데이터로 하면 후보가 전부 '유의'해 보인다. 사전은 한 번 넣으면
    전 화면의 축이 되고 `TAGSET_VER`가 물려 과거 분석까지 다시 계산되니, 다이제스트보다
    한 단계 더 보수적이어야 한다.

    같은 말을 앞 기간엔 1.5배로, 뒤 기간엔 0.67배로 심는다 — 앞만 보면 반드시 뽑히고
    홀드아웃을 보면 반드시 탈락한다.
    """
    up = A.phrase_candidates(_frame(lift_train=1.5, lift_hold=1.5))[0]
    assert "타임딜" in _words(up), f"심은 말을 앞에서도 못 찾았어요 — {_words(up)}"

    flip, steps = A.phrase_candidates(_frame(lift_train=1.5, lift_hold=0.67))
    assert steps["q"] > 0, f"뒤집히는 말이 q까지도 못 갔어요 — 픽스처가 약해요: {steps}"
    assert "타임딜" not in _words(flip), (
        f"홀드아웃에서 뒤집힌 말이 후보로 올라왔어요 — {_words(flip)}")


@case
def t_the_holdout_split_is_by_date_not_by_row():
    """홀드아웃은 **날짜**로 가른다 — 행으로 가르면 같은 날이 양쪽에 걸린다.

    같은 날 발송이 앞뒤로 쪼개지면 그날의 행사·시즌이 양쪽에 다 들어가, 홀드아웃이
    '다른 기간'이 아니라 '같은 기간의 다른 절반'이 된다. 그러면 확인이 아니라
    **같은 데이터로 두 번 재는 것**이다.
    """
    df = _frame(days=100, per_day=8)
    days = pd.to_datetime(df["date"]).dt.normalize()
    uniq = days.drop_duplicates().sort_values()
    cut = uniq.iloc[max(1, int(len(uniq) * (1 - A.PHRASE_HOLDOUT))) - 1]
    tr, ho = df[days <= cut], df[days > cut]
    assert len(tr) and len(ho), "한쪽이 비었어요"
    assert set(pd.to_datetime(tr["date"]).dt.normalize()) & \
           set(pd.to_datetime(ho["date"]).dt.normalize()) == set(), \
        "같은 날이 앞뒤에 같이 들어갔어요"


# ══════════════════════════════════════════════════════════════════════
# 신호 있음 · 신호 없음 — 둘 다 본다
# ══════════════════════════════════════════════════════════════════════

@case
def t_a_real_signal_is_found():
    """진짜로 갈리는 말은 찾아야 한다 — 못 찾는 발굴은 안전한 게 아니라 고장이다."""
    out, steps = A.phrase_candidates(_frame())
    assert "타임딜" in _words(out), f"심은 말을 못 찾았어요 — {_words(out)} / {steps}"
    r = out[out["단어"] == "타임딜"].iloc[0]
    assert r["변화율"] > 0.3, f"변화율이 너무 작아요 — {r['변화율']:.3f}"
    assert r["홀드아웃 변화율"] > 0, f"홀드아웃이 음수예요 — {r['홀드아웃 변화율']:.3f}"


@case
def t_noise_is_not_reported_as_a_finding():
    """효과가 없으면 **아무것도 안 내놔야** 한다. 매주 뭔가 찾아내면 그게 거짓 발견이다."""
    out, steps = A.phrase_candidates(_flat())
    assert out.empty, f"없는 신호를 찾았어요 — {_words(out)} / {steps}"


# ══════════════════════════════════════════════════════════════════════
# 문턱 넷
# ══════════════════════════════════════════════════════════════════════

@case
def t_fdr_is_actually_applied():
    """보정을 **실제로 거는지** 본다 — `bh_qvalues` 단위 검사로는 못 보는 자리다.

    p를 그대로 q에 넣어도 단위 검사는 통과한다(다이제스트에서 겪은 그 자리). 그래서
    보정을 **항등으로 갈아 끼운 뒤** 결과가 달라지는지로 본다.

    픽스처는 **효과가 없는** 쪽이어야 한다. 효과를 크게 심으면 p가 전부 정확히 0이
    되어 `BH(0,0,…) = 0`이라 **보정을 걸든 안 걸든 값이 같다** — 실제로 그렇게 짰다가
    이 검사가 «보정을 안 걸었다»고 잘못 외쳤다.
    """
    df = _flat()
    real, _ = A.phrase_candidates(df, rel=0.0, qmax=1.0)
    assert len(real) >= 2, f"검정이 둘 미만이라 비교할 게 없어요 — {len(real)}"

    orig = A.bh_qvalues
    try:
        A.bh_qvalues = lambda ps: np.asarray(list(ps), dtype=float)   # 보정 끄기
        raw, _ = A.phrase_candidates(df, rel=0.0, qmax=1.0)
    finally:
        A.bh_qvalues = orig

    a = real.set_index("단어")["q"].astype(float)
    b = raw.set_index("단어")["q"].astype(float)
    both = a.index.intersection(b.index)
    assert len(both) >= 2, f"맞댈 단어가 모자라요 — {list(both)}"
    assert (a[both] >= b[both] - 1e-12).all(), "q가 p보다 작아요 — 보정이 거꾸로예요"
    assert (a[both] > b[both] + 1e-12).any(), \
        f"q가 p와 전부 같아요 — 보정을 안 걸었어요\n{a[both]}\n{b[both]}"


@case
def t_a_word_in_almost_every_campaign_is_not_tested():
    """거의 모든 건에 있는 말은 **비교군이 얇아서** 검정에서 뺀다.

    **「전 건에 있는 말」로는 이 규칙을 못 잡는다.** 그땐 안 들어간 캠페인이 0건이라
    최소 표본 문턱에 먼저 걸려서, `PHRASE_MAX_SHARE`를 꺼도 결과가 같다 — 실제로 그
    주입이 안 잡혔다. 그래서 **90%에만** 심는다. 남는 10%는 최소 표본은 넘지만
    비교군으로 쓰기엔 얇아서, 문턱이 없으면 그 말이 목록에 올라온다.
    """
    df = _frame(word="거의전부", every=10, lift_train=0.5, lift_hold=0.5)
    # `every=10`이면 그 말이 1/10에만 들어간다 — 뒤집어서 9/10에 넣는다
    hit = df["title"].str.contains("거의전부")
    df.loc[hit, "title"] = "오늘 일반 안내"
    df.loc[~hit, "title"] = "오늘 거의전부 안내"
    share = float((~hit).mean())
    assert share > A.PHRASE_MAX_SHARE, f"픽스처가 문턱을 안 넘어요 — {share:.2f}"

    out, _ = A.phrase_candidates(df, rel=0.0, qmax=1.0)
    assert "거의전부" not in _words(out), \
        f"{share:.0%}에 있는 말이 올라왔어요 — {_words(out)}"


@case
def t_small_samples_are_not_tested():
    """표본이 작은 말은 아예 검정에 안 올린다 — `rank_adjusted`와 같은 이유다."""
    df = _frame(days=60, per_day=2, send=100)       # 발송이 `PHRASE_MIN_SEND`에 한참 못 미친다
    out, steps = A.phrase_candidates(df)
    assert out.empty, f"작은 표본이 통과했어요 — {_words(out)}"
    assert steps["표본"] == 0, f"표본 문턱이 안 걸렸어요 — {steps}"


@case
def t_the_effect_gate_narrows_the_list():
    """변화폭 문턱을 올리면 후보가 **줄어야** 한다 — 문턱이 실제로 걸리는지 본다."""
    df = _frame(days=220, per_day=8, every=4)
    loose, _ = A.phrase_candidates(df, rel=0.0, qmax=1.0)
    tight, _ = A.phrase_candidates(df, rel=0.40, qmax=1.0)
    assert len(tight) < len(loose), f"문턱을 올려도 그대로예요 — {len(loose)} vs {len(tight)}"


# ══════════════════════════════════════════════════════════════════════
# 목록 모양
# ══════════════════════════════════════════════════════════════════════

@case
def t_prefix_fragments_collapse_to_the_longest():
    """같은 칸을 집는 접두어 조각은 **가장 긴 것만** 남긴다.

    `무료배송`을 심으면 `무료`·`무료배`·`무료배송`이 전부 같은 캠페인을 집어 성과가
    똑같이 찍힌다. 다 올리면 목록이 조각으로 차서 정작 볼 게 안 보인다.
    """
    out, _ = A.phrase_candidates(_frame(word="무료배송"))
    ws = _words(out)
    assert "무료배송" in ws, f"심은 말이 없어요 — {ws}"
    assert "무료배" not in ws, f"접두어 조각이 같이 올라왔어요 — {ws}"
    assert "무료" not in ws, f"접두어 조각이 같이 올라왔어요 — {ws}"


@case
def t_two_different_words_on_the_same_campaigns_both_survive():
    """**늘 같이 쓰는 말 둘은 둘 다 남긴다** — 어느 쪽이 먹힌 건지 코드는 모른다.

    접두어 겹침만 정리하고 동시출현은 정리하지 않는다. 한쪽을 임의로 지우면 화면이
    「이 말이 효과였다」고 단정하게 되는데, 그건 데이터가 말해 주지 않는 것이다.
    같은 칸에 말이 둘이면 **둘 중 아무거나가 아니라 둘 다** 올라와야 한다.
    """
    out, _ = A.phrase_candidates(_frame(word="타임딜"))
    ws = _words(out)
    assert "타임딜" in ws and "시작" in ws, (
        f"같이 다니는 말 중 하나가 사라졌어요 — {ws}")


@case
def t_an_empty_result_keeps_its_columns():
    """후보가 하나도 없어도 **컬럼은 그대로**여야 한다.

    한쪽 경로만 컬럼 없는 프레임을 돌려주면 소비처가 `out["변화율"]`에서 죽는데,
    그 경로는 「후보가 하나도 없을 때」라 평소엔 안 밟혀 영영 안 드러난다
    (다이제스트에서 적대적 재독으로 찾은 자리와 같다).
    """
    for df in (_flat(), pd.DataFrame(), _frame(days=1, per_day=2)):
        out, _ = A.phrase_candidates(df)
        assert list(out.columns) == list(A.PHRASE_COLS), \
            f"컬럼이 달라요 — {list(out.columns)}"


@case
def t_words_already_in_the_dictionary_are_marked_not_hidden():
    """이미 `KW`에 있는 말은 **지우지 말고 어느 속성인지 밝힌다.**

    거르면 목록이 '새 말'인지 '못 찾은 말'인지 구분이 안 가 믿을 수가 없다.
    """
    assert A.kw_owner("타임딜") == "마감임박", A.kw_owner("타임딜")
    assert A.kw_owner("무료배송비") == "가격소구", A.kw_owner("무료배송비")
    assert A.kw_owner("가가가") == "", A.kw_owner("가가가")
    out, _ = A.phrase_candidates(_frame())
    r = out[out["단어"] == "타임딜"]
    assert len(r) and r.iloc[0]["사전"] == "마감임박", \
        f"이미 있는 말에 속성이 안 붙었어요 — {r.to_dict('records')}"


# ══════════════════════════════════════════════════════════════════════
# 화면
# ══════════════════════════════════════════════════════════════════════

def _store_with(word="타임딜", weeks=40, seed=3, lift=1.6, every=3):
    """스모크 합성본에 말 하나를 심는다 — 화면이 보는 것과 같은 모양이어야 한다."""
    d = synth_store(weeks=weeks, seed=seed).reset_index(drop=True)
    hit = (d.index % every == 0)
    d.loc[hit, "title"] = f"오늘 {word} 시작"
    d.loc[~hit, "title"] = "오늘 일반 안내"
    base = pd.to_numeric(d["uv"], errors="coerce").fillna(0)
    d.loc[hit, "uv"] = (base[hit] * lift).round()
    return d


def _open(store):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = store
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
def t_the_screen_shows_how_many_survived_each_gate():
    """단계마다 몇 개가 남았는지 **화면에 찍는다.**

    안 보이면 「왜 하나도 없지」만 남는다 — 문턱에 걸린 건지 데이터가 없는 건지
    구분이 안 간다(다이제스트와 같은 규칙).
    """
    at = _open(_store_with())
    txt = " ".join(_texts(at))
    for k in ("후보", "표본", "변화폭", "홀드아웃"):
        assert k in txt, f"단계 「{k}」가 화면에 없어요"


@case
def t_the_screen_says_nothing_found_plainly():
    """신호가 없으면 **빈 표가 아니라 「없다」**고 말한다."""
    d = synth_store(weeks=40, seed=5)
    d = d.reset_index(drop=True)
    d["title"] = "오늘 일반 안내"                    # 전 건이 같은 말 — 갈릴 게 없다
    at = _open(d)
    txt = " ".join(_texts(at))
    assert "안 보여요" in txt or "부족해요" in txt, f"왜 비었는지 안 말해요 — {txt[:400]}"


@case
def t_the_screen_does_not_touch_the_dictionary():
    """**화면은 `KW`를 안 고친다.** 사전은 사람이 고친다.

    자동 반영은 전 화면의 축을 말없이 바꾸는 일이고, `TAGSET_VER`가 물려 과거
    분석까지 다시 계산된다. 정적으로 금지해 둔다 — 다음 사람이 「편하게」 넣고 싶어질
    자리다.
    """
    MUT = {"update", "setdefault", "pop", "popitem", "clear",
           "append", "extend", "insert", "remove", "add", "discard"}

    def _roots_at_kw(node):
        """`KW` · `KW[x]` · `KW[x][y]`처럼 **뿌리가 `KW`인지** 따라간다.

        `KW[...]`만 보면 `KW["할인율소구"].append(...)`를 놓친다 — 메서드를 부르는
        대상이 Name이 아니라 Subscript라서다. 실제로 그 주입이 안 잡혔다.
        """
        while isinstance(node, (ast.Subscript, ast.Attribute)):
            node = node.value
        return isinstance(node, ast.Name) and node.id == "KW"

    tree = ast.parse(pathlib.Path(APP).read_text(encoding="utf-8"))
    bad = []
    for n in ast.walk(tree):
        if isinstance(n, (ast.Assign, ast.AugAssign)):
            tg = n.targets if isinstance(n, ast.Assign) else [n.target]
            bad += [n.lineno for t in tg
                    if isinstance(t, ast.Subscript) and _roots_at_kw(t)]
        if isinstance(n, ast.Delete):
            bad += [n.lineno for t in n.targets if _roots_at_kw(t)]
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute) and \
                n.func.attr in MUT and _roots_at_kw(n.func.value):
            bad.append(n.lineno)
    assert not bad, f"`KW`를 코드에서 고치고 있어요 — 줄 {sorted(set(bad))}"


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
    print(f"소구 사전 후보 테스트 {len(CASES) if not want else '선택'}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
