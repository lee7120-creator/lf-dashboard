"""「4. 성과 진단 › 전환·AOV 진단」의 목표 CTR 역산 테스트.

거래액 = 발송 × CTR × 주문CR × 객단가라, 나머지를 붙들면 목표 거래액에 필요한 CTR이
나눗셈 한 번으로 나온다. 위험한 건 그 숫자가 **달성 가능한 수준인지 말하지 않는 것**이다.
최근 1년 최고치보다 높은 값을 목표로 내걸면 그대로 실행 계획이 된다.

로컬 실행:
    python tests/test_target_ctr.py
"""
import pathlib
import re
import sys

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest        # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300
PAGE, TAB = "4. 성과 진단", "전환·AOV 진단"


def synth(weeks=70, ctr=0.030, cr=0.010, aov=130_000, drop=0.5, jitter=0.0, seed=8):
    """전년 대비 거래액이 줄어든 데이터.

    최근 4주만 발송량을 drop배로 줄인다 — CTR·주문CR·객단가는 그대로라, 거래액 격차는
    전부 발송량에서 온다. 그러면 'CTR만으로 메우기'는 산술적으로 1/drop배가 필요하다.

    jitter>0이면 주별 CTR에 흔들림을 준다. '최근 1년 최고치'와 비교하는 판정을 보려면
    분포에 폭이 있어야 해서다(흔들림이 0이면 최고치=현재값이라 어떤 목표든 '초과'가 된다).
    비율이 정확히 맞아떨어져야 하는 검사는 jitter=0으로 부른다.
    """
    rng = np.random.default_rng(seed)
    end = pd.Timestamp.utcnow().tz_localize(None).normalize()
    rows = []
    for i, d in enumerate(pd.date_range(end - pd.Timedelta(weeks=weeks), end, freq="D")):
        _recent = d > end - pd.Timedelta(weeks=4)
        for j in range(2):
            send = int(600_000 * (drop if _recent else 1.0))
            _c = ctr * (float(rng.normal(1.0, jitter)) if jitter else 1.0)
            uv = max(int(send * _c), 1)
            oc = max(int(uv * cr), 1)
            rows.append(dict(
                date=d.strftime("%Y%m%d"), af=f"AF{i:04d}{j}", hour="1000", target="전체",
                stype="기본발송", bpu="1BPU", prio="1", cat="패션", attr="정상", owner="김",
                brand="닥스", send=send, uv=uv, visit=uv + 3, cust=uv // 3, oc=oc,
                amt=int(oc * aov), infl_cr=uv / send, ord_cr=oc / uv, promo="",
                title=f"제목{i}", body="본문", matched=True))
    return pd.DataFrame(rows)


def _open(store):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = store
    at.run()
    assert not at.exception, at.exception[0].value
    at.sidebar.radio[0].set_value(PAGE)
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    assert subs and TAB in subs[0].options, f"{TAB} 탭이 없어요"
    subs[0].set_value(TAB)
    at.run()
    assert not at.exception, at.exception[0].value
    return at


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success):
        out += [str(e.value) for e in coll]
    return out


def _kpi(at):
    return {m.label: m.value for m in at.metric}


def _tbl(at, first_col):
    for t in at.dataframe:
        cols = [str(c) for c in getattr(t.value, "columns", [])]
        if cols and cols[0] == first_col:
            return t.value
    return None


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_block_renders():
    """역산 블록이 뜨고 목표 CTR KPI가 나와야 한다."""
    at = _open(synth())
    txt = " ".join(_texts(at))
    assert "목표 CTR 역산" in txt, f"블록이 없어요 — {txt[:200]}"
    k = _kpi(at)
    for want in ("현재 거래액", "현재 CTR", "목표 CTR"):
        assert want in k, f"{want} KPI가 없어요 — {sorted(k)}"


@case
def t_target_ctr_matches_formula():
    """목표 CTR = 목표 거래액 ÷ (발송 × 주문CR × 객단가).

    합성 데이터는 최근 4주 발송량만 절반이라, 나머지가 그대로면 필요한 CTR은 정확히
    2배가 되어야 한다.
    """
    at = _open(synth(drop=0.5))
    k = _kpi(at)
    _cur = float(k["현재 CTR"].rstrip("%"))
    _need = float(k["목표 CTR"].rstrip("%"))
    assert abs(_need / _cur - 2.0) < 0.06, f"현재 {_cur}% · 목표 {_need}% — 2배여야 해요"


@case
def t_all_levers_need_same_relative_change():
    """거래액은 네 레버의 곱이라, 한 레버만 움직일 때 필요 변화율은 넷이 같다."""
    at = _open(synth(drop=0.5))
    tv = _tbl(at, "레버")
    assert tv is not None, "레버 표를 못 찾았어요"
    assert set(tv["레버"].astype(str)) == {"CTR", "주문CR", "발송", "객단가"}, list(tv["레버"])
    pct = [float(re.sub(r"[^0-9.\-+]", "", str(v))) for v in tv["필요 변화"]]
    assert max(pct) - min(pct) < 0.3, f"필요 변화율이 서로 달라요 — {pct}"


@case
def t_unreachable_target_is_called_out():
    """최근 1년 최고치보다 높은 목표면 'CTR만으로는 못 메워요'라고 말해야 한다.

    숫자만 내놓고 달성 가능성을 말하지 않으면 그대로 실행 목표가 된다.
    """
    at = _open(synth(drop=0.35))                  # 발송이 1/3 → 필요 CTR이 3배 가까이
    txt = " ".join(_texts(at))
    assert "CTR만으로는 못 메워요" in txt, f"비현실 경고가 없어요 — {txt[-500:]}"


@case
def t_reachable_target_is_encouraged():
    """목표가 최근 분포 안쪽이면 노려볼 만하다고 말해야 한다."""
    at = _open(synth(drop=0.97, jitter=0.06))     # 격차가 거의 없음 + 주별 폭 있음
    txt = " ".join(_texts(at))
    assert "CTR만으로는 못 메워요" not in txt, "달성 가능한데 비현실이라고 했어요"
    assert any(k in txt for k in ("노려볼 만한", "도달한 적은 있지만", "이미 목표를 넘었어요")), \
        f"달성 가능 판정이 없어요 — {txt[-400:]}"


@case
def t_scenario_table_offers_other_levers():
    """CTR 단독이 무리일 때 무엇을 같이 움직이면 되는지 보여줘야 한다."""
    at = _open(synth(drop=0.35, jitter=0.06))
    tv = _tbl(at, "같이 움직이는 것")
    assert tv is not None, "시나리오 표를 못 찾았어요"
    names = set(tv["같이 움직이는 것"].astype(str))
    assert {"지금 그대로", "발송을 기준 수준으로", "발송·주문CR 둘 다"} <= names, sorted(names)

    def _row(nm):
        return tv[tv["같이 움직이는 것"] == nm].iloc[0]

    def _pct(v):
        return float(re.sub(r"[^0-9.\-+]", "", str(v)))

    # CTR 하나로는 못 닿는 상황이어야 이 표가 의미가 있다
    assert str(_row("지금 그대로")["달성 이력"]) == "없음", _row("지금 그대로").to_dict()
    # 발송을 되돌리면 CTR은 거의 그대로여도 목표에 닿는다 (합성 데이터 설계)
    assert abs(_pct(_row("발송을 기준 수준으로")["지금 대비"])) < 8, _row("발송을 기준 수준으로").to_dict()
    _both = _row("발송·주문CR 둘 다")
    assert abs(_pct(_both["지금 대비"])) < 8 and str(_both["달성 이력"]) == "있음", _both.to_dict()


@case
def t_manual_target_switches_basis():
    """'직접 입력'을 고르면 그 금액을 목표로 쓴다."""
    at = _open(synth(drop=0.5))
    sel = [s for s in at.selectbox if s.label == "목표 기준"]
    assert sel, f"목표 기준 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
    sel[0].set_value("직접 입력")
    at.run()
    assert not at.exception, at.exception[0].value
    ni = [n for n in at.number_input if n.label == "목표 거래액(억원)"]
    assert ni, "목표 거래액 입력이 없어요"
    _before = _kpi(at)["목표 CTR"]
    ni[0].set_value(float(ni[0].value) * 2)
    at.run()
    assert not at.exception, at.exception[0].value
    _after = _kpi(at)["목표 CTR"]
    assert float(_after.rstrip("%")) > float(_before.rstrip("%")), (_before, _after)


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
    print(f"목표 CTR 역산 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
