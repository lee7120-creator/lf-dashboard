"""「4. 성과 진단 › 전환·AOV 진단」의 목표 CTR 역산 테스트.

거래액 = 발송 × CTR × 주문CR × 객단가라, 나머지를 붙들면 목표 거래액에 필요한 CTR이
나눗셈 한 번으로 나온다. 위험한 건 그 숫자가 **달성 가능한 수준인지 말하지 않는 것**이다.
최근 1년 최고치보다 높은 값을 목표로 내걸면 그대로 실행 계획이 된다.

로컬 실행:
    python tests/test_target_ctr.py
"""
import json
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


def _funnel_sent(at):
    """④ 퍼널 차트의 첫 막대 라벨에서 발송 건수를 읽는다.

    AppTest는 plotly 요소에 `.value`를 안 준다 — `el.proto.spec`이 figure JSON 통째다.
    가로 막대가 이 페이지에 또 생길 수 있으니 **레이아웃 제목으로** 그 차트를 집는다.
    """
    for el in at.get("plotly_chart"):
        spec = json.loads(el.proto.spec)
        ttl = str(((spec.get("layout") or {}).get("title") or {}).get("text", ""))
        if "단계별 전환" not in ttl:
            continue
        txt = (spec.get("data") or [{}])[0].get("text") or []
        if txt:
            return int(str(txt[0]).replace(",", "").replace("건", ""))
    return None


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
def t_prior_window_does_not_overlap_current():
    """'직전 같은 기간'은 현재 창과 겹치면 안 된다.

    시차를 28일로 고정해 두면 8주·13주 창에서 비교 구간이 현재 창을 파고들어,
    '지금'과 '직전'이 같은 발송을 나눠 갖게 된다.
    """
    at = _open(synth(drop=0.5))
    sel = [s_ for s_ in at.selectbox if s_.label == "목표 기준"]
    assert sel and "직전 같은 기간" in list(sel[0].options), \
        f"직전 같은 기간 옵션이 없어요 — {sel[0].options if sel else None}"
    sel[0].set_value("직전 같은 기간")
    at.run()
    assert not at.exception, at.exception[0].value
    for _wk in ("최근 4주", "최근 8주", "최근 13주"):
        _w = [s_ for s_ in at.selectbox if s_.label == "현재 구간"][0]
        _w.set_value(_wk)
        at.run()
        assert not at.exception, f"{_wk}에서 실패: {at.exception[0].value}"
        _cap = [c for c in _texts(at) if "현재 " in c and "기준 " in c and "~" in c]
        assert _cap, f"{_wk}: 구간 캡션이 없어요"
        _m = re.findall(r"(\d{2}/\d{2}/\d{2})~(\d{2}/\d{2}/\d{2})", _cap[0])
        assert len(_m) >= 2, f"{_wk}: 구간을 못 읽었어요 — {_cap[0][:120]}"
        _cur = [pd.Timestamp("20" + x.replace("/", "-")) for x in _m[0]]
        _ref = [pd.Timestamp("20" + x.replace("/", "-")) for x in _m[1]]
        assert _ref[1] < _cur[0], \
            f"{_wk}: 직전 구간({_ref[0].date()}~{_ref[1].date()})이 현재({_cur[0].date()}~)와 겹쳐요"
        assert (_cur[1] - _cur[0]).days == (_ref[1] - _ref[0]).days, \
            f"{_wk}: 두 구간 길이가 달라요"


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


# ── ④ 퍼널 — 「전체 효율·추이」에서 옮겨 온 블록 ──────────────────────────
# 그 탭은 여섯 블록 중 다섯이 다른 화면과 겹쳐서 없앴고, 안 겹친 이 퍼널만 여기로 왔다.
# 옮기면서 **프레임이 바뀌면 값이 조용히 달라진다** — ①②③은 '최소 발송수'를 넘긴
# 캠페인(`fdf`)을 보는데 퍼널은 작은 발송까지 더한 `dff_all`을 봐야 전체 잔존이 된다.

@case
def t_funnel_counts_every_campaign_not_just_the_big_ones():
    """④ 퍼널만 '최소 발송수' 필터를 안 거친다.

    `dff_all` → `fdf`로 바꿔 놓으면 발송 합계가 작아지는데 **화면은 멀쩡히 뜬다**.
    픽스처는 최소 발송수(기본 5,000) **아래** 캠페인을 일부러 섞어야 둘이 갈린다 —
    전부 큰 발송이면 두 프레임이 같아서 규칙을 깨도 값이 안 변한다.
    """
    big = synth(weeks=10)
    small = big.head(60).copy()
    small["af"] = [f"SM{i:04d}" for i in range(len(small))]
    small["send"] = 1_000                          # 기본 최소 발송수 5,000 미만
    small["uv"] = 40
    small["cust"] = 12
    small["oc"] = 4
    small["amt"] = 4 * 130_000
    small["infl_cr"] = small["uv"] / small["send"]
    small["ord_cr"] = small["oc"] / small["uv"]
    store = pd.concat([big, small], ignore_index=True)

    at = _open(store)
    sent = _funnel_sent(at)
    assert sent is not None, "④ 퍼널 차트를 못 찾았어요"

    all_sum = int(store["send"].sum())
    big_sum = int(big["send"].sum())
    assert all_sum != big_sum, "픽스처로는 규칙을 못 깨요 — 작은 발송이 안 섞였어요"
    assert sent == all_sum, (
        f"퍼널 발송 합계가 {sent:,} — 전체 {all_sum:,}가 아니라 "
        f"최소 발송수 통과분 {big_sum:,}만 센 것 같아요")


@case
def t_funnel_says_it_skips_the_min_send_filter():
    """같은 화면의 다른 블록과 기준이 다르면 **화면이 그렇게 말해야** 한다."""
    at = _open(synth(weeks=10))
    txt = " ".join(_texts(at))
    assert "최소 발송수" in txt and "안 거쳐" in txt, \
        "퍼널이 최소 발송수를 안 거친다는 안내가 없어요"


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
