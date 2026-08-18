"""「6. 효율·피로도 › 발송량 최적 구간」 렌더·계산 테스트.

이 탭은 전사 MTD가 있어야 열리고, 손익분기 블록은 앱푸시 수신동의까지 있어야
계산된다. 스모크는 캠페인 데이터만 넣어 돌리므로 MTD 경로가 통째로 안 밟힌다 —
그래서 따로 둔다. 회귀·나눗셈이 들어가 있어 빈 데이터·0 분모에서 조용히 죽기 쉽다.

로컬 실행:
    python tests/test_send_volume_band.py
"""
import pathlib
import sys

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest        # noqa: E402

import send_perf_dashboard as S                 # noqa: E402
from smoke_pages import synth_store             # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300
TAB = "발송량 최적 구간"


def synth_mtd(days=200, seed=3):
    """일자별 전사 MTD — 발송을 늘리면 총거래액은 늘고 RPS는 떨어지게 만든다."""
    rng = np.random.default_rng(seed)
    end = pd.Timestamp.utcnow().tz_localize(None).normalize()
    dates = pd.date_range(end - pd.Timedelta(days=days - 1), end, freq="D")
    per = np.clip(rng.normal(3.2, 0.8, len(dates)), 1.2, 5.5)
    cust = rng.integers(400_000, 600_000, len(dates)).astype(float)
    send = per * cust
    rps = 900 - 80 * per + rng.normal(0, 40, len(dates))     # 더 보낼수록 건당 효율 하락
    rev = send * rps                                          # 총량은 그래도 우상향
    inflow = send * np.clip(0.02 - 0.002 * per, 0.004, None)
    pcust = cust * np.clip(0.004 - 0.0003 * per, 0.001, None)
    return pd.DataFrame(dict(
        date=dates.strftime("%Y-%m-%d"), perSend=per, revenue=rev, rps=rps,
        totalSend=send, customers=cust, ctr=inflow / cust, uniqueInflow=inflow * 0.8,
        totalInflow=inflow, visitPerPerson=rng.uniform(1.1, 1.6, len(dates)),
        purchaseCust=pcust, purchaseCnt=pcust * 1.2,
        purchasePerPerson=rng.uniform(1.0, 1.4, len(dates)),
        avgOrderVal=rev / (pcust * 1.2), unitPrice=rng.uniform(50_000, 90_000, len(dates)),
        mRevenue=rng.uniform(1e6, 3e6, len(dates)), pointM=rng.uniform(1e5, 5e5, len(dates)),
    ))


def synth_push(mtd, churn_slope=0.0, seed=5):
    """앱푸시 수신동의 일자별 — churn_slope로 '발송↑ → 이탈↑' 강도를 조절한다."""
    rng = np.random.default_rng(seed)
    d = pd.to_datetime(mtd["date"])
    removed = np.clip(3_000 + churn_slope * mtd["perSend"].values
                      + rng.normal(0, 200, len(d)), 0, None)
    added = np.clip(rng.normal(3_500, 300, len(d)), 0, None)
    return pd.DataFrame(dict(
        date=d, group="Total", consent=np.linspace(5e6, 5.2e6, len(d)),
        added=added, removed=removed, diff=added - removed, is_outlier=False))


def _open_tab(camp=None, mtd=None, push=None):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=10) if camp is None else camp
    if mtd is not None:
        at.session_state["mtd_store_df"] = mtd
    if push is not None:
        at.session_state["push_consent_df"] = push
    at.run()
    assert not at.exception, at.exception[0].value
    at.sidebar.radio[0].set_value("6. 효율·피로도")
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    assert subs, "하위탭 라디오를 찾지 못했어요"
    assert TAB in subs[0].options, f"{TAB} 탭이 없어요 — {subs[0].options}"
    subs[0].set_value(TAB)
    at.run()
    assert not at.exception, at.exception[0].value
    return at


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success):
        out += [str(e.value) for e in coll]
    return out


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_tab_exists_in_group():
    """페이지 그룹에 탭이 등록돼 있어야 스모크 커버리지에도 들어간다."""
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=10)
    at.run()
    at.sidebar.radio[0].set_value("6. 효율·피로도")
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    assert subs and TAB in subs[0].options, "탭이 하위탭 목록에 없어요"


@case
def t_renders_without_mtd():
    """MTD가 없으면 죽지 말고 올려 달라고 안내해야 한다."""
    at = _open_tab()
    assert any("MTD" in t for t in _texts(at)), "MTD 안내 문구가 없어요"


@case
def t_renders_with_mtd_only():
    """MTD만 있어도 ①·③은 그려지고, ②는 앱푸시를 요청해야 한다."""
    at = _open_tab(mtd=synth_mtd())
    txt = " ".join(_texts(at))
    assert "① " in txt and "③ " in txt, "①·③ 블록이 안 보여요"
    assert "앱푸시 수신동의 파일" in txt, "② 앱푸시 안내가 없어요"
    assert any(m.label.startswith("인당발송 +1건당") for m in at.metric), \
        f"회귀 KPI가 없어요 — {[m.label for m in at.metric]}"
    cols = [list(getattr(t.value, "columns", [])) for t in at.dataframe]
    assert any("인당발송 구간" in c for c in cols), f"구간 표가 없어요 — {cols}"
    assert any("발송유형" in c for c in cols), f"발송유형 표가 없어요 — {cols}"
    assert "일평균 거래액" in txt, "총량 기준 최적 구간 문장이 없어요"


@case
def t_slopes_have_expected_sign():
    """합성 데이터의 설계(총량↑·효율↓)가 화면 KPI 부호로 그대로 나와야 한다."""
    at = _open_tab(mtd=synth_mtd())
    got = {m.label: m.value for m in at.metric}
    rev = got.get("인당발송 +1건당 일 거래액", "")
    rps = got.get("인당발송 +1건당 RPS", "")
    assert not rev.startswith("-") and rev != "–", f"거래액 기울기가 음수/결측이에요 — {rev}"
    assert rps.startswith("-"), f"RPS 기울기가 음수여야 해요 — {rps}"


@case
def t_breakeven_needs_enough_overlap():
    """겹치는 날짜가 적으면 손익분기를 지어내지 말고 부족하다고 말해야 한다."""
    mtd = synth_mtd()
    push = synth_push(mtd).head(5)
    at = _open_tab(mtd=mtd, push=push)
    assert any("부족해요" in t for t in _texts(at)), "표본 부족 안내가 없어요"


@case
def t_flat_churn_reports_no_signal():
    """발송량과 이탈이 무관하면 '신호가 안 보인다'로 말해야 한다 (가짜 손익분기 금지)."""
    mtd = synth_mtd()
    at = _open_tab(mtd=mtd, push=synth_push(mtd, churn_slope=0.0))
    txt = " ".join(_texts(at))
    assert "이탈이 더 늘어나는 신호가 안 보여요" in txt, "무신호 판정이 안 나왔어요"
    assert "보다 크면 손해" not in txt, "신호가 없는데 손익분기를 계산했어요"


@case
def t_rising_churn_gives_breakeven():
    """발송량과 이탈이 같이 오르면 손익분기 금액을 내놔야 한다."""
    mtd = synth_mtd()
    at = _open_tab(mtd=mtd, push=synth_push(mtd, churn_slope=900.0))
    txt = " ".join(_texts(at))
    assert "보다 크면 손해" in txt, f"손익분기 문장이 없어요 — {txt[-400:]}"
    got = {m.label: m.value for m in at.metric}
    assert got.get("발송 +1건당 이탈 증가", "–") != "–", "이탈 기울기가 비어 있어요"


@case
def t_stype_ranking_uses_normalized_types():
    """발송유형은 norm_stype으로 묶은 뒤 비교해야 한다 (컨틴전시 A/B가 따로 놀면 안 됨)."""
    camp = synth_store(weeks=10)
    rng = np.random.default_rng(7)
    camp = camp.copy()
    camp["stype"] = rng.choice(["컨틴전시 A", "컨틴전시 B", "우수발송 1", "우수발송 3", "남은발송"],
                               len(camp))
    at = _open_tab(camp=camp, mtd=synth_mtd())
    tbls = [t for t in at.dataframe
            if "발송유형" in list(getattr(t.value, "columns", []))]
    assert tbls, "발송유형 표를 찾지 못했어요"
    types = set(tbls[0].value["발송유형"].astype(str))
    assert "컨틴" in types, f"컨틴전시 A/B가 '컨틴'으로 안 묶였어요 — {types}"
    assert not any(t.startswith("컨틴전시") for t in types), f"원 표기가 남아 있어요 — {types}"
    assert "우수발송" in types and "우수발송 1" not in types, f"우수발송 묶기 실패 — {types}"


@case
def t_top_type_sentence_uses_right_particle():
    """받침 없는 유형 이름에 '이'가 붙으면 안 된다 ('시그니처이 가장 나아요')."""
    camp = synth_store(weeks=10)
    rng = np.random.default_rng(9)
    camp = camp.copy()
    camp["stype"] = rng.choice(["시그니처", "컨틴"], len(camp))
    # 시그니처가 1등이 되도록 거래액을 몰아 준다
    camp.loc[camp["stype"] == "시그니처", "amt"] = camp["amt"] * 8
    at = _open_tab(camp=camp, mtd=synth_mtd())
    txt = " ".join(_texts(at))
    assert "시그니처이 가장" not in txt, "받침 없는 이름에 '이'가 붙었어요"
    assert "가장 나아요" in txt, f"1등 문장이 없어요 — {txt[-300:]}"


@case
def t_metric_switch_keeps_rendering():
    """판단 지표를 바꿔도 정상 렌더돼야 한다."""
    at = _open_tab(mtd=synth_mtd())
    sels = [s for s in at.selectbox if s.label == "판단 지표"]
    assert sels, f"'판단 지표' 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
    for opt in list(sels[0].options):
        tgt = [s for s in at.selectbox if s.label == "판단 지표"][0]
        tgt.set_value(opt)
        at.run()
        assert not at.exception, f"판단 지표={opt}에서 실패: {at.exception[0].value}"


@case
def t_policy_toggle_switches_scope():
    """정책 변경 이후 체크를 끄면 전체 기간으로 넓어져야 한다."""
    at = _open_tab(mtd=synth_mtd())
    box = [c for c in at.checkbox if "정책 변경" in c.label]
    assert box, f"정책 변경 체크박스가 없어요 — {[c.label for c in at.checkbox]}"
    assert box[0].value is True, "기본값이 켜짐이어야 해요"
    box[0].set_value(False)
    at.run()
    assert not at.exception, at.exception[0].value


@case
def t_norm_stype_groups_contingency():
    """헬퍼 단위 확인 — 화면과 같은 규칙으로 묶이는지."""
    assert S.norm_stype("컨틴전시 A") == "컨틴"
    assert S.norm_stype("우수발송 3") == "우수발송"
    assert S.norm_stype("우수발송 3", group=False) == "우수발송 3"
    assert S.norm_stype("") is None and S.norm_stype(np.nan) is None


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
    print(f"발송량 최적 구간 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
