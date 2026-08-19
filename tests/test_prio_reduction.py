"""우선순위 정규화(0순위→1순위)와 「발송 감축 효과」 탭 테스트.

이 탭은 '남은발송을 빼서 피로도를 낮췄으니 앞 순위 지표가 올랐을 것'이라는 가설을
전후로 검증한다. 조용히 틀리기 쉬운 곳이 두 군데다.

1. **구성 변화**: 남은발송이 2~4순위에만 몰려 있어서, 그냥 빼기만 해도 그 순위 평균이
   저절로 올라간다. 「남은발송 빼고 비교」가 실제로 양쪽에서 다 빼는지 확인한다.
2. **0순위**: 실백업에서 2건뿐이라 1순위에 합쳤다. 사이드바 세션에 남은 옛 '0' 선택이
   multiselect를 예외로 죽이지 않아야 한다.

로컬 실행:
    python tests/test_prio_reduction.py
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
TAB = "발송 감축 효과"
CUT = pd.Timestamp("2026-06-15")


def synth_cut_store(seed=4, lift=0.0, remain_prio=(2, 3, 4)):
    """감축 시점 전후 데이터.

    감축 전에는 2~4순위 발송의 절반이 남은발송이고, 감축 후에는 남은발송이 사라진다.
    남은발송은 CTR이 낮게 깔려 있어, 빼기만 해도 그 순위 평균이 올라간다(구성 효과).
    lift>0이면 감축 후 '남은발송이 아닌' 발송의 CTR을 실제로 끌어올린다(진짜 효과).
    """
    rng = np.random.default_rng(seed)
    rows = []
    for d in pd.date_range(CUT - pd.Timedelta(days=60), CUT + pd.Timedelta(days=45), freq="D"):
        after = d >= CUT
        for prio in (1, 2, 3, 4):
            for j in range(3):
                remain = (not after) and prio in remain_prio and j == 0
                base = {1: 0.040, 2: 0.030, 3: 0.025, 4: 0.020}[prio]
                ctr = 0.006 if remain else base * (1 + (lift if after else 0.0))
                ctr = float(max(rng.normal(ctr, ctr * 0.10), 0.001))
                send = int(rng.integers(40_000, 60_000))
                uv = max(int(send * ctr), 1)
                oc = max(int(uv * 0.05), 1)
                rows.append(dict(
                    date=d.strftime("%Y%m%d"), af=f"AP{d:%m%d}{prio}{j}",
                    hour="1000", target="전체",
                    stype="남은발송" if remain else "기본발송",
                    bpu="1BPU", prio=str(prio), cat="패션", attr="정상", owner="김",
                    brand="닥스", send=send, uv=uv, visit=uv + 10, cust=uv // 3,
                    oc=oc, amt=int(oc * 130_000),
                    infl_cr=uv / send, ord_cr=oc / uv, promo="",
                    title=f"제목{prio}", body="본문", matched=True))
    return pd.DataFrame(rows)


def _open_tab(store, tab=TAB):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = store
    at.run()
    assert not at.exception, at.exception[0].value
    at.sidebar.radio[0].set_value("11. 개선 효과 검증")
    at.run()
    subs = [r for r in at.radio if r.label != "페이지"]
    assert subs and tab in subs[0].options, f"{tab} 탭이 없어요 — {subs[0].options if subs else None}"
    subs[0].set_value(tab)
    at.run()
    assert not at.exception, at.exception[0].value
    return at


def synth_slot_store(seed=6):
    """컨틴 구좌와 영업(BPU) 세일즈 푸시가 같이 있는 데이터 — 평일 16시 슬롯 포함.

    smoke_pages의 합성 데이터는 16시 발송도 컨틴도 없어서 컨틴 탭의 본문 경로를
    통째로 안 밟는다 (실제로 그 경로에 남아 있던 NameError를 스모크가 못 잡았다).
    """
    rng = np.random.default_rng(seed)
    end = pd.Timestamp.utcnow().tz_localize(None).normalize()
    # 컨틴 탭은 정책 변경일(2026-08-01) 이후만 본다 — 그 뒤로 넉넉히 깔아 둔다
    _start = max(pd.Timestamp(S.POLICY_CHANGE_DATE), end - pd.Timedelta(days=180))
    rows = []
    for d in pd.date_range(_start, end, freq="D"):
        if d.weekday() >= 5:
            continue
        for stype, bpu, hour in (("컨틴", "1BPU", "1600"), ("기본발송", "2BPU", "1600"),
                                 ("기본발송", "3BPU", "1000"), ("기본발송", "마케팅", "2000")):
            send = int(rng.integers(30_000, 80_000))
            uv = max(int(send * float(rng.uniform(0.02, 0.05))), 1)
            oc = max(int(uv * float(rng.uniform(0.004, 0.012))), 1)
            rows.append(dict(
                date=d.strftime("%Y%m%d"), af=f"AF{d:%m%d}{hour}{bpu}", hour=hour,
                target="전체", stype=stype, bpu=bpu, prio="1", cat="패션", attr="정상",
                owner="김", brand="닥스", send=send, uv=uv, visit=uv + 5, cust=uv // 3,
                oc=oc, amt=int(oc * 130_000), infl_cr=uv / send, ord_cr=oc / uv,
                promo="", title=f"제목{hour}", body="본문", matched=True))
    return pd.DataFrame(rows)


def _texts(at):
    out = []
    for coll in (at.markdown, at.caption, at.info, at.warning, at.success):
        out += [str(e.value) for e in coll]
    return out


def _cmp_table(at):
    for t in at.dataframe:
        cols = list(getattr(t.value, "columns", []))
        if "우선순위" in cols and "변화율" in cols:
            return t.value
    return None


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_contin_tab_excludes_pre_policy():
    """2025년 컨틴전시(주말 17시)는 성격이 다른 구좌라 이 화면에 들어오면 안 된다."""
    store = synth_slot_store()
    old = store.head(40).copy()                          # 정책 변경 전 옛 컨틴
    old["date"] = "20250715"
    old["stype"] = "컨틴전시 A"
    old["af"] = "OLD" + old["af"]
    at = _open_tab(pd.concat([old, store], ignore_index=True), tab="컨틴 구좌 효율")
    txt = " ".join(_texts(at))
    assert "이전 발송" in txt and "뺐어요" in txt, f"제외 안내가 없어요 — {txt[:300]}"
    labs = {m.label: m.value for m in at.metric}
    _n = int(str(labs["컨틴"]).replace("건", "").replace(",", ""))
    _want = len(store[store["stype"] == "컨틴"])
    assert _n == _want, f"컨틴 표본이 {_n}건 — 옛 컨틴을 빼면 {_want}건이어야 해요"


@case
def t_contin_tab_compares_against_sales_bpu():
    """컨틴 탭 ①은 컨틴 구좌와 영업(BPU) 요청 세일즈 푸시를 맞대야 한다.

    카테고리 순위만 봐서는 '컨틴에 뭘 넣을까'는 알아도 '컨틴이 나은가'는 알 수 없다.
    """
    at = _open_tab(synth_slot_store(), tab="컨틴 구좌 효율")
    txt = " ".join(_texts(at))
    assert "컨틴 구좌 vs 영업 세일즈 푸시" in txt, f"비교 블록이 없어요 — {txt[:300]}"
    assert "카테고리가 잘 됐나" not in txt, "카테고리 블록이 아직 남아 있어요"
    labs = [m.label for m in at.metric]
    assert "컨틴" in labs and "영업 세일즈 푸시" in labs, f"비교 KPI가 없어요 — {labs}"
    tbl = [t.value for t in at.dataframe
           if "컨틴" in list(getattr(t.value, "columns", []))]
    assert tbl, "비교 표를 찾지 못했어요"
    assert set(tbl[0]["지표"].astype(str)) >= {"CTR", "주문CR", "RPS", "객단가"}, \
        sorted(set(tbl[0]["지표"].astype(str)))


@case
def t_contin_sales_group_excludes_non_sales_bpu():
    """영업 세일즈 푸시는 1~4BPU만 — 마케팅·편성 요청분이 섞이면 안 된다."""
    store = synth_slot_store()
    at = _open_tab(store, tab="컨틴 구좌 효율")
    _p = [s_ for s_ in at.selectbox if s_.label == "기간"]
    assert _p, f"기간 셀렉트가 없어요 — {[s_.label for s_ in at.selectbox]}"
    assert "2026-08 이후 전체" in list(_p[0].options), _p[0].options
    _p[0].set_value("2026-08 이후 전체")
    at.run()
    assert not at.exception, at.exception[0].value
    labs = {m.label: m.value for m in at.metric}
    _n = int(str(labs["영업 세일즈 푸시"]).replace("건", "").replace(",", ""))
    _want = len(store[store["bpu"].isin(["1BPU", "2BPU", "3BPU", "4BPU"])
                      & (store["stype"] != "컨틴")
                      & (store["date"] >= S.POLICY_CHANGE_DATE)])
    assert _n == _want, f"영업 표본이 {_n}건 — 1~4BPU 기준이면 {_want}건이어야 해요"
    assert _n < len(store), "전체가 다 들어갔어요 — 마케팅 요청분이 안 걸러졌어요"


@case
def t_norm_prio_merges_zero():
    """0순위는 1순위로. 못 읽는 값은 None."""
    assert S.norm_prio(0) == 1 and S.norm_prio("0") == 1
    assert S.norm_prio(1) == 1 and S.norm_prio("1.0") == 1
    assert S.norm_prio(3) == 3 and S.norm_prio("  4 ") == 4
    assert S.norm_prio("") is None and S.norm_prio(None) is None and S.norm_prio("없음") is None


@case
def t_finalize_adds_prio_g():
    """_finalize가 prio_g를 붙이고, 원본 prio는 건드리지 않는다."""
    d = pd.DataFrame([dict(date="20260601", prio="0", send=1, uv=1, visit=1, cust=1,
                           oc=1, amt=1, hour="1000")])
    r = S._finalize(d)
    assert int(r.loc[0, "prio_g"]) == 1, r["prio_g"].tolist()
    assert str(r.loc[0, "prio"]) == "0", "원본 prio가 바뀌었어요"


@case
def t_prio_series_falls_back_without_prio_g():
    """prio_g가 없는 프레임(구버전 캐시·옛 백업)에서도 우선순위를 되살려야 한다.

    `prepare_raw`가 @st.cache_data라 `_finalize`를 고쳐도 캐시는 그걸 모른다. 옛 프레임이
    남으면 우선순위 필터·9번 페이지·감축 효과 탭이 통째로 비고, 증상은 '20건 미만'으로만
    보여서 원인이 안 드러난다.
    """
    d = pd.DataFrame({"prio": ["0", "1", "2", "", "3.0"]})
    got = [None if pd.isna(v) else int(v) for v in S.prio_series(d)]
    assert got == [1, 1, 2, None, 3], got
    # prio_g가 있지만 통째로 비어 있어도(구버전 저장본) 원본에서 복구한다
    d2 = d.assign(prio_g=np.nan)
    got2 = [None if pd.isna(v) else int(v) for v in S.prio_series(d2)]
    assert got2 == [1, 1, 2, None, 3], got2
    # 정상 프레임이면 prio_g를 그대로 쓴다
    d3 = pd.DataFrame({"prio": ["9", "9"], "prio_g": [1, 2]})
    assert [int(v) for v in S.prio_series(d3)] == [1, 2]
    assert len(S.prio_series(pd.DataFrame())) == 0


@case
def t_cache_marker_mentions_prio():
    """_finalize에 파생을 추가하면 TAGSET_VER 표식도 같이 올려야 한다.

    표식을 안 올리면 캐시가 옛 프레임을 그대로 돌려준다 — hour:norm1이 같은 이유로 붙어 있다.
    실제로 prio_g를 넣고 표식을 안 올려서 화면이 통째로 빈 적이 있다.
    """
    src = (ROOT / "send_perf_dashboard.py").read_text(encoding="utf-8")
    _i = src.index("TAGSET_VER = ")
    _marker = src[_i:_i + 400]
    assert "prio:" in _marker, f"TAGSET_VER에 prio 표식이 없어요 — {_marker[:200]}"


@case
def t_zero_prio_absent_from_filter():
    """사이드바 우선순위 선택지에 0순위가 남아 있으면 안 된다."""
    store = synth_store(weeks=8)
    store = store.copy()
    store.loc[store.index[:5], "prio"] = "0"
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = store
    at.run()
    assert not at.exception, at.exception[0].value
    sels = [m for m in at.sidebar.multiselect if m.label == "우선순위"]
    assert sels, f"우선순위 필터가 없어요 — {[m.label for m in at.sidebar.multiselect]}"
    assert "0" not in list(sels[0].options), f"0순위가 남아 있어요 — {sels[0].options}"


@case
def t_stale_zero_selection_does_not_crash():
    """세션에 남은 옛 '0' 선택이 multiselect를 죽이면 안 된다 (guard_multi)."""
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=8)
    at.session_state["flt_prio"] = ["0"]                  # 합치기 전 세션이 그대로 남은 상황
    at.run()
    assert not at.exception, at.exception[0].value
    assert at.session_state["flt_prio"] == [], "선택지 밖 값이 그대로 남았어요"


@case
def t_tab_renders_and_confirms_premise():
    """전제(남은발송 감소)를 데이터에서 확인하고 화면에 말해야 한다."""
    at = _open_tab(synth_cut_store())
    txt = " ".join(_texts(at))
    assert "① " in txt and "④ " in txt, "블록이 다 안 보여요"
    assert "전제는 확인됐어요" in txt, f"전제 확인 문장이 없어요 — {txt[:400]}"
    labels = [m.label for m in at.metric]
    assert "남은발송 비중" in labels, f"전제 KPI가 없어요 — {labels}"


@case
def t_like_for_like_removes_remainder_both_sides():
    """「남은발송 빼고 비교」는 양쪽 구간에서 다 빼야 한다.

    한쪽만 빼면 감축 전 표본만 줄어 비교가 통째로 기울어진다. 남은발송이 없던
    1순위는 켜든 끄든 값이 같아야 하고, 남은발송이 있던 2순위는 달라져야 한다.
    """
    at = _open_tab(synth_cut_store())
    box = [c for c in at.checkbox if "남은발송" in c.label]
    assert box, f"체크박스가 없어요 — {[c.label for c in at.checkbox]}"
    assert box[0].value is True, "기본값이 켜짐이어야 해요 (구성 효과 방어)"
    on = _cmp_table(at)
    assert on is not None, "비교 표를 찾지 못했어요"

    box[0].set_value(False)
    at.run()
    assert not at.exception, at.exception[0].value
    off = _cmp_table(at)

    def _pre(t, p):
        r = t[t["우선순위"] == p]
        return float(r.iloc[0][[c for c in t.columns if c.startswith("전 ")][-1]]) if len(r) else np.nan

    assert abs(_pre(on, "1순위") - _pre(off, "1순위")) < 1e-9, \
        "남은발송이 없던 1순위 값이 바뀌었어요 — 필터가 엉뚱한 데 걸렸어요"
    assert abs(_pre(on, "2순위") - _pre(off, "2순위")) > 1e-9, \
        "남은발송이 섞인 2순위 값이 그대로예요 — 감축 전에서 안 빠졌어요"
    assert _pre(off, "2순위") < _pre(on, "2순위"), \
        "남은발송을 넣으면 감축 전 CTR이 더 낮아야 해요 (합성 데이터 전제)"


@case
def t_composition_effect_is_not_read_as_improvement():
    """진짜 개선이 없는데 구성 변화만 있으면 '가설 지지'라고 하면 안 된다."""
    at = _open_tab(synth_cut_store(lift=0.0))       # 남은발송만 사라지고 실효율은 그대로
    txt = " ".join(_texts(at))
    assert "가설을 지지해요" not in txt, \
        f"구성 효과를 개선으로 읽었어요 — {txt[-500:]}"


@case
def t_real_lift_is_detected():
    """감축 후 실제로 효율이 오르면 '가설을 지지해요'가 떠야 한다."""
    at = _open_tab(synth_cut_store(lift=0.35))
    txt = " ".join(_texts(at))
    assert "가설을 지지해요" in txt, f"진짜 개선을 못 잡았어요 — {txt[-500:]}"


@case
def t_windows_are_equal_length():
    """전후 창은 같은 길이로 잘라야 한다 (긴 쪽이 유리해지는 것 방지)."""
    at = _open_tab(synth_cut_store())
    caps = [c for c in _texts(at) if "같은" in c and "길이로 잘랐어요" in c]
    assert caps, f"창 길이 캡션이 없어요 — {_texts(at)[:6]}"
    import re
    m = re.search(r"(\d{4}-\d{2}-\d{2}) ~ .*?(\d{4}-\d{2}-\d{2}).*?(\d{4}-\d{2}-\d{2}) ~ .*?(\d{4}-\d{2}-\d{2})",
                  caps[0])
    assert m, f"날짜를 못 읽었어요 — {caps[0]}"
    b0, b1, a0, a1 = (pd.Timestamp(m.group(i)) for i in (1, 2, 3, 4))
    assert (b1 - b0).days == (a1 - a0).days, f"창 길이가 달라요 — 전 {(b1-b0).days}일 / 후 {(a1-a0).days}일"
    assert b1 < a0, "전 구간이 감축 시점을 넘겼어요"


@case
def t_metric_switch_keeps_rendering():
    """지표를 바꿔도 정상 렌더돼야 한다."""
    at = _open_tab(synth_cut_store())
    sels = [s for s in at.selectbox if s.label == "지표"]
    assert sels, f"'지표' 셀렉트가 없어요 — {[s.label for s in at.selectbox]}"
    assert "CTR" in str(sels[0].value), f"기본 지표가 CTR이어야 해요 — {sels[0].value}"
    for opt in list(sels[0].options):
        tgt = [s for s in at.selectbox if s.label == "지표"][0]
        tgt.set_value(opt)
        at.run()
        assert not at.exception, f"지표={opt}에서 실패: {at.exception[0].value}"


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
    print(f"우선순위·발송 감축 효과 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
