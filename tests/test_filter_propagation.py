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
import os
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


def perf_page(tab, weeks=60, seed=5, store=None, **fixed):
    """발송성과 화면 하나를 여는 함수를 만든다.

    픽스처가 **1년을 넘어야** 한다 — 60주 미만이면 전년 짝이 없어 「전년 동요일
    비교」의 소구 블록이 아예 안 그려지고, 계약이 덮는 게 0인 채로 초록이 된다.
    """
    store = synth_store(weeks=weeks, seed=seed) if store is None else store

    def _o(**sess):
        return _open(tab, store, **{**fixed, **sess})
    return _o


def copy_perf_store(weeks=60, seed=5):
    """소구 속성이 실제로 «붙는» 픽스처.

    `synth_store`의 제목은 「제목N 세일&특가 <단독>」이라 `KW` 사전에 거의 안 걸린다.
    그대로 쓰면 「문구 성과」가 '속성 비교에 필요한 표본이 없어요'로 끝나 **컨트롤
    아래에 값 블록이 하나도 없다** — 계약이 아무것도 못 보는 채로 초록이 될 뻔했고,
    `_snap`의 「블록이 하나도 안 잡혔으면 멈춘다」 가드가 그 자리에서 잡았다.
    """
    store = synth_store(weeks=weeks, seed=seed).copy()
    tpl = ["오늘만 최대 50% 할인 — 마감임박",
           "신상 입고 · 선착순 쿠폰 드려요",
           "베스트 인기상품 한정수량",
           "무료배송 이벤트 응모하세요",
           "장바구니에 담아두신 상품 특가"]
    store["title"] = [tpl[i % len(tpl)] for i in range(len(store))]
    # **본문도 길이를 흔든다.** `synth_store`의 본문은 전부 「본문 내용입니다」라
    # 본문길이 고유값이 1인데, 그러면 `qcut`이 구간을 0개 주고 「본문 길이 구간별」이
    # 값 없는 빈 차트가 된다 — 지표를 바꿔도 «둘 다 비어 있어» 안 따라오는 걸로 잡힌다.
    # 블록이 진짜로 따라오는지 보려면 픽스처가 그걸 반증할 수 있어야 한다.
    store["body"] = [("본문 내용입니다 " * (1 + i % 5)).strip() for i in range(len(store))]
    return store


def weekly_funnel_page(unit, store=None, orgcat=None):
    """주간보고 「02. 첫구매 퍼널별 상세 실적」을 여는 함수.

    저쪽 앱은 저장소가 **cwd 기준 CSV**라 하네스가 임시 폴더로 들어간다
    (`test_funnel_page._open`). 돌아오는 건 `check`가 맡는다 — 중간에 돌아오면
    다음 `at.run()`부터 저장소가 통째로 사라져 「데이터가 없어요」 화면을 보고
    통과해 버린다.
    """
    def _o(**sess):
        import test_funnel_page as F
        at = F._open(store=store, orgcat=orgcat, unit=unit)
        for k, v in sess.items():
            at.session_state[k] = v
        if sess:
            at.run()
        assert not at.exception, at.exception[0].value
        return at
    return _o


def _diff(a, b):
    """두 스냅샷을 **자리별로** 맞댄다. 블록 수가 달라지면 그것부터 말한다."""
    assert len(a) == len(b), (
        f"컨트롤을 바꿨더니 블록 개수가 달라졌어요 ({len(a)} → {len(b)}). "
        "블록이 생기거나 사라지면 자리 대조가 무의미해져요")
    same, moved = [], []
    for (la, va), (lb, vb) in zip(a, b):
        (same if va == vb else moved).append((la, lb))
    return same, moved


def check(tab, opener, ctrl_key, values, stays_below=(), follows_above=()):
    """컨트롤 하나의 «전파 계약»을 확인한다.

    `values` — 넣어 볼 두 값. 화면에서 읽어야 하면 `(at) -> (a, b)` 함수로 준다
    (범위 슬라이더처럼 선택지가 데이터에서 나오는 경우).
    `stays_below` — 아래인데 **안 바뀌어도 되는** 블록(이름 조각).
    `follows_above` — 위인데 **바뀌어도 되는** 블록. 둘 다 적은 게 실제로 화면에
    나와야 한다. 블록을 지우고 목록만 남으면 그 줄이 썩은 채로 규칙을 느슨하게
    만드는데, 증상이 없어서 영영 안 드러난다.
    """
    cwd = os.getcwd()
    try:
        if callable(values):
            values = values(opener())
        assert len(values) == 2, f"값 두 개가 필요해요 — {values}"
        snaps = [_snap(opener(**{ctrl_key: v}), ctrl_key) for v in values]
    finally:
        os.chdir(cwd)
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

# 「기준 기간 자체를 보여 주는 블록」은 **비교 기준을 안 따라간다.** 비교 기준은
# '무엇과 맞댈지'를 고르는 것이라, 기준 기간에 나간 문구·캠페인 목록은 그대로다.
# 그 자리가 `stays_below`에 반복해서 나오는 이유다.
_REF_ONLY = ("날짜|시간|카테고리",     # 「그날 나간 소재」 문구 표 (render_messages)
             "날짜|카테고리|제목",      # 「기준 기간에 나간 소재」
             "날짜|카테고리|브랜드")    # 캠페인 목록


@case
def t_yoy_metric_picker_drives_everything_below():
    """「비교 지표」(`p05b_met`) — 시간대 차트·표·일자 합산·소구 블록을 다 끈다.

    2026-10 사고 자리다. 소구 블록만 CTR에 박혀 있었고 제목에도 「소구 속성별 CTR」
    이라 적혀 있어 눈으로도 안 잡혔다.
    """
    moved, _ = check("전년 동요일 비교", perf_page("전년 동요일 비교"),
                     "p05b_met", ("CTR", "거래액"), stays_below=_REF_ONLY[:1])
    assert moved >= 4, f"따라간 블록이 {moved}개뿐이에요 — 픽스처가 너무 얕아요"


@case
def t_yoy_ratio_basis_drives_everything_below():
    """「비율 기준」(`p05b_basis`) — 합산↔캠페인 평균. 소구 블록까지 같이 간다."""
    check("전년 동요일 비교", perf_page("전년 동요일 비교"),
          "p05b_basis", ("합산", "캠페인 평균"), stays_below=_REF_ONLY[:1])


@case
def t_copy_metric_drives_everything_below():
    """「성과 지표」(`p02_metric`) — 소구·조합·길이·키워드·이모지를 다 끈다.

    옛 「키워드·이모지 성과」 탭을 이 화면으로 합치면서 생긴 자리다. 합치기 전엔
    두 화면이 각자 「성과 지표」를 들고 있어서 양쪽에서 따로 골라야 했는데,
    합치고 나면 **하나가 전부를 끌어야** 한다. 옮긴 블록이 자기 셀렉트를 그대로
    들고 있었으면 같은 화면이 두 지표를 말하게 된다.
    """
    moved, _ = check("문구 성과",
                     perf_page("문구 성과", store=copy_perf_store()),
                     "p02_metric", ("주문전환율", "거래액"))
    assert moved >= 4, f"따라간 블록이 {moved}개뿐이에요 — 픽스처가 너무 얕아요"


@case
def t_weekly_comparison_basis_drives_the_blocks_below():
    """「비교 기준」(`wr_wf_cmp_월`) — 증감 기여 분해 두 블록과 카테고리 표를 같이 끈다.

    처음엔 위젯을 탭 ① 안에 두고 아래 카테고리 블록은 안 따라가게 짰다가
    「전년 동월로 바꿔도 카테고리 표가 안 바뀐다」는 지적을 받은 자리다. 둘은 같은
    질문('어느 카테고리가 빠졌나')이라 기준이 갈리면 한 화면이 두 기간을 말한다.
    """
    moved, _ = check("주간보고", perf_page("주간보고", weeks=70, wr_unit="월별"),
                     "wr_wf_cmp_월", ("전월", "전년 동월"),
                     stays_below=_REF_ONLY[1:] + ("지표|당월 MTD",))
    assert moved >= 3, f"따라간 블록이 {moved}개뿐이에요"


@case
def t_weekly_comparison_basis_leaves_the_cards_above_alone():
    """**반대 방향도 계약이다.** 위 KPI 카드·「주요 지표 현황」·추이는 그대로여야 한다.

    위쪽은 「비교」 칩으로 셋을 한꺼번에 보는 화면이라 기준 하나를 고르는 개념이
    없다. 같이 바뀌면 어디서 바뀐 건지 되짚을 수가 없다. `check`가 `follows_above`
    를 비워 둔 채로 이미 그걸 보는데, 규칙이 눈에 띄게 **따로** 한 줄 둔다.
    """
    check("주간보고", perf_page("주간보고", weeks=70, wr_unit="월별"),
          "wr_wf_cmp_월", ("전월", "전년 동월"),
          stays_below=_REF_ONLY[1:] + ("지표|당월 MTD",), follows_above=())


# 일별 값을 **시간에 따라 흔든** 마스터 픽스처. 원본 `synth_store()`는 세그먼트마다
# 상수라(세그먼트×연도 고유값이 전부 1) 창을 바꿔도 값이 같을 수밖에 없다 — 그걸
# 그대로 쓰면 「안 따라온다」가 전부 픽스처 탓이라 계약이 아무것도 못 본다.
# **비율 지표는 안 건드린다** — 가입율 = 가입자수 ÷ 비회원트래픽 같은 항등식을 깨면
# 화면이 딴소리를 하고, 카운트만 흔들어도 비율은 저절로 보존된다.
_RATIO_METRICS = ("가입율", "당일가입CR", "첫구매 객단가", "상품CR")


def _wavy_funnel_store():
    import test_funnel_page as F
    st = F.synth_store().copy()
    day = st[st["gran"] == "일"]
    labs = (day[["label", "sortkey"]].drop_duplicates()
            .sort_values("sortkey")["label"].tolist())
    f = {l: 0.5 + 1.5 * (i / max(len(labs) - 1, 1)) for i, l in enumerate(labs)}
    m = (st["gran"] == "일") & (~st["metric"].isin(_RATIO_METRICS))
    st.loc[m, "value"] = [float(v) * f.get(l, 1.0)
                          for v, l in zip(st.loc[m, "value"], st.loc[m, "label"])]
    return st


def _wavy_orgcat():
    """조직×카테고리 픽스처도 같이 흔든다 — 원본은 **라벨마다 값이 같다**.

    `synth_orgcat()`은 노드(조직·카테고리·채널)로는 값이 갈리지만 **기간으로는
    상수**라, 일자 라벨 36개의 고유값이 1이다. 창을 좁혀도 일평균이 같은 값이라
    ⑤의 «레벨 표» 셋은 규칙을 지키든 안 지키든 값이 안 변한다 — 그래서 오래
    「⑤ 레벨 표는 창을 안 따라간다」로 읽혔는데 **앱이 아니라 픽스처 탓**이었다.
    같은 블록의 연중 추이 차트만 잡힌 건 그쪽 지문엔 x축 라벨이 같이 들어가서다.

    카운트 셋(거래액·고객수·상품UV)만 흔든다. 객단가·상품CR은 그대로 둬야
    `거래액 = 상품UV × 상품CR × 객단가`와 `거래액 = 고객수 × 객단가`가 둘 다
    살아남는다 — 셋을 같은 배수로 밀면 두 항등식이 저절로 보존된다.
    """
    import test_funnel_page as F
    oc = F.synth_orgcat().copy()
    day = oc[oc["gran"] == "일"]
    labs = (day[["label", "sortkey"]].drop_duplicates()
            .sort_values("sortkey")["label"].tolist())
    f = {l: 0.5 + 1.5 * (i / max(len(labs) - 1, 1)) for i, l in enumerate(labs)}
    m = (oc["gran"] == "일") & (~oc["metric"].isin(_RATIO_METRICS))
    oc.loc[m, "value"] = [float(v) * f.get(l, 1.0)
                          for v, l in zip(oc.loc[m, "value"], oc.loc[m, "label"])]
    # **자기 반증 가드** — 흔든 뒤에도 라벨별 값이 하나뿐이면 계약이 아무것도 못 본다.
    chk = oc[(oc["gran"] == "일") & (oc["metric"] == "첫구매 거래액")
             & (oc["org"] == "*TOTAL") & (oc["cat"] == "*TOTAL")
             & (oc["ch"] == F.W.ORGCAT_CH_ALL) & (oc["year"] == 2026)]
    assert chk["value"].nunique() > 1, "조직×카테고리 픽스처가 아직 평평해요"
    return oc


def _day_range(at):
    """「조회 기간」 선택지에서 «연중 전체»와 «최근 7칸»을 뽑는다."""
    sl = [x for x in at.get("select_slider")
          if getattr(x, "key", None) == "wr_fn_day"]
    assert sl, "「조회 기간」 슬라이더가 화면에 없어요 — 일자별이 아닌가요"
    o = list(sl[0].options)
    assert len(o) >= 8, f"선택지가 {len(o)}칸뿐이라 창을 못 좁혀요"
    return ((o[0], o[-1]), (o[-7], o[-1]))


@case
def t_funnel_lookup_window_scopes_the_blocks_below():
    """주간보고 「02」의 「조회 기간」 — 위에서 3주를 골라 놓고 아래만 연중이면 안 된다.

    세 번째 사고 자리다. 창은 `gran_ref`가 한 번 정하고(`clabel`·`win`) 아래는 받아
    쓰기만 한다. **예외 없이 16개 블록이 전부 따라간다.**

    오래 ⑤ 조직×카테고리의 «레벨 표» 셋이 예외로 적혀 있었는데 **앱이 아니라
    픽스처 탓이었다.** `synth_orgcat()`은 노드로는 값이 갈리지만 기간으로는 상수라
    (일자 라벨 36개의 고유값이 1), 창을 좁혀도 일평균이 같은 값이다 — 규칙을 지키든
    안 지키든 표가 안 변하니 「안 따라간다」로 읽혔다. 같은 블록의 연중 추이 차트만
    잡힌 건 그쪽 지문에 x축 라벨이 같이 들어가서고, 그래서 「픽스처 탓이 아니다」로
    잘못 읽었다. `_wavy_orgcat()`으로 흔들어 보니 셋 다 따라간다.
    """
    moved, _ = check(
        "주간보고 02", weekly_funnel_page("일자별", store=_wavy_funnel_store(),
                                         orgcat=_wavy_orgcat()),
        "wr_fn_day", _day_range)
    assert moved >= 16, f"따라간 블록이 {moved}개뿐이에요 — 픽스처가 너무 얕아요"


def main():
    want = sys.argv[1:]
    fails = []
    picked = [f for f in CASES if not want or any(w in f.__name__ for w in want)]
    # **하나도 안 걸리면 멈춘다.** 이름은 ASCII 인데 한글 조각으로 거르면 0건이
    # 뽑히고, 예전엔 그래도 「통과 ✅」가 찍혔다 — 아무것도 안 돌린 걸 초록으로 읽는다.
    if want and not picked:
        print(f"이름에 {want} 가 든 계약이 없어요 — "
              + ", ".join(f.__name__ for f in CASES))
        return 2
    for fn in picked:
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
