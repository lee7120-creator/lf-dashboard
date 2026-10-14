"""사이드바 필터 프리셋 — 저장·복원·백업, 그리고 «조용히 틀리는» 자리들.

자주 쓰는 필터 조합(「영업1 브랜드만」·「컨틴 16시만」)을 매번 손으로 다시 맞추던 걸
이름 붙여 저장한다. 여기서 조용히 틀리는 방식이 넷이다.

1. **새 필터를 `FLT_KEYS`에 안 넣는다** → 그 칸만 저장·복원에서 빠진다. 화면은 멀쩡히
   뜨고 「프리셋을 불러왔는데 브랜드만 안 돌아오네」로만 보인다.
2. **복원이 합치기가 된다** → 프리셋 A를 눌렀다가 B를 누르면 «둘을 섞은» 상태가 된다.
   숫자만 틀리고 화면은 멀쩡하다.
3. **지금 없는 값이 세션에 들어온다** → multiselect가 그 자리에서 예외로 죽는다.
   사이드바라 페이지가 전부 같이 내려간다(우선순위 0순위 때 겪은 그 자리).
4. **뺀 값을 말 안 한다** → 「프리셋이 다르게 복원됐다」로만 보인다.

로컬 실행:
    python tests/test_filter_preset.py
    python tests/test_filter_preset.py t_loading     # 이름이 든 케이스만
"""
import ast
import datetime
import io
import json
import os
import pathlib
import shutil
import sys
import tempfile
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from streamlit.testing.v1 import AppTest            # noqa: E402

import send_perf_dashboard as A                     # noqa: E402
from smoke_pages import group_of, synth_store       # noqa: E402

APP = str(ROOT / "send_perf_dashboard.py")
TIMEOUT = 300
SRC = pathlib.Path(APP).read_text(encoding="utf-8")

# 값을 가진 위젯만 — 버튼·다운로드는 필터가 아니라서 뺀다.
VALUE_WIDGETS = {
    "multiselect", "selectbox", "text_input", "number_input", "checkbox",
    "date_input", "radio", "slider", "select_slider", "pills",
    "segmented_control", "text_area", "toggle",
}

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# ══════════════════════════════════════════════════════════════════════
# 순수 함수 — 앱을 안 띄우고 본다
# ══════════════════════════════════════════════════════════════════════

@case
def t_round_trip_keeps_every_kind_of_value():
    """위젯 값은 꼴이 제각각이다 — 문자열·불리언·정수·목록·날짜 튜플.

    날짜만 JSON에 못 실려 ISO로 오가는데, 되돌릴 때 문자열로 두면 `date_input`이
    예외로 죽는다. 꼴마다 한 칸씩 넣어 **원래 값과 같은지**까지 본다.
    """
    want = {
        "flt_q": "세일",
        "flt_matched": False,
        "flt_minsend": 0,
        "flt_date": (datetime.date(2026, 1, 1), datetime.date(2026, 2, 1)),
        "flt_st": ["컨틴", "우수발송"],
        "flt_tags": [],
        "tags_mode": "하나라도(OR)",
    }
    got = A.preset_read(A.preset_dump(dict(want)))
    assert got == want, f"왕복이 값을 바꿨어요\n  넣은 값: {want}\n  나온 값: {got}"
    for k, v in want.items():
        assert type(got[k]) is type(v), \
            f"{k}의 꼴이 달라졌어요 — {type(v).__name__} → {type(got[k]).__name__}"


@case
def t_dump_does_not_invent_cells_that_are_not_there():
    """세션에 없는 칸은 지어내지 않는다 — **이게 «갈아 끼우기»의 전제다.**

    운영에선 위젯이 한 번 그려지면 그 키가 세션에 남아서 프리셋이 대개 전 칸을
    싣는다. 그래서 복원이 합치기여도 평소엔 티가 안 난다. 칸이 **빠진** 프리셋
    (그 필터가 생기기 전에 저장했거나 옛 백업에서 복원한 것)에서만 둘이 갈리고,
    아래 `t_loading_a_preset_replaces_rather_than_merges`가 그걸로 본다.
    """
    d = json.loads(A.preset_dump({"flt_q": "세일"}))
    assert set(d) == {"flt_q"}, f"없는 칸까지 지어냈어요: {sorted(d)}"


@case
def t_unknown_keys_are_dropped():
    """지금 없는 필터가 든 옛 프리셋도 열려야 한다 — 모르는 칸만 버린다."""
    got = A.preset_read(json.dumps({"flt_q": "세일", "flt_없어진필터": [1, 2]}))
    assert got == {"flt_q": "세일"}, got


@case
def t_broken_payload_opens_as_empty_not_an_exception():
    """저장소가 손상돼도 사이드바가 죽으면 안 된다 — 사이드바는 전 페이지 공용이다."""
    for bad in ("", "깨진 JSON", "[1, 2]", "null", None):
        assert A.preset_read(bad) == {}, f"{bad!r} 에서 {A.preset_read(bad)!r}"


@case
def t_a_broken_date_loses_only_that_cell():
    """날짜 한 칸이 깨졌다고 프리셋이 통째로 안 열리면 안 된다."""
    got = A.preset_read(json.dumps({"flt_q": "세일", "flt_date": ["어제", "오늘"]}))
    assert got == {"flt_q": "세일"}, got


# ══════════════════════════════════════════════════════════════════════
# 정적 규칙 — 앱 소스를 읽어서 본다
# ══════════════════════════════════════════════════════════════════════

def _widget_keys():
    """앱이 실제로 쓰는 «필터 위젯 키» 집합.

    `st.<위젯>(..., key="flt_x")`의 리터럴과, 가드를 거치는 `_msel(라벨, 키, …)`의
    둘째 인자를 같이 모은다. 버튼은 `VALUE_WIDGETS`에 없어 저절로 빠진다.
    """
    found = set()
    for n in ast.walk(ast.parse(SRC)):
        if not isinstance(n, ast.Call):
            continue
        f = n.func
        if isinstance(f, ast.Attribute) and f.attr in VALUE_WIDGETS:
            for kw in n.keywords:
                if kw.arg == "key" and isinstance(kw.value, ast.Constant) \
                        and isinstance(kw.value.value, str):
                    found.add(kw.value.value)
        elif isinstance(f, ast.Name) and f.id == "_msel" and len(n.args) >= 2 \
                and isinstance(n.args[1], ast.Constant):
            found.add(n.args[1].value)
    return {k for k in found if k.startswith("flt_") or k == "tags_mode"}


@case
def t_every_filter_widget_key_is_registered():
    """`FLT_KEYS`와 앱의 필터 위젯이 **정확히 같아야** 한다.

    한쪽만 늘면 그 필터가 저장·복원·「필터 전체 초기화」에서 조용히 빠지고,
    한쪽만 줄면 없는 칸을 세션에 심는다. 규약은 **키에 `flt_` 접두어**를 붙이는
    것이다(`tags_mode`는 그 규약 이전부터 있던 하나라 이름으로 같이 본다).
    """
    have, want = _widget_keys(), set(A.FLT_KEYS)
    assert have == want, (
        f"\n  앱에만 있음(= FLT_KEYS에 안 넣음): {sorted(have - want)}"
        f"\n  FLT_KEYS에만 있음(= 없어진 필터): {sorted(want - have)}")


@case
def t_sidebar_multiselects_go_through_the_guard():
    """필터 다중선택은 전부 `_msel`을 거쳐야 한다 — 거기서 `guard_multi`가 돈다.

    가드 없이 만들면 프리셋에서 돌아온 «지금 없는 값»이 그 자리에서 예외를 내고,
    사이드바라 페이지가 전부 같이 내려간다.
    """
    bare = []
    for n in ast.walk(ast.parse(SRC)):
        if not (isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                and n.func.attr == "multiselect"):
            continue
        for kw in n.keywords:
            if kw.arg == "key" and isinstance(kw.value, ast.Constant) \
                    and str(kw.value.value).startswith("flt_"):
                bare.append((kw.value.value, n.lineno))
    assert not bare, f"가드를 안 거친 필터 다중선택: {bare} — `_msel`을 쓰세요"


@case
def t_backup_routes_presets_before_the_campaign_fallback():
    """복원 갈래에서 `preset`이 **캠페인 폴백(`else`)보다 앞**에 있어야 한다.

    `_apply_backup`의 마지막 `else`가 캠페인이라, 안 걸린 CSV는 전부 거기로 간다 —
    `presets.csv`가 캠페인 파서로 흘러 들어가면 엉뚱한 저장소를 건드린다.
    """
    i_p, i_c = SRC.find('elif "preset" in base'), SRC.find("m = merge_store(stored, df)")
    assert i_p > 0, "복원 갈래에 preset이 없어요"
    assert i_p < i_c, "preset 갈래가 캠페인 폴백보다 뒤에 있어요"
    assert '_z.writestr("presets.csv"' in SRC, "통합 백업 ZIP에 프리셋이 안 실려요"


# ══════════════════════════════════════════════════════════════════════
# 앱을 띄워서 본다
# ══════════════════════════════════════════════════════════════════════

def _btn(at, key):
    for b in at.button:
        if getattr(b, "key", None) == key:
            return b
    raise AssertionError(f"버튼 {key!r} 이(가) 화면에 없어요")


def _sv(at, key, default=None):
    try:
        return at.session_state[key]
    except Exception:                                     # noqa: BLE001
        return default


def _open(**sess):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    at.session_state["camp_store"] = synth_store(weeks=8)
    for k, v in sess.items():
        at.session_state[k] = v
    at.run()
    assert not at.exception, at.exception[0].value
    return at


def _in_tmp(fn):
    """저장소 CSV가 저장소 폴더를 더럽히지 않게 임시 폴더에서 돌린다."""
    cwd, tmp = os.getcwd(), tempfile.mkdtemp()
    try:
        os.chdir(tmp)
        return fn()
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)


def _brands(at):
    """지금 화면의 브랜드 선택지 — 픽스처마다 달라서 읽어 쓴다."""
    for m in at.multiselect:
        if m.label == "브랜드":
            return list(m.options)
    raise AssertionError("브랜드 필터가 화면에 없어요")


@case
def t_saving_captures_the_current_filters():
    """지금 걸린 필터가 그대로 저장돼야 한다."""
    def _go():
        at = _open(flt_q="세일", flt_minsend=0)
        at.text_input(key="preset_name").set_value("내조합").run()
        _btn(at, "preset_put").click().run()
        assert not at.exception, at.exception[0].value
        saved = _sv(at, "flt_presets", {})
        assert "내조합" in saved, f"저장이 안 됐어요 — {sorted(saved)}"
        got = A.preset_read(saved["내조합"])
        assert got.get("flt_q") == "세일", got
        assert got.get("flt_minsend") == 0, got
        assert os.path.exists(A.PRESET_STORE), "저장소 파일이 안 생겼어요"
    _in_tmp(_go)


@case
def t_loading_restores_what_was_saved():
    """저장했다 다른 값으로 바꾼 뒤 불러오면 저장 시점으로 돌아와야 한다."""
    def _go():
        at = _open(flt_q="세일", flt_minsend=0)
        at.text_input(key="preset_name").set_value("내조합").run()
        _btn(at, "preset_put").click().run()

        at.session_state["flt_q"] = "딴거"
        at.session_state["flt_minsend"] = 9999
        at.run()
        _btn(at, "preset_load").click().run()
        assert not at.exception, at.exception[0].value
        assert _sv(at, "flt_q") == "세일", f"검색어가 안 돌아왔어요 — {_sv(at, 'flt_q')!r}"
        assert _sv(at, "flt_minsend") == 0, f"최소 발송수가 안 돌아왔어요 — {_sv(at, 'flt_minsend')}"
    _in_tmp(_go)


@case
def t_loading_a_preset_replaces_rather_than_merges():
    """프리셋 A → B 를 눌렀을 때 A에서 고른 값이 남으면 안 된다.

    **처음엔 버튼으로 두 프리셋을 저장해 봤는데 그 주입이 안 잡혔다.** 위젯이 한 번
    그려지면 그 키가 세션에 남아서 B에도 `flt_brand2: []`가 실리고, 합치기여도
    덮어써져 값이 같아진다 — 규칙을 깨도 결과가 안 변하는 픽스처였다.

    둘이 갈리는 건 **칸이 빠진 프리셋**뿐이다(그 필터가 생기기 전에 저장했거나 옛
    백업에서 복원한 것). 그래서 payload를 직접 심어 B에서 브랜드 칸을 뺀다.
    합치기로 짜면 화면은 멀쩡히 뜨고 **브랜드 필터만 A인 채로 남는다.**
    """
    def _go():
        at = _open()
        br = _brands(at)
        assert br, "픽스처에 브랜드가 없어요 — 이 픽스처론 규칙을 못 깨요"
        at.session_state["flt_presets"] = {
            "A": json.dumps({"flt_brand2": [br[0]], "flt_q": "에이"}, ensure_ascii=False),
            "B": json.dumps({"flt_q": "비"}, ensure_ascii=False),   # ← 브랜드 칸이 없다
        }
        at.run()

        at.selectbox(key="preset_pick").set_value("A").run()
        _btn(at, "preset_load").click().run()
        assert _sv(at, "flt_brand2") == [br[0]], \
            f"A가 안 걸렸어요 — {_sv(at, 'flt_brand2')}"

        at.selectbox(key="preset_pick").set_value("B").run()
        _btn(at, "preset_load").click().run()
        assert not at.exception, at.exception[0].value
        assert _sv(at, "flt_q") == "비", _sv(at, "flt_q")
        assert _sv(at, "flt_brand2") == [], \
            (f"A에서 고른 브랜드가 B에 남았어요 — {_sv(at, 'flt_brand2')}. "
             "불러오기는 «덮어쓰기»가 아니라 «갈아 끼우기»예요")
    _in_tmp(_go)


@case
def t_dropped_values_are_announced():
    """지금 데이터에 없는 값이 든 프리셋 — 죽지 말고, 뭘 뺐는지 말해야 한다."""
    def _go():
        at = _open()
        br = _brands(at)
        at.session_state["flt_presets"] = {
            "옛조합": json.dumps({"flt_brand2": [br[0], "없어진브랜드"]}, ensure_ascii=False)}
        at.run()
        _btn(at, "preset_load").click().run()
        assert not at.exception, \
            f"없는 값이 든 프리셋에서 죽었어요 — {at.exception[0].value}"
        assert _sv(at, "flt_brand2") == [br[0]], _sv(at, "flt_brand2")
        warn = " ".join(w.value for w in at.warning)
        assert "없어진브랜드" in warn, \
            f"뺀 값을 화면에서 말하지 않아요 — 경고: {warn[:200]!r}"
        assert "브랜드" in warn, f"어느 필터인지 안 밝혀요 — {warn[:200]!r}"
    _in_tmp(_go)


@case
def t_a_preset_dated_outside_the_data_does_not_kill_the_sidebar():
    """범위 밖 날짜가 든 프리셋 — `date_input`이 예외로 죽으면 전 페이지가 내려간다."""
    def _go():
        at = _open()
        at.session_state["flt_presets"] = {
            "옛날": json.dumps({"flt_date": ["2000-01-01", "2000-01-31"]})}
        at.run()
        _btn(at, "preset_load").click().run()
        assert not at.exception, \
            f"범위 밖 날짜에서 죽었어요 — {at.exception[0].value}"
        lo, hi = _sv(at, "flt_date")
        assert datetime.date(2000, 1, 1) < lo <= hi, f"경계 안으로 안 당겼어요 — {lo}~{hi}"
    _in_tmp(_go)


@case
def t_presets_ride_the_backup_zip():
    """재배포하면 로컬 CSV가 날아간다 — 통합 백업이 유일한 복구 수단이다."""
    def _go():
        at = _open(flt_q="세일")
        at.text_input(key="preset_name").set_value("내조합").run()
        _btn(at, "preset_put").click().run()

        # 페이지를 번호로 가리키지 않는다 — 재정렬하면 조용히 썩는다(개발 규칙).
        at.sidebar.radio[0].set_value(group_of("데이터·다운로드")).run()
        _btn(at, "bak_make").click().run()
        assert not at.exception, at.exception[0].value
        bak = _sv(at, "_bak_zip")
        assert bak, "백업 ZIP이 안 만들어졌어요"
        names = zipfile.ZipFile(io.BytesIO(bak[1])).namelist()
        assert "presets.csv" in names, f"ZIP에 프리셋이 없어요 — {names}"
    _in_tmp(_go)


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
    print(f"필터 프리셋 테스트 {len(CASES) if not want else '선택'}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
