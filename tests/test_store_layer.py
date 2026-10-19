"""저장소 계층 방어 테스트 — dtype이 문자열로 되돌아와도 앱이 죽지 않아야 한다.

구글시트 라운드트립은 모든 값을 문자열로 되돌린다. push 저장소는 `finalize_push()`로
복구하는 게 원칙인데, 한 경로라도 빠뜨리면 `get_push_stats`가 'str vs Timestamp'
비교로 터진다. 이 함수는 사이드바 렌더 경로라 **11개 페이지가 전부 다운**된다.

실백업으로 전 페이지를 돌리다 실제로 재현돼서 방어를 넣고 여기에 고정한다.

로컬 실행:
    python tests/test_store_layer.py
"""
import ast
import pathlib
import re
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import send_perf_dashboard as S  # noqa: E402

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def raw_push():
    """구글시트에서 막 읽어온 모양 — 전부 문자열."""
    return pd.DataFrame([
        {"date": "2026-07-01", "group": "앱푸시_동의자수", "consent": "700000",
         "added": "1000", "removed": "400", "diff": "600", "is_outlier": "False"},
        {"date": "2026-07-02", "group": "앱푸시_동의자수", "consent": "700600",
         "added": "1200", "removed": "500", "diff": "700", "is_outlier": "False"},
        {"date": "2026-07-03", "group": "앱푸시_동의자수", "consent": "701300",
         "added": "900", "removed": "300", "diff": "600", "is_outlier": "True"},
    ])


@case
def t_get_push_stats_survives_string_dtypes():
    """문자열 dtype이 들어와도 죽지 않고, finalize된 것과 같은 값을 낸다."""
    got = S.get_push_stats(raw_push(), "2026-07-01", "2026-07-31", "앱푸시_동의자수")
    want = S.get_push_stats(S.finalize_push(raw_push()), "2026-07-01", "2026-07-31",
                            "앱푸시_동의자수")
    assert got is not None, "문자열 dtype에서 None이 나왔어요"
    assert got == want, f"{got} != {want}"


@case
def t_outlier_excluded():
    """is_outlier가 문자열 'True'여도 제외돼야 한다."""
    r = S.get_push_stats(raw_push(), "2026-07-01", "2026-07-31", "앱푸시_동의자수")
    assert r["tot_added"] == 2200, r          # 이상치 행(900) 제외
    assert r["last_consent"] == 700600, r     # 이상치 제외 후 마지막 날


@case
def t_empty_and_none_are_safe():
    """빈 프레임·None·컬럼 부재에서 예외 대신 None."""
    assert S.get_push_stats(None, "2026-07-01", "2026-07-31", "g") is None
    assert S.get_push_stats(pd.DataFrame(), "2026-07-01", "2026-07-31", "g") is None
    assert S.get_push_stats(pd.DataFrame({"x": [1]}), "2026-07-01", "2026-07-31", "g") is None


@case
def t_no_match_returns_none():
    """기간·그룹이 안 맞으면 None (0으로 위장하지 않는다)."""
    assert S.get_push_stats(raw_push(), "2020-01-01", "2020-12-31", "앱푸시_동의자수") is None
    assert S.get_push_stats(raw_push(), "2026-07-01", "2026-07-31", "없는그룹") is None


@case
def t_finalize_push_is_idempotent():
    """두 번 통과시켜도 같은 결과 — 여러 경로에서 중복 호출된다."""
    once = S.finalize_push(raw_push())
    twice = S.finalize_push(once)
    assert len(once) == len(twice)
    assert list(once["date"]) == list(twice["date"])
    assert list(once["is_outlier"]) == list(twice["is_outlier"])


# ── 저장소 «종류» 레지스트리가 서로 맞는지 ──────────────────────────────
# 저장소 한 종류를 더하려면 dict 여섯을 같이 고쳐야 한다. 다섯은 빠뜨리면 KeyError로
# 시끄럽게 죽지만 `_KIND_LABEL`만 `.get`이라 **조용히** 통과하고, 저장이 막힌 날
# 「⛔ preset 저장을 차단했어요」처럼 영문 키가 그대로 찍힌다. 실제로 `site`가 거기
# 빠져 있었다.
KIND_REGISTRIES = ("GS_TITLES", "_STORE_COLS_BY_KIND", "_LOCAL_LOAD",
                   "_LOCAL_SAVE", "_LOCAL_FILE", "_KIND_LABEL")


@case
def t_storage_kinds_agree_across_registries():
    src = (ROOT / "send_perf_dashboard.py").read_text(encoding="utf-8")
    keys = {}
    for nm in KIND_REGISTRIES:
        m = re.search(rf"^\s*{nm} = \{{(.*?)\}}", src, re.S | re.M)
        assert m, f"{nm} 를 못 찾았어요 — 이름이 바뀌었으면 KIND_REGISTRIES도 고치세요"
        keys[nm] = set(re.findall(r'"(\w+)":', m.group(1)))
    base = keys["GS_TITLES"]
    bad = {nm: (sorted(base - ks), sorted(ks - base))
           for nm, ks in keys.items() if ks != base}
    assert not bad, "\n".join(
        f"  {nm}: 빠진 종류 {miss} / 남는 종류 {extra}" for nm, (miss, extra) in bad.items())


# ── 업로드 저장 큐 — 올린 종류가 몇이든 버튼은 하나다 ──────────────────────
# 예전엔 발송실적·전사 MTD·기획전이 각자 버튼을 들고 있었다. 셋을 같이 올린 날 하나만
# 누르고 떠나면 나머지는 조용히 안 남는데, 화면엔 「합쳐서 N건 (미리보기)」가 그대로
# 떠 있어서 저장된 줄로 읽힌다. 이 저장소의 전형적인 조용한 실패다.

def _sb_button_labels(src):
    """`st.sidebar.button("…")` 의 첫 인자 리터럴 목록 — AST로 읽는다.

    문자열로 찾으면 주석·문서에 적힌 버튼 이름까지 걸린다.
    """
    out = []
    for n in ast.walk(ast.parse(src)):
        if not isinstance(n, ast.Call) or not isinstance(n.func, ast.Attribute):
            continue
        if n.func.attr != "button":
            continue
        tgt = n.func.value
        if not (isinstance(tgt, ast.Attribute) and tgt.attr == "sidebar"
                and isinstance(tgt.value, ast.Name) and tgt.value.id == "st"):
            continue
        if n.args and isinstance(n.args[0], ast.Constant) and isinstance(n.args[0].value, str):
            out.append(n.args[0].value)
    return out


@case
def t_upload_has_one_save_button():
    src = (ROOT / "send_perf_dashboard.py").read_text(encoding="utf-8")
    saves = [b for b in _sb_button_labels(src) if "저장하기" in b]
    assert len(saves) == 1, (
        f"사이드바 저장 버튼이 {len(saves)}개예요 — {saves}. 저장소 종류를 더해도 버튼은 "
        "하나입니다. `save_queue_add`로 큐에 넣으세요.")


def _entry(kind, applied):
    """저장 대기 한 칸 — `build`가 만든 값을 `applied`에 남긴다."""
    return dict(kind=kind, label=kind,
                build=lambda k=kind: f"{k}-df",
                apply_=lambda df, k=kind: applied.__setitem__(k, df))


def _queue(kinds, applied):
    pend = []
    for k in kinds:
        e = _entry(k, applied)
        S.save_queue_add(pend, e["kind"], e["label"], e["build"], e["apply_"])
    return pend


@case
def t_one_press_saves_every_queued_store():
    """한 번 눌러서 큐에 있는 **전부**가 저장돼야 한다."""
    applied, saved = {}, []
    pend = _queue(["campaign", "mtd", "promo"], applied)
    ok, ng = S.save_queue_flush(pend, lambda kind, df: (saved.append((kind, df)) or True))
    assert [k for k, _ in saved] == ["campaign", "mtd", "promo"], f"저장 호출이 {saved}"
    assert saved == [(k, f"{k}-df") for k in ("campaign", "mtd", "promo")], \
        f"엉뚱한 프레임을 저장했어요 — {saved}"
    assert applied == {k: f"{k}-df" for k in ("campaign", "mtd", "promo")}, \
        f"세션 반영이 빠졌어요 — {applied}"
    assert (ok, ng) == (["campaign", "mtd", "promo"], []), f"결과가 {(ok, ng)}"


@case
def t_a_blocked_store_does_not_swallow_the_others():
    """한 종류가 막혀도 나머지는 끝까지 간다 — 그리고 막힌 건 세션에 안 넣는다.

    중간에 멈추면 사용자는 「저장했어요」를 보고 떠나는데 일부만 남는다. 막힌 걸
    세션에만 넣으면 화면은 저장된 것처럼 보이고 저장소엔 없다.
    """
    applied, saved = {}, []
    pend = _queue(["campaign", "mtd", "promo"], applied)
    ok, ng = S.save_queue_flush(
        pend, lambda kind, df: (saved.append(kind) or (kind != "mtd")))
    assert saved == ["campaign", "mtd", "promo"], f"중간에 멈췄어요 — {saved}"
    assert ok == ["campaign", "promo"] and ng == ["mtd"], f"결과가 {(ok, ng)}"
    assert "mtd" not in applied, "저장이 막혔는데 세션엔 넣었어요"
    assert set(applied) == {"campaign", "promo"}, f"세션 반영이 {sorted(applied)}"


@case
def t_a_raising_build_is_reported_not_swallowed():
    """프레임을 만들다 터져도 나머지는 저장하고, 터진 건 실패로 보고한다."""
    applied, saved = {}, []
    pend = _queue(["campaign", "promo"], applied)
    pend.insert(1, {"kind": "mtd", "label": "전사 MTD",
                    "build": lambda: (_ for _ in ()).throw(ValueError("깨진 프레임")),
                    "apply": lambda df: applied.__setitem__("mtd", df)})
    ok, ng = S.save_queue_flush(pend, lambda kind, df: (saved.append(kind) or True))
    assert saved == ["campaign", "promo"], f"저장 호출이 {saved}"
    assert ok == ["campaign", "promo"], f"성공 목록이 {ok}"
    assert len(ng) == 1 and ng[0].startswith("전사 MTD") and "깨진 프레임" in ng[0], \
        f"실패를 삼켰어요 — {ng}"
    assert "mtd" not in applied, "터졌는데 세션엔 넣었어요"


def main():
    fails = []
    for fn in CASES:
        try:
            fn()
            print(f"  OK   {fn.__name__}")
        except Exception as e:                            # noqa: BLE001
            print(f"  FAIL {fn.__name__}: {e}")
            fails.append(fn.__name__)
    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print(f"저장소 계층 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
