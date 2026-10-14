"""골든 회귀 — **파이프라인이 내는 숫자가 어제와 같은가.**

이 저장소에서 고장은 조용히 난다. 「값이 좀 다르네」가 실은 같은 지표를 두 화면이
다른 규칙으로 읽는 것이고, 「금액이 좀 적네」가 파서가 못 읽은 행을 합계에서 뺀
것이다. 기존 검사들은 **규칙 하나씩**을 본다 — 리팩터링이 파이프라인 어딘가의 숫자를
0.3% 움직여도 전부 통과한다.

골든은 그 자리를 메운다. 날짜를 고정한 합성 데이터를 **실제 함수들에 통째로 흘려**
집계값을 찍고, 커밋된 스냅샷과 **한 자리까지** 맞댄다.

| 보는 것 | 어떤 사고를 막나 |
|---|---|
| `_finalize` 파생(시간대 보정·우선순위 묶기) | 보정 규칙이 조용히 바뀌어 발송이 두 줄로 갈리는 것 |
| 문구 태깅(`add_tags`) 속성별 건수 | `KW`를 손봤을 때 다른 속성까지 움직이는 것 |
| 브랜드 분류(`brand2`·`sales_org`·`brand_kind`) | 사전·불용어를 고쳤을 때의 오탐 번짐 |
| `compute_mtd` 지표 합 | 파생 하나를 고치다 다른 지표를 흔드는 것 |
| `rank_adjusted` 순위 | 수축 상수를 바꿔 추천 순서가 뒤집히는 것 |
| `change_digest` · `phrase_candidates` | 통계 경로가 말없이 달라지는 것 |

**값이 달라지는 게 늘 버그는 아니다.** 의도한 변경이면 `--update`로 갱신하되,
**무엇이 왜 달라졌는지 커밋 메시지에 남길 것.** 그 설명이 이 검사의 값이다.

로컬 실행:
    python tests/test_golden.py            # 대조
    python tests/test_golden.py --update   # 스냅샷 갱신(이유를 커밋에 남길 것)
"""
import json
import math
import pathlib
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import send_perf_dashboard as A                            # noqa: E402
from golden_fixture import golden_mtd, golden_store        # noqa: E402

GOLDEN = ROOT / "tests" / "golden" / "pipeline.json"


# ──────────────────────────────────────────────────────────────────────
# 값 고르기 — **부동소수 잡음에 안 걸리게** 자리를 맞춰 둔다
# ──────────────────────────────────────────────────────────────────────

def _r(x, nd=6):
    """숫자를 스냅샷에 넣을 꼴로. NaN·inf는 문자열로 박아 **조용히 0이 되지 않게** 한다."""
    if x is None:
        return None
    try:
        v = float(x)
    except (TypeError, ValueError):
        return str(x)
    if math.isnan(v):
        return "NaN"
    if math.isinf(v):
        return "inf" if v > 0 else "-inf"
    return round(v, nd)


def _counts(s, top=None):
    """값 분포 → {값: 건수}. 키 순서는 **값 이름 순**으로 고정한다(dict 순서가 흔들리면 diff가 거짓말한다)."""
    vc = pd.Series(s).astype(str).value_counts()
    if top:
        vc = vc.head(top)
    return {k: int(v) for k, v in sorted(vc.items())}


# ──────────────────────────────────────────────────────────────────────
# 스냅샷 만들기
# ──────────────────────────────────────────────────────────────────────

def build(raw=None, mtd_raw=None):
    """집계 스냅샷을 만든다. 인자를 안 주면 **날짜 고정 합성본**을 쓴다.

    실백업으로도 같은 모양을 뽑을 수 있게 인자를 받는다(`tools/golden_realdata.py`).
    실데이터는 깃에 못 넣으니 그쪽 스냅샷은 로컬에만 둔다 — 그래도 **같은 코드로**
    찍어야 두 골든이 서로 다른 걸 재지 않는다.
    """
    out = {}
    raw = golden_store() if raw is None else raw

    # ① `_finalize` — 파생과 보정
    fin = A._finalize(raw.copy())
    out["finalize"] = {
        "행": int(len(fin)),
        "발송합": _r(fin["send"].sum(), 0),
        "UV합": _r(fin["uv"].sum(), 0),
        "거래액합": _r(fin["amt"].sum(), 0),
        "시간대": _counts(fin["hour"]),
        "우선순위(prio_g)": _counts(fin["prio_g"]) if "prio_g" in fin else {},
        "요일키": _counts(fin["dow_k"]) if "dow_k" in fin else {},
    }

    # ② 문구 태깅 — 속성마다 몇 건이 걸리나
    tag = A.add_tags(fin.copy())
    out["tags"] = {t: int(pd.Series(tag[t]).fillna(False).astype(bool).sum())
                   for t in A.TAG_BOOLS if t in tag.columns}
    out["tags_meta"] = {
        "제목길이 평균": _r(pd.to_numeric(tag.get("제목길이"), errors="coerce").mean(), 3),
        "이모지수 합": _r(pd.to_numeric(tag.get("이모지수"), errors="coerce").sum(), 0),
    }

    # ③ 브랜드 자동 분류
    out["brand"] = {
        "brand2": _counts(tag["brand2"], top=12),
        "sales_org": _counts(tag["sales_org"]),
        "brand_kind": _counts(tag["brand_kind"]),
    }

    # ④ 발송유형 정규화
    out["stype"] = _counts([A.norm_stype(v) for v in raw["stype"]])

    # ⑤ 전사 MTD — `compute_mtd`는 프레임이 아니라 **묶음(dict)** 을 돌려준다
    mt = A.compute_mtd(golden_mtd() if mtd_raw is None else mtd_raw)
    md = mt["df"]
    out["mtd"] = {c: _r(pd.to_numeric(md[c], errors="coerce").sum(), 3)
                  for c in sorted(md.columns) if c != "date"}
    out["mtd_monthly"] = {
        str(k): _r(pd.to_numeric(v, errors="coerce").sum(), 3)
        for k, v in mt["monthly"].items()
        if k != "date" and pd.api.types.is_numeric_dtype(mt["monthly"][k])
    } if isinstance(mt.get("monthly"), pd.DataFrame) else {}

    # ⑥ 순위 보정 — 순서가 바뀌면 «추천»이 바뀐다.
    # `rank_adjusted`는 정렬된 프레임이 아니라 **점수 Series**를 돌려준다. 컬럼 이름도
    # `RANK_ND`의 키(`infl_cr`·`ord_cr`·`rps`·`aov`)여야 보정 경로를 탄다 — 아무 이름이나
    # 주면 조용히 원값으로 떨어져 「보정을 껐는데도 통과」가 된다.
    g = (fin.groupby("cat", as_index=False)
            .agg(send=("send", "sum"), uv=("uv", "sum"),
                 oc=("oc", "sum"), amt=("amt", "sum")))
    out["rank"] = {}
    for met in ("infl_cr", "ord_cr", "rps", "aov"):
        sc = A.rank_adjusted(g, met)
        out["rank"][met] = [str(x) for x in
                            g.assign(_s=sc).sort_values("_s", ascending=False)["cat"]]
        out["rank"][met + "(점수)"] = [_r(v) for v in
                                       sc.sort_values(ascending=False).tolist()]

    # ⑦ 변화 다이제스트 — 뒤 60일 vs 그 앞 60일
    d = fin.dropna(subset=["date"]).copy()
    end = pd.to_datetime(d["date"]).max()
    cur = d[pd.to_datetime(d["date"]) > end - pd.Timedelta(days=60)]
    prev = d[(pd.to_datetime(d["date"]) <= end - pd.Timedelta(days=60)) &
             (pd.to_datetime(d["date"]) > end - pd.Timedelta(days=120))]
    dig = A.change_digest(cur, prev)
    out["digest"] = {
        "행": int(len(dig)),
        "칼럼": list(A.DIGEST_COLS),
        "상위": [{"축": str(r["축"]), "값": str(r["값"]), "지표": str(r["지표"]),
                  "변화율": _r(r["변화율"]), "q": _r(r["q"])}
                 for _i, r in dig.head(5).iterrows()],
    }

    # ⑧ 소구 사전 후보
    cand, steps = A.phrase_candidates(fin)
    out["phrase"] = {
        "단계": {k: int(v) for k, v in steps.items()},
        "상위": [{"단어": str(r["단어"]), "변화율": _r(r["변화율"]),
                  "홀드아웃": _r(r["홀드아웃 변화율"]), "q": _r(r["q"])}
                 for _i, r in cand.head(5).iterrows()],
    }

    # ⑨ 플레이북 추천 — 슬롯 목록이 곧 운영 지시다(순서까지 본다)
    out["playbook"] = {met: [f"{a}·{b}" for a, b in A.playbook_slots(fin, met)]
                       for met in ("ord_cr", "infl_cr", "rps")}

    return out


# ──────────────────────────────────────────────────────────────────────
# 대조
# ──────────────────────────────────────────────────────────────────────

def _flat(obj, pre=""):
    """중첩 dict → {경로: 값}. **어디가 달라졌는지**를 경로로 짚어 준다."""
    out = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            out.update(_flat(v, f"{pre}.{k}" if pre else str(k)))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            out.update(_flat(v, f"{pre}[{i}]"))
    else:
        out[pre] = obj
    return out


# 숫자는 **상대 허용오차**로 맞댄다. scipy·numpy가 판올림되면 마지막 비트가
# 흔들리는데, 그걸 회귀로 읽으면 **코드를 안 고쳐도 빨간불**이 난다. 진짜 회귀는
# 값을 훨씬 크게 움직이므로 이 정도 여유로도 다 잡힌다. 문자열·개수는 정확 비교다.
GOLDEN_RTOL = 1e-9


def _same(x, y):
    if isinstance(x, bool) or isinstance(y, bool):
        return x is y
    if isinstance(x, (int, float)) and isinstance(y, (int, float)):
        return math.isclose(float(x), float(y), rel_tol=GOLDEN_RTOL, abs_tol=1e-12)
    return x == y


def diff(got, want):
    a, b = _flat(got), _flat(want)
    msgs = []
    for k in sorted(set(a) | set(b)):
        if k not in b:
            msgs.append(f"  + {k} = {a[k]!r}   (스냅샷에 없던 값)")
        elif k not in a:
            msgs.append(f"  - {k} = {b[k]!r}   (이번엔 안 나옴)")
        elif not _same(a[k], b[k]):
            msgs.append(f"  ! {k}: {b[k]!r} → {a[k]!r}")
    return msgs


def main():
    update = "--update" in sys.argv
    got = build()
    GOLDEN.parent.mkdir(parents=True, exist_ok=True)

    if update or not GOLDEN.exists():
        GOLDEN.write_text(json.dumps(got, ensure_ascii=False, indent=2,
                                     sort_keys=True) + "\n", encoding="utf-8")
        print(f"스냅샷을 {'갱신' if update else '처음 생성'}했어요 — {GOLDEN.relative_to(ROOT)}")
        print("무엇이 왜 달라졌는지 커밋 메시지에 남길 것.")
        return 0

    want = json.loads(GOLDEN.read_text(encoding="utf-8"))
    msgs = diff(got, want)
    if msgs:
        print(f"골든과 {len(msgs)}곳이 달라요:\n" + "\n".join(msgs[:60]))
        if len(msgs) > 60:
            print(f"  … 외 {len(msgs) - 60}곳")
        print("\n의도한 변경이면 `python tests/test_golden.py --update` 뒤 "
              "**이유를 커밋 메시지에 남기세요.**")
        return 1
    n = len(_flat(got))
    print(f"골든 회귀 통과 ✅ — 값 {n:,}개가 스냅샷과 같아요")
    return 0


if __name__ == "__main__":
    sys.exit(main())
