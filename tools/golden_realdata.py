"""실백업으로 골든 회귀를 돌린다 — **스냅샷은 깃에 안 들어간다.**

`tests/test_golden.py`는 날짜를 고정한 합성본으로 돈다. 그걸로도 집계 규칙이 바뀌면
잡히지만, **합성본이 안 밟는 자리**가 있다 — 수기 입력의 흔들림(금액 표기·시간대
오타·AF코드 변형), 실제 브랜드 어휘, 한쪽만 비어 오는 칸. 거기서 나는 고장이 이
저장소에서 제일 자주 나는 종류다.

그래서 같은 집계를 **실백업으로도** 찍어 두고 다음 번과 맞댄다. 실적은 사업 데이터라
스냅샷을 깃에 넣지 않는다 — 기본 저장 위치가 `.gitignore`에 걸려 있다.

쓰는 법:
    python tools/golden_realdata.py <백업.zip>              # 처음이면 스냅샷 생성
    python tools/golden_realdata.py <백업.zip>              # 다음부터 대조
    python tools/golden_realdata.py <백업.zip> --update     # 의도한 변경이면 갱신
    python tools/golden_realdata.py <백업.zip> --out 경로   # 스냅샷 위치 지정

**대조가 달라졌다고 곧 버그는 아니다.** 백업을 새로 받았으면 데이터가 달라서 값도
달라진다 — **같은 백업 파일**로 코드 전후를 맞댈 때만 뜻이 있다.
"""
import argparse
import io
import json
import pathlib
import sys
import zipfile

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from test_golden import build, diff                        # noqa: E402

DEFAULT_OUT = ROOT / "golden_realdata.json"                # .gitignore에 걸어 둔다

# 백업 ZIP 안 CSV를 이름으로 가른다 — 앱의 `_apply_backup`과 같은 규칙이다.
# 캠페인은 **나머지 전부**(폴백)라 여기서도 마지막에 둔다. 순서를 바꾸면 MTD가
# 캠페인으로 읽혀 집계가 통째로 엉킨다.
OTHERS = ("mtd", "promo", "site", "push", "preset", "note")


def read_backup(path):
    """백업 ZIP → (캠페인 프레임, MTD 프레임). 못 찾으면 그 자리는 None."""
    camp = mtd = None
    with zipfile.ZipFile(path) as z:
        for nm in z.namelist():
            base = nm.split("/")[-1].lower()
            if not base.endswith(".csv"):
                continue
            raw = io.BytesIO(z.read(nm))
            if "mtd" in base:
                mtd = pd.read_csv(raw, encoding="utf-8-sig")
            elif not any(k in base for k in OTHERS):
                camp = pd.read_csv(raw, encoding="utf-8-sig",
                                   dtype={"date": str, "af": str})
    return camp, mtd


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("backup", help="통합 백업 ZIP")
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    ap.add_argument("--update", action="store_true")
    a = ap.parse_args()

    camp, mtd = read_backup(a.backup)
    if camp is None:
        print("캠페인 CSV를 못 찾았어요. 통합 백업 ZIP이 맞는지 확인해 주세요.")
        return 2
    print(f"읽음 — 캠페인 {len(camp):,}행" + (f" · MTD {len(mtd):,}행" if mtd is not None else " · MTD 없음"))

    got = build(camp, mtd)
    out = pathlib.Path(a.out)
    if a.update or not out.exists():
        out.write_text(json.dumps(got, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
                       encoding="utf-8")
        print(f"스냅샷을 {'갱신' if a.update else '처음 생성'}했어요 — {out}")
        print("이 파일은 깃에 안 올라가요(.gitignore). 같은 백업으로 코드 전후를 맞댈 때 쓰세요.")
        return 0

    msgs = diff(got, json.loads(out.read_text(encoding="utf-8")))
    if msgs:
        print(f"스냅샷과 {len(msgs)}곳이 달라요:\n" + "\n".join(msgs[:60]))
        if len(msgs) > 60:
            print(f"  … 외 {len(msgs) - 60}곳")
        print("\n백업 파일이 바뀌었으면 값도 달라져요 — 같은 파일로 맞대고 있는지 먼저 보세요.")
        return 1
    print("실데이터 골든 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
