"""영업별 운영브랜드 원본(UTF-16 CSV) → 대시보드가 읽는 data/brand_map.csv 생성.

원본은 `BPU, ADMIN브랜드명` 2열이고 ADMIN브랜드명은 라인 단위라 같은 브랜드가
'닥스 골프' · '닥스 남성 세트' · '닥스여성세트'처럼 여러 줄로 쪼개져 있다.
분석은 브랜드 단위로 봐야 하므로 라인 접미어를 떼어 base 브랜드로 굴린다.

사용:
    python tools/build_brand_map.py <원본.csv> [출력경로]
"""
import pathlib
import re
import sys

import pandas as pd

OUT_DEFAULT = pathlib.Path(__file__).resolve().parent.parent / "data" / "brand_map.csv"

# 라인 접미어 — 브랜드가 아니라 '그 브랜드의 어떤 라인인가'를 나타내는 꼬리표.
# 긴 것부터 떼어야 '액세서리 남성'처럼 겹친 꼬리를 제대로 벗긴다.
SUFFIXES = [
    "세트코드", "액세서리", "악세서리", "액세사리", "악세사리", "아이웨어", "언더웨어",
    "컬렉션", "주니어", "주얼리", "보야지", "에뚜왈", "슈즈", "키즈", "우먼", "여성",
    "남성", "세트", "골프", "용품", "가방", "잡화", "신발", "모자", "벨트", "지갑",
]
_SUF_RE = re.compile(r"(?:\s*(?:" + "|".join(SUFFIXES) + r"))+$")


def strip_line(name):
    """라인 접미어를 반복해서 떼고 base 브랜드명만 남긴다."""
    s = " ".join(str(name).split())
    prev = None
    while s and s != prev:
        prev = s
        s = _SUF_RE.sub("", s).strip()
    return s or " ".join(str(name).split())


def read_src(path):
    """원본은 UTF-16 탭구분으로 저장돼 있다(엑셀에서 내보낸 형태)."""
    last = None
    for enc in ("utf-16", "utf-8-sig", "cp949"):
        try:
            df = pd.read_csv(path, dtype=str, keep_default_na=False,
                             encoding=enc, sep=None, engine="python")
            if df.shape[1] >= 2:
                return df
        except Exception as e:                            # noqa: BLE001
            last = e
    raise SystemExit(f"원본을 읽지 못했어요: {last}")


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 1
    src = pathlib.Path(argv[1])
    out = pathlib.Path(argv[2]) if len(argv) > 2 else OUT_DEFAULT
    df = read_src(src)
    df = df.iloc[:, :2]
    df.columns = ["org", "admin"]
    df = df.apply(lambda s: s.str.strip())
    df = df[(df["org"] != "") & (df["admin"] != "")].drop_duplicates()

    df["base"] = df["admin"].map(strip_line)
    df = df[df["base"] != ""]

    # base 브랜드가 여러 영업에 걸치면 자사(e-영업1·2)를 우선한다 — CRM 발송은 대부분
    # 자사 브랜드 행사라, 동명이 있으면 자사 쪽으로 붙는 게 오분류가 적다.
    df["_pri"] = df["org"].map(lambda o: 0 if o in ("e-영업1", "e-영업2") else 1)
    df = df.sort_values(["base", "_pri", "org"])
    base = (df.groupby("base", as_index=False)
              .agg(org=("org", "first"), admin_n=("admin", "size")))

    base = base.rename(columns={"base": "brand"})
    base["brand_len"] = base["brand"].str.replace(r"\s+", "", regex=True).str.len()
    base = base.sort_values(["brand_len", "brand"], ascending=[False, True])
    out.parent.mkdir(parents=True, exist_ok=True)
    base[["brand", "org", "admin_n"]].to_csv(out, index=False, encoding="utf-8")

    print(f"원본 {len(df):,}행 → base 브랜드 {len(base):,}개  →  {out}")
    print("\n영업별 base 브랜드 수:")
    print(base["org"].value_counts().to_string())
    print("\n자사(e-영업1·2) base 브랜드 샘플:")
    own = base[base["org"].isin(["e-영업1", "e-영업2"])]["brand"].sort_values().tolist()
    print(f"  ({len(own)}개) " + ", ".join(own[:80]))
    print("\n2글자 이하라 문구 매칭에서 제외될 브랜드 수:",
          int((base["brand_len"] <= 2).sum()))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
