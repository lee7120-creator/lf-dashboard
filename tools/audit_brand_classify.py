"""브랜드 자동 분류 대량 검증 — 실제 누적 백업으로 돌린다.

사전이 9천 개라 오탐(평범한 문구가 브랜드로 잡히는 것)이 제일 무섭다.
합성 문구로는 그게 안 잡혀서, 대시보드 백업 ZIP(또는 campaign.csv)을 그대로 태워
분류 분포와 브랜드별 실제 문구를 눈으로 확인할 수 있게 한다.

사용:
    python tools/audit_brand_classify.py <백업.zip | campaign.csv> [--min N]

보는 법:
  · 구분 분포 — '브랜드' 비중이 갑자기 떨어지면 사전/경계 규칙이 망가진 것
  · 브랜드 목록 — 일반어가 섞여 있으면 BRAND_STOP 후보다
    ('클리어'←클리어런스, '화이트'←화이트데이, '크리스마'←크리스마스)
  · 예시 문구 — 그 브랜드로 잡힌 실제 제목을 보고 오탐인지 판단한다
"""
import io
import pathlib
import sys
import zipfile
from collections import Counter

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import send_perf_dashboard as S  # noqa: E402


def load(path):
    p = pathlib.Path(path)
    if p.suffix.lower() == ".zip":
        with zipfile.ZipFile(p) as z:
            name = next((n for n in z.namelist()
                         if n.split("/")[-1].lower().startswith("campaign")), None)
            if not name:
                raise SystemExit(f"ZIP 안에 campaign.csv가 없어요: {z.namelist()}")
            data = io.BytesIO(z.read(name))
    else:
        data = p
    return pd.read_csv(data, dtype=str, keep_default_na=False, encoding="utf-8-sig")


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 1
    min_n = 5
    if "--min" in argv:
        min_n = int(argv[argv.index("--min") + 1])
    df = load(argv[1])
    print(f"사전: 부분일치 {len(S.BRAND_MAP):,} · 정확일치 {len(S.BRAND_EXACT):,} · "
          f"코드 {len(S.BRAND_CODE):,} · 제외어 {len(S.BRAND_STOP)}")
    print(f"검증 대상 {len(df):,}건 (문구 있는 행 "
          f"{int((df.get('title', pd.Series(dtype=str)).astype(str).str.strip() != '').sum()):,})\n")

    out = S.add_tags(df)
    kinds = out["brand_kind"].value_counts()
    print("── 구분 분포 ──")
    for k, v in kinds.items():
        print(f"   {k:<10} {v:>6,}  ({v / len(out) * 100:4.1f}%)")
    nb = int(kinds.get(S.KIND_BRAND, 0))
    print(f"\n브랜드로 특정: {nb:,}/{len(out):,} ({nb / max(len(out), 1) * 100:.1f}%)")

    bx = out[out["brand_kind"] == S.KIND_BRAND]
    vc = bx["brand2"].value_counts()
    print(f"고유 브랜드 {len(vc)}개 · 영업 분포 {dict(out['sales_org'].value_counts())}\n")

    print(f"── 브랜드별 건수 (≥{min_n}건) + 실제 문구 예시 ──")
    for b, n in vc[vc >= min_n].items():
        ex = bx[bx["brand2"] == b]
        t = ""
        for _, r in ex.iterrows():
            t = (str(r.get("title", "")) or str(r.get("body", "")) or str(r.get("brand", ""))).strip()
            if t:
                break
        print(f"   {b:<18} {n:>5,}  {t[:52]}")

    tail = vc[vc < min_n]
    if len(tail):
        print(f"\n── 롱테일 {len(tail)}개 (일반어가 섞였는지 눈으로 확인) ──")
        print("   " + ", ".join(f"{b}({n})" for b, n in tail.items()))

    print("\n── 브랜드로 못 집은 발송 상위 (거래액 순) ──")
    un = out[out["brand_kind"] != S.KIND_BRAND].copy()
    if "amt" in un:
        un["_amt"] = pd.to_numeric(un["amt"], errors="coerce").fillna(0)
        g = (un.groupby(un["brand"].replace("", "(빈칸)"))
               .agg(건수=("brand_kind", "size"), 거래액=("_amt", "sum"),
                    구분=("brand_kind", "first"),
                    예시=("title", lambda s: next((x for x in s if str(x).strip()), "")))
               .sort_values("거래액", ascending=False).head(20))
        for k, r in g.iterrows():
            print(f"   {str(k)[:20]:<22} {int(r['건수']):>5,}건 "
                  f"{int(r['거래액']):>12,}원  {r['구분']:<8} {str(r['예시'])[:34]}")
    print("\n구분 요약:", dict(Counter(out["brand_kind"])))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
