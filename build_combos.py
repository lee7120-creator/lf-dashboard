"""③단계 — [CEP]×[카테고리] 롱테일 조합 후보 생성.

왜? CEP(상황·맥락)와 카테고리(상품)를 곱해 롱테일 키워드를 만든다.
   예: [결혼식 하객] × [원피스] → "결혼식 하객 원피스"
이 후보가 ④단계(롱테일 수요조사)의 입력이 된다.

전략: 각 CEP 축은 '의미가 통하는' 카테고리 섹션하고만 조합한다.
   (선물 축은 가방·주얼리와, 골프 축은 골프 카테고리와…)
   고검색량 우선: CEP 상위 × 카테고리 상위만 뽑아 폭발(n×m)을 억제.

입력:  data/cep_keyword_research.csv, data/lfmall_keyword_research.csv
출력:  data/combo_candidates.csv
   CEP, CEP축, CEP검색량, 카테고리, 섹션, 카테고리검색량, 조합키워드, 조합점수
"""
import csv
import os
import re

import pandas as pd

ROOT = os.path.dirname(os.path.abspath(__file__))

# CEP 축 → 조합할 카테고리 섹션 (의미 매칭)
AXIS_TO_SECTIONS = {
    "이벤트/TPO": ["의류", "신발", "가방", "주얼리/시계", "액세서리"],
    "시즌/날씨":  ["의류", "신발", "가방", "액세서리"],
    "휴가/여행":  ["의류", "신발", "가방", "수영/비치", "뷰티/향수"],
    "선물/관계":  ["가방", "주얼리/시계", "지갑/벨트", "뷰티/향수", "액세서리"],
    "활동/스포츠": ["골프", "신발", "의류", "가방", "캠핑/등산"],
    "체형/핏":    ["의류", "신발"],
    "연령/세대":  ["의류", "가방", "신발", "액세서리"],
    "스타일/무드": ["의류", "신발", "가방", "액세서리"],
}

# CEP 키워드 → '수식어 핵심'만 추출 (접미 제거). 카테고리를 붙여 자연스러운 검색어로.
CEP_SUFFIX = ["룩", "코디", "패션", "옷차림", "옷", "웨어", "복장", "정장",
              "원피스", "선물", "신발", "복"]


def cep_core(kw):
    s = kw.strip()
    for suf in CEP_SUFFIX:
        if s.endswith(suf) and len(s) > len(suf):
            s = s[:-len(suf)].strip()
            break
    return s


TOP_CEP = 8     # 축별 상위 CEP
TOP_CAT = 12    # 섹션별 상위 카테고리


def main():
    cep = pd.read_csv(os.path.join(ROOT, "data", "cep_keyword_research.csv"),
                      encoding="utf-8-sig")
    cat = pd.read_csv(os.path.join(ROOT, "data", "lfmall_keyword_research.csv"),
                      encoding="utf-8-sig")
    for df in (cep, cat):
        df["대표검색량"] = pd.to_numeric(df["대표검색량"], errors="coerce").fillna(0).astype(int)

    rows = []
    for axis, sections in AXIS_TO_SECTIONS.items():
        top_cep = (cep[(cep["섹션"] == axis) & (cep["대표검색량"] > 0)]
                   .nlargest(TOP_CEP, "대표검색량"))
        # 매핑 섹션의 상위 카테고리 키워드 풀
        cat_pool = (cat[(cat["섹션"].isin(sections)) & (cat["대표검색량"] > 0)]
                    .sort_values("대표검색량", ascending=False))
        cat_top = (cat_pool.groupby("섹션", group_keys=False)
                   .apply(lambda g: g.nlargest(TOP_CAT, "대표검색량")))
        for _, c in top_cep.iterrows():
            core = cep_core(c["키워드"])
            for _, k in cat_top.iterrows():
                combo = f"{core} {k['키워드']}".strip()
                combo = re.sub(r"\s+", " ", combo)
                # 조합점수 = √(CEP검색량) × √(카테고리검색량) — 둘 다 큰 쌍 우대
                score = round((c["대표검색량"] ** 0.5) * (k["대표검색량"] ** 0.5))
                rows.append({
                    "CEP": c["키워드"], "CEP축": axis, "CEP검색량": c["대표검색량"],
                    "카테고리": k["키워드"], "섹션": k["섹션"],
                    "카테고리검색량": k["대표검색량"],
                    "조합키워드": combo, "조합점수": score,
                })

    out_df = pd.DataFrame(rows).drop_duplicates("조합키워드")
    out_df = out_df.sort_values("조합점수", ascending=False).reset_index(drop=True)
    out_df["우선순위"] = out_df.index + 1

    out = os.path.join(ROOT, "data", "combo_candidates.csv")
    out_df.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"조합 후보 {len(out_df):,}개 → {out}")
    print(f"CEP축 {len(AXIS_TO_SECTIONS)}개 기준 · 축별 상위 {TOP_CEP} × 섹션별 상위 {TOP_CAT}")
    print("\n=== 조합점수 TOP 15 (④단계 우선 조사 대상) ===")
    for _, r in out_df.head(15).iterrows():
        print(f"  {r['우선순위']:>3}. {r['조합키워드']}  "
              f"(CEP {r['CEP']} {r['CEP검색량']:,} × {r['카테고리']} {r['카테고리검색량']:,})")


if __name__ == "__main__":
    main()
