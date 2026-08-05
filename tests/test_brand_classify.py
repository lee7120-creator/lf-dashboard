"""문구·브랜드칸 기반 브랜드 자동 태깅 테스트.

실적 엑셀의 '브랜드' 칸은 담당자 수기 입력이라 캠페인명·행사명이 섞여 들어온다
('위켄드세일 추가 적립', 'L+DAY 쇼핑핫타임 1차', '주문서이탈'…). 그대로 차원으로 쓰면
브랜드별 비교가 안 되므로, 제목·내용·브랜드칸을 같이 훑어 영업별 운영브랜드 시트
(data/brand_map.csv) 기준으로 다시 태깅한다.

사전이 커서 오탐(일반 문장이 브랜드로 잡히는 것)이 제일 무섭다 — 그 방어를 여기서 건다.

로컬 실행:
    python tests/test_brand_classify.py
"""
import pathlib
import sys

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import send_perf_dashboard as S  # noqa: E402


def BC(title="", body="", brand=""):
    return S.brand_from_copy(title, body, brand)[0]


def ORG(title="", body="", brand=""):
    return S.brand_from_copy(title, body, brand)[1]


CASES = []


def case(fn):
    CASES.append(fn)
    return fn


@case
def t_map_loaded():
    """정본 사전(data/brand_map.csv)이 실제로 로드돼야 한다."""
    assert len(S.BRAND_MAP) > 1000, f"부분일치 사전이 {len(S.BRAND_MAP)}개뿐이에요"
    assert "헤지스" in {v[0] for v in S.BRAND_MAP.values()}
    assert "닥스" in {v[0] for v in S.BRAND_MAP.values()}


@case
def t_brand_in_raw_column():
    """담당자가 적은 브랜드칸에 브랜드가 섞여 있어도 뽑아낸다."""
    assert BC(brand="헤지스남성 아울렛") == "헤지스"
    assert BC(brand="닥스키즈(재)") == "닥스"
    assert BC(brand="L:ABLE - DAKS LADIES") == "닥스"


@case
def t_sales_org_attached():
    """브랜드가 잡히면 영업(e-영업N)도 같이 붙는다."""
    assert ORG(brand="헤지스남성 아울렛") in ("e-영업1", "e-영업2")
    assert ORG(brand="위켄드세일 추가 적립") == S.ORG_UNKNOWN


@case
def t_promo_not_treated_as_brand():
    """L+DAY·엘프 계열은 전사 프로모션이지 브랜드가 아니다."""
    for v in ("L+DAY 쇼핑핫타임 1차", "L+DAY 오픈 알림", "L+DAY 아울렛",
              "L+DAY 나이트마켓 1차", "엘프 룰렛 1일차", "위켄드세일 추가 적립",
              "럭키먼데이", "럭키먼데이 장바구니"):
        assert BC(brand=v) == S.BRAND_PROMO, f"{v} → {BC(brand=v)}"


@case
def t_trigger_sends_separated():
    """발송 트리거(주문서이탈 등)는 브랜드도 행사도 아닌 별도 묶음."""
    assert BC(brand="주문서이탈") == S.BRAND_TRIGGER
    assert BC(brand="남은모수") == S.BRAND_TRIGGER


@case
def t_trigger_only_from_brand_column():
    """트리거 키워드를 문구에까지 적용하면 흔한 표현이 죄다 트리거로 잡힌다."""
    assert BC(title="장바구니에 담아두신 상품이 곧 품절돼요") == S.BRAND_UNKNOWN
    assert BC(title="찜해두신 상품 재입고됐어요") == S.BRAND_UNKNOWN


@case
def t_brand_found_in_copy_only():
    """브랜드칸이 비어도 문구에 브랜드가 있으면 잡는다."""
    assert BC(title="[헤지스] 시즌오프 최대 50%") == "헤지스"
    assert BC(body="티엔지티 단독 특가", brand="전사행사") == "티엔지티"


@case
def t_spacing_and_case_normalized():
    """공백·중점·영문 표기 변형을 같은 브랜드로 흡수한다."""
    for v in ("질스튜어트", "질 스튜어트", "질·스튜어트", "JILL STUART"):
        assert BC(title=v).startswith("질"), f"{v} → {BC(title=v)}"
    for v in ("HAZZYS", "hazzys", "헤지스", "해지스"):
        assert BC(title=v) == "헤지스", f"{v} → {BC(title=v)}"


@case
def t_longest_match_wins():
    """더 구체적인 브랜드가 짧은 이름에 먹히면 안 된다."""
    assert BC(title="질 스튜어트 뉴욕 신상") == "질 스튜어트 뉴욕"
    assert BC(title="닥스런던골프 여성 신상") == "닥스런던"


@case
def t_no_false_positive_on_plain_copy():
    """브랜드가 없는 평범한 문구를 엉뚱한 브랜드로 붙이면 안 된다.

    사전이 9천 개라 '객'·'갭'·'고요' 같은 짧은 입점 브랜드가 일반 문장에 그대로 박힌다.
    """
    for t in ("오늘까지 무료배송! 지금 주문하세요",
              "여름 휴가 필수템 모음",
              "고객님을 위한 맞춤 추천",
              "이번 주말 단 3일, 최대 70% 할인",
              "새로 나온 신상 컬렉션을 만나보세요",
              "지금 응모하면 100% 당첨 기회"):
        assert BC(title=t) == S.BRAND_UNKNOWN, f"{t!r} → {BC(title=t)}"


@case
def t_short_brand_needs_exact_brand_column():
    """2글자 이하 입점 브랜드는 브랜드칸 정확일치일 때만 인정한다."""
    short = [n for n, (b, o) in S.BRAND_EXACT.items()
             if len(n) <= 2 and o not in ("e-영업1", "e-영업2")]
    assert short, "정확일치 전용 짧은 브랜드가 하나도 없어요"
    n = short[0]
    assert BC(brand=S.BRAND_EXACT[n][0]) == S.BRAND_EXACT[n][0]


@case
def t_own_two_char_brand_matches_in_copy():
    """자사 2글자 브랜드(닥스·리복 등)는 문구 안에서도 잡혀야 한다."""
    assert BC(title="리복 운동화 30% 할인") == "리복"
    assert BC(title="빠투 신상 입고") == "빠투"


@case
def t_nan_safe():
    """NaN·None 입력에서 죽지 않는다 (실데이터에 NaN이 섞인다)."""
    for v in (None, float("nan"), np.nan, ""):
        assert BC(v, v, v) == S.BRAND_UNKNOWN


@case
def t_add_tags_adds_brand2_and_org():
    """add_tags가 brand2·sales_org 컬럼을 붙이고 값이 분류기와 일치한다."""
    df = pd.DataFrame([
        {"title": "[헤지스] 신상", "body": "", "brand": "헤지스남성 아울렛"},
        {"title": "룰렛 돌리고 당첨", "body": "", "brand": "L+DAY 쇼핑핫타임 1차"},
        {"title": "다시 보러 오세요", "body": "", "brand": "주문서이탈"},
    ])
    out = S.add_tags(df)
    assert "brand2" in out.columns and "sales_org" in out.columns
    assert list(out["brand2"]) == ["헤지스", S.BRAND_PROMO, S.BRAND_TRIGGER], list(out["brand2"])
    assert out.loc[0, "sales_org"] in ("e-영업1", "e-영업2")


@case
def t_admin_brand_code_resolved():
    """브랜드칸에 ADMIN브랜드코드만 적혀 있어도 브랜드로 푼다 (HZ·DM·JN)."""
    assert len(S.BRAND_CODE) > 1000, f"코드 표가 {len(S.BRAND_CODE)}개뿐이에요"
    for code, want in (("HZ", "헤지스"), ("DM", "닥스"), ("JN", "질스튜어트 뉴욕"),
                       ("TG", "티엔지티")):
        b, o, k = S.brand_from_copy(brand_raw=code)
        assert b == want and k == S.KIND_BRAND, f"{code} → {b}"
        assert o in ("e-영업1", "e-영업2"), f"{code} 영업 {o}"


@case
def t_brand_code_not_matched_inside_copy():
    """코드는 짧고 의미 없는 문자열이라 문구 안에서 부분 일치하면 안 된다."""
    assert S.brand_from_copy(title="HZ 같은 글자가 섞인 문구 DM TG")[2] != S.KIND_BRAND


@case
def t_kind_buckets():
    """구분(brand_kind)이 성격별로 갈린다."""
    assert S.brand_from_copy(brand_raw="헤지스남성 아울렛")[2] == S.KIND_BRAND
    assert S.brand_from_copy(brand_raw="L+DAY 아울렛")[2] == S.KIND_PROMO
    assert S.brand_from_copy(brand_raw="주문서이탈")[2] == S.KIND_TRIGGER
    assert S.brand_from_copy(brand_raw="DD", cat="골프")[2] == S.KIND_CATEGORY
    assert S.brand_from_copy(brand_raw="???")[2] == S.KIND_ETC
    # 시트에 등록된 이름이면 제휴처라도 브랜드로 잡는다(한국금거래소=e-영업3)
    assert S.brand_from_copy(brand_raw="한국금거래소")[2] == S.KIND_BRAND


@case
def t_category_fallback_labels():
    """브랜드를 못 집어도 카테고리를 알면 그 단위로 묶어 준다."""
    b, _o, k = S.brand_from_copy(brand_raw="AR", cat="남성")
    assert b == "남성 외" and k == S.KIND_CATEGORY, (b, k)
    # '통합'은 카테고리 정보가 사실상 없는 값이라 묶지 않는다
    assert S.brand_from_copy(brand_raw="???", cat="통합")[0] == S.BRAND_UNKNOWN


@case
def t_partner_keeps_its_name():
    """시트에 없는 제휴처는 적어 둔 이름을 살려 둔다 — 어디 제휴인지가 정보다."""
    b, _o, k = S.brand_from_copy(brand_raw="쿠팡 제휴 프로모션")
    assert k == S.KIND_PARTNER and b == "쿠팡 제휴 프로모션", (b, k)


@case
def t_no_unclassified_left_when_cat_known():
    """카테고리가 있으면 미분류로 남기지 않는다 (전부 어딘가에 묶인다)."""
    df = pd.DataFrame([
        {"title": "", "body": "", "brand": "DD", "cat": "골프"},
        {"title": "", "body": "", "brand": "HZ", "cat": "남성"},
        {"title": "", "body": "", "brand": "L+DAY 아울렛", "cat": "통합"},
    ])
    out = S.add_tags(df)
    assert S.BRAND_UNKNOWN not in set(out["brand2"]), list(out["brand2"])
    assert "brand_kind" in out.columns


@case
def t_brandset_ver_changes_with_dict():
    """사전이 바뀌면 캐시 무효화 키도 바뀌어야 한다 (안 그러면 옛 분류가 화면에 남는다).

    사전 요소를 새로 추가할 때 brandset_ver()에 넣는 걸 빠뜨리면 여기서 걸린다.
    """
    assert S.brandset_ver() == S.BRANDSET_VER, "BRANDSET_VER가 현재 사전과 어긋나요"
    variants = {
        "alias": dict(alias={**S.BRAND_ALIAS, "ZZTEST": "닥스"}),
        "promo": dict(promo=S.PROMO_KW + ["ZZTEST행사"]),
        "trigger": dict(trigger=S.TRIGGER_KW + ["ZZTEST트리거"]),
        "partner": dict(partner=S.PARTNER_KW + ["ZZTEST제휴"]),
        "stop": dict(stop=set(S.BRAND_STOP) | {"ZZTEST"}),
        "brand_map": dict(nmap=len(S.BRAND_MAP) + 1),
        "brand_exact": dict(nexact=len(S.BRAND_EXACT) + 1),
        "brand_code": dict(ncode=len(S.BRAND_CODE) + 1),
    }
    for name, kw in variants.items():
        assert S.brandset_ver(**kw) != S.BRANDSET_VER, f"{name}을 바꿔도 버전이 그대로예요"


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
    print(f"브랜드 분류 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
