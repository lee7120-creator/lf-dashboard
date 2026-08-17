"""문구 개선 효과 — 기획시트 W·X열 파싱과 발송유형 정규화.

기획 시트 오른쪽 블록의 **문구 변경 여부(W)**·**잔여모수 추가 여부(X)**를 읽어
'우리측이 다듬은 문구가 실제로 나았는지'를 비교한다. 이 두 칸을 놓치면 화면은
멀쩡히 뜨고 비교만 통째로 비므로 스모크로는 안 잡힌다.

담당자가 손으로 적는 칸이라 표기가 흔들린다('문구변경'·'문구 변경'·'추가').
빈칸은 **검토했지만 고칠 필요가 없었던 건**이라 대조군으로 쓴다 — 그래서
빈칸을 True로 잘못 읽으면 대조군이 사라져 비교 자체가 무너진다.

발송유형도 같은 유형이 여러 표기로 온다(2026-08부터 '컨틴전시 A/B' → '컨틴').
정규화하지 않으면 층화 비교에서 같은 유형이 두세 칸으로 갈려 표본이 쪼개진다.

로컬 실행:
    python tests/test_copy_effect.py
"""
import pathlib
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import send_perf_dashboard as S  # noqa: E402

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


# 실제 기획시트 구조 — 2줄 헤더 + 오른쪽 별도 블록(AF코드가 두 번 나온다)
HDR1 = ['', '일자(8자리)', '요일', '시간대', '타겟 구분', '발송유형', 'BPU', '우선순위',
        '카테고리', '속성', '담당자', '브랜드', 'AF코드', '기획전 No.', '발송모수',
        '랜딩 링크(모바일)', '타겟팅 조건', 'RED CRM 고객군명', '요청문구', '', '', '', '', '']
HDR2 = [''] * 18 + ['제목', '내용', 'AF코드', '기획전 번호', '문구 변경 여부', '잔여모수 추가 여부']


def prow(d, af, title="제목", body="내용", fix="", rem="", stype="기본발송"):
    return ['세팅완료', d, '월', '1400', '타겟', stype, '1BPU_A', '1', '여성', '정상',
            '김은지', 'HS', af, '113344', '', 'https://x', '조건', 'REAL_x',
            title, body, af, '113344', fix, rem]


@case
def t_flags_parsed():
    """W·X열을 읽어 4-튜플로 넣는다."""
    lk = {}
    S._parse_plan_sheet([HDR1, HDR2,
                         prow("20260810", "AP100", fix="문구변경"),
                         prow("20260813", "AP64", rem="추가"),
                         prow("20260814", "AP85", fix="문구변경", rem="추가")], lk)
    assert lk[("20260810", "AP100")][2] is True, lk[("20260810", "AP100")]
    assert lk[("20260810", "AP100")][3] is False
    assert lk[("20260813", "AP64")][3] is True
    assert lk[("20260814", "AP85")][2:] == (True, True)


@case
def t_blank_is_control_group():
    """빈칸은 '검토했지만 그대로 감' — 반드시 False여야 대조군이 성립한다."""
    lk = {}
    S._parse_plan_sheet([HDR1, HDR2, prow("20260812", "AP122")], lk)
    assert lk[("20260812", "AP122")][2] is False, "빈칸을 변경으로 읽으면 대조군이 사라져요"
    assert lk[("20260812", "AP122")][3] is False


@case
def t_spacing_variants():
    """'문구변경'·'문구 변경'이 섞여 들어온다 (수기 입력)."""
    lk = {}
    S._parse_plan_sheet([HDR1, HDR2,
                         prow("20260810", "AP10", fix="문구변경"),
                         prow("20260811", "AP11", fix="문구 변경"),
                         prow("20260812", "AP12", fix=" 문구변경 ")], lk)
    assert all(lk[k][2] for k in lk), lk


@case
def t_explicit_negatives_are_false():
    """명시적 부정 표기는 False로 되돌린다."""
    for v in ("X", "-", "없음", "미변경", "해당 없음", "nan"):
        lk = {}
        S._parse_plan_sheet([HDR1, HDR2, prow("20260810", "AP10", fix=v)], lk)
        assert lk[("20260810", "AP10")][2] is False, f"{v!r}를 변경으로 읽었어요"


@case
def t_duplicate_key_ors_flags():
    """같은 (날짜, AF)가 두 번 나오면(원발송 + 잔여모수) 체크를 OR로 합친다.

    덮어쓰면 뒤 행의 빈칸이 앞 행의 체크를 지운다.
    """
    lk = {}
    S._parse_plan_sheet([HDR1, HDR2,
                         prow("20260813", "AP64", title="원발송", fix="문구변경"),
                         prow("20260813", "AP64", title="잔여모수", rem="추가")], lk)
    v = lk[("20260813", "AP64")]
    assert v[2] is True and v[3] is True, v


@case
def t_merge_writes_flag_columns():
    """merge_perf_plan이 copy_fix·remain_add 칼럼을 만든다."""
    perf = pd.DataFrame({"date": ["20260810", "20260811"], "af": ["AP100", "AP210"]})
    lk = {("20260810", "AP100"): ("t", "b", True, False),
          ("20260811", "AP210"): ("t", "b", False, True)}
    m = S.merge_perf_plan(perf, lk, keep_unmatched=True)
    assert list(m["copy_fix"]) == [True, False], list(m["copy_fix"])
    assert list(m["remain_add"]) == [False, True]
    assert {"copy_fix", "remain_add"} <= set(S.STORE_COLS), "저장소에 안 실리면 다음 세션에 사라져요"


@case
def t_legacy_two_tuple_lookup():
    """옛 2-튜플 lookup을 넘겨도 죽지 않는다 (외부에서 만들어 넘기는 경로가 있다)."""
    perf = pd.DataFrame({"date": ["20260810"], "af": ["AP100"]})
    m = S.merge_perf_plan(perf, {("20260810", "AP100"): ("제목", "내용")}, keep_unmatched=True)
    assert m["matched"].all()
    assert list(m["copy_fix"]) == [False] and list(m["remain_add"]) == [False]


@case
def t_unmatched_rows_get_false():
    """문구를 못 찾은 행은 플래그도 False — NaN이 섞이면 비교에서 조용히 빠진다."""
    perf = pd.DataFrame({"date": ["20260810", "20260901"], "af": ["AP100", "AP999"]})
    m = S.merge_perf_plan(perf, {("20260810", "AP100"): ("t", "b", True, True)},
                          keep_unmatched=True)
    assert list(m["copy_fix"]) == [True, False], list(m["copy_fix"])
    assert m["copy_fix"].dtype == bool


@case
def t_stype_normalized():
    """발송유형 표기 흔들림을 하나로 묶는다 (2026-08부터 컨틴전시 → 컨틴)."""
    assert S.norm_stype("컨틴전시 A") == "컨틴"
    assert S.norm_stype("컨틴전시 B") == "컨틴"
    assert S.norm_stype("컨틴") == "컨틴"
    assert S.norm_stype("시그니쳐") == "시그니처"
    assert S.norm_stype("우수발송 1") == "우수발송"
    assert S.norm_stype("우수발송 8") == "우수발송"
    assert S.norm_stype("기본발송") == "기본발송"
    assert S.norm_stype("남은발송") == "남은발송"


@case
def t_stype_ungrouped_keeps_number():
    """번호별로 봐야 하는 화면을 위해 group=False면 원 번호를 남긴다."""
    assert S.norm_stype("우수발송 3", group=False) == "우수발송 3"
    assert S.norm_stype("컨틴전시 A", group=False) == "컨틴 A"


@case
def t_stype_blank_safe():
    """빈 값·NaN에서 죽지 않는다."""
    for v in ("", None, "nan", "  "):
        assert S.norm_stype(v) is None, v


@case
def t_gsheet_roundtrip_strings_recovered():
    """구글시트 왕복은 True/False를 문자열로 되돌린다 — bool로 복구돼야 한다.

    문자열 'False'는 파이썬에서 참이라, 복구를 빠뜨리면 전 행이 '변경함'이 된다.
    """
    assert S._to_bool("True") is True
    assert S._to_bool("False") is False
    assert S._to_bool("TRUE") is True
    assert S._to_bool("") is False


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
    print(f"문구 개선 효과 테스트 {len(CASES)}건 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
