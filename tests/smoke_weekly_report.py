"""주간보고(weekly_report.py) 전 페이지 렌더 스모크.

send_perf_dashboard와 달리 이 앱은 자동 렌더 테스트가 없어서, 2026-08 디버깅 때
week_like()의 라벨 파싱 크래시(앱 초기 렌더부터 다운)를 아무도 못 잡고 있었다.
그 공백을 메우는 테스트다.

스토어가 로컬 CSV(wr_data_store.csv)라 임시 디렉터리에 합성 데이터를 깔고 cwd를 옮겨 실행한다.

로컬 실행:
    python tests/smoke_weekly_report.py
"""
import os
import pathlib
import shutil
import sys
import tempfile

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
APP = ROOT / "weekly_report.py"
TIMEOUT = 300

STORE_COLS = ["gran", "metric", "segment", "year", "label", "close", "sortkey", "value"]
METRICS = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가",
           "비회원트래픽", "가입자수", "가입율", "당일가입CR"]
CHANNELS = ["*TOTAL", "네이버", "카카오", "구글", "직접유입"]
PUSH = ["앱푸시수신동의", "앱푸시_신규추가", "앱푸시_이탈",
        "앱푸시_동의자수", "앱푸시_유효회원", "앱푸시_수신동의전체"]


def _val(metric, i):
    rng = np.random.default_rng(abs(hash(metric)) % 9999 + i)
    if metric == "첫구매 거래액":
        return float(rng.integers(5e7, 2e8))
    if metric in ("가입율", "당일가입CR"):
        return float(rng.uniform(0.5, 5.0))
    if metric == "첫구매 객단가":
        return float(rng.integers(80_000, 160_000))
    if metric.startswith("앱푸시"):
        return float(rng.integers(500_000, 800_000))
    return float(rng.integers(500, 40_000))


def synth_store(with_5th=True, years=(2025, 2026)):
    """업로드 누적본과 같은 모양의 합성 스토어.

    with_5th=True면 일부 달에 5주차를 만들어 '전월에 같은 주차가 없어 마지막 주로
    대체'하는 week_like 경로까지 태운다.
    """
    rows, i = [], 0
    for year in years:
        for mo in range(1, 13):
            for met in METRICS:
                for seg in CHANNELS:
                    i += 1
                    rows.append(dict(gran="월", metric=met, segment=seg, year=year,
                                     label=f"{mo}월", close="final",
                                     sortkey=year * 10000 + mo * 100, value=_val(met, i)))
            nweeks = 5 if (with_5th and mo % 3 == 1) else 4
            for wk in range(1, nweeks + 1):
                for met in METRICS:
                    for seg in CHANNELS:
                        i += 1
                        rows.append(dict(gran="주", metric=met, segment=seg, year=year,
                                         label=f"{mo:02d}월 {wk}주차", close="final",
                                         sortkey=year * 10000 + mo * 100 + wk,
                                         value=_val(met, i)))
            for dd in (1, 8, 15, 22):
                for met in METRICS[:3] + PUSH:
                    i += 1
                    rows.append(dict(gran="일", metric=met, segment="*TOTAL", year=year,
                                     label=f"{mo}/{dd}", close="final",
                                     sortkey=year * 10000 + mo * 100 + dd, value=_val(met, i)))
    return pd.DataFrame(rows)[STORE_COLS]


def run_pages(store, tag):
    from streamlit.testing.v1 import AppTest
    tmp = tempfile.mkdtemp()
    app = os.path.join(tmp, "weekly_report.py")
    shutil.copy(APP, app)
    # 표 엑셀 내보내기 공용 모듈 — 같이 안 옮기면 다운로드 버튼 경로가 통째로 미검증
    _te = ROOT / "table_export.py"
    if _te.exists():
        shutil.copy(_te, os.path.join(tmp, "table_export.py"))
    store.to_csv(os.path.join(tmp, "wr_data_store.csv"), index=False, encoding="utf-8-sig")
    cwd = os.getcwd()
    os.chdir(tmp)
    fails = []
    try:
        at = AppTest.from_file(app, default_timeout=TIMEOUT)
        at.run()
        if at.exception:
            print(f"  FAIL [{tag}] 초기 렌더: {at.exception[0].value}")
            return [f"{tag}:init"]
        radios = [r for r in at.radio if r.label == "페이지"] or \
                 [r for r in at.sidebar.radio if r.label == "페이지"]
        if not radios:
            print(f"  FAIL [{tag}] 페이지 라디오를 찾지 못했어요")
            return [f"{tag}:nav"]
        for p in list(radios[0].options):
            try:
                a = AppTest.from_file(app, default_timeout=TIMEOUT)
                a.run()
                rs = [r for r in a.radio if r.label == "페이지"] or \
                     [r for r in a.sidebar.radio if r.label == "페이지"]
                rs[0].set_value(p)
                a.run()
                if a.exception:
                    raise RuntimeError(a.exception[0].value)
                print(f"  OK   [{tag}] {p}")
            except Exception as e:                       # noqa: BLE001
                print(f"  FAIL [{tag}] {p}: {str(e)[:200]}")
                fails.append(f"{tag}:{p}")
                continue

            # 페이지 안의 다른 라디오(비교 기준·차트 보기 기준 등)도 전부 눌러 본다.
            # '주간 ↔ 월누적(MTD)'처럼 코드 경로를 크게 가르는 것이 있어서,
            # 기본값만 보면 나머지 경로가 통째로 미검증으로 남는다.
            for r in [x for x in a.radio if x.label != "페이지"]:
                opts = list(r.options)
                for v in opts[1:]:
                    sub = f"{p} › {r.label}={v}"
                    try:
                        c = AppTest.from_file(app, default_timeout=TIMEOUT)
                        c.run()
                        _cr = [x for x in c.radio if x.label == "페이지"] or \
                              [x for x in c.sidebar.radio if x.label == "페이지"]
                        _cr[0].set_value(p)
                        c.run()
                        _tgt = [x for x in c.radio if x.label == r.label]
                        if not _tgt:
                            continue
                        _tgt[0].set_value(v)
                        c.run()
                        if c.exception:
                            raise RuntimeError(c.exception[0].value)
                        print(f"  OK   [{tag}] {sub}")
                    except Exception as e:               # noqa: BLE001
                        print(f"  FAIL [{tag}] {sub}: {str(e)[:200]}")
                        fails.append(f"{tag}:{sub}")
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    return fails


def check_yoy_summary():
    """실적 요약 표 — 컬럼 구성과 전월비(당월) 계산을 직접 확인한다.

    렌더 스모크는 '표가 떴다'까지만 본다. 컬럼이 빠지거나 엉뚱한 두 값을 맞대도
    화면은 멀쩡히 뜨므로 여기서 값으로 잡는다.
    """
    sys.path.insert(0, str(ROOT))
    import weekly_report as W

    # 7월 → 8월이 정확히 0.9배, 전년 대비 정확히 0.8배가 되게 깔아 둔다
    rows = []
    for y in (2025, 2026):
        for m in (7, 8):
            for met in METRICS:
                base = 0.5 if met in ("가입율", "당일가입CR") else 1000.0
                v = base * (1.0 if y == 2025 else 0.8) * (1.0 if m == 7 else 0.9)
                for close in ("final", "mtd"):
                    rows.append(dict(gran="월", metric=met, segment="*TOTAL", year=y,
                                     label=W.month_label(m), close=close,
                                     sortkey=y * 10000 + m * 100, value=v))
    df = pd.DataFrame(rows)
    tbl, (pm_y, pm_m) = W.yoy_summary_table(df, 2026, 8, METRICS)
    fails = []
    want = ["2025년 7월", "2026년 7월", "전년비(전월)",
            "2025년 8월", "2026년 8월", "전년비(당월)", "전월비(당월)"]
    if list(tbl.columns) != want:
        print(f"  FAIL [요약표] 컬럼 구성 — {list(tbl.columns)}")
        fails.append("요약표:컬럼")
    elif (pm_y, pm_m) != (2026, 7):
        print(f"  FAIL [요약표] 전월 라벨 — {(pm_y, pm_m)}")
        fails.append("요약표:전월")
    else:
        # 전월비(당월)은 같은 해 8월 ↔ 7월. 0.9배로 깔았으니 △10.0%.
        # (비율 지표는 %p라 0.4 - 0.5 = △0.10%p… 단위가 %면 △0.05%p)
        got = str(tbl.loc["첫구매 고객수", "전월비(당월)"])
        if got != "△10.0%":
            print(f"  FAIL [요약표] 전월비(당월) 계산 — {got} (기대 △10.0%)")
            fails.append("요약표:전월비")
        # 전년비 칼럼이 전월비로 덮이지 않았는지도 같이 본다
        gotyoy = str(tbl.loc["첫구매 고객수", "전년비(당월)"])
        if gotyoy != "△20.0%":
            print(f"  FAIL [요약표] 전년비(당월)가 바뀌었어요 — {gotyoy} (기대 △20.0%)")
            fails.append("요약표:전년비")
    if not fails:
        print("  OK   [요약표] 컬럼 구성·전월비(당월) 계산")
    return fails


def main():
    if not APP.exists():
        print(f"앱 파일을 찾을 수 없어요: {APP}")
        return 1
    fails = []

    print("── 실적 요약 표(전년비·전월비) ──")
    fails += check_yoy_summary()

    print("── 5주차 포함 2개년 ──")
    fails += run_pages(synth_store(with_5th=True), "5주차")

    # 비표준 주간 라벨 방어 — 백업 CSV 복원은 컬럼만 보고 내용은 검증하지 않아서
    # 'MM월 N주차'가 아닌 행이 섞일 수 있다. 예전엔 week_like가 여기서 죽었다.
    print("\n── 비표준 주간 라벨 섞임(백업 복원 경로) ──")
    base = synth_store(with_5th=True)
    base = base[~((base["gran"] == "주") & (base["year"] == 2026)
                  & (base["sortkey"] > 2026 * 10000 + 1 * 100 + 5))]
    odd = base[(base["gran"] == "주") & (base["year"] == 2025)].head(2).copy()
    odd["label"] = "12월"
    odd["sortkey"] = 2025 * 10000 + 12 * 100
    fails += run_pages(pd.concat([base, odd], ignore_index=True)[STORE_COLS], "비표준라벨")

    print()
    if fails:
        print(f"실패 {len(fails)}건: {fails}")
        return 1
    print("주간보고 전 페이지 스모크 통과 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
