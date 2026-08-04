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


def main():
    if not APP.exists():
        print(f"앱 파일을 찾을 수 없어요: {APP}")
        return 1
    fails = []

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
