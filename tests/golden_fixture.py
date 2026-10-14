"""골든 회귀용 **날짜 고정** 합성 데이터.

`smoke_pages.synth_store`를 쓰면 안 된다 — 그건 **오늘을 끝으로** 만들어서 날마다
값이 달라진다. 골든은 「어제와 같은 답이 나오는가」를 보는 검사라, 픽스처가 달력에
걸리면 **코드를 안 고쳐도 내일 빨간불**이 난다(이 저장소가 두 번 밟은 자리다).

그래서 여기 있는 건 전부 **절대 날짜**고 난수 씨앗도 고정이다. 이 파일을 고치면
골든 값이 당연히 달라지니, **고친 이유를 커밋 메시지에 남길 것.**
"""
import numpy as np
import pandas as pd

GOLDEN_FROM = "2025-01-06"        # 월요일
GOLDEN_DAYS = 420                 # 1년을 넘긴다 — 전년비 경로를 밟아야 한다
GOLDEN_SEED = 20260101

CATS = ["여성", "남성", "슈즈", "잡화", "골프", "리빙"]
STYPES = ["우수발송 1", "우수발송 2", "시그니처", "컨틴전시 A", "컨틴", "남은발송"]
BPUS = ["1BPU", "2BPU", "3BPU", "4BPU", "마케팅", "편성"]
BRANDS = ["헤지스", "질스투어트", "닥스", "마에스트로", "TNGT", "위켄드세일 추가 적립"]
OWNERS = ["김담당", "이담당", "박담당"]
HOURS = [800, 1000, 1200, 1600, 1900, 2100]
PRIOS = [0, 1, 2, 3, 4]

TITLES = [
    "오늘만 {b} 최대 50% 할인",
    "{b} 신상 입고 · 선착순 쿠폰",
    "[{b}] 타임딜 마감임박",
    "회원님께 드리는 {b} 단독 특가",
    "{b} 무료배송 이벤트 응모",
    "장바구니에 담아두신 {b} 상품",
]
BODIES = [
    "지금 확인하고 적립 받아가세요 🎁",
    "수량 한정이라 조기 소진될 수 있어요",
    "오늘 자정까지만 드리는 혜택입니다",
    "베스트 상품을 MD가 골라 담았어요",
]


def golden_store():
    """실적 저장소 모양의 프레임 — 컬럼·값 분포가 운영과 같은 꼴이다."""
    rng = np.random.default_rng(GOLDEN_SEED)
    dates = pd.date_range(GOLDEN_FROM, periods=GOLDEN_DAYS, freq="D")
    rows = []
    for i, d in enumerate(dates):
        for j in range(3):
            k = i * 3 + j
            # **축마다 다른 난수로 뽑는다.** 전부 `k % 6`으로 돌리면 카테고리·발송유형·
            # BPU가 한 몸이 되어, 다이제스트가 축 여섯을 훑어도 **같은 칸을 여섯 번**
            # 세게 된다 — 스냅샷이 축을 구분하지 못한다.
            cat = CATS[int(rng.integers(len(CATS)))]
            brand = BRANDS[int(rng.integers(len(BRANDS)))]
            stype = STYPES[int(rng.integers(len(STYPES)))]
            hour = HOURS[int(rng.integers(len(HOURS)))]
            send = int(rng.integers(40_000, 400_000))
            # 요일·시간대·카테고리로 CTR이 갈리게 둔다 — 전부 같은 값이면
            # 「달라졌다」를 보는 검사가 아무것도 못 본다
            ti = int(rng.integers(len(TITLES)))
            base = 0.018 + 0.004 * (d.dayofweek < 5) + 0.0015 * (hour in (1600, 1900))
            base += 0.002 * (cat in ("여성", "슈즈"))
            # 문구 하나에 **진짜 효과**를 심는다. 안 심으면 소구 발굴이 늘 빈 목록이라
            # 골든이 그 경로를 아예 안 밟는다 — 스냅샷에 0만 남는다.
            base += 0.006 * (ti == 2)              # "[{b}] 타임딜 마감임박"
            uv = int(rng.binomial(send, min(max(base, 0.001), 0.999)))
            oc = int(rng.binomial(uv, 0.035)) if uv else 0
            rows.append({
                # **저장소의 날짜는 `YYYYMMDD` 문자열이다.** `_finalize`가
                # `format="%Y%m%d"`로 파싱하므로 datetime을 넣으면 `dt`가 통째로 NaT가
                # 되고, 거기서 파생되는 `dow_k`가 전부 비어 플레이북이 빈 목록을 낸다.
                "date": d.strftime("%Y%m%d"),
                "af": f"AP{(k % 90) + 10:02d}",
                "cat": cat, "brand": brand, "stype": stype,
                "bpu": BPUS[int(rng.integers(len(BPUS)))],
                "owner": OWNERS[int(rng.integers(len(OWNERS)))],
                "hour": hour, "prio": PRIOS[int(rng.integers(len(PRIOS)))],
                "send": send, "uv": uv, "oc": oc,
                "amt": int(oc * rng.integers(45_000, 95_000)),
                "cust": int(send * 0.92), "visit": uv,
                "target": send, "promo": "",
                "title": TITLES[ti].format(b=brand),
                "body": BODIES[int(rng.integers(len(BODIES)))],
                "matched": True, "copy_fix": (k % 7 == 0), "remain_add": (k % 11 == 0),
            })
    df = pd.DataFrame(rows)
    # 저장소가 실제로 들고 있는 파생 둘 — 없으면 `playbook_slots`가 `KeyError`로 죽는다
    df["infl_cr"] = df["uv"] / df["send"].replace(0, np.nan)
    df["ord_cr"] = df["oc"] / df["uv"].replace(0, np.nan)
    # 시간대 오타 — `norm_hhmm`이 복구하는 두 가지를 일부러 심는다
    df.loc[df.index[5], "hour"] = 8000
    df.loc[df.index[9], "hour"] = 2400
    return df


def golden_mtd():
    """전사 MTD 모양 — `MTD_METRICS`의 **키를 그대로** 칼럼으로 쓴다.

    칼럼 이름을 적당히 지으면 `compute_mtd`가 `KeyError`로 죽는다. 파서가 그 키로
    값을 집으므로, 픽스처도 같은 이름이어야 운영과 같은 길을 탄다.
    """
    rng = np.random.default_rng(GOLDEN_SEED + 1)
    dates = pd.date_range(GOLDEN_FROM, periods=GOLDEN_DAYS, freq="D")
    rows = []
    for d in dates:
        cust = int(rng.integers(1_500_000, 2_200_000))
        tsend = int(cust * float(rng.uniform(1.2, 1.8)))
        uinf = int(rng.integers(40_000, 90_000))
        tinf = int(uinf * float(rng.uniform(1.3, 2.1)))
        pcust = int(rng.integers(1_500, 4_000))
        pcnt = int(pcust * float(rng.uniform(1.05, 1.4)))
        rev = int(rng.integers(90_000_000, 260_000_000))
        rows.append({
            "date": d,
            "perSend": round(tsend / cust, 4), "revenue": rev,
            "rps": round(rev / tsend, 4), "totalSend": tsend, "customers": cust,
            "ctr": round(uinf / cust, 6), "uniqueInflow": uinf, "totalInflow": tinf,
            "visitPerPerson": round(tinf / uinf, 4),
            "purchaseCust": pcust, "purchaseCnt": pcnt,
            "purchasePerPerson": round(pcnt / pcust, 4),
            "avgOrderVal": round(rev / pcust, 1), "unitPrice": round(rev / pcnt, 1),
            "mRevenue": int(rev * 0.012), "pointM": int(rev * 0.004),
        })
    return pd.DataFrame(rows)
