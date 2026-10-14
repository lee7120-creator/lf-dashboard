"""CI 샤드 표 — 검사를 몇 갈래로 나눠 «동시에» 돌린다.

단일 잡 직렬로 24단계를 돌던 때 **30~36분**이 걸렸고 제한이 40분이었다. 여유가
4분뿐이라 검사를 하나만 더 붙여도, 혹은 러너가 느린 날이면 **전부 통과한 PR이
벽에 잘려 빨간불**로 끝난다. 실제로 전에 한 번 물려서(2026-09-09, 20분 → 40분으로
올림) CLAUDE.md에 기록이 남아 있다.

샤드를 나눠도 **벽시계의 바닥은 «제일 느린 파일 하나»**다. `test_weekly_perf`가
10.8분이라 3샤드든 6샤드든 11.5분에서 더 안 내려간다 — 그래서 러너를 적게 쓰는
3샤드로 둔다(실측 기반, 아래 표 주석 참고).

**새 검사 파일이 샤드에 안 들어가면 CI에서 조용히 빠진다.** 이 저장소에서 제일
흔한 고장 모양(화면은 멀쩡히 뜨는데 값만 틀림)과 같은 자리라, `--verify`가
`tests/` 디렉터리와 이 표를 대조해 **빠진 것·중복·없는 파일**을 전부 잡는다.
CI의 첫 관문(gate)에서 이걸 돌리므로 샤드에 안 넣으면 머지가 막힌다.

쓰는 법:
    python tests/ci_shards.py --verify       # 표 ↔ 디렉터리 대조
    python tests/ci_shards.py --list 1       # 1번 샤드가 돌릴 파일
    python tests/ci_shards.py --run 1        # 1번 샤드 실행
    python tests/ci_shards.py --plan         # 샤드별 예상 시간
    python tests/ci_shards.py --matrix       # 워크플로 매트릭스(JSON)
"""
import argparse
import json
import pathlib
import re
import subprocess
import sys
import time

ROOT = pathlib.Path(__file__).resolve().parent.parent
TESTS = ROOT / "tests"

# 최근 성공 런 실측(초). 균형을 다시 잡을 때 쓰는 **근거**지 실행에는 안 쓴다 —
# 숫자가 낡아도 검사는 그대로 돈다. 크게 틀어지면 `--plan`으로 다시 나눈다.
SECONDS = {
    "test_weekly_perf": 650, "test_funnel_page": 266, "smoke_pages": 210,
    "test_send_volume_band": 101, "test_site_metrics": 93, "test_inflow_funnel": 90,
    "test_filter_follows_upload": 84, "test_prio_reduction": 83, "test_target_ctr": 71,
    "smoke_weekly_report": 45, "test_orgcat": 38, "test_detail": 22,
    "test_rerun_cost": 19, "test_summary_top": 12, "test_plan_merge": 11,
    "check_shadowing": 1, "test_store_layer": 1, "test_brand_classify": 1,
    "test_note_editor": 1, "test_hour_norm": 1, "test_copy_effect": 1,
    "test_table_export": 1,
}

# 첫 관문 — 몇 초짜리 정적 검사. 샤드 셋을 띄우기 «전»에 돌려서, 문법이 깨졌으면
# 러너 셋을 11분씩 태우지 않고 30초에 끝낸다.
GATE = ["check_shadowing"]

# 샤드 — LPT(긴 것부터 제일 한가한 쪽에)로 나눈 결과. `test_weekly_perf`가 혼자
# 한 샤드인 건 그게 바닥이라서다. 다른 둘을 그 시간 안에 채워 넣었다.
SHARDS = {
    "1": ["test_weekly_perf"],
    "2": ["test_funnel_page", "smoke_pages", "test_prio_reduction",
          "test_detail", "test_summary_top", "test_plan_merge",
          "test_table_export", "test_hour_norm"],
    "3": ["test_send_volume_band", "test_site_metrics", "test_inflow_funnel",
          "test_filter_follows_upload", "test_target_ctr", "smoke_weekly_report",
          "test_orgcat", "test_rerun_cost", "test_store_layer",
          "test_brand_classify", "test_note_editor", "test_copy_effect"],
}


def discovered():
    """`tests/`에서 실제로 돌릴 수 있는 검사 — 실행 진입점이 있는 것만."""
    out = set()
    for p in sorted(TESTS.glob("*.py")):
        if p.name == "ci_shards.py":
            continue
        if 'if __name__ == "__main__"' in p.read_text(encoding="utf-8"):
            out.add(p.stem)
    return out


WORKFLOW = ROOT / ".github" / "workflows" / "dashboard-ci.yml"


def verify_workflow():
    """워크플로가 이 표를 «실제로» 본다는 것까지 확인한다.

    검사가 CI에서 조용히 빠지는 길이 둘이다 — ①표에 안 넣거나 ②표에는 넣었는데
    그 샤드를 워크플로가 안 띄우거나. ①은 아래 `verify()`가 잡고, ②는 지금
    매트릭스를 `--matrix`에서 받아 오니 생길 수가 없다. 그 전제가 깨지는 것
    (매트릭스에 샤드 번호를 도로 박는 것)만 여기서 막는다.
    """
    if not WORKFLOW.exists():
        return []
    # **줄을 보고 판단한다 — 파일 어딘가에 그 글자가 있는지가 아니라.** 위 머리말
    # 주석이 둘 다 언급하고 있어서, 그냥 `in w`로 보면 주석이 규칙을 대신 만족시켜
    # 진짜로 지운 걸 못 잡는다(실제로 못 잡았다).
    lines = WORKFLOW.read_text(encoding="utf-8").splitlines()
    bad = []
    if not any(re.match(r"\s*shard:\s*\$\{\{\s*fromJSON\(needs\.gate\.outputs"
                        r"\.shards\)\s*\}\}\s*$", ln) for ln in lines):
        bad.append("워크플로 매트릭스가 `--matrix`를 안 봅니다.\n"
                   "     → 샤드 번호를 거기 박으면 표에만 있고 «안 도는» 샤드가 생깁니다.")
    if not any(re.match(r"\s*(-\s*)?run:.*ci_shards\.py\s+--verify", ln) for ln in lines):
        bad.append("워크플로가 `ci_shards.py --verify`를 안 돌립니다.\n"
                   "     → 샤드에 안 들어간 새 검사를 아무도 못 잡습니다.")
    return bad


def verify():
    """표 ↔ 디렉터리 대조. 빠진 것·중복·없는 파일을 전부 말한다."""
    listed = GATE + [m for s in SHARDS.values() for m in s]
    have, want = set(listed), discovered()
    bad = []
    if len(listed) != len(have):
        dup = sorted({m for m in listed if listed.count(m) > 1})
        bad.append(f"샤드 표에 중복: {dup}")
    missing = sorted(want - have)
    if missing:
        bad.append(f"**샤드에 안 들어간 검사**: {missing}\n"
                   f"     → CI에서 조용히 빠집니다. `SHARDS`에 넣어 주세요.")
    ghost = sorted(have - want)
    if ghost:
        bad.append(f"표에만 있고 파일이 없음: {ghost}")
    unmeasured = sorted(want - set(SECONDS))
    bad += verify_workflow()
    if bad:
        print("샤드 표가 디렉터리·워크플로와 안 맞아요:")
        for b in bad:
            print(f"  ❌ {b}")
        return 1
    print(f"샤드 표 OK ✅ — 관문 {len(GATE)}개 + 샤드 {len(SHARDS)}개에 검사 {len(want)}개")
    if unmeasured:
        print(f"  (참고: 실측 시간이 없는 검사 {unmeasured} — `--plan`이 0초로 셉니다)")
    return 0


def plan():
    print(f"{'샤드':<6}{'예상':>8}  검사")
    tot = 0
    for sid, mods in SHARDS.items():
        s = sum(SECONDS.get(m, 0) for m in mods)
        tot = max(tot, s)
        print(f"  {sid:<4}{s/60:>6.1f}분  {' '.join(mods)}")
    g = sum(SECONDS.get(m, 0) for m in GATE)
    print(f"\n  관문 {g}초 + 제일 긴 샤드 {tot/60:.1f}분 + 셋업 ≈ "
          f"벽시계 {(g + tot + 40)/60:.1f}분")
    print(f"  (직렬로 다 돌리면 {sum(SECONDS.values())/60:.1f}분)")
    return 0


def matrix():
    """워크플로가 띄울 샤드 목록(JSON). **워크플로에 샤드를 또 적지 않는다** —
    따로 적으면 `SHARDS`에 넣었는데 매트릭스에 안 넣은 샤드가 생기고, 그 샤드의
    검사는 `--verify`를 통과한 채 CI에서 통째로 안 돈다."""
    print(json.dumps(list(SHARDS)))
    return 0


def run(names):
    fails, t0 = [], time.perf_counter()
    for m in names:
        print(f"\n{'=' * 64}\n▶ {m}\n{'=' * 64}", flush=True)
        t = time.perf_counter()
        r = subprocess.run([sys.executable, str(TESTS / f"{m}.py")], cwd=ROOT)
        el = time.perf_counter() - t
        print(f"◀ {m} — {'OK' if r.returncode == 0 else 'FAIL'} ({el:.0f}초)", flush=True)
        if r.returncode != 0:
            fails.append(m)
    print(f"\n{'=' * 64}\n전체 {time.perf_counter() - t0:.0f}초")
    if fails:
        print(f"실패: {fails}")
        return 1
    print(f"통과 ✅ ({len(names)}개)")
    return 0


def main():
    ap = argparse.ArgumentParser(description="CI 샤드")
    ap.add_argument("--verify", action="store_true", help="표 ↔ 디렉터리 대조")
    ap.add_argument("--plan", action="store_true", help="샤드별 예상 시간")
    ap.add_argument("--matrix", action="store_true", help="워크플로 매트릭스(JSON)")
    ap.add_argument("--list", metavar="ID", help="그 샤드가 돌릴 파일")
    ap.add_argument("--run", metavar="ID", help="그 샤드 실행 (gate도 가능)")
    a = ap.parse_args()
    if a.verify:
        return verify()
    if a.plan:
        return plan()
    if a.matrix:
        return matrix()
    for want, fn in ((a.list, lambda ms: print(" ".join(ms)) or 0), (a.run, run)):
        if not want:
            continue
        if want != "gate" and want not in SHARDS:
            print(f"그런 샤드가 없어요: {want!r} (있는 것: gate, {', '.join(SHARDS)})")
            return 2
        return fn(GATE if want == "gate" else SHARDS[want])
    ap.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
