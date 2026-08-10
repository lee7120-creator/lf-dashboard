"""모듈 전역 헬퍼를 지역변수로 덮어쓰는지 정적 검사 (AST).

왜 필요한가 — 2026-07 실제 사고:
    main() 안 어딘가에서 `_s = ...` 한 줄을 쓰는 순간 파이썬은 `_s`를 main() 전체의
    지역변수로 승격시킨다. main() 안의 중첩 함수(render_messages 등)에서 부르던
    모듈 전역 헬퍼 `_s()`는 클로저 자유변수로 묶이고, 그 대입 줄이 실행되기 전에는
    미바인딩이라 NameError로 죽는다. 이 한 줄 때문에 11개 페이지가 전부 다운됐다.

    문법 오류도, 린트 경고도 아니고, 그 페이지를 안 열면 재현도 안 되는 사고라
    정적 검사로 막는 게 가장 확실하다.

로컬 실행:
    python tests/check_shadowing.py
"""
import ast
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
# 활발히 개발 중인 대시보드 — 새 파일을 추가하면 여기에 넣어 주세요
TARGETS = ["send_perf_dashboard.py", "weekly_report.py"]


def _bound_names(fn):
    """함수 fn의 '자기 스코프'에 바인딩되는 이름 → 첫 등장 줄번호.

    중첩 함수/람다의 본문은 별도 스코프라 들어가지 않는다(그쪽은 따로 검사됨).
    """
    bound = {}

    def add(name, lineno):
        if name is not None:
            bound.setdefault(name, lineno)

    for arg in [*fn.args.posonlyargs, *fn.args.args, *fn.args.kwonlyargs]:
        add(arg.arg, fn.lineno)
    add(getattr(fn.args.vararg, "arg", None), fn.lineno)
    add(getattr(fn.args.kwarg, "arg", None), fn.lineno)

    def walk(node):
        for child in ast.iter_child_nodes(node):
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                add(child.name, child.lineno)        # 이름만 바인딩, 본문은 별도 스코프
                continue
            if isinstance(child, ast.Lambda):
                continue
            if isinstance(child, (ast.ListComp, ast.SetComp, ast.DictComp, ast.GeneratorExp)):
                continue                             # 파이썬3에서 컴프리헨션은 별도 스코프
            if isinstance(child, ast.Name) and isinstance(child.ctx, ast.Store):
                add(child.id, child.lineno)
            elif isinstance(child, (ast.Import, ast.ImportFrom)):
                for a in child.names:
                    add((a.asname or a.name).split(".")[0], child.lineno)
            elif isinstance(child, ast.ExceptHandler):
                add(child.name, child.lineno)
            elif isinstance(child, ast.Global) or isinstance(child, ast.Nonlocal):
                for n in child.names:                # 명시적으로 전역/비지역 선언한 건 안전
                    bound.pop(n, None)
            walk(child)

    walk(fn)
    return bound


def _module_names(tree):
    """모듈 스코프에 바인딩되는 이름 → (종류, 줄번호).

    함수뿐 아니라 **임포트한 모듈**과 **대문자 상수**도 본다. 실제로
    `main()` 안에서 `import itertools` 한 줄 때문에 그 함수 앞쪽의
    `itertools.count(...)`가 UnboundLocalError로 죽은 적이 있다 —
    함수만 보던 예전 검사기는 이걸 놓쳤다.
    """
    names = {}
    for n in tree.body:
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
            names[n.name] = ("전역 함수", n.lineno)
        elif isinstance(n, (ast.Import, ast.ImportFrom)):
            for a in n.names:
                names[(a.asname or a.name).split(".")[0]] = ("임포트", n.lineno)
        elif isinstance(n, ast.Assign):
            for t in n.targets:                      # 대문자 상수만 (일반 변수는 오탐이 많다)
                if isinstance(t, ast.Name) and t.id.isupper() and len(t.id) > 2:
                    names.setdefault(t.id, ("전역 상수", n.lineno))
    return names


def check(path):
    tree = ast.parse(path.read_text(encoding="utf-8"))
    mod = _module_names(tree)

    problems = []
    for node in ast.walk(tree):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        for name, lineno in _bound_names(node).items():
            if name in mod and mod[name][1] != lineno:
                kind, def_line = mod[name]
                problems.append((lineno, name, node.name, def_line, kind))
    return sorted(problems)


def main():
    total = 0
    for fname in TARGETS:
        path = ROOT / fname
        if not path.exists():
            print(f"  건너뜀 {fname} (파일 없음)")
            continue
        problems = check(path)
        if not problems:
            print(f"  OK   {fname}")
            continue
        total += len(problems)
        for lineno, name, fn, def_line, kind in problems:
            print(f"  FAIL {fname}:{lineno} — '{fn}()' 안의 '{name}' 이(가) "
                  f"모듈 {kind} '{name}' (line {def_line}) 을(를) 가려요. "
                  + ("모듈 상단 import를 쓰고 함수 안에서 다시 import하지 마세요."
                     if kind == "임포트" else "다른 이름으로 바꿔 주세요."))

    print()
    if total:
        print(f"전역 이름 섀도잉 {total}건 발견 ❌")
        return 1
    print("전역 헬퍼 섀도잉 없음 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
