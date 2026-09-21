# -*- coding: utf-8 -*-
"""도메인 추출 계획 자동 산출 (일회성 도구).

주어진 route 함수 집합에 대해:
  - 모든 자유 이름(free name)을 계산
  - main 에 Import/ImportFrom 로 들어온 이름 → 그 import 문을 route 모듈에 복제
  - main 에서 정의된 helper/전역/상수/상태 → register() 로 주입
  - 커버 안 된 이름이 남으면 경고(반드시 0 이어야 안전)

출력: inject_params, module_imports(복제할 import 문), private_helpers, 커버리지.
"""
from __future__ import annotations

import ast
import builtins
import sys

SRC = open("대조 서버.py", encoding="utf-8").read()
TREE = ast.parse(SRC)
FUNCS = {n.name: n for n in TREE.body if isinstance(n, ast.FunctionDef)}

ROUTE_FUNCS = {}
for n in TREE.body:
    if isinstance(n, ast.FunctionDef):
        for d in n.decorator_list:
            if (isinstance(d, ast.Call) and isinstance(d.func, ast.Attribute)
                    and isinstance(d.func.value, ast.Name) and d.func.value.id == "app"):
                ROUTE_FUNCS[n.name] = d.args[0].value if d.args and isinstance(d.args[0], ast.Constant) else "?"
HELPERS = {n for n in FUNCS if n not in ROUTE_FUNCS}

# main 의 이름 → import 문 텍스트 (복제용)
IMPORT_OF = {}
ASSIGNED_NAMES = set()
for n in TREE.body:
    if isinstance(n, ast.Import):
        for a in n.names:
            top = (a.asname or a.name).split(".")[0]
            IMPORT_OF[top] = f"import {a.name}" + (f" as {a.asname}" if a.asname else "")
    elif isinstance(n, ast.ImportFrom):
        for a in n.names:
            nm = a.asname or a.name
            IMPORT_OF[nm] = f"from {n.module} import {a.name}" + (f" as {a.asname}" if a.asname else "")
    elif isinstance(n, ast.Assign):
        for t in n.targets:
            if isinstance(t, ast.Name):
                ASSIGNED_NAMES.add(t.id)
    elif isinstance(n, ast.AnnAssign) and isinstance(n.target, ast.Name):
        ASSIGNED_NAMES.add(n.target.id)

BUILTIN = set(dir(builtins))


def freenames(fn):
    assigned = set()
    for node in ast.walk(fn):
        if isinstance(node, ast.arg):
            assigned.add(node.arg)
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
            assigned.add(node.id)
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            for a in node.names:
                assigned.add((a.asname or a.name).split(".")[0])
        if isinstance(node, ast.ExceptHandler) and node.name:
            assigned.add(node.name)
        if isinstance(node, ast.FunctionDef):
            assigned.add(node.name)
    return {node.id for node in ast.walk(fn)
            if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load)
            and node.id not in BUILTIN and node.id not in assigned}


def helper_domains(classifier):
    from collections import defaultdict
    hu = defaultdict(set)
    for f in ROUTE_FUNCS:
        for h in freenames(FUNCS[f]) & HELPERS:
            hu[h].add(classifier(ROUTE_FUNCS[f], f))
    return hu


def plan(route_names, this_domain, classifier):
    allfree = set()
    for f in route_names:
        allfree |= freenames(FUNCS[f])
    hu = helper_domains(classifier)
    used_helpers = allfree & HELPERS
    private = sorted(h for h in used_helpers if hu[h] == {this_domain})
    shared = sorted(h for h in used_helpers if hu[h] != {this_domain})
    # private 헬퍼가 참조하는 이름도 커버 대상에 포함
    for h in private:
        allfree |= freenames(FUNCS[h])
    allfree -= {"app"}
    allfree -= set(route_names) | set(private)
    imports = {}
    inject = []
    uncovered = []
    for nm in sorted(allfree):
        if nm in HELPERS:
            if nm in private:
                continue
            inject.append(nm)
        elif nm in IMPORT_OF:
            imports[nm] = IMPORT_OF[nm]
        elif nm in ASSIGNED_NAMES:
            inject.append(nm)
        else:
            uncovered.append(nm)
    print(f"=== {this_domain} ===")
    print("route_names =", sorted(route_names, key=lambda n: 0))
    print("private_helpers =", private)
    print("inject_params =", inject)
    print("module_imports (복제):")
    for nm, imp in sorted(imports.items()):
        print("   ", imp)
    print("uncovered (반드시 0):", uncovered)
    print()
    return dict(route_names=list(route_names), private_helpers=private,
                inject_params=inject, imports=sorted(set(imports.values())),
                uncovered=uncovered)


def classify(path, name):
    if path.startswith("/api/remote30/system"): return "r30_system"
    if path.startswith("/api/remote30/machineroom"): return "r30_machineroom"
    if path.startswith("/api/remote30/gnn"): return "r30_gnn"
    if path.startswith("/api/remote30/has"): return "r30_has"
    if path.startswith("/api/remote30/inspect"): return "r30_inspect"
    if path.startswith("/api/remote30/combined"): return "r30_combined"
    if "feedback" in path: return "feedback"
    if path.startswith("/api/cad") or "cad-compare" in path: return "cad_compare"
    return "misc"


if __name__ == "__main__":
    target = sys.argv[1]
    routes = [f for f, p in ROUTE_FUNCS.items() if classify(p, f) == target]
    plan(routes, target, classify)
