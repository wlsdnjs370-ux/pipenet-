# -*- coding: utf-8 -*-
"""모듈 F 리팩터링 사전 조사 — 눈이 아니라 기계로 훑는다.

4,900줄(라우트) + 4,170줄(JS)을 사람이 훑으면 반드시 놓친다. 고치기 전에
**무엇이 문제인지 목록으로** 만든다:

    ① 죽은 함수      정의는 있는데 아무도 안 부르는 것 (JS · 파이썬)
    ② 중복 덩이      6줄 이상 똑같이 반복되는 블록
    ③ 큰 함수        한 함수가 60줄을 넘는 곳 — 쪼갤 후보
    ④ 반복 조회      같은 `$("...")` 를 한 함수에서 여러 번 부르는 곳
    ⑤ 넓은 except    `except Exception` 이 사유 없이 삼키는 곳

★고칠 목록이 아니라 **볼 목록**이다. 여기 뜬 것을 그대로 고치면 안 된다 —
  죽은 것처럼 보이는 함수가 문자열로 불릴 수 있고, 중복처럼 보이는 것이
  일부러 갈라 둔 것일 수 있다. 하나씩 확인하고 고친다.

    python scripts/_probe_f_refactor_scan.py
"""
from __future__ import annotations

import ast
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PY_DIR = ROOT / "routes" / "module_f"
JS = ROOT / "static" / "module_f.js"
HTML = ROOT / "templates" / "module_f.html"
BIG_FN = 60
DUP_MIN = 6


def py_files() -> list:
    return sorted(PY_DIR.glob("*.py"))


def scan_python() -> None:
    print("\n" + "=" * 78)
    print("■ 파이썬 — routes/module_f")
    print("=" * 78)
    defs: dict[str, tuple] = {}
    calls: Counter = Counter()
    big: list = []
    wide: list = []
    for f in py_files():
        src = f.read_text(encoding="utf-8")
        try:
            tree = ast.parse(src)
        except SyntaxError as exc:
            print(f"   ★{f.name} 파싱 실패: {exc}")
            continue
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                n = node.end_lineno - node.lineno + 1
                defs.setdefault(node.name, (f.name, node.lineno, n))
                if n > BIG_FN:
                    big.append((n, f.name, node.lineno, node.name))
            elif isinstance(node, ast.Name):
                calls[node.id] += 1
            elif isinstance(node, ast.Attribute):
                calls[node.attr] += 1
            elif isinstance(node, ast.ExceptHandler):
                t = node.type
                if isinstance(t, ast.Name) and t.id in ("Exception",
                                                        "BaseException"):
                    # 바로 다음 줄에 사유 주석이 있으면 «설명된 것» 으로 본다.
                    line = src.splitlines()[node.lineno - 1]
                    if "#" not in line:
                        wide.append((f.name, node.lineno))
    # 라우트 데코레이터로 등록되는 함수는 이름으로 안 불린다 — 죽은 것이 아니다.
    routed = set()
    for f in py_files():
        src = f.read_text(encoding="utf-8")
        for m in re.finditer(r"@app\.(?:get|post|route)\([^)]*\)\s*\n"
                             r"(?:\s*@[^\n]*\n)*\s*def (\w+)", src):
            routed.add(m.group(1))

    dead = [(n, v) for n, v in sorted(defs.items())
            if calls[n] == 0 and n not in routed
            and not n.startswith("__") and n not in ("register", "job",
                                                     "worker", "main")]
    print(f"\n   ① 아무도 안 부르는 함수 {len(dead)}개")
    for n, (fn, ln, sz) in dead:
        print(f"      {fn}:{ln:<5} {n}  ({sz}줄)")
    print(f"\n   ③ {BIG_FN}줄 넘는 함수 {len(big)}개")
    for sz, fn, ln, n in sorted(big, reverse=True)[:14]:
        print(f"      {sz:>4}줄  {fn}:{ln:<5} {n}")
    print(f"\n   ⑤ 사유 주석 없는 넓은 except {len(wide)}개")
    for fn, ln in wide[:12]:
        print(f"      {fn}:{ln}")


def scan_js() -> None:
    print("\n" + "=" * 78)
    print(f"■ JS — {JS.name}")
    print("=" * 78)
    src = JS.read_text(encoding="utf-8")
    lines = src.splitlines()
    # 함수 정의 — `function 이름(` 과 `const 이름 = (`/`async 이름(`
    defs: dict[str, int] = {}
    for m in re.finditer(r"^\s*(?:async\s+)?function\s+(\w+)\s*\(", src,
                         re.M):
        defs[m.group(1)] = src[:m.start()].count("\n") + 1
    for m in re.finditer(r"^\s*const\s+(\w+)\s*=\s*(?:async\s*)?\(", src,
                         re.M):
        defs.setdefault(m.group(1), src[:m.start()].count("\n") + 1)
    html = HTML.read_text(encoding="utf-8")
    dead = []
    for n, ln in sorted(defs.items()):
        # 정의 자체를 뺀 나머지에서 쓰이나 — 마크업의 onclick 도 본다.
        uses = len(re.findall(rf"\b{re.escape(n)}\b", src)) - 1
        if uses <= 0 and n not in html:
            dead.append((n, ln))
    print(f"\n   ① 아무도 안 부르는 함수 {len(dead)}개")
    for n, ln in dead:
        print(f"      {JS.name}:{ln:<5} {n}")

    # ③ 큰 함수 — 다음 정의까지의 줄 수로 어림한다.
    order = sorted(defs.items(), key=lambda kv: kv[1])
    big = []
    for i, (n, ln) in enumerate(order):
        end = order[i + 1][1] if i + 1 < len(order) else len(lines)
        if end - ln > BIG_FN:
            big.append((end - ln, n, ln))
    print(f"\n   ③ {BIG_FN}줄 넘는 함수 {len(big)}개")
    for sz, n, ln in sorted(big, reverse=True)[:14]:
        print(f"      {sz:>4}줄  {JS.name}:{ln:<5} {n}")

    # ② 중복 덩이 — 공백·주석 뺀 뒤 연속 DUP_MIN 줄이 그대로 반복되는가
    norm = []
    for i, ln in enumerate(lines):
        s = ln.strip()
        if not s or s.startswith("//"):
            norm.append(None)
        else:
            norm.append(s)
    seen: dict[tuple, list] = defaultdict(list)
    for i in range(len(norm) - DUP_MIN):
        blk = tuple(norm[i:i + DUP_MIN])
        if any(x is None for x in blk):
            continue
        seen[blk].append(i + 1)
    dups = [(v, k) for k, v in seen.items() if len(v) > 1]
    print(f"\n   ② {DUP_MIN}줄 이상 그대로 반복 {len(dups)}곳")
    shown = set()
    for locs, blk in sorted(dups, key=lambda t: -len(t[0]))[:8]:
        if locs[0] in shown:
            continue
        shown.update(locs)
        print(f"      줄 {locs} — {blk[0][:64]}")

    # ④ 같은 $("...") 를 여러 번 — 한 함수 안에서
    ids = Counter(re.findall(r'\$\("([\w-]+)"\)', src))
    hot = [(c, i) for i, c in ids.items() if c >= 8]
    print(f"\n   ④ 8번 이상 조회되는 요소 {len(hot)}개")
    for c, i in sorted(hot, reverse=True)[:12]:
        print(f"      {c:>4}회  #{i}")


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    scan_python()
    scan_js()
    print("\n  ★이 목록은 «볼 것» 이지 «고칠 것» 이 아니다 — 하나씩 확인한다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
