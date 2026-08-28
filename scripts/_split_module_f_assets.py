# -*- coding: utf-8 -*-
"""module_f.html 의 인라인 <style>/<script> 를 정적 파일로 떼어낸다.

측정: 템플릿 171,704자 중 script 127,785 + style 21,045 = **87%** 가 인라인
자산이다. 이건 매 페이지 로드마다 다시 내려가고 브라우저가 캐시할 수 없다.
두 블록 다 Jinja 표현이 **0개** 라 글자 그대로 옮길 수 있다(치환 없음).

실행 시점을 바꾸지 않는다 — <script> 는 </body> 바로 앞에 있었고 외부 파일도
같은 자리에 둔다. `defer` 를 붙이지 않는 이유가 그것이다(붙이면 DOM 준비
시점이 달라진다).

    python scripts/_split_module_f_assets.py [--check]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TPL = ROOT / "templates" / "module_f.html"
CSS = ROOT / "static" / "module_f.css"
JS = ROOT / "static" / "module_f.js"
VER = "20260828-f9"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="쓰지 않고 계획만 본다")
    a = ap.parse_args()

    src = TPL.read_text(encoding="utf-8")
    lines = src.split("\n")

    def find(tag: str, start: int = 0) -> int:
        for i in range(start, len(lines)):
            if lines[i].strip() == tag:
                return i
        raise SystemExit(f"태그를 못 찾음: {tag}")

    s0 = find("<style>")
    s1 = find("</style>", s0)
    j0 = find("<script>", s1)
    j1 = find("</script>", j0)

    css_body = "\n".join(lines[s0 + 1:s1])
    js_body = "\n".join(lines[j0 + 1:j1])

    print(f"  <style>   {s0 + 1:>5}~{s1 + 1:<5} → {len(css_body):>8,}자")
    print(f"  <script>  {j0 + 1:>5}~{j1 + 1:<5} → {len(js_body):>8,}자")
    if "{{" in css_body or "{%" in css_body:
        raise SystemExit("★style 안에 Jinja 표현이 있다 — 그대로 못 옮긴다")
    if "{{" in js_body or "{%" in js_body:
        raise SystemExit("★script 안에 Jinja 표현이 있다 — 그대로 못 옮긴다")
    if "</script" in js_body or "</style" in css_body:
        raise SystemExit("★블록 안에 닫는 태그가 또 있다 — 경계가 틀렸다")

    out = (lines[:s0]
           + [f'<link rel="stylesheet" href="/static/module_f.css?v={VER}">']
           + lines[s1 + 1:j0]
           + [f'<script src="/static/module_f.js?v={VER}"></script>']
           + lines[j1 + 1:])
    new = "\n".join(out)

    print(f"\n  템플릿 {len(src):,}자 → {len(new):,}자 "
          f"({(1 - len(new) / len(src)) * 100:.0f}% 감소)")
    if a.check:
        print("  --check 라 쓰지 않았다.")
        return 0

    CSS.write_text(css_body + "\n", encoding="utf-8", newline="\n")
    JS.write_text(js_body + "\n", encoding="utf-8", newline="\n")
    TPL.write_text(new, encoding="utf-8", newline="\n")
    print(f"  썼다 — {CSS.name} · {JS.name} · {TPL.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
