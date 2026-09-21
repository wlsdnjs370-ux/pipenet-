# -*- coding: utf-8 -*-
"""Phase 2b 공유 코어 추출기 (일회성 도구, 커밋 안 함).

대조 서버.py 에서 지정한 top-level 함수(+상수 대입문)의 정확한 소스 라인을
슬라이스해 core/<module>.py 로 옮기고, 원본에서는 제거한 뒤 그 자리에
`from <module> import (<names>)` 를 삽입한다. 함수 본문·이름은 원본 그대로 유지되므로
모든 주입 배선(register kwargs)과 golden 의 mod.<name> 접근이 그대로 유지된다.

사용: CONFIG dict 를 인자로 전달. repo 루트에서 PYTHONPATH=scripts 로 실행.
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path as P

SERVER = P("대조 서버.py")


def _func_spans(tree, names):
    out = {}
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.ClassDef)) and node.name in names:
            start = node.decorator_list[0].lineno if node.decorator_list else node.lineno
            out[node.name] = (start, node.end_lineno)
    return out


def _assign_spans(tree, names):
    """상수 대입문 (NAME = ...) -> (start, end) inclusive."""
    out = {}
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for tgt in node.targets:
                if isinstance(tgt, ast.Name) and tgt.id in names:
                    out[tgt.id] = (node.lineno, node.end_lineno)
    return out


def extract(module_name, out_path, func_names, const_names,
            module_imports, header_doc, source=None):
    src = P(source) if source else SERVER
    text = src.read_text(encoding="utf-8")
    lines = text.splitlines(keepends=True)
    tree = ast.parse(text)

    fspans = _func_spans(tree, func_names)
    cspans = _assign_spans(tree, const_names)
    missing = (set(func_names) - set(fspans)) | (set(const_names) - set(cspans))
    if missing:
        raise SystemExit(f"못 찾음: {sorted(missing)}")

    def block(span):
        s, e = span
        return "".join(lines[s - 1:e]).rstrip("\n")

    # ── core 모듈 생성 (상수 먼저, 그 뒤 원본 순서대로 함수) ──
    parts = [header_doc.rstrip() + "\n", "from __future__ import annotations\n\n"]
    if module_imports:
        parts.append("\n".join(module_imports) + "\n\n\n")
    for name in sorted(const_names, key=lambda n: cspans[n][0]):
        parts.append(block(cspans[name]) + "\n")
    if const_names:
        parts.append("\n\n")
    ordered_funcs = sorted(func_names, key=lambda n: fspans[n][0])
    parts.append("\n\n".join(block(fspans[n]) for n in ordered_funcs) + "\n")
    P(out_path).write_text("".join(parts), encoding="utf-8")

    # ── 원본에서 제거 + import 삽입 ──
    all_spans = list(fspans.values()) + list(cspans.values())
    first_line = min(s for s, _ in all_spans)
    for s, e in sorted(all_spans, key=lambda t: t[0], reverse=True):
        del lines[s - 1:e]
    # 삽입: import 문 (함수는 원본 정의 순, import 는 알파벳 순서 무관하나 정의순 유지)
    imp_names = ", ".join(sorted(const_names, key=lambda n: cspans[n][0])
                          + ordered_funcs)
    stmt = f"from {module_name} import ({imp_names})  # noqa: E501  (Phase2b core)\n"
    ins = min(first_line - 1, len(lines))
    lines.insert(ins, stmt)
    src.write_text("".join(lines), encoding="utf-8")
    print(f"OK: {out_path} 생성, 원본에서 {len(all_spans)} 블록 제거, import 삽입 @ line {ins+1}")


if __name__ == "__main__":
    CONFIG = eval(sys.argv[1]) if len(sys.argv) > 1 else None
    if CONFIG is None:
        raise SystemExit("CONFIG dict 인자 필요")
    extract(**CONFIG)
