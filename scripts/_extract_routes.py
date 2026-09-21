# -*- coding: utf-8 -*-
"""Phase 2 도메인 라우트 추출기 (일회성 도구, 커밋 안 함).

대조 서버.py 에서 지정한 route 함수(+선택적 private 헬퍼)의 정확한 소스 라인을
슬라이스해 routes/<module>.py 로 옮기고, 원본에서는 제거한다. 손 전사 오류 없이
185줄짜리 함수도 안전하게 이동. register(app, *, <inject>) 클로저로 감싼다.

사용: 스크립트 하단 CONFIG 를 수정하고 실행. dry-run 으로 먼저 확인.
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path as P

SERVER = P("대조 서버.py")


def _spans(src_tree, names):
    """name -> (deco_start_1idx, end_1idx) inclusive."""
    out = {}
    for node in src_tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in names:
            start = node.decorator_list[0].lineno if node.decorator_list else node.lineno
            out[node.name] = (start, node.end_lineno)
    return out


def extract(module_name, out_path, route_names, private_helpers,
            inject_params, module_imports, header_doc, marker, nest_private=False):
    text = SERVER.read_text(encoding="utf-8")
    lines = text.splitlines(keepends=True)
    tree = ast.parse(text)

    route_spans = _spans(tree, route_names)
    priv_spans = _spans(tree, private_helpers)
    missing = (set(route_names) - set(route_spans)) | (set(private_helpers) - set(priv_spans))
    if missing:
        raise SystemExit(f"함수 못 찾음: {sorted(missing)}")

    def block(span):
        s, e = span
        return "".join(lines[s - 1:e]).rstrip("\n")

    def indent(body):
        return "\n".join(("    " + ln) if ln.strip() else ln
                         for ln in body.split("\n"))

    # ── 라우트 모듈 생성 ──
    parts = [header_doc.rstrip() + "\n", "from __future__ import annotations\n\n"]
    parts.append("\n".join(module_imports) + "\n\n\n")
    # nest_private=False: private 헬퍼는 모듈 레벨(원본 순서). True: register 안에 중첩
    # (주입 이름을 클로저로 참조해야 하는 헬퍼 다수일 때).
    if not nest_private:
        for name in sorted(private_helpers, key=lambda n: priv_spans[n][0]):
            parts.append(block(priv_spans[name]) + "\n\n\n")
    # register 래퍼
    params = ", ".join(inject_params)
    parts.append(f"def register(app, *, {params}):\n")
    if nest_private:
        for name in sorted(private_helpers, key=lambda n: priv_spans[n][0]):
            parts.append("\n" + indent(block(priv_spans[name])) + "\n")
    for name in sorted(route_names, key=lambda n: route_spans[n][0]):
        parts.append("\n" + indent(block(route_spans[name])) + "\n")
    P(out_path).write_text("".join(parts), encoding="utf-8")

    # ── 원본에서 제거 (라우트 + private 헬퍼 라인 범위) ──
    remove = sorted(list(route_spans.values()) + list(priv_spans.values()),
                    key=lambda t: t[0], reverse=True)
    first_line = min(t[0] for t in route_spans.values())
    for s, e in remove:
        del lines[s - 1:e]
    # 첫 라우트 위치에 마커 삽입
    ins = first_line - 1
    # 삭제로 인해 인덱스가 당겨졌으니, 가장 가까운 안전 위치(원본 first_line 자리)에 삽입
    ins = min(ins, len(lines))
    lines.insert(ins, marker + "\n")
    SERVER.write_text("".join(lines), encoding="utf-8")
    print(f"OK: {out_path} 생성, 원본에서 {len(remove)} 블록 제거")


if __name__ == "__main__":
    CONFIG = eval(sys.argv[1]) if len(sys.argv) > 1 else None
    if CONFIG is None:
        raise SystemExit("CONFIG dict 인자 필요")
    extract(**CONFIG)
