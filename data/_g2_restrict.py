# -*- coding: utf-8 -*-
"""[G2] `_restrict_to_worst` 이식 + 제한 전개 진입점(design/restrict.py)."""
import io
import re

SRC = "routes/module_f/remote30.py"
lines = io.open(SRC, encoding="utf-8").read().split("\n")
i = next(k for k, ln in enumerate(lines) if ln.startswith("def _restrict_to_worst("))
j = len(lines)
for k in range(i + 1, len(lines)):
    if re.match(r"^(def |class |# ─)", lines[k]):
        j = k
        break
while j - 1 > i and lines[j - 1].strip() == "":
    j -= 1
body = "\n".join(lines[i:j]).replace(
    "def _restrict_to_worst(", "def restrict_to_worst(", 1)
# 모듈 F 의 화면 문구 → G 의 것. 로직이 아니라 로그 문자열이다.
body = body.replace("[변환] 최불리", "[G2] 최불리")

head = '''# -*- coding: utf-8 -*-
"""[G2] corridor 제한 payload 와 «두 번째 전개».

두 번 전개 원칙(지시서 §0.3)::

    EditBoard ─┬─ 전체망 전개 ─────────────→ .kfp   (기존 경로 · 손대지 않는다)
               └─ 최불리 제한 → 제한 전개 → 5표 → .sdf   (이쪽)

같은 손질 결과·같은 치수 입력에서 두 산출이 나오므로 설계 내용이 어긋나지 않는다.
"""
from __future__ import annotations

import math


'''

tail = '''


def expand_worst(payload: dict, board, worst: dict, *,
                 selected_source=None, key: str | None = None) -> dict:
    """제한 payload 로 **두 번째 전개**를 돌린다. 파일을 쓰지 않는다.

    기존 `convert_to_kfp` 의 저장 경로·반환 규약은 건드리지 않는다(§3) — 이쪽은
    메모리 상의 망만 돌려주는 별개 진입점이다.

    반환에는 G3 이 쓸 역참조가 함께 실린다::

        {"ok", "kfp", "edge_ref", "node_ref", "hcov", "head_kinds", "sources", …}

    `edge_ref` 는 «kfp 배관 → 원 board 간선» 이다. 관경 매칭은 평면 mm 좌표에서
    해야 하는데 전개 결과는 m 이고 한 간선이 여러 배관으로 쪼개지므로, 이 표가
    없으면 관경이 엉뚱한 배관에 붙는다(§T1).
    """
    from services.cad_import.convert.planar import build_planar_graph

    limited = restrict_to_worst(payload, board, worst)
    built = build_planar_graph(
        key or limited.get("key") or "worst",
        write=False,
        selected_source=selected_source or limited.get("selected_source"),
        pts=limited.get("pts"),
        edges=limited.get("edges"),
        hcov=limited.get("hcov"),
        ups=limited.get("ups"),
        head_kinds=limited.get("head_kinds"),
        user_sources=limited.get("sources"),
        ho=limited.get("ho"),
    )
    if not built.get("ok") or built.get("kfp") is None:
        return {"ok": False,
                "error": built.get("error") or "제한 전개가 .kfp 를 내지 못했습니다.",
                "code": built.get("code")}

    kfp = built["kfp"]
    pipes = kfp.get("pipe_data") or {}
    edge_ref = built.get("edge_ref") or {}
    # 덮지 못한 배관은 조용히 넘기지 않는다 — 관경이 엉뚱해질 자리다(§G2 수용 기준).
    uncovered = [pid for pid in pipes if pid not in edge_ref]
    if uncovered:
        print(f"[G2] 역참조 미포함 배관 {len(uncovered)}개 — "
              f"예: {uncovered[:5]}")
    print(f"[G2] 제한 전개 · 노드 {len(kfp.get('nodes_meta_runtime') or {})} · "
          f"배관 {len(pipes)} · 역참조 {len(edge_ref)}/{len(pipes)}")
    return {
        "ok": True,
        "kfp": kfp,
        "edge_ref": edge_ref,
        "node_ref": built.get("node_ref") or {},
        "uncovered_pipes": uncovered,
        "hcov": built.get("hcov"),
        "head_kinds": built.get("head_kinds"),
        "node_head_kinds": built.get("node_head_kinds"),
        "origin_mm": built.get("origin_mm"),
        "sources": built.get("sources") or [],
    }
'''

io.open("cad_project_editor_g/services/cad_import/design/restrict.py",
        "w", encoding="utf-8", newline="\n").write(
    head + body.strip("\n") + tail)
print("restrict.py 작성")
