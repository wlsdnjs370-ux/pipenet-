# -*- coding: utf-8 -*-
"""[G2] corridor 제한 payload 와 «두 번째 전개».

두 번 전개 원칙(지시서 §0.3)::

    EditBoard ─┬─ 전체망 전개 ─────────────→ .kfp   (기존 경로 · 손대지 않는다)
               └─ 최불리 제한 → 제한 전개 → 5표 → .sdf   (이쪽)

같은 손질 결과·같은 치수 입력에서 두 산출이 나오므로 설계 내용이 어긋나지 않는다.
"""
from __future__ import annotations

import math


def restrict_to_worst(payload: dict, board, worst: dict) -> dict:
    """변환 대상을 최불리 K 헤드로 좁힌다 — 헤드만 지우고 배관은 안 자른다.

    간선을 직접 잘라내고 싶은 유혹이 있지만 그러면 안 된다. 모듈 E 의
    `build_planar_graph` 는 이미 «급수원에서 물 닿는 간선만 남기고, 헤드로
    가지 않는 막다른관을 쳐내는» 단계를 갖고 있다(실측 로그: 물길 필터 →
    막다른관 삭제). 그러니 남길 헤드만 남겨 두면 그 배관은 E 가 제 규칙으로
    정리한다. 손으로 자르면 E 가 지키는 불변식(티 겹침·노드정리)을 깬다.

    hcov / disk_kinds / head_kinds 는 같은 디스크 집합을 가리키므로 함께 건다.
    ups 는 좌표 집합으로만 쓰여 남아 있어도 해가 없다.
    """
    from services.cad_import.kinds import disk_key

    keep_idx = {int(i) for i in (worst or {}).get("heads") or ()}
    disks = list(board.disks)
    kept = [disks[i] for i in sorted(keep_idx) if 0 <= i < len(disks)]
    if not kept:
        return payload

    keys = {disk_key(d[0], d[1], d[2]) for d in kept}
    out = dict(payload)
    out["hcov"] = [list(d) for d in kept]
    dk = payload.get("disk_kinds") or []
    out["disk_kinds"] = [dk[i] for i in sorted(keep_idx) if 0 <= i < len(dk)]

    fresh = []
    for rec in payload.get("head_kinds") or ():
        if not isinstance(rec, dict) or "c" not in rec:
            continue
        c = rec["c"]
        r = rec.get("head_r")
        if r is None and rec.get("tri_side"):
            r = float(rec["tri_side"]) / math.sqrt(3.0)
        if r is None:
            continue
        if disk_key(c[0], c[1], r) in keys:
            fresh.append(dict(rec))
    out["head_kinds"] = fresh
    print(f"[G2] 최불리 {len(kept)} 헤드로 범위를 좁힘 "
          f"(도면 헤드 {len(disks)} · 종류표 {len(fresh)}행)")
    return out


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
