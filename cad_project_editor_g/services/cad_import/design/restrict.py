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


def attachable_heads(payload: dict, *, selected_source=None,
                     key: str | None = None) -> dict:
    """전개가 **배관에 붙일 수 있는** 헤드 번호. 전체망 전개를 한 번 돌려 얻는다.

    선정은 board 그래프 도달로 헤드를 세지만, 전개는 «헤드 중심 노드가 물길
    필터 안» 이라야 인정한다. 후자가 더 엄격해 B1F 실측으로 868 대 619 다.
    이 차이를 모른 채 최불리를 고르면 30개가 통째로 전개 밖으로 떨어진다.

    반환: {"ok", "wet": set(hcov 번호), "total": 도면 헤드 수, "dropped": 제외 수}
    """
    from services.cad_import.convert.planar import build_planar_graph

    built = build_planar_graph(
        key or payload.get("key") or "probe",
        write=False,
        selected_source=selected_source or payload.get("selected_source"),
        pts=payload.get("pts"),
        edges=payload.get("edges"),
        hcov=payload.get("hcov"),
        ups=payload.get("ups"),
        head_kinds=payload.get("head_kinds"),
        user_sources=payload.get("sources"),
        ho=payload.get("ho"),
    )
    if not built.get("ok"):
        return {"ok": False,
                "error": built.get("error") or "전개가 실패했습니다.",
                "wet": set(), "total": 0, "dropped": 0}
    wet = set(built.get("wet_head_idx") or [])
    total = len(payload.get("hcov") or [])
    return {"ok": True, "wet": wet, "total": total,
            "dropped": max(0, total - len(wet))}


def _ends(pr):
    """배관 행의 양 끝. 전개 결과는 start/end, 표는 from/to 를 쓴다."""
    return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")


def tree_loads(kfp) -> dict:
    """배관마다 «그 아래에 달린 헤드 수». 급수원을 뿌리로 한 트리에서 센다.

    ★board 간선 역참조(`edge_ref`)로는 셀 수 없는 배관이 있다 — 헤드 접속관과
      가지 상승관은 도면에 그려진 선이 아니라 변환이 만든 세로 구간이라 대응할
      board 간선이 없다. 그런 배관의 담당 헤드 수를 0 으로 두면 별표1 이 전부
      25A 를 주고, 가지 상승관이 제 아래 헤드 스무 개를 25A 로 받게 된다.
    """
    nodes = (kfp or {}).get("nodes_meta_runtime") or {}
    pipes = (kfp or {}).get("pipe_data") or {}
    adj: dict = {}
    for pid, pr in pipes.items():
        a, b = _ends(pr)
        if a is None or b is None:
            continue
        adj.setdefault(a, []).append((b, pid))
        adj.setdefault(b, []).append((a, pid))
    root = next((n for n, m in nodes.items()
                 if str((m or {}).get("type_id", "")) == "pump"), None)
    if root is None:
        root = next(iter(nodes), None)
    if root is None:
        return {}
    heads = {n for n, m in nodes.items()
             if str((m or {}).get("type_id", "")) == "head"}

    # 뿌리에서 뻗는 순서를 먼저 잡고, 거꾸로 접으면서 헤드를 센다.
    order, parent_pipe, parent = [], {}, {root: None}
    stack = [root]
    seen = {root}
    while stack:
        cur = stack.pop()
        order.append(cur)
        for nxt, pid in adj.get(cur, ()):
            if nxt in seen:
                continue
            seen.add(nxt)
            parent[nxt] = cur
            parent_pipe[nxt] = pid
            stack.append(nxt)
    count = {n: (1 if n in heads else 0) for n in seen}
    out: dict = {}
    for n in reversed(order):
        p = parent.get(n)
        if p is None:
            continue
        count[p] = count.get(p, 0) + count.get(n, 0)
        out[parent_pipe[n]] = count[n]
    return out


def apply_vertical(payload, built, *, convert_kwargs=None):
    """제한 전개 평면망에 **정상 변환과 같은** 세로 처리를 얹는다.

    ★이걸 안 하면 설계 SDF 는 «평면 그래프» 그대로 나간다 — 헤드 접속관(①②③④)도
      가지 상승도 없고 **모든 표고가 0** 이다. 그러면 헤드가 가지배관 위의 통과점이
      되어 화면에서 세울 스텁이 없고, 무엇보다 정수두 차가 통째로 빠진 채
      수리계산이 나간다. `.kfp` 는 이 처리를 거치는데 `.sdf` 만 안 거쳤다.

    변환 창이 받은 값(`convert_kwargs`)을 그대로 쓴다 — 사람이 고른 헤드 접속관
    길이가 `.kfp` 와 `.sdf` 에서 달라지면 두 산출물이 다른 망이 된다.
    """
    from services.cad_import.convert.engine import convert_to_kfp
    from services.cad_import.dto import default_dto, dto_to_convert_kwargs

    kw = dict(convert_kwargs or dto_to_convert_kwargs(default_dto()))
    sub = dict(payload or {})
    sub["kfp"] = built["kfp"]
    sub.pop("kfp_path", None)
    for k in ("hcov", "head_kinds", "node_head_kinds", "origin_mm"):
        if built.get(k) is not None:
            sub[k] = built[k]
    if built.get("sources"):
        sub["sources"] = built["sources"]
    r = convert_to_kfp(sub, None, **kw)
    if not r.get("ok"):
        codes = [b.get("code") for b in (r.get("blockers") or [])]
        return None, f"세로 처리 실패: {codes}"
    return r["kfp"], None


def expand_worst(payload: dict, board, worst: dict, *,
                 selected_source=None, key: str | None = None,
                 convert_kwargs=None, vertical: bool = True) -> dict:
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

    # ★평면 그래프 «그대로» 내면 헤드 접속관도 가지 상승도 없고 표고가 전부 0 이다.
    #   `.kfp` 는 이 처리를 거치는데 `.sdf` 만 안 거쳐서, 같은 도면인데 두 산출물이
    #   다른 망이었다. 여기서 같은 처리를 얹는다.
    flat = built["kfp"]
    loads_by_pipe = {}
    if vertical:
        raised, verr = apply_vertical(payload, built,
                                      convert_kwargs=convert_kwargs)
        if raised is None:
            return {"ok": False, "error": verr}
        n0 = len(flat.get("nodes_meta_runtime") or {})
        p0 = len(flat.get("pipe_data") or {})
        built["kfp"] = raised
        loads_by_pipe = tree_loads(raised)
        print(f"[G19] 세로 처리 · 노드 {n0} → "
              f"{len(raised.get('nodes_meta_runtime') or {})} · 배관 {p0} → "
              f"{len(raised.get('pipe_data') or {})} "
              f"(헤드 접속관·가지 상승이 여기서 생긴다)")

    kfp = built["kfp"]
    pipes = kfp.get("pipe_data") or {}
    edge_ref = built.get("edge_ref") or {}
    # 덮지 못한 배관은 조용히 넘기지 않는다 — 관경이 엉뚱해질 자리다(§G2 수용 기준).
    uncovered = [pid for pid in pipes if pid not in edge_ref]
    if uncovered and not vertical:
        # 세로 처리를 안 켰는데 덮이지 않는 배관이 있으면 그건 이상한 일이다.
        print(f"[G2] 역참조 미포함 배관 {len(uncovered)}개 — "
              f"예: {uncovered[:5]}")
    elif uncovered:
        # 세로 구간은 도면에 그려진 선이 아니라 대응할 board 간선이 없다.
        # 이상한 것이 아니라 «당연히 없는» 것이므로 그렇게 말한다.
        print(f"[G2] 세로 구간 {len(uncovered)}개는 도면 선이 아니라 역참조가 "
              f"없습니다 — 관경은 담당 헤드 수로 정합니다.")
    print(f"[G2] 제한 전개 · 노드 {len(kfp.get('nodes_meta_runtime') or {})} · "
          f"배관 {len(pipes)} · 역참조 {len(edge_ref)}/{len(pipes)}")
    return {
        "ok": True,
        "kfp": kfp,
        "edge_ref": edge_ref,
        "node_ref": built.get("node_ref") or {},
        "uncovered_pipes": uncovered,
        "tree_loads": loads_by_pipe,
        "hcov": built.get("hcov"),
        "head_kinds": built.get("head_kinds"),
        "node_head_kinds": built.get("node_head_kinds"),
        "origin_mm": built.get("origin_mm"),
        "sources": built.get("sources") or [],
    }


def select_and_expand(payload: dict, board, *, k=None, only_heads=None,
                      selected_source=None, key: str | None = None,
                      convert_kwargs=None, vertical: bool = True) -> dict:
    """최불리 선정 → corridor 제한 전개를 한 번에. G7 창이 부르는 진입점.

    ★선정 후보를 **전개가 붙일 수 있는 헤드**로 먼저 좁힌다(BLOCKED B4 · 1안).
    선정은 board 도달로 헤드를 세고 전개는 더 엄격한 규칙을 쓰므로, 좁히지
    않으면 최불리 K 가 통째로 전개 밖으로 떨어져 빈 망이 나온다.

    `only_heads`(도면 장 제한)가 함께 오면 **교집합**을 쓴다 — 장 제한이 먼저고
    그 안에서 붙일 수 있는 것만 남긴다.

    반환의 `excluded_heads` 는 «붙지 못해 후보에서 뺀 헤드 수» 다. 이 값이 크면
    그 도면의 배관이 끊겨 있다는 뜻이므로 **화면이 반드시 보여야 한다** —
    조용히 빼면 더 불리한 헤드를 못 본 채 수리계산이 나간다.
    """
    from services.cad_import.design.worst import REMOTE_K_DEFAULT, worst_k_heads

    k = REMOTE_K_DEFAULT if k is None else k
    probe = attachable_heads(payload, selected_source=selected_source, key=key)
    if not probe["ok"]:
        return {"ok": False, "error": probe.get("error")}

    wet = probe["wet"]
    cand = wet if only_heads is None else (set(only_heads) & wet)
    if not cand:
        return {"ok": False,
                "error": ("전개가 배관에 붙일 수 있는 헤드가 없습니다. "
                          "손질 단계에서 배관을 먼저 이어 주세요.")}

    b = board
    worst = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources,
                          k=k, only_heads=cand)
    if not worst.get("heads"):
        return {"ok": False, "error": "급수원에서 닿는 헤드가 없습니다."}

    got = expand_worst(payload, b, worst,
                       selected_source=selected_source, key=key,
                       convert_kwargs=convert_kwargs, vertical=vertical)
    if not got.get("ok"):
        return got

    got["worst"] = worst
    got["excluded_heads"] = probe["dropped"]
    got["candidate_heads"] = len(cand)
    got["total_heads"] = probe["total"]
    if probe["dropped"]:
        print(f"[G2] 전개가 붙이지 못한 헤드 {probe['dropped']}개는 후보에서 "
              f"제외했습니다 — 그만큼 배관이 끊겨 있을 수 있습니다 "
              f"(후보 {len(cand)} / 도면 {probe['total']}).")
    return got
