# -*- coding: utf-8 -*-
"""손질 망의 셈 — 덩이 통계와 자동 이음.

«A 처럼 재고, E 처럼 붙인다». 후보를 고르는 것만 여기서 하고, 실제로 붙일지는
모듈 E 의 `board.join` 이 모양을 보고 정한다.
"""
from __future__ import annotations

import math

from routes.module_f.common import (
    AUTOJOIN_ANG_TOL_DEG, AUTOJOIN_LADDER_MM, AUTOJOIN_MAX_PAIRS,
    AUTOJOIN_PLATEAU, _r1)

# ────────────────────────────────────────────── 자동 이음 · 도면 장 · 덩이
def _body_index(board):
    """노드 → 덩이 번호, 덩이별 노드 수."""
    bodies = board.bodies()
    body_of: dict[int, int] = {}
    for bi, nodes in enumerate(bodies):
        for n in nodes:
            body_of[n] = bi
    return body_of, [len(nodes) for nodes in bodies]


def _body_stat(board) -> dict:
    """덩이 수와 «급수원이 닿는 헤드». 물흐름을 돌리기 전에 알려주려는 것이다.

    B1F 실측처럼 헤드 3,163개 중 264개만 물이 닿는 도면에서, 그 사실을 물흐름을
    돌린 뒤에야 아는 것은 너무 늦다. 배관이 몇 조각인지, 급수원이 그중 어느
    조각에 있는지는 지금 바로 셀 수 있다.
    """
    body_of, sizes = _body_index(board)
    n_bodies = len(sizes)
    heads_in = [0] * max(1, n_bodies)
    for nodes in board.hnodes:
        bs = sorted(body_of[n] for n in nodes if n in body_of)
        if bs:
            heads_in[bs[0]] += 1
    src_bodies = {body_of[n] for n in board.sources if n in body_of}
    return {
        "bodies": n_bodies,
        "total_heads": len(board.disks),
        "biggest_heads": max(heads_in) if n_bodies else 0,
        "source_heads": sum(heads_in[bi] for bi in src_bodies),
        "has_source": bool(board.sources),
    }


def _autojoin_scan(board, force_eps: float | None = None) -> dict:
    """끊긴 관 끝을 짝짓고, 이을 여유(eps)를 도면에서 잰다. 아직 붙이지 않는다.

    목적함수는 모듈 A 와 같다 — **가장 큰 덩이의 노드 수**. 여유가 모자라면 망이
    조각난 채로 남고, 지나치면 남의 배관을 삼킨다. 그래서 사다리를 훑어 정점을
    고르고, 같은 값이면 **작은 쪽**을 택한다(A 는 최대치를 택하지만, F 는 붙이는
    판정을 E 가 다시 하므로 후보를 넓히는 쪽보다 좁히는 쪽이 안전하다).

    짝은 «차수 1인 관 끝» 끼리, 그리고 **서로 다른 덩이**끼리만 만든다. 같은
    덩이를 다시 이어봐야 조각 수가 줄지 않는다.
    """
    pts, edges = board.pts, board.edges
    body_of, sizes = _body_index(board)

    deg: dict[int, int] = {}
    nb: dict[int, int] = {}
    for a, b in edges:
        deg[a] = deg.get(a, 0) + 1
        deg[b] = deg.get(b, 0) + 1
        nb.setdefault(a, b)
        nb.setdefault(b, a)
    ends = [n for n, d in deg.items() if d == 1]

    cap = AUTOJOIN_LADDER_MM[-1]
    grid: dict[tuple[int, int], list[int]] = {}
    for n in ends:
        x, y = pts[n]
        grid.setdefault((int(x // cap), int(y // cap)), []).append(n)

    def stub_dir(n: int) -> tuple[float, float] | None:
        """관 끝에서 «바깥» 을 가리키는 단위벡터 — 그 관이 가던 방향."""
        o = nb.get(n)
        if o is None:
            return None
        ux, uy = pts[n][0] - pts[o][0], pts[n][1] - pts[o][1]
        h = math.hypot(ux, uy)
        return None if h < 1e-9 else (ux / h, uy / h)

    sin_tol = math.sin(math.radians(AUTOJOIN_ANG_TOL_DEG))

    def plausible(a: int, b: int) -> bool:
        """틈이 «관이 이어지던 자리» 로 보이나 — E 의 이음 규칙을 싸게 흉내낸다.

        E 는 일직선·T자·직각만 붙인다. 셋 다 다리(bridge)가 두 관 중 한쪽 축과
        나란하다. 그래서 «다리가 어느 한쪽 관이 가던 방향과 나란하고 그 바깥을
        향한다» 를 통과 조건으로 쓴다. 이 걸름이 없으면 여유를 넓힐수록 아무
        상관없는 관 끝끼리 짝지어져 사다리가 늘 상한으로 밀린다.
        """
        bx, by = pts[b][0] - pts[a][0], pts[b][1] - pts[a][1]
        h = math.hypot(bx, by)
        if h < 1e-9:
            return True
        bx, by = bx / h, by / h
        for n, sign in ((a, 1.0), (b, -1.0)):
            u = stub_dir(n)
            if u is None:
                continue
            cx, cy = bx * sign, by * sign
            if abs(u[0] * cy - u[1] * cx) <= sin_tol and (u[0] * cx + u[1] * cy) > 0:
                return True
        return False

    raw: list[dict] = []
    n_near = 0

    # ── ① 끝 ↔ 끝. 부속 기호가 관 «가운데» 를 끊어놓은 자리.
    for a in ends:
        ax, ay = pts[a]
        gx, gy = int(ax // cap), int(ay // cap)
        ba = body_of.get(a)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for b in grid.get((gx + dx, gy + dy), ()):
                    if b <= a or body_of.get(b) == ba:
                        continue
                    d = math.hypot(pts[b][0] - ax, pts[b][1] - ay)
                    if d > cap:
                        continue
                    n_near += 1
                    if plausible(a, b):
                        raw.append({"d": d, "kind": "끝↔끝", "ba": ba,
                                    "bb": body_of[b],
                                    "seg_a": [list(pts[a]), list(pts[nb[a]])],
                                    "seg_b": [list(pts[b]), list(pts[nb[b]])],
                                    "line": [pts[a][0], pts[a][1],
                                             pts[b][0], pts[b][1]]})

    # ── ② 끝 ↔ 다른 덩이의 관 «몸통»(T분기).
    #
    # ★여기가 핵심이다. 끝점끝점만 보면 큰 조각들이 안 붙는다 — 연결복원
    # 링크예측이 v1 에서 겪은 그 함정이고(전이 0건), T분기를 1급으로 올려서야
    # 갭이 풀렸다. B1F 실측도 같다: 끝↔끝만으로는 덩이 271 → 248 에 그치는데,
    # 5,614·2,561·2,118… 짜리 큰 조각들은 서로 «끝» 이 아니라 «옆구리» 로 만난다.
    ecell: dict[tuple[int, int], set[tuple[int, int]]] = {}
    step = cap * 0.5
    for u, v in edges:
        x0, y0 = pts[u]
        x1, y1 = pts[v]
        n = int(math.hypot(x1 - x0, y1 - y0) / step) + 1
        for s in range(n + 1):
            t = s / n
            key = (int((x0 + (x1 - x0) * t) // cap),
                   int((y0 + (y1 - y0) * t) // cap))
            ecell.setdefault(key, set()).add((u, v))

    for a in ends:
        ax, ay = pts[a]
        ua = stub_dir(a)
        if ua is None:
            continue
        ba = body_of.get(a)
        gx, gy = int(ax // cap), int(ay // cap)
        seen: set[tuple[int, int]] = set()
        best_by_body: dict[int, tuple[float, tuple[int, int], tuple[float, float]]] = {}
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for e in ecell.get((gx + dx, gy + dy), ()):
                    if e in seen:
                        continue
                    seen.add(e)
                    u, v = e
                    bb = body_of.get(u)
                    if bb == ba:
                        continue
                    x0, y0 = pts[u]
                    x1, y1 = pts[v]
                    ex, ey = x1 - x0, y1 - y0
                    den = ex * ex + ey * ey
                    if den < 1e-9:
                        continue
                    t = ((ax - x0) * ex + (ay - y0) * ey) / den
                    if not (1e-6 < t < 1 - 1e-6):
                        continue  # 끝점은 ①의 몫 — E 도 T자는 몸통 안일 때만
                    fx, fy = x0 + ex * t, y0 + ey * t
                    d = math.hypot(fx - ax, fy - ay)
                    if d > cap or d < 1e-9:
                        continue
                    n_near += 1
                    # 다리는 그 관이 가던 방향으로 나 있어야 하고(∥ stub),
                    # 몸통과는 직각이어야 한다 — E 의 tee_bridges 와 같은 조건.
                    cx, cy = (fx - ax) / d, (fy - ay) / d
                    if abs(ua[0] * cy - ua[1] * cx) > sin_tol:
                        continue
                    if ua[0] * cx + ua[1] * cy <= 0:
                        continue
                    eh = math.hypot(ex, ey)
                    if abs(ua[0] * ex / eh + ua[1] * ey / eh) > sin_tol:
                        continue
                    cur = best_by_body.get(bb)
                    if cur is None or d < cur[0]:
                        best_by_body[bb] = (d, e, (fx, fy))
        for bb, (d, (u, v), foot) in best_by_body.items():
            raw.append({"d": d, "kind": "T분기", "ba": ba, "bb": bb,
                        "seg_a": [list(pts[a]), list(pts[nb[a]])],
                        "seg_b": [list(pts[u]), list(pts[v])],
                        "line": [pts[a][0], pts[a][1], foot[0], foot[1]]})

    raw.sort(key=lambda c: c["d"])
    pairs = [(c["d"], c["ba"], c["bb"]) for c in raw]

    def sweep(eps: float) -> tuple[int, int, int]:
        parent = list(range(len(sizes)))
        size = list(sizes)

        def find(i: int) -> int:
            while parent[i] != i:
                parent[i] = parent[parent[i]]
                i = parent[i]
            return i

        used = 0
        for d, ba, bb in pairs:
            if d > eps:
                break
            used += 1
            ra, rb = find(ba), find(bb)
            if ra == rb:
                continue
            parent[rb] = ra
            size[ra] += size[rb]
        roots = {find(i) for i in range(len(sizes))}
        largest = max((size[r] for r in roots), default=0)
        return largest, len(roots), used

    trials = []
    for eps in AUTOJOIN_LADDER_MM:
        largest, comps, used = sweep(eps)
        trials.append({"eps_mm": eps, "largest": largest,
                       "bodies": comps, "pairs": used})
    # 걸름을 통과한 짝만 세므로 곡선은 «늘다가 멎는다». 정점 그 자체가 아니라
    # 정점에 사실상 닿는 **첫 여유**를 고른다 — 같은 결과면 좁은 쪽이 안전하다.
    peak = max((t["largest"] for t in trials), default=0)
    chosen = AUTOJOIN_LADDER_MM[0]
    for t in trials:
        if peak and t["largest"] >= peak * AUTOJOIN_PLATEAU:
            chosen = t["eps_mm"]
            break
    auto = chosen
    # 사다리를 화면에 그대로 펴 두고 사람이 다른 칸을 고를 수 있게 한다 —
    # 마지막 판단은 사람 몫이라는 것이 모듈 E 의 방식이다.
    if force_eps is not None and float(force_eps) in AUTOJOIN_LADDER_MM:
        chosen = float(force_eps)

    at_eps = [c for c in raw if c["d"] <= chosen]
    dropped = max(0, len(at_eps) - AUTOJOIN_MAX_PAIRS)
    cands = []
    for c in at_eps[:AUTOJOIN_MAX_PAIRS]:
        cands.append({
            "d": round(c["d"], 1), "kind": c["kind"],
            "ba": c["ba"], "bb": c["bb"],
            "seg_a": c["seg_a"], "seg_b": c["seg_b"],
            "line": [_r1(v) for v in c["line"]],
        })
    by_kind: dict[str, int] = {}
    for c in cands:
        by_kind[c["kind"]] = by_kind.get(c["kind"], 0) + 1
    return {"eps_mm": chosen, "auto_eps_mm": auto, "ladder": AUTOJOIN_LADDER_MM,
            "trials": trials, "cands": cands,
            "ends": len(ends), "dropped": dropped, "near": n_near,
            "kept": len(pairs), "by_kind": by_kind,
            "bodies_before": len(sizes)}


def _autojoin_apply(board, scan: dict) -> dict:
    """후보를 E 의 이음 판정에 하나씩 태운다. 붙일지 말지는 E 가 정한다.

    ★되돌리기는 한 번에 되돌아가야 한다. `board.join` 은 부를 때마다 스냅샷을
    쌓는데(스냅샷 하나가 노드 2만 + 헤드 종류표 3천 행), 수백 번이면 메모리도
    UI 도 감당이 안 된다. 그래서 묶음이 끝난 뒤 **첫 스냅샷 하나만 남기고
    잘라낸다** — 그 하나가 곧 «묶음 시작 직전 상태» 이므로 되돌리기 한 번이
    자동 이음 전체를 정확히 되돌린다.
    """
    cands = scan.get("cands") or []
    h0 = len(board.history)
    parent: dict[int, int] = {}

    def find(i: int) -> int:
        while parent.get(i, i) != i:
            parent[i] = parent.get(parent[i], parent[i])
            i = parent[i]
        return i

    made = blocked = skipped = 0
    kinds: dict[str, int] = {}
    total = len(cands)
    print(f"[자동이음] 여유 {scan.get('eps_mm')}mm 로 후보 {total}곳을 훑습니다…")
    for i, c in enumerate(cands, start=1):
        ra, rb = find(int(c["ba"])), find(int(c["bb"]))
        if ra == rb:
            skipped += 1
            continue
        seg_a = (tuple(c["seg_a"][0]), tuple(c["seg_a"][1]))
        seg_b = (tuple(c["seg_b"][0]), tuple(c["seg_b"][1]))
        try:
            n, _blk, _cov, kind = board.join(seg_a, seg_b)
        except Exception as exc:  # noqa: BLE001 — 한 곳이 튀어도 나머지는 붙인다
            print(f"[자동이음] 건너뜀({i}/{total}): {type(exc).__name__} {exc}")
            blocked += 1
            continue
        if n:
            made += n
            kinds[kind] = kinds.get(kind, 0) + 1
            parent[rb] = ra
        else:
            blocked += 1
        if i % 100 == 0 or i == total:
            print(f"[자동이음] {i}/{total} · 붙임 {made} · 막힘 {blocked}"
                  f" · 이미이어짐 {skipped}")
    if len(board.history) > h0 + 1:
        del board.history[h0 + 1:]
    after = len(board.bodies())
    print(f"[자동이음] 완료 · 붙임 {made} · 막힘 {blocked} · "
          f"덩이 {scan.get('bodies_before')} → {after} · {kinds}")
    return {"made": made, "blocked": blocked, "skipped": skipped,
            "kinds": kinds, "eps_mm": scan.get("eps_mm"),
            "bodies_before": scan.get("bodies_before"), "bodies_after": after}
