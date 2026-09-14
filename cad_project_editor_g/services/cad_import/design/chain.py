# -*- coding: utf-8 -*-
"""[회랑 사슬좌표 §3] 회랑 나무의 좌표를 «재지 않고 만든다».

지시서 `ModuleF_회랑_사슬좌표_지시서.md` · 그림 7·8.

■ 규칙 한 줄 (§1-1)

    시작 노드(급수원)의 좌표만 두고, 거기서부터 배관마다
    (8 방향 중 가장 가까운 방향, board 좌표에서 잰 **실제 길이**) 의 벡터를
    앞 배관의 끝점에 **이어 붙여** 모든 노드의 좌표를 정한다.

    p(자식) = p(부모) + L_e · u_e        (부모→자식 방향으로 잰 u_e)

■ 정본이 바뀐다

  종전은 «좌표» 를 지키고 길이를 좌표에서 다시 쟀다(격자 스냅 뒤의 맨해튼
  거리). 이제는 **«위상 + 길이»** 를 지키고 좌표를 사슬로 만든다 —
  마찰손실을 정하는 것은 좌표가 아니라 길이이기 때문이다.

  실측(조치 전): 손질이 고른 최원 유하거리와 표가 쓰는 길이가 대명동
  70 mm · B1F 745 mm 어긋나 있었다. 이 단계가 그 둘을 같은 수로 만든다.

■ 범위 (§1-2)

  **최불리 회랑만.** 전체망 `.kfp` 는 종전 그대로다 — 이 함수는
  `expand_worst` 안에서만 불린다. 회랑은 급수원에서 뻗은 최단경로의
  합집합이라 **나무**이고, 이 규칙은 나무에서만 정의된다.

■ 안 건드리는 것 (§5)

  `planar.py`(격자 스냅·맨해튼 길이) · `engine.py`(①②③④) · `sdf_post.py` ·
  `worst.py` · `api_edit.py`. 여기서는 **이미 만들어진 kfp 의 좌표와 평면
  배관 길이만** 다시 쓴다. 노드·간선 집합과 연결(위상)은 손대지 않는다.
"""
from __future__ import annotations

import math
from collections import defaultdict, deque

# 8 방향 — 0·45·90·135·180·225·270·315°. 같으면 작은 각(정렬 순서가 그것이다).
_DIRS8 = [(round(math.cos(math.radians(a)), 12),
           round(math.sin(math.radians(a)), 12), a)
          for a in (0, 45, 90, 135, 180, 225, 270, 315)]

VERT_TOL_M = 1e-6      # 평면에서 길이 0 이면 세로 토막
ANG_TIE = 1e-9


def _ends(pr):
    return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")


def snap_dir8(dx, dy):
    """(dx, dy) → 가장 가까운 8 방향의 단위벡터와 편차(도). 같으면 작은 각."""
    L = math.hypot(dx, dy)
    if L < 1e-12:
        return 0.0, 0.0, 0.0
    ux, uy = dx / L, dy / L
    best = None
    for cx, cy, a in _DIRS8:
        dot = max(-1.0, min(1.0, ux * cx + uy * cy))
        dev = math.degrees(math.acos(dot))
        if best is None or dev < best[2] - ANG_TIE:
            best = (cx, cy, dev, a)
    return best[0], best[1], best[2]


def _board_path_len(adj, board_pts, vi, vj, cache):
    """회랑(나무) 위에서 board 노드 vi→vj 의 **실제 길이 합**(m).

    ★일직선 중간 노드가 병합되면 kfp 배관 하나가 board 간선 **여럿**을 덮는다.
      그때 길이를 「양 끝 사이 직선 거리」로 넣으면 꺾인 만큼 짧아진다 —
      실측 B1F: 최원 경로에서 736 mm 가 사라졌다(353 m 중 0.2 %).
      관의 «실제 길이» 는 그 조각들의 합이다(§1-1 「실제 길이」).
      회랑은 나무라 두 노드 사이 길이 하나뿐이다.
    """
    key = (vi, vj) if vi <= vj else (vj, vi)
    if key in cache:
        return cache[key]
    prev = {vi: None}
    q = deque([vi])
    while q:
        cur = q.popleft()
        if cur == vj:
            break
        for nxt in adj.get(cur, ()):
            if nxt in prev:
                continue
            prev[nxt] = cur
            q.append(nxt)
    if vj not in prev:
        cache[key] = None
        return None
    tot, cur = 0.0, vj
    while prev[cur] is not None:
        par = prev[cur]
        a, b = board_pts[par], board_pts[cur]
        tot += math.hypot(float(b[0]) - float(a[0]),
                          float(b[1]) - float(a[1])) / 1000.0
        cur = par
    cache[key] = tot
    return tot


def chain_coords(kfp, *, edge_ref, node_ref, board_pts, origin_mm,
                 declared_pipes=(), root=None, corridor_edges=(),
                 chain_len_mode="euclid") -> dict:
    """회랑 kfp 의 좌표를 사슬로 다시 만든다. **제자리에서** 고친다.

    반환 `chain_report` — 화면·프로브가 읽는다. 판정(C1~C4)은 부르는 쪽이
    이 값으로 한다. 조용히 넘기는 자리가 없어야 한다(§3-2 6).
    """
    nodes = (kfp or {}).get("nodes_meta_runtime") or {}
    pipes = (kfp or {}).get("pipe_data") or {}
    if not nodes or not pipes:
        return {"ok": False, "error": "빈 망"}

    declared = {str(p) for p in (declared_pipes or ())}
    minx = float((origin_mm or (0.0, 0.0))[0])
    miny = float((origin_mm or (0.0, 0.0))[1])

    def board_frame(vid):
        """board mm → kfp 와 **같은 좌표계**(m). 스냅은 하지 않는다."""
        p = board_pts[vid]
        return ((float(p[0]) - minx) / 1000.0 + 1.0,
                (float(p[1]) - miny) / 1000.0 + 1.0)

    nref = {str(k): int(v) for k, v in (node_ref or {}).items()}
    eref = {str(k): v for k, v in (edge_ref or {}).items() if v is not None}
    # 회랑(board 간선) — 병합된 배관의 «실제 길이 합» 을 재는 데 쓴다.
    cadj = defaultdict(list)
    for a, b in (corridor_edges or ()):
        cadj[int(a)].append(int(b))
        cadj[int(b)].append(int(a))
    plen_cache: dict = {}

    # ── 뿌리 = 급수원(Input 경계). 없으면 여기서 멈춘다 — 지어내지 않는다.
    if root is None:
        from services.cad_import.design.anchor import require_anchor
        root = require_anchor(nodes, what="사슬 좌표")
    root = str(root)

    adj = defaultdict(list)
    for pid, p in pipes.items():
        a, b = _ends(p)
        if a is None or b is None or a == b:
            continue
        adj[str(a)].append((str(b), str(pid)))
        adj[str(b)].append((str(a), str(pid)))

    def xyz(nid):
        c = (nodes.get(nid) or {}).get("coords") or (0.0, 0.0, 0.0)
        return (float(c[0]), float(c[1]),
                float(c[2]) if len(c) > 2 else 0.0)

    # ── 배관 분류 · 벡터 (부모→자식은 걸으면서 정한다)
    kinds = {"plane": 0, "vert": 0, "engine": 0, "no_ref": 0, "skew": 0}

    def leg(par, chi, pid):
        """(부모→자식) 의 (dx, dy, dz, L, 편차°, 갈래)."""
        pa, pb = xyz(par), xyz(chi)
        dz = pb[2] - pa[2]
        flat = math.hypot(pb[0] - pa[0], pb[1] - pa[1])
        if flat <= VERT_TOL_M:
            return 0.0, 0.0, dz, abs(dz), 0.0, "vert"
        vi, vj = nref.get(par), nref.get(chi)
        if vi is None or vj is None:
            org = eref.get(pid)
            if org is not None:
                # 노드정리로 양 끝 중 하나가 병합된 배관 — edge_ref 의 board 쌍을
                # 쓰되 «지금 좌표» 와 가까운 쪽을 부모로 맞춘다.
                oi, oj = int(org[0]), int(org[1])
                ci, cj = board_frame(oi), board_frame(oj)
                if (math.dist(ci, pa[:2]) + math.dist(cj, pb[:2])
                        > math.dist(cj, pa[:2]) + math.dist(ci, pb[:2])):
                    oi, oj = oj, oi
                vi, vj = oi, oj
        if vi is None or vj is None:
            # board 쌍을 모르는 가로 배관 = 엔진이 만든 것(상하향식 ③).
            #   방향은 엔진이 정한 것을 8 방향으로, 길이는 그 값 그대로.
            ux, uy, dev = snap_dir8(pb[0] - pa[0], pb[1] - pa[1])
            L = float((pipes.get(pid) or {}).get("length_m") or flat)
            return ux * L, uy * L, dz, L, dev, "engine"
        bi, bj = board_frame(vi), board_frame(vj)
        # ★지시서 §1-3 그대로 — **양 끝 board 노드의 유클리드 거리**.
        #
        #   ★★여기에 «둘 다는 안 되는» 자리가 있다. 일직선 정리가 **꺾인 점**을
        #     지우면(앞 지시서 §6-5) 그 kfp 배관 하나가 덮는 board 조각들의 합이
        #     직선 거리보다 길다. 그 차가 곧 W(최원 경로 길이)의 어긋남이다 —
        #     실측: 대명동 0.5mm · 영역1 35mm · **B1F 736mm**(353m 중 0.2%).
        #
        #     조각 합을 쓰면 W 는 맞지만 좌표가 그만큼 **튀어나간다**(실측 B1F
        #     이탈 86mm → 465mm). 어느 쪽을 정본으로 할지는 오너가 정할 자리라,
        #     여기서는 지시서에 적힌 것(직선)을 쓰고 차이를 보고한다.
        #     `chain_len_mode="path"` 로 바꾸면 조각 합이 된다.
        if chain_len_mode == "path":
            Lp = _board_path_len(cadj, board_pts, vi, vj, plen_cache)
            L = math.dist(bi, bj) if Lp is None else Lp
        else:
            L = math.dist(bi, bj)
        ux, uy, dev = snap_dir8(bj[0] - bi[0], bj[1] - bi[1])
        return ux * L, uy * L, dz, L, dev, "plane"

    # ── S 에서 BFS — 나무이므로 노드마다 부모가 하나다(그림 8 ①)
    newxy = {root: xyz(root)}
    seen = {root}
    order = []
    q = deque([root])
    legs = {}
    while q:
        cur = q.popleft()
        for nxt, pid in adj.get(cur, ()):
            if nxt in seen:
                continue
            seen.add(nxt)
            dx, dy, dz, L, dev, kind = leg(cur, nxt, pid)
            kinds[kind] = kinds.get(kind, 0) + 1
            if kind == "plane" and abs(dz) > VERT_TOL_M:
                kinds["skew"] += 1
            p0 = newxy[cur]
            newxy[nxt] = (p0[0] + dx, p0[1] + dy, p0[2] + dz)
            legs[pid] = {"par": cur, "chi": nxt, "L": L, "dev": dev,
                         "kind": kind, "vec": (dx, dy, dz)}
            order.append(nxt)
            q.append(nxt)

    unreached = [n for n in nodes if n not in seen]

    # ── 좌표를 덮어쓴다 (반올림하지 않는다 — 반올림은 출력에서)
    for nid, p in newxy.items():
        m = nodes.get(nid)
        if m is None:
            continue
        m["coords"] = [p[0], p[1], p[2]]
        m["elevation_m"] = p[2]

    # ── 평면 배관의 length_m 을 board 유클리드로 덮는다 (선언 배관은 그대로)
    #
    #   ★기울어진 평면 배관이 있다. 하향식 ① 이 팔을 들어 올릴 때 **한쪽 끝만**
    #     올라간 자리가 생겨서(실측 대명동 5개 · Δz 0.3m), 그 배관은 가로로
    #     board 거리만큼 가면서 세로로도 0.3 올라간다. 그때 길이를 가로 거리로만
    #     두면 «길이 = 좌표 거리»(C2)가 217mm 어긋난다.
    #     `engine.py` 는 못 건드리므로(§5) 사슬이 그 기울기를 안고 간다 —
    #     길이는 **3차원 거리**로 둔다. 수리계산이 쓰는 값으로도 그쪽이 맞다
    #     (관은 실제로 그만큼 길다). Δz=0 인 절대다수는 board 거리 그대로다.
    n_len, n_skew = 0, 0
    for pid, lg in legs.items():
        if lg["kind"] != "plane" or pid in declared:
            continue
        pr = pipes.get(pid)
        if pr is None:
            continue
        a, b = _ends(pr)
        dz = xyz(str(b))[2] - xyz(str(a))[2]
        L3 = math.hypot(float(lg["L"]), dz)
        if abs(dz) > VERT_TOL_M:
            n_skew += 1
        pr["length_m"] = round(L3, 6)
        n_len += 1

    # ── 조각 합 − 직선 = W 가 어긋날 수 있는 양 (오너 결정용 · §6)
    gap_m = 0.0
    if cadj:
        for pid, lg in legs.items():
            if lg["kind"] != "plane":
                continue
            vi, vj = nref.get(lg["par"]), nref.get(lg["chi"])
            if vi is None or vj is None:
                continue
            Lp = _board_path_len(cadj, board_pts, vi, vj, plen_cache)
            if Lp is not None:
                gap_m += abs(Lp - float(lg["L"]))

    # ── 보고 — 회전각·이탈. 오너가 허용한 어긋남이라 «경고» 가 아니라 «참고».
    devs = [lg["dev"] for lg in legs.values() if lg["kind"] == "plane"]
    slip = []
    for nid, vid in nref.items():
        if nid not in newxy or not (0 <= vid < len(board_pts)):
            continue
        b = board_frame(vid)
        p = newxy[nid]
        slip.append((math.dist((p[0], p[1]), b) * 1000.0, nid, vid))
    slip.sort(reverse=True)
    r0 = xyz(root)
    rv = nref.get(root)
    root_snap = (math.dist((r0[0], r0[1]), board_frame(rv)) * 1000.0
                 if rv is not None and 0 <= rv < len(board_pts) else None)
    return {
        "ok": True, "root": root,
        "nodes": len(newxy), "unreached": unreached,
        "kinds": kinds, "len_overwritten": n_len, "skew": n_skew,
        "len_mode": chain_len_mode, "path_minus_line_m": gap_m,
        "dev_max": max(devs, default=0.0),
        "dev_over_10": sum(1 for d in devs if d > 10.0),
        "dev_over_22_5": sum(1 for d in devs if d > 22.5),
        "slip_max_mm": (slip[0][0] if slip else 0.0),
        "slip_top": [(round(s, 1), n, v) for s, n, v in slip[:20]],
        "root_snap_mm": root_snap,
        "legs": legs,
    }


def merge_straight_runs(kfp, *, edge_ref=None, node_ref=None,
                        declared_pipes=(), tol_deg=0.05) -> dict:
    """한 직선은 **배관 하나** — 시작점과 끝점만 남긴다 (오너 2026-09-14).

    ■ 왜 지금 할 수 있나

      사슬이 방향을 8 방향으로 맞춰 놓았으므로 한 줄 위의 토막들은 **정확히**
      같은 방향이다. 그래서 합쳐도 잃는 것이 없다 — 길이는 더하면 되고
      (L = L1 + L2), 좌표가 일직선이라 `|p(끝) − p(시작)| = L` 도 그대로다.
      스냅 시절에는 토막마다 ±50mm 씩 어긋나 이 합침이 길이를 왜곡했다.

    ■ 무엇을 남기나 — 「끝점」의 뜻

      가운데 절점을 지우는 것은 **차수 2 이고 아무 뜻도 없는** 자리뿐이다.
      아래는 절대 안 지운다(지우면 망이 뜻을 잃는다):

        · 헤드(노즐)            · 급수원·알람밸브(접속점)
        · 차수 ≠ 2 (티·크로스·관말)  · 방향이 꺾이는 자리
        · 선언 길이 배관(신축배관 접기)의 끝  · 속성이 다른 배관 사이

    ■ 무엇이 따라 바뀌나

      `edge_ref` 는 살아남은 배관이 **바깥 두 끝**의 board 쌍을 가리키도록
      고쳐 준다. 관경·부속·담당 헤드 수는 이 함수 **뒤에** 정해지므로
      (`tree_loads` · `decide_bores` · `build_fittings`) 자동으로 새 배관을
      본다 — 여기서 먼저 손대면 두 벌이 된다.
    """
    nodes = (kfp or {}).get("nodes_meta_runtime") or {}
    pipes = (kfp or {}).get("pipe_data") or {}
    if not nodes or not pipes:
        return {"ok": False, "merged": 0}
    declared = {str(p) for p in (declared_pipes or ())}
    eref = dict(edge_ref or {})
    nref = {str(k): int(v) for k, v in (node_ref or {}).items()}

    def xyz(nid):
        c = (nodes.get(nid) or {}).get("coords") or (0.0, 0.0, 0.0)
        return (float(c[0]), float(c[1]),
                float(c[2]) if len(c) > 2 else 0.0)

    def keepable(nid):
        m = nodes.get(nid) or {}
        t = str(m.get("type_id") or "").lower()
        if t in ("head", "pump", "valve"):
            return True
        if str(m.get("io_node") or "No").lower() not in ("no", ""):
            return True
        return False

    SAME = ("type", "diameter", "nominal_mm", "C", "roughness_mm", "schedule")

    def same_attrs(a, b):
        return all((a or {}).get(k) == (b or {}).get(k) for k in SAME)

    merged = 0
    changed = True
    while changed:
        changed = False
        adj: dict = {}
        for pid, pr in pipes.items():
            s, e = _ends(pr)
            if s is None or e is None or s == e:
                continue
            adj.setdefault(str(s), []).append(str(pid))
            adj.setdefault(str(e), []).append(str(pid))
        for nid, pids in adj.items():
            if len(pids) != 2 or keepable(nid):
                continue
            p1, p2 = pipes.get(pids[0]), pipes.get(pids[1])
            if p1 is None or p2 is None:
                continue
            if pids[0] in declared or pids[1] in declared:
                continue
            if not same_attrs(p1, p2):
                continue
            a1, b1 = (str(x) for x in _ends(p1))
            a2, b2 = (str(x) for x in _ends(p2))
            far1 = b1 if a1 == nid else a1
            far2 = b2 if a2 == nid else a2
            if far1 == far2:
                continue                      # 고리 — 합치지 않는다
            pa, pm, pb = xyz(far1), xyz(nid), xyz(far2)
            v1 = (pm[0] - pa[0], pm[1] - pa[1], pm[2] - pa[2])
            v2 = (pb[0] - pm[0], pb[1] - pm[1], pb[2] - pm[2])
            n1 = math.sqrt(sum(c * c for c in v1))
            n2 = math.sqrt(sum(c * c for c in v2))
            if n1 < 1e-12 or n2 < 1e-12:
                continue
            dot = sum(v1[i] * v2[i] for i in range(3)) / (n1 * n2)
            ang = math.degrees(math.acos(max(-1.0, min(1.0, dot))))
            if ang > tol_deg:
                continue                      # 꺾이는 자리 — 끝점이다
            # ── 합친다: p1 을 살리고 p2 와 가운데 절점을 지운다
            L = (float(p1.get("length_m") or 0.0)
                 + float(p2.get("length_m") or 0.0))
            p1["start"], p1["end"] = far1, far2
            p1["length_m"] = round(L, 6)
            p1["equivalent_length"] = (
                float(p1.get("equivalent_length") or 0.0)
                + float(p2.get("equivalent_length") or 0.0))
            fit = list(p1.get("fittings") or []) + list(p2.get("fittings") or [])
            if fit:
                p1["fittings"] = fit
            pipes.pop(pids[1], None)
            nodes.pop(nid, None)
            eref.pop(pids[1], None)
            if far1 in nref and far2 in nref:
                eref[pids[0]] = (nref[far1], nref[far2])
            else:
                eref.pop(pids[0], None)
            merged += 1
            changed = True
            break
    return {"ok": True, "merged": merged, "edge_ref": eref,
            "nodes": len(nodes), "pipes": len(pipes)}


def relay_from_lengths(kfp, *, root=None, length_overrides=None) -> dict:
    """길이가 바뀌면 좌표를 **다시 놓는다** — 방향은 그대로, 길이만.

    ■ 왜 (오너 2026-09-14)

        「20m 배관을 2m로 바꾸었는데, 아이소가 그대로인건 말이안되니까」

      회랑 좌표는 사슬로 만든다 — `p(자식) = p(부모) + L·u`(사슬좌표 §1-1).
      길이가 좌표를 정하므로, 덮은 길이를 넣고 사슬을 다시 걸으면 좌표가
      따라 움직이고 **아이소가 따라 변한다.** 표·`.sdf`·그림이 한꺼번에 맞는다.

    ■ 왜 `chain_coords` 를 다시 부르지 않는가

      `chain_coords` 는 **합치기 전** 배관을 board 노드쌍으로 잰다. 그런데
      카드가 주는 안정 키는 **합친 뒤**의 배관을 가리킨다(직선 합치기가 가운데
      절점을 지우므로 `edge_ref` 가 바깥 두 끝으로 바뀐다). 거기서 다시 board
      직선거리로 재면 합쳐진 구간이 짧아진다 — 길이가 조용히 바뀐다.
      그래서 여기서는 **지금 길이를 그대로 쓰고**(덮인 것만 갈아 끼우고)
      방향만 지금 좌표에서 읽어 다시 놓는다. 길이를 재지 않으므로 잃는 것이
      없고, `|p(끝)−p(시작)| = length_m`(C2)도 그대로 성립한다.

    반환: {"ok", "moved": 다시 놓은 절점 수, "applied": 갈아 끼운 배관 수,
           "unreached": 뿌리에서 못 닿은 절점}
    """
    nodes = (kfp or {}).get("nodes_meta_runtime") or {}
    pipes = (kfp or {}).get("pipe_data") or {}
    lov = {str(k): float(v) for k, v in (length_overrides or {}).items()
           if v is not None and float(v) > 0}
    if not nodes or not pipes or not lov:
        return {"ok": True, "moved": 0, "applied": 0, "unreached": []}

    if root is None:
        from services.cad_import.design.anchor import require_anchor
        root = require_anchor(nodes, what="길이 덮기")
    root = str(root)

    def xyz(nid):
        c = (nodes.get(str(nid)) or {}).get("coords") or (0.0, 0.0, 0.0)
        return (float(c[0]), float(c[1]),
                float(c[2]) if len(c) > 2 else 0.0)

    adj = defaultdict(list)
    for pid, p in pipes.items():
        a, b = _ends(p)
        if a is None or b is None or a == b:
            continue
        adj[str(a)].append((str(b), str(pid)))
        adj[str(b)].append((str(a), str(pid)))

    old = {nid: xyz(nid) for nid in nodes}
    new = {root: old[root]}
    seen = {root}
    q = deque([root])
    applied = 0
    while q:
        cur = q.popleft()
        for nxt, pid in adj.get(cur, ()):
            if nxt in seen:
                continue
            seen.add(nxt)
            pa, pb = old[cur], old[nxt]
            d = math.dist(pa, pb)
            pr = pipes.get(pid) or {}
            L = lov.get(pid)
            if L is None:
                L = float(pr.get("length_m") or d)
            else:
                pr["length_m"] = round(float(L), 6)
                applied += 1
            if d < 1e-12:
                u = (0.0, 0.0, 0.0)
            else:
                u = ((pb[0] - pa[0]) / d, (pb[1] - pa[1]) / d,
                     (pb[2] - pa[2]) / d)
            p0 = new[cur]
            new[nxt] = (p0[0] + u[0] * L, p0[1] + u[1] * L, p0[2] + u[2] * L)
            q.append(nxt)

    moved = 0
    for nid, p in new.items():
        m = nodes.get(nid)
        if m is None:
            continue
        if math.dist(p, old[nid]) > 1e-9:
            moved += 1
        m["coords"] = [p[0], p[1], p[2]]
        m["elevation_m"] = p[2]
    return {"ok": True, "moved": moved, "applied": applied,
            "unreached": [n for n in nodes if n not in seen]}
