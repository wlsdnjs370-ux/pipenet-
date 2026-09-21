# -*- coding: utf-8 -*-
"""호(원호) 기호로 읽는 «우회 구간» — 덱 「z축 및 피팅류」 3장 그림 1 (2026-09-20 오너 확정).

도면은 위에서 본 그림이라 높이가 없다. 배관이 잠깐 올라갔다 내려오는 자리는 설계자가
배관 위에 호(⊂ … ⊃, 또는 C 꼴) 두 개를 마주 보게 그려 표시한다. 세 기호(괄호 쌍·사발·
마주 보는 호)는 전부 «하나의 호» 로 정의하고, 호가 앉은 자리의 위상으로 가른다:

  · 접속 3개 이상인 노드의 호  → 갈래(통과 T · 관말 E). walk_main 이 «열린 쪽» 을 가지로
    떼고 _apply_vertical 이 branch_rise 만큼 올린다(종전 규칙 · 여기서 다루지 않음).
  · 접속 2개인 노드의 호      → 종전에는 「메인 위 기호 — 통과」로 버렸다. 이제 같은 배관
    위에서 짝을 찾는다: 호의 «열린 쪽» 이 배관을 따라 안쪽을 가리키고, 그쪽으로 걸어가
    만나는 다음 호의 열린 쪽이 되돌아 가리키면 그 사이가 우회 구간이다.

우회 구간은 양 끝에서 rise(가지 상승값과 같은 값 · 오너 «0.5 는 세 경우 모두 같다») 만큼
오르내린다: 끝 노드 바로 위에 새 노드를 하나씩 세우고 세로관을 넣고, 사이 노드는 전부
올린다. 꺾이는 자리 넷은 부속 판정이 2접속·90° 로 읽어 엘보 4개가 된다(오너 확정). 길이는
도면 길이 + rise×2. 배관을 꺾는 자리는 호의 중심점(그 호가 앉은 노드)이다.

짝을 못 찾은 호는 손대지 않고 세어 돌려준다 — 지어내지 않는다.
"""
from __future__ import annotations

import math

from services.cad_import.convert.main_walk import is_open

MAX_STEPS = 64          # 짝을 찾아 걷는 최대 노드 수
MAX_RUN_M = 60.0        # 짝을 찾아 걷는 최대 길이 (m)


def _ang(a, b):
    return math.degrees(math.atan2(b[1] - a[1], b[0] - a[0]))


def _is_flat(nodes, pid, pipes):
    p = pipes[pid]
    za = float(nodes[p["start"]]["coords"][2])
    zb = float(nodes[p["end"]]["coords"][2])
    return abs(za - zb) <= 1e-9


def apply_arc_jogs(kfp, node_arcs, rise, *, protected=(), exclude=(), degree=None,
                   next_node_id, next_pipe_id, clone_base, make_vert_pipe):
    """우회 구간을 kfp 에 제자리에서 넣는다. 반환: 판독 보고(dict).

    kfp        : nodes_meta_runtime / pipe_data 를 가진 dict (m 단위)
    node_arcs  : {node_id: [호…]} — sit_arcs 결과 (sa·sweep 있는 호만 쓴다)
    rise       : 오르내림 높이 (m). 0 이면 아무것도 안 한다
    protected  : 건드리지 않을 노드(펌프·밸브·표고 있는 노드와 그 이웃)
    exclude    : 갈래 처리에 쓰인 노드(주배관 접속점·세로 꼭대기) — 우회 후보에서 뺀다
    degree     : {노드: 지우기 전 접속 수} — 회랑이 갈래를 잘라 2개로 보여도 건물에서 갈래인
                 노드는 우회 끝이 아니다. 없으면 지금 접속 수만 본다
    나머지 넷  : _apply_vertical 의 노드/배관 생성기 (같은 규약으로 만든다)
    """
    nodes = kfp["nodes_meta_runtime"]
    pipes = kfp["pipe_data"]
    report = {"pairs": [], "unpaired": [], "skipped": {}, "n_vert": 0, "n_raised": 0}
    if not node_arcs or abs(float(rise)) <= 1e-9:
        return report
    rise = float(rise)

    def skip(why):
        report["skipped"][why] = report["skipped"].get(why, 0) + 1

    adj: dict = {}
    for pid, p in pipes.items():
        adj.setdefault(p["start"], []).append(pid)
        adj.setdefault(p["end"], []).append(pid)

    def xy(nid):
        c = nodes[nid]["coords"]
        return (float(c[0]), float(c[1]))

    def other(pid, nid):
        p = pipes[pid]
        return p["end"] if p["start"] == nid else p["start"]

    protected = set(protected)
    exclude = set(exclude)

    def inward(nid):
        """접속 2개 노드의 호가 가리키는 «안쪽» 배관. (pid, 이유) — pid None 이면 후보 아님."""
        arcs = [a for a in node_arcs.get(nid, ()) if a.get("sa") is not None and a.get("sweep") is not None]
        if not arcs:
            return None, "각 없는 호"
        links = adj.get(nid, [])
        if len(links) != 2 or int((degree or {}).get(nid, 2)) > 2:
            return None, "접속 2개 아님"
        if any(not _is_flat(nodes, pid, pipes) for pid in links):
            return None, "세로관 붙은 노드"
        if nodes[nid].get("type_id") == "head":
            return None, "헤드 노드"
        opens = []
        here = xy(nid)
        for pid in links:
            o = other(pid, nid)
            if is_open(arcs, _ang(here, xy(o))):
                opens.append(pid)
        if len(opens) != 1:
            return None, ("열린 쪽 없음" if not opens else "양쪽 열림")
        return opens[0], None

    cand = {}
    for nid in list(node_arcs):
        if nid in protected or nid in exclude or nid not in nodes:
            continue
        pid, why = inward(nid)
        if pid is None:
            if why not in ("접속 2개 아님",):     # 갈래 노드의 호는 이 규칙의 대상이 아니다
                skip(why)
            continue
        cand[nid] = pid

    used = set()
    for nid in sorted(cand, key=str):
        if nid in used:
            continue
        first = cand[nid]
        prev, cur = nid, other(first, nid)
        path = []                      # nid 와 짝 사이의 노드
        walked = float(pipes[first].get("length_m") or 0.0)
        partner = None
        steps = 0
        arrive_pid = first
        while steps < MAX_STEPS and walked <= MAX_RUN_M:
            steps += 1
            if cur in protected or nodes[cur].get("type_id") == "head":
                break
            if cur in cand and cur not in used:
                if cand[cur] == arrive_pid:      # 되돌아 가리킨다 — 짝
                    partner = cur
                break
            links = adj.get(cur, [])
            if len(links) != 2:
                break
            nxt_pid = links[0] if links[1] == arrive_pid else links[1]
            if nxt_pid == arrive_pid:
                break
            if not _is_flat(nodes, nxt_pid, pipes):
                break
            path.append(cur)
            walked += float(pipes[nxt_pid].get("length_m") or 0.0)
            prev, cur, arrive_pid = cur, other(nxt_pid, cur), nxt_pid
        if partner is None:
            report["unpaired"].append(str(nid))
            continue
        # 양 끝에 세로관 · 사이는 올린다
        for end_nid, end_pid in ((nid, first), (partner, cand[partner])):
            c = nodes[end_nid]["coords"]
            top = next_node_id()
            nodes[top] = clone_base(end_nid if nodes[end_nid].get("type_id") == "base" else next(
                (k for k, n in nodes.items() if n.get("type_id") == "base"), end_nid),
                top, [c[0], c[1], float(c[2]) + rise])
            pipes[next_pipe_id()] = make_vert_pipe(end_nid, top, pipes[end_pid], rise)
            p = pipes[end_pid]
            if p["start"] == end_nid:
                p["start"] = top
            elif p["end"] == end_nid:
                p["end"] = top
            report["n_vert"] += 1
        for mid in path:
            c = list(nodes[mid]["coords"])
            c[2] = float(c[2]) + rise
            nodes[mid]["coords"] = c
            nodes[mid]["elevation_m"] = float(c[2])
        report["n_raised"] += len(path)
        report["pairs"].append({"a": str(nid), "b": str(partner), "between": len(path)})
        used.add(nid)
        used.add(partner)
    return report
