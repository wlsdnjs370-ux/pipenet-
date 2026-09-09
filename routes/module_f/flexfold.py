# -*- coding: utf-8 -*-
"""[신축배관 접기] 한 가닥(선분+호)을 직선 배관 하나로 접는다.

지시서 `ModuleF_신축배관_접기_지시서.md`.

■ 왜 «빼기» 가 아니라 «접기» 인가

  어제는 「도면의 후렉시블 선은 변환이 만드는 수직 접속관과 중복이니 빼라」고
  했다. 틀렸다 — 빼면 물닿음 헤드가 111 → 5 가 된다(실측). 도면의 후렉시블은
  가지관 → 헤드를 **평면에서 수평으로** 잇는 300~1,000 mm 구간이고, 변환이
  만드는 접속관은 헤드를 **수직으로** 세우는 0.3 m 다. 다른 구간이다.

  그 45° 꺾임은 실제 배관 형상이 맞다(유연 호스). 없애는 게 아니라 **편다.**

■ ★한 가닥은 «선분 + 호» 다 — 이것이 이 파일의 핵심 사실

  대명동 실측:

      선분  320개 · 44,554 mm      ← `Board.by_bundle` 에 들어가는 것
      호    115개 · 14,057 mm      ← **아무 데도 안 들어간다**
      ────────────────────────
      선분만으로 사슬을 이으면   200토막 (4~375 mm · 중앙값 200)
      호를 현으로 넣어 이으면     85가닥 (663~723 mm · 중앙값 665 · 합 57,271)

  85 · 663~723 · 665 · 57,271 은 지시서가 원도면을 직접 재서 낸 값과 **정확히
  같다.** 즉 「가닥당 4~9 선분」은 선분과 호를 합한 조각 수였다.

  그러므로 호를 안 보면 가닥을 못 만든다. `Board` 는 `w.segs` 만 담으므로
  여기서 `w.arcs` 를 직접 읽는다.

■ 길이 — 좌표가 아니라 «선언» 이 권위다

  접으면 좌표가 줄어든다. 그냥 두면 표 length 도 줄어든다:

      L1(맨해튼) 합 65,990 → 현의 L1 60,017   (−5,974 mm · −9.1 %)

  85가닥 중 38가닥이 되돌아 꺾여(단조롭지 않아) 그렇다. 길이가 줄면 마찰손실이
  과소 산정되고 그것은 **비보수측**이다. 그래서 접기 전 총연장을 «선언» 해
  `planar` 로 보낸다 — 나르는 길은 `_snap_origin` 이 이미 놓은 것을 쓴다.
"""
from __future__ import annotations

import math

# 자동 사전 — **추천만** 한다. 도면마다 관례가 다르므로 반드시 놓친다(S340).
FLEX_WORDS = ("후렉시블", "후렉", "플렉", "flex", "fx")
# 끝점을 뭉치는 자(mm). ★격자 반올림이 아니라 eps 클러스터다 — 4.9 와 5.1 이
#   다른 칸으로 갈라지면 한 폴리라인이 토막 난다(실측: 85가닥이 196토막).
JOIN_EPS = 0.5
# 접힌 가닥의 끝을 손질판 노드에 붙이는 자(mm). 찍기 → 손질에서 좌표가 미세하게
# 움직일 수 있어 join 보다 넉넉히 잡는다.
SNAP_EPS = 30.0


def is_flex(layer) -> bool:
    s = str(layer or "").lower()
    return any(w.lower() in s for w in FLEX_WORDS)


class _Cluster:
    """좌표를 eps 로 뭉친다 — 같은 점이면 같은 번호."""

    def __init__(self, eps=JOIN_EPS):
        self.eps = float(eps)
        self.cell = max(self.eps, 1e-9)
        self.grid: dict = {}
        self.at: list = []

    def key(self, p, add=True):
        gx, gy = int(p[0] // self.cell), int(p[1] // self.cell)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for i in self.grid.get((gx + dx, gy + dy), ()):
                    q = self.at[i]
                    if abs(q[0] - p[0]) <= self.eps \
                            and abs(q[1] - p[1]) <= self.eps:
                        return i
        if not add:
            return None
        i = len(self.at)
        self.at.append((float(p[0]), float(p[1])))
        self.grid.setdefault((gx, gy), []).append(i)
        return i


def flex_pieces(world, layers):
    """지정 레이어의 «조각» — 선분은 그대로, 호는 **현** 으로.

    호를 현으로 바꾸는 것은 지시서의 셈과 같은 규약이다(그 값이 57,271 mm).
    실제 호길이로 세면 58,611 mm (+2.3 %) 로 더 보수측이 되지만, 수용 기준이
    57,271 이므로 여기서는 현을 쓴다 — 바꾸려면 이 함수 한 곳만 고치면 된다.
    """
    want = {str(x) for x in (layers or ())}
    out = []
    for ly, _col, a, b in (getattr(world, "segs", None) or ()):
        if str(ly) in want:
            out.append(((float(a[0]), float(a[1])),
                        (float(b[0]), float(b[1]))))
    angs = list(getattr(world, "arc_ang", None) or ())
    for i, (ly, _col, x, y, r) in enumerate(getattr(world, "arcs", None) or ()):
        if str(ly) not in want:
            continue
        ang = angs[i] if i < len(angs) else None
        if not ang:
            continue
        sa, sw = float(ang[0]), float(ang[1])
        out.append(((x + r * math.cos(math.radians(sa)),
                     y + r * math.sin(math.radians(sa))),
                    (x + r * math.cos(math.radians(sa + sw)),
                     y + r * math.sin(math.radians(sa + sw)))))
    return out


def strands(pieces, eps=JOIN_EPS):
    """조각 → 가닥. 차수 1 끝점 둘을 가진 사슬만 «가닥» 이다.

    폴리라인 정체성은 `stage1` 이 쪼갠 뒤라 이미 없다(지시서 §2). 끝점으로 이어
    붙이는 이 방식이 오히려 낫다 — 도면이 후렉시블을 여러 LINE 으로 그렸어도
    같은 결과가 나온다.

    돌려주는 것: (가닥 목록, 사슬이 아닌 것, 어디에도 안 든 조각 번호)
    """
    cl = _Cluster(eps)
    adj: dict = {}
    at: dict = {}
    for i, (a, b) in enumerate(pieces):
        ka, kb = cl.key(a), cl.key(b)
        if ka == kb:
            continue
        at.setdefault(ka, a)
        at.setdefault(kb, b)
        adj.setdefault(ka, []).append((kb, i))
        adj.setdefault(kb, []).append((ka, i))

    seen = set()
    chains, tangled = [], []
    for start, links in adj.items():
        if len(links) != 1:
            continue
        for _nxt, si in links:
            if si in seen:
                continue
            path, cur, prev = [start], start, None
            ok = True
            while True:
                nxts = [(n, s) for n, s in adj[cur] if s != prev]
                if not nxts:
                    break
                if len(nxts) > 1:          # 갈래 — 사슬이 아니다
                    ok = False
                    break
                n, s = nxts[0]
                seen.add(s)
                path.append(n)
                prev, cur = s, n
                if len(path) > 400:        # 고리 방어
                    ok = False
                    break
            (chains if ok else tangled).append([at[k] for k in path])
    rest = [i for i in range(len(pieces)) if i not in seen]
    return chains, tangled, rest


def run_length(path) -> float:
    """가닥 한 구간의 «표에 들어갈» 길이 — **유클리드** 총연장이다.

    ★한 번 L1(맨해튼)으로 바꿨다가 되돌렸다. `planar.py` 가 배관을 만들 때
      L1 으로 재는 것은 맞지만, 그 값은 바로 뒤 «마무리 1/2 — 배관 길이 갱신»
      (`PipeManager.update_all_pipe_lengths` → `compute_length` → `dist_3d`)
      이 **유클리드로 전부 덮는다.** 표에 실제로 실리는 규약은 유클리드다.

      그 사실을 모른 채 L1 로 선언하면 접힌 배관만 다른 자로 재게 된다.
      실측으로 걸렸다 — 접힌 배관 23개의 표 값 8,819 mm 를 좌표 L1 11,211 mm
      와 견주고 「길이를 잃었다」고 읽었는데, 견줄 자가 틀렸던 것이다.

    지시서 §1 의 「가닥 총연장 약 665 mm」가 바로 이 유클리드 값이다.
    """
    return sum(math.dist(path[i], path[i + 1]) for i in range(len(path) - 1))


def fold_board(board, chains, *, eps=SNAP_EPS):
    """손질판의 간선을 접는다 — 가닥 하나를 직선 간선 하나로.

    ★`pts` 는 건드리지 않는다. 노드 번호를 다시 매기면 `sources`·`valves` 가
      가리키는 자리가 통째로 어긋난다(그 둘은 인덱스로 산다). 안 쓰이게 된
      중간 점은 남겨 둔다 — `build_planar_graph` 는 간선이 가리키는 점만 본다.

    접지 않는 경우를 **세어서** 돌려준다(지어내지 않는다):

        끝점을 손질판에서 못 찾음      좌표가 eps 밖
        중간 노드에 다른 배관이 붙음   접으면 그 배관이 고아가 된다

    돌려주는 것: {"folded", "edge_len_mm", "skipped", "why", "before", "after"}
    """
    pts = list(getattr(board, "pts", None) or ())
    edges = set(tuple(sorted(e)) for e in (getattr(board, "edges", None) or ()))
    idx_of: dict = {}
    for i, p in enumerate(pts):
        idx_of.setdefault(_grid_key(p, eps), []).append(i)

    deg: dict = {}
    for i, j in edges:
        deg[i] = deg.get(i, 0) + 1
        deg[j] = deg.get(j, 0) + 1

    def near(p):
        best, bd = None, eps
        for k in _around(p, eps):
            for i in idx_of.get(k, ()):
                d = math.dist(pts[i], p)
                if d <= bd:
                    best, bd = i, d
        return best

    folded, declared, why = 0, {}, {}
    done_strand = 0
    before_edges = len(edges)
    drop, add = set(), set()
    n_cut = 0
    for path in chains:
        nodes = [near(p) for p in path]
        if any(n is None for n in nodes):
            why["점을 손질판에서 못 찾음"] = why.get(
                "점을 손질판에서 못 찾음", 0) + 1
            continue
        # ★가닥 «전체» 를 접으면 안 된다 — 가닥의 끝은 헤드만이 아니라 **가지관
        #   접속점**이기도 하고, 그 너머로 짧은 꼬리가 더 있기도 하다. 실측:
        #
        #     151(차수1·헤드) – 152 – 153 – 128(차수3·가지관) – 154   ← 꼬리 30 mm
        #
        #   128 을 삼켜 버리면 거기 붙은 가지관이 통째로 떨어진다. 그래서 «다른
        #   배관이 붙은 노드» 에서 **자르고**, 잘린 구간마다 따로 접는다.
        #   지시서 §1 의 「가지관 접속점 → 헤드, 직선」이 바로 이것이다.
        cuts = [0, len(nodes) - 1]
        for t in range(1, len(nodes) - 1):
            n = nodes[t]
            here = {tuple(sorted((n, nodes[t - 1]))),
                    tuple(sorted((n, nodes[t + 1])))}
            own = sum(1 for e in here if e in edges)
            if deg.get(n, 0) > own:
                cuts.append(t)
        cuts = sorted(set(cuts))
        n_cut += len(cuts) - 2
        done_strand += 1
        for s, e_ in zip(cuts, cuts[1:]):
            sub = nodes[s:e_ + 1]
            if len(sub) < 2 or sub[0] == sub[-1]:
                continue
            mine = set()
            for a, b in zip(sub, sub[1:]):
                if a == b:
                    continue
                ee = tuple(sorted((a, b)))
                if ee in edges:
                    mine.add(ee)
            key = tuple(sorted((sub[0], sub[-1])))
            if key in edges and len(mine) <= 1:
                continue          # 이미 직선 하나 — 접을 것이 없다
            drop |= mine
            add.add(key)
            declared[key] = run_length(path[s:e_ + 1])
            folded += 1

    edges = (edges - drop) | add
    board.edges = frozenset(edges)
    # ★«접힌 구간 수» 와 «다룬 가닥 수» 는 다르다 — 한 가닥이 분기에서 잘려
    #   여러 구간이 된다. 둘을 한 이름으로 세면 못 접은 수가 음수가 된다.
    return {"folded": folded, "strands": done_strand,
            "edge_len_mm": declared,
            "skipped": len(chains) - done_strand, "why": why, "cuts": n_cut,
            "before": before_edges, "after": len(edges)}


def _grid_key(p, cell):
    return (int(p[0] // cell), int(p[1] // cell))


def _around(p, cell):
    gx, gy = int(p[0] // cell), int(p[1] // cell)
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            yield (gx + dx, gy + dy)
