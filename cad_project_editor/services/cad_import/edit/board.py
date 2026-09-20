# -*- coding: utf-8 -*-
"""손질 판 — 이음·삭제·종류·급수. 화면 없음.

망 계산은 user_net, 헤드 중심은 flow.attach_heads_center.
이음 규칙·물길 규칙을 여기서 다시 쓰지 않는다.
"""
import math
from collections import defaultdict

from services.cad_import.colors import BODY_COLORS, KIND_COLORS, KIND_COLORS_DRY
from services.cad_import.kinds import (
    CONFIRMED_KINDS, disk_key, disk_kind_list, normalize_head_kind,
)
from services.cad_import.pipeline.expand import gnear, gput
from services.cad_import.pipeline.flow import SRC_SNAP, attach_heads_center, head_nodes
from services.cad_import.pipeline.user_net import (
    apply_deletes, apply_head_bridge, apply_joins, apply_kind_overrides,
    apply_tee_bridge, colinear_bridges, colinear_chain_segs, corner_bridges,
    head_bridge, head_cover_ok, insert_node_on_pipe, recompute_bodies,
    tee_bridges, wet_heads,
)

PICK_PX = 10.0


def head_face_colors(kinds, wet_set=None):
    """종류 채움색. wet_set=None(물흐름 전)은 밝은 종류색, 이후는 젖음=밝음·마름=어두운 톤."""
    cols = []
    for i, kind in enumerate(kinds):
        if wet_set is not None and i not in wet_set:
            cols.append(KIND_COLORS_DRY.get(kind, KIND_COLORS_DRY["미지정"]))
        else:
            cols.append(KIND_COLORS.get(kind, KIND_COLORS["미지정"]))
    return cols


def wet_disk_keys(disks, wet_set):
    """젖은 헤드 인덱스 → 좌표 키. 편집 후에도 동일 헤드를 다시 칠할 때 쓴다."""
    keys = set()
    for i in wet_set or ():
        if not isinstance(i, int) or i < 0 or i >= len(disks):
            continue
        d = disks[i]
        keys.add(disk_key(d[0], d[1], d[2]))
    return keys


def wet_set_from_disk_keys(disks, keys):
    """좌표 키 → 현재 disks 인덱스. 삭제된 헤드는 자연 소멸."""
    if keys is None:
        return None
    out = set()
    for i, d in enumerate(disks or ()):
        if disk_key(d[0], d[1], d[2]) in keys:
            out.add(i)
    return out


def _seg_key(a, b):
    return tuple(sorted((tuple(a), tuple(b))))


def _edge_segments(pts, edges):
    return [(pts[i], pts[j]) for i, j in edges]


def _head_crosses(bridge, disks):
    """bridge 내부가 헤드 원판을 지나면 True. 통과 시 head_cover_ok 로 재심사."""
    (a, b) = bridge
    dx, dy = b[0] - a[0], b[1] - a[1]
    den = dx * dx + dy * dy
    for x, y, r in disks:
        if den == 0:
            d = math.hypot(a[0] - x, a[1] - y)
        else:
            t = max(0.0, min(1.0, ((x - a[0]) * dx + (y - a[1]) * dy) / den))
            d = math.hypot(a[0] + t * dx - x, a[1] + t * dy - y)
        if d < r - 1e-7:
            return True
    return False


def body_seg_groups(pts, edges, bodies):
    """물덩이별 (선분들, 색) — 한 덩이 = 한 색.

    색은 크기(간선 수) 순위. 상위 len(BODY_COLORS)만 고유색, 나머지는 회색.
    wrap 하면 큰 덩이 둘이 같은 색 → 「이어진 것」으로 오인.
    """
    body_of = {n: bi for bi, body in enumerate(bodies) for n in body}
    by_body = defaultdict(list)
    for i, j in edges:
        bi = body_of.get(i, -1)
        by_body[bi].append((pts[i], pts[j]))
    ranked = sorted(by_body.items(), key=lambda kv: -len(kv[1]))
    groups = []
    rank = 0
    ncol = len(BODY_COLORS)
    for bi, segs in ranked:
        if bi < 0 or rank >= ncol:
            color = "#666666"
        else:
            color = BODY_COLORS[rank]
        if bi >= 0:
            rank += 1
        groups.append((segs, color))
    return groups


class EditBoard:
    """저장 좌표와 화면 망을 함께 관리한다."""

    def __init__(self, key, pts, edges, disks, original_edges=None, hnodes=None,
                 ups=None, head_kinds=None, ho=None):
        self.key, self.pts = key, list(pts)
        self.base_edges = frozenset(edges)
        self.original_edges = frozenset(original_edges if original_edges is not None
                                        else edges)
        self.edges, self.disks = frozenset(edges), list(disks)
        self.ups = list(ups or ())
        self.head_kinds = [dict(r) if isinstance(r, dict) else r
                           for r in (head_kinds or ())]
        self.kind_overrides = []
        self.ho = [dict(s) for s in (ho or ()) if isinstance(s, dict)]
        self.joins, self.deletes, self.history = [], [], []
        self.sources = []
        self.valves = []
        self._complete_heads()
        _ = hnodes
        self._head_nodes()
        self._source_nodes_cache = None
        self.pending = None
        self._refresh_kind_views()

    def segments(self):
        return _edge_segments(self.pts, self.edges)

    def _kind(self, seg):
        key = _seg_key(*seg)
        if key in {_seg_key(self.pts[i], self.pts[j]) for i, j in self.original_edges}:
            return "original"
        if key in {_seg_key(tuple(b[0]), tuple(b[1]))
                   for rec in self.joins for b in rec["bridges"]}:
            return "user_join"
        return "auto_join"

    def join(self, first, second):
        """선택 두 관 사이만 잇는다. 헤드 원 통과는 헤드걸침만 허용.

        순서: 일직선 → T자(끝→몸통 수선) → 직각 코너(끝↔끝).
        """
        raw = colinear_bridges(first, second, segs=self.segments())
        kind = "일직선"
        tee_raw = None
        if not raw:
            chain_a = colinear_chain_segs(self.pts, self.edges, first)
            chain_b = colinear_chain_segs(self.pts, self.edges, second)
            tee_raw = tee_bridges(first, second, chain_a=chain_a,
                                 chain_b=chain_b)
            if tee_raw:
                raw = [(a, b) for a, b, _tr in tee_raw]
                kind = "T자"
            else:
                raw = corner_bridges(first, second)
                kind = "직각"
        bridges, blocked, covered = [], 0, 0
        for b in raw:
            if not _head_crosses(b, self.disks):
                bridges.append(b)
            elif head_cover_ok(b, self.disks):
                bridges.append(b)
                covered += 1
            else:
                blocked += 1
        if not bridges:
            return 0, blocked, covered, kind
        before = self._snapshot()
        trunks_saved = []
        if kind == "T자":
            keep = {tuple(sorted((tuple(a), tuple(b)))) for a, b in bridges}
            for a, b, trunk in tee_raw:
                if tuple(sorted((tuple(a), tuple(b)))) not in keep:
                    continue
                self.pts, self.edges = apply_tee_bridge(
                    self.pts, self.edges, a, b, trunk)
                trunks_saved.append([list(trunk[0]), list(trunk[1])])
        else:
            self.pts, self.edges = apply_joins(self.pts, self.edges, bridges)
        rec = {"a": [list(p) for p in first], "b": [list(p) for p in second],
               "bridges": [[list(a), list(b)] for a, b in bridges],
               "kind": kind}
        if trunks_saved:
            rec["trunks"] = trunks_saved
        self.joins.append(rec)
        self.history.append(before)
        self._invalidate_source_cache()
        return len(bridges), blocked, covered, kind

    def join_head(self, seg, disk):
        """선택 관 자유단을 헤드 «중심»에 붙인다(관말). 배관↔배관과 같은 모드 1."""
        bridge = head_bridge(seg, disk)
        if bridge is None:
            return 0, 0, 0, "헤드"
        hx, hy, hr = float(disk[0]), float(disk[1]), float(disk[2])
        others = [d for d in self.disks
                  if abs(d[0] - hx) > 1e-7 or abs(d[1] - hy) > 1e-7
                  or abs(d[2] - hr) > 1e-7]
        if _head_crosses(bridge, others):
            return 0, 1, 0, "헤드"
        before = self._snapshot()
        self.pts, self.edges = apply_head_bridge(self.pts, self.edges, bridge)
        self.joins.append({
            "a": [list(p) for p in seg],
            "b": [[hx, hy], [hx, hy]],
            "head": [hx, hy, hr],
            "bridges": [[list(bridge[0]), list(bridge[1])]],
            "kind": "헤드",
        })
        self.history.append(before)
        self._invalidate_source_cache()
        self._head_nodes()
        return 1, 0, 0, "헤드"

    def delete(self, seg):
        """원본·자동 이음·유저 이음 모두 현재 망에서 지운다."""
        before = self._snapshot()
        kind = self._kind(seg)
        n0 = len(self.edges)
        self.edges = apply_deletes(self.pts, self.edges, [seg])
        n = n0 - len(self.edges)
        self.deletes.append({"a": list(seg[0]), "b": list(seg[1]), "kind": kind,
                             "n": n})
        self.history.append(before)
        self._invalidate_source_cache()
        return kind, n

    def _refresh_kind_views(self):
        self.disk_kinds = disk_kind_list(self.disks, self.head_kinds)

    def _disk_index(self, disk):
        hx, hy, hr = float(disk[0]), float(disk[1]), float(disk[2])
        for i, d in enumerate(self.disks):
            if (abs(float(d[0]) - hx) <= 1e-7
                    and abs(float(d[1]) - hy) <= 1e-7
                    and abs(float(d[2]) - hr) <= 1e-7):
                return i
        return None

    def _upsert_kind_override(self, hx, hy, hr, kind):
        key = disk_key(hx, hy, hr)
        xy = (round(float(hx), 1), round(float(hy), 1))
        for rec in self.kind_overrides:
            c = rec.get("c") or ()
            if len(c) < 2:
                continue
            r = rec.get("r", rec.get("head_r"))
            if r is not None and disk_key(c[0], c[1], r) == key:
                rec["kind"] = kind
                return
            if (r is None
                    and (round(float(c[0]), 1), round(float(c[1]), 1)) == xy):
                rec["kind"] = kind
                rec["r"] = float(hr)
                return
        self.kind_overrides.append(
            {"c": [float(hx), float(hy)], "r": float(hr), "kind": kind})

    def set_head_kind(self, disk, kind):
        """고른 헤드를 단추에 적힌 종류로 덮는다. 이음·ups·물길 불변."""
        kind = normalize_head_kind(kind)
        if kind not in CONFIRMED_KINDS:
            return None
        di = self._disk_index(disk)
        if di is None:
            return None
        before = self._snapshot()
        hx, hy, hr = float(disk[0]), float(disk[1]), float(disk[2])
        self._upsert_kind_override(hx, hy, hr, kind)
        self.head_kinds = apply_kind_overrides(
            self.head_kinds, [{"c": [hx, hy], "r": hr, "kind": kind}])
        self._refresh_kind_views()
        if di >= len(self.disk_kinds) or self.disk_kinds[di] != kind:
            while len(self.disk_kinds) <= di:
                self.disk_kinds.append("미지정")
            self.disk_kinds[di] = kind
        self.history.append(before)
        return kind

    def undo(self):
        if not self.history:
            return False
        (self.pts, self.edges, self.joins, self.deletes, self.sources,
         self.valves, self.head_kinds, self.kind_overrides) = self.history.pop()
        self.pending = None
        self._invalidate_source_cache()
        self._refresh_kind_views()
        self._complete_heads()
        self._head_nodes()
        return True

    def _snapshot(self):
        return (list(self.pts), self.edges, list(self.joins), list(self.deletes),
                list(self.sources), list(self.valves),
                [dict(r) if isinstance(r, dict) else r for r in self.head_kinds],
                [dict(r) for r in self.kind_overrides])

    def bodies(self):
        return recompute_bodies(self.pts, self.edges)

    def _invalidate_source_cache(self):
        self._source_nodes_cache = None

    def _complete_heads(self):
        """헤드 접속 완성 — water/flow.attach_heads_center SSOT."""
        (self.pts, self.edges, self.head_centers,
         n_wire, self.multi_arm_heads) = attach_heads_center(
            self.pts, self.edges, self.disks)
        return n_wire

    def _head_nodes(self):
        self.hnodes = [set(x) for x in head_nodes(
            self.pts, self.disks, edges=self.edges, upright=self.ups)]
        return self.hnodes

    def _source_nodes(self, hnodes=None):
        """현재 망에서 급수원이 될 수 있는 배관 노드. 헤드 원 안 노드는 제외."""
        if hnodes is None and self._source_nodes_cache is not None:
            return self._source_nodes_cache
        if hnodes is None:
            hnodes = self._head_nodes()
        forbidden = {n for nodes in hnodes for n in nodes}
        incident = {n for edge in self.edges for n in edge}
        pool = incident - forbidden
        if not pool or not self.disks:
            candidates = sorted(pool)
            self._source_nodes_cache = candidates
            return candidates
        cell = 500.0
        ng = defaultdict(list)
        for n in pool:
            px, py = self.pts[n]
            gput(ng, cell, px, py, n)
        inside = set()
        for hx, hy, hr in self.disks:
            lim = float(hr) + 1e-7
            rings = 1 + int(lim // cell)
            for n in set(gnear(ng, cell, hx, hy, rings=rings)):
                if n in inside:
                    continue
                px, py = self.pts[n]
                if math.hypot(px - hx, py - hy) <= lim:
                    inside.add(n)
        candidates = [n for n in pool if n not in inside]
        self._source_nodes_cache = candidates
        return candidates

    def snap_source(self, x, y, hnodes=None, source_nodes=None):
        """헤드/헤드원 안을 빼고 2,500mm 안의 배관 노드 하나만 고른다."""
        if source_nodes is None:
            source_nodes = self._source_nodes(hnodes)
        candidates = [(math.hypot(self.pts[n][0] - x, self.pts[n][1] - y), n)
                      for n in source_nodes]
        if not candidates:
            return None
        dist, node = min(candidates)
        return node if dist <= SRC_SNAP else None

    def pick_source(self, x, y, max_d=None):
        if not self.sources:
            return None
        lim = SRC_SNAP if max_d is None else float(max_d)
        npts = len(self.pts)
        best, best_d = None, lim
        for node in self.sources:
            if not isinstance(node, int) or node < 0 or node >= npts:
                continue
            d = math.hypot(self.pts[node][0] - x, self.pts[node][1] - y)
            if d <= best_d:
                best, best_d = node, d
        return best

    def remove_source(self, node):
        if node not in self.sources:
            return False
        before = self._snapshot()
        self.sources.remove(node)
        self.history.append(before)
        return True

    def toggle_source(self, x, y):
        """배관 노드를 급수원으로 토글. 같은 노드를 다시 찍으면 해제."""
        node = self.snap_source(x, y)
        if node is None:
            return None, False
        if node in self.sources:
            self.remove_source(node)
            return node, False
        before = self._snapshot()
        self.sources.append(node)
        self.history.append(before)
        return node, True

    def pick_valve(self, x, y, max_d=None):
        if not self.valves:
            return None
        lim = SRC_SNAP if max_d is None else float(max_d)
        npts = len(self.pts)
        best, best_d = None, lim
        for node in self.valves:
            if not isinstance(node, int) or node < 0 or node >= npts:
                continue
            d = math.hypot(self.pts[node][0] - x, self.pts[node][1] - y)
            if d <= best_d:
                best, best_d = node, d
        return best

    def remove_valve(self, node):
        if node not in self.valves:
            return False
        before = self._snapshot()
        self.valves.remove(node)
        self.history.append(before)
        return True

    def toggle_valve(self, x, y, max_d):
        existing = self.pick_valve(x, y, max_d)
        if existing is not None:
            self.remove_valve(existing)
            return existing, False
        before = self._snapshot()
        pts, edges, node = insert_node_on_pipe(
            self.pts, self.edges, x, y, max_d=max_d)
        if node is None:
            return None, False
        self.pts, self.edges = pts, edges
        if node not in self.valves:
            self.valves.append(node)
        self._invalidate_source_cache()
        self.history.append(before)
        return node, True

    def water_state(self):
        """현재 급수원에서 간선 BFS. 물길 규칙은 추가하지 않는다."""
        self._head_nodes()
        adj = defaultdict(set)
        for a, b in self.edges:
            adj[a].add(b)
            adj[b].add(a)
        hop, todo = {}, list(self.sources)
        for n in self.sources:
            hop[n] = 0
        pos = 0
        while pos < len(todo):
            u = todo[pos]
            pos += 1
            for v in adj[u]:
                if v not in hop:
                    hop[v] = hop[u] + 1
                    todo.append(v)
        reach = set(hop)
        wet = wet_heads(reach, self.hnodes)
        wet_edges = {(min(a, b), max(a, b)) for a, b in self.edges
                     if a in reach and b in reach}
        return dict(reach=reach, hop=hop, wet_edges=wet_edges, wet_heads=wet,
                    total_heads=len(self.disks))

    def payload(self):
        return {"version": 1, "joins": self.joins, "deletes": self.deletes,
                "sources": [{"tag": f"Z{i + 1}", "xy": list(self.pts[n])}
                            for i, n in enumerate(self.sources)],
                "valve_picks": [{"xy": list(self.pts[n])}
                                for n in self.valves],
                "kind_overrides": list(self.kind_overrides),
                "head_kinds": [dict(r) if isinstance(r, dict) else r
                               for r in self.head_kinds],
                "ho": [dict(s) for s in self.ho]}
