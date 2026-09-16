"""Spatial acceleration must preserve attachment, join order and angular limits."""
from __future__ import annotations

import copy
import math
from pathlib import Path
import random
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "cad_project_editor_g"))
from services.cad_import.pipeline import flow, heads, stage45
from services.cad_import.pipeline.stage1 import DEFAULT_KNOBS, Graph, World, ang_between
from services.cad_import.pick.board import Board


@pytest.mark.parametrize("assume_unattached", [False, True])
def test_head_node_index_matches_exhaustive_search(monkeypatch, assume_unattached):
    rng = random.Random(1971)
    pts, edges, heads = [], set(), []
    for i in range(120):
        x, y = rng.uniform(-20000, 20000), rng.uniform(-20000, 20000)
        r = rng.choice([5.0, 49.999, 50.0, 250.0, 1000.0, 2500.0])
        n = len(pts)
        pts.extend([(x - 4000, y), (x + 4000, y)])
        edges.add((n, n + 1))
        heads.append((x, y, r))
        if i % 3:
            # Centre/rim boundaries, including negative and cross-cell positions.
            d = rng.choice([5.0, 5.000001, r + 50.0, r + 50.000001])
            n = len(pts)
            pts.extend([(x, y + d), (x, y + d + 4000)])
            edges.add((n, n + 1))
    fast = flow.stage5_split_through_uprights(pts, edges, heads, assume_unattached)
    original = flow.gnear

    def exhaustive_nodes(grid, cell, x, y, rings=1):
        sample = next(iter(grid.values()), ())
        if sample and isinstance(sample[0], int):
            return (n for bucket in grid.values() for n in bucket)
        return original(grid, cell, x, y, rings=rings)

    monkeypatch.setattr(flow, "gnear", exhaustive_nodes)
    slow = flow.stage5_split_through_uprights(pts, edges, heads, assume_unattached)
    assert fast == slow
    assert fast[2] > 0


@pytest.mark.parametrize("angle", [0.0, 17.0, 45.0, 89.0, 135.0])
def test_head_cover_index_preserves_joins_and_ties(monkeypatch, angle):
    g, ebundle, heads = Graph(), {}, []
    a = math.radians(angle)

    def xy(x, y):
        return (x * math.cos(a) - y * math.sin(a) - 501,
                x * math.sin(a) + y * math.cos(a) - 999)

    for row in range(12):
        y = row * 3500
        start = len(g.pts)
        g.pts.extend(xy(x, y) for x in [-1200, -200, 200, 1200])
        for e in [(start, start + 1), (start + 2, start + 3)]:
            g.edges.add(e)
            ebundle[e] = ("pipe", 1)
        # Tied centres with different radii: the first matching disk must win.
        heads.extend([(*xy(0, y), 100), (*xy(0, y), 250)])
        heads.extend([(*xy(0, y + 101), 100), (*xy(10000, y), 1800)])
    old_g, old_b = copy.deepcopy(g), dict(ebundle)
    fast = stage45.join_by_head_cover(g, ebundle, heads)
    monkeypatch.setattr(stage45, "_bbox_ids", lambda *args, **kw: range(len(heads)))
    slow = stage45.join_by_head_cover(old_g, old_b, heads)
    assert fast == slow
    assert g.edges == old_g.edges
    assert ebundle == old_b
    assert len(fast) == 12
    assert all(j["sym_r"] == 100 for j in fast)


def test_arrival_angle_shortcut_preserves_boundaries_and_neighbour_order():
    def exhaustive(g, neighbours):
        best = None
        for k in neighbours:
            wx, wy = g.pts[k]
            if ang_between((wx, wy), (1, 0)) <= 30:
                return None
            a = ang_between((-wx, -wy), (1, 0))
            if a <= 8 and (best is None or a < best[0]):
                best = (a, k)
        return best[1] if best else None

    rng = random.Random(123)
    for angle in [0, 30 - 1e-9, 30, 30 + 1e-9, 171, 172 - 1e-9, 172, 172 + 1e-9, 180]:
        for _ in range(10):
            g = Graph()
            angles = [angle, rng.uniform(0, 180), 180 - angle]
            g.pts = [(0, 0)] + [(100 * math.cos(math.radians(a)),
                                  100 * math.sin(math.radians(a))) for a in angles]
            ns = [1, 2, 3]
            rng.shuffle(ns)
            assert stage45._arr_edge(g, {0: ns}, 0, 1, 0) == exhaustive(g, ns)


@pytest.mark.parametrize("tolerance", [0.0, 0.5, 1.0, 2.0])
def test_triangle_endpoint_index_matches_exhaustive_pairs(monkeypatch, tolerance):
    rng = random.Random(682)
    strokes = []
    for i in range(35):
        x, y = rng.uniform(-800, 800), rng.uniform(-800, 800)
        a, b, c = (x, y), (x + 20, y), (x + 10, y + 17)
        strokes.extend([(a, b), (b, c), (c, (x + rng.choice([0, 0.5, 1, 1.000001]), y))])
    # Same-vertex, duplicates and zero-length strokes preserve legacy semantics.
    strokes += [((0, 0), (0, 0))] * 3
    fast = heads.closed_tris(strokes, tolerance)
    monkeypatch.setattr(heads, "_grid_near", lambda grid, *args, **kw:
                        (item for bucket in grid.values() for item in bucket))
    slow = heads.closed_tris(strokes, tolerance)
    assert fast == slow
    assert fast


def test_triangle_geometry_reused_but_each_ruler_still_checked(monkeypatch):
    world = World()
    for i, side in enumerate([100, 170, 250]):
        x = i * 4000.0
        a, b, c = (x, 0), (x + side, 0), (x + side / 2, side * math.sqrt(3) / 2)
        world.segs.extend(("head", 2, p, q) for p, q in [(a, b), (b, c), (c, a)])
    world.segs.append(("head", 2, (-1000, 0), (-800, 0)))
    world.circles.append(("head", 2, 4000, 0, 40))
    knobs = dict(DEFAULT_KNOBS)
    board = Board(world, knobs)
    picks = [{"bundle": ("head", 2), "tri_side": side, "label": "측벽형"}
             for side in [100, 170, 250, 110, 300, 100]]
    expected = []
    for pick in picks:
        clusters = heads.collect_head_clusters(world, {"heads": [pick]}, knobs)
        found, _, _ = heads.split_head_circles(clusters, knobs)
        expected.append([s for h in found if "tri_side" in h for s in h["tri_segs"]])
    calls = []
    original = heads.head_clusters

    def counted(*args, **kwargs):
        calls.append(1)
        return original(*args, **kwargs)

    monkeypatch.setattr(heads, "head_clusters", counted)
    assert [board.tri_head_segs(pick) for pick in picks] == expected
    assert len(calls) == 1
    assert expected[0] and expected[1] and expected[2]
    assert expected[4] == []
    # Tolerance is part of the geometry key; a changed configuration recomputes.
    board.kn["a1_lat"] = 2.0
    board.tri_head_segs(picks[0])
    assert len(calls) == 2
