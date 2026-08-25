# -*- coding: utf-8 -*-
"""찍기 단계가 브라우저로 내려보내는 도면 — 직렬화와 상한만 다룬다."""
from __future__ import annotations

import json
import os
import time

from routes.module_f.common import (
    MAX_ARCS, MAX_CIRCLES, MAX_SEGS, _layer_category, _r1)

def _world_payload(world) -> dict:
    """DXF 세계 → 캔버스가 그릴 수 있는 묶음별 좌표 다발.

    레이어×색(bundle) 단위로 접는다. 찍기가 재료를 그 단위로 고르므로
    화면 토글·강조도 같은 단위여야 손으로 맞출 필요가 없다.
    """
    from services.cad_import.colors import cname, rgb_dark

    bundles: dict[tuple, dict] = {}
    cat_cache: dict[str, str] = {}

    def slot(ly, c) -> dict:
        k = (ly, int(c) if isinstance(c, int) else c)
        b = bundles.get(k)
        if b is None:
            name = str(ly)
            cat = cat_cache.get(name)
            if cat is None:
                cat = _layer_category(name)
                cat_cache[name] = cat
            b = {"layer": name, "color": c, "name": cname(c),
                 "css": rgb_dark(c), "cat": cat,
                 "segs": [], "circles": [], "arcs": [],
                 "n_seg": 0, "n_circle": 0, "n_arc": 0}
            bundles[k] = b
        return b

    minx = miny = float("inf")
    maxx = maxy = float("-inf")

    def grow(x, y):
        nonlocal minx, miny, maxx, maxy
        if x < minx:
            minx = x
        if x > maxx:
            maxx = x
        if y < miny:
            miny = y
        if y > maxy:
            maxy = y

    n_seg = n_cir = n_arc = 0
    shown_seg = shown_cir = shown_arc = 0

    for ly, c, a, b in world.segs:
        n_seg += 1
        grow(a[0], a[1])
        grow(b[0], b[1])
        if shown_seg >= MAX_SEGS:
            continue
        s = slot(ly, c)
        s["segs"] += [_r1(a[0]), _r1(a[1]), _r1(b[0]), _r1(b[1])]
        s["n_seg"] += 1
        shown_seg += 1

    for ly, c, cx, cy, r in world.circles:
        n_cir += 1
        grow(cx - r, cy - r)
        grow(cx + r, cy + r)
        if shown_cir >= MAX_CIRCLES:
            continue
        s = slot(ly, c)
        s["circles"] += [_r1(cx), _r1(cy), _r1(r)]
        s["n_circle"] += 1
        shown_cir += 1

    angs = list(getattr(world, "arc_ang", ()) or ())
    for i, (ly, c, cx, cy, r) in enumerate(world.arcs):
        n_arc += 1
        grow(cx - r, cy - r)
        grow(cx + r, cy + r)
        if shown_arc >= MAX_ARCS:
            continue
        ang = angs[i] if i < len(angs) else None
        sa, sweep = (float(ang[0]), float(ang[1])) if ang else (0.0, 360.0)
        s = slot(ly, c)
        s["arcs"] += [_r1(cx), _r1(cy), _r1(r), round(sa, 2), round(sweep, 2)]
        s["n_arc"] += 1
        shown_arc += 1

    if minx == float("inf"):
        minx = miny = 0.0
        maxx = maxy = 1.0

    ordered = sorted(bundles.items(),
                     key=lambda kv: -(kv[1]["n_seg"] + kv[1]["n_circle"]))
    out = []
    for i, ((ly, c), b) in enumerate(ordered):
        b = dict(b)
        b["i"] = i
        b["id"] = f"{ly}{c}"
        out.append(b)

    cats: dict[str, int] = {}
    for b in out:
        cats[b["cat"]] = cats.get(b["cat"], 0) + 1

    return {
        "bounds": {"minx": minx, "miny": miny, "maxx": maxx, "maxy": maxy},
        "bundles": out,
        "cats": cats,
        "counts": {"segs": n_seg, "circles": n_cir, "arcs": n_arc},
        "shown": {"segs": shown_seg, "circles": shown_cir, "arcs": shown_arc},
        "dropped": {"segs": n_seg - shown_seg, "circles": n_cir - shown_cir,
                    "arcs": n_arc - shown_arc},
    }


def _pts_bounds(pts) -> dict:
    if not pts:
        return {"minx": 0.0, "miny": 0.0, "maxx": 1.0, "maxy": 1.0}
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return {"minx": min(xs), "miny": min(ys),
            "maxx": max(xs), "maxy": max(ys)}


def _saved_keys() -> list[dict]:
    """이미 찍어 둔 도면들 — 데스크톱 E 로 찍은 것도 여기 그대로 보인다."""
    from services.cad_import.pipeline import handoff
    out = []
    d = handoff.pick_out_dir()
    if not os.path.isdir(d):
        return out
    for name in sorted(os.listdir(d)):
        if not name.endswith("_찍은스펙.json") or "자동백업" in name:
            continue
        key = name[: -len("_찍은스펙.json")]
        path = os.path.join(d, name)
        src = ""
        try:
            with open(path, encoding="utf-8") as f:
                src = json.load(f).get("source_dxf") or ""
        except Exception:  # noqa: BLE001 — 목록이므로 한 건 실패로 멈추지 않는다
            pass
        out.append({
            "key": key,
            "source_dxf": src,
            "source_exists": bool(src) and os.path.isfile(src),
            "picked_at": time.strftime("%Y-%m-%d %H:%M",
                                       time.localtime(os.path.getmtime(path))),
        })
    return out
