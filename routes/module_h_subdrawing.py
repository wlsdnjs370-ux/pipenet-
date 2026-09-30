"""H source display adapter. No display primitive is fed into path extraction."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path

from routes.module_f.subdrawing import _EntWorld
from routes.module_f.world import _world_payload
from routes.module_h_plan_display import source_color


def _grow(bounds: dict | None, other: dict) -> dict:
    if bounds is None:
        return dict(other)
    return {key: (min if key.startswith("min") else max)(bounds[key], other[key])
            for key in ("minx", "miny", "maxx", "maxy")}


def display_payload(entities: list[dict], colors: dict, path: str | Path,
                    *, source: dict | None = None) -> dict:
    """Adapt a shared-read snapshot; leave extraction entities untouched.

    Hidden layers/parents stay in toggleable bundles. Initial fit follows CAD's
    visible contents, not invisible remote objects. Never crop calculation data.
    """
    if source is None:
        from routes.module_f.sub_fastread import read_view
        source = read_view(path, source_display=True)["parsed"]["source_display"]
    payload = {"bundles": [], "texts": [], "counts": {}, "shown": {}, "dropped": {}, "cats": {}}
    all_bounds = visible_bounds = None
    for hidden in (False, True):
        world = _EntWorld()
        for item in source["primitives"]:
            if bool(item["hidden"]) != hidden:
                continue
            layer, color, xy = item["layer"], item["color"], item["xy"]
            if item["kind"] == "segs":
                world.segs.append((layer, color, tuple(xy[:2]), tuple(xy[2:])))
            elif item["kind"] == "circles":
                world.circles.append((layer, color, *xy))
            elif item["kind"] == "arcs":
                world.arcs.append((layer, color, *xy[:3]))
                world.arc_ang.append(tuple(xy[3:]))
        part = _world_payload(world, complete=True)
        bundles = {(b["layer"], b["color"]): b for b in part["bundles"]}
        for original in source["texts"]:
            if bool(original.get("hidden")) != hidden:
                continue
            text = deepcopy(original)
            key = (text["layer"], text["color"])
            if key not in bundles:
                bundles[key] = dict(layer=key[0], color=key[1], name=str(key[1]), cat="TEXT",
                                    segs=[], circles=[], arcs=[], n_seg=0, n_circle=0, n_arc=0,
                                    n_all=0, n_circle_all=0, n_arc_all=0, len_m=0, len_mid=0,
                                    display_partial=False)
            bundles[key].setdefault("source_text_points", []).append((text["x"], text["y"]))
            identity = f"{key[0]}\x1f{key[1]}" + ("\x1fhidden" if hidden else "")
            text.update(bundle_id=identity, css=source_color(key[1]))
            payload["texts"].append(text)
        for key, bundle in bundles.items():
            bounds = None
            def add(x: float, y: float, radius: float = 0) -> None:
                nonlocal bounds
                bounds = _grow(bounds, dict(minx=x-radius, maxx=x+radius,
                                           miny=y-radius, maxy=y+radius))
            for x, y in bundle.pop("source_text_points", []):
                add(x, y)
            for j in range(0, len(bundle["segs"]), 2):
                add(*bundle["segs"][j:j+2])
            for name, stride in (("circles", 3), ("arcs", 5)):
                for j in range(0, len(bundle[name]), stride):
                    add(*bundle[name][j:j+3])
            bundle.update(id=f"{key[0]}\x1f{key[1]}" + ("\x1fhidden" if hidden else ""),
                          i=len(payload["bundles"]), css=source_color(key[1]),
                          source_hidden=hidden, bounds=bounds)
            payload["bundles"].append(bundle)
            if bounds:
                all_bounds = _grow(all_bounds, bounds)
                if not hidden:
                    visible_bounds = _grow(visible_bounds, bounds)
        for name, value in part["counts"].items():
            payload["counts"][name] = payload["counts"].get(name, 0) + value
    payload["counts"]["texts"] = len(payload["texts"])
    payload["shown"] = dict(payload["counts"])
    payload["dropped"] = dict.fromkeys(payload["counts"], 0)
    for b in payload["bundles"]:
        payload["cats"][b["cat"]] = payload["cats"].get(b["cat"], 0) + 1
    payload.update(bounds=visible_bounds or dict(minx=0, miny=0, maxx=1, maxy=1),
                   full_bounds=all_bounds, source_all_hidden=visible_bounds is None,
                   h_source_display=True, source_visibility=True,
                   source_unsupported=dict(source["unsupported"]))
    saved = source.get("saved_view")
    if saved and visible_bounds and not (saved["maxx"] < visible_bounds["minx"] or
            saved["minx"] > visible_bounds["maxx"] or saved["maxy"] < visible_bounds["miny"] or
            saved["miny"] > visible_bounds["maxy"]):
        payload["source_view_bounds"] = dict(saved)
    return payload
