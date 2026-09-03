# -*- coding: utf-8 -*-
"""찍기 단계가 브라우저로 내려보내는 도면 — 직렬화와 상한만 다룬다."""
from __future__ import annotations

import json
import os
import time

from routes.module_f.common import (
    IMPORT_WORK_ROOT, MAX_ARCS, MAX_CIRCLES, MAX_SEGS, _layer_category, _r1)


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

    # ── 묶음마다 «도면에 실제로 몇 개 있나» 를 따로 센다.
    #
    # ★아래 `n_seg` 는 **그려 보낸** 수다(MAX_SEGS 에서 잘린다). 그것을 목록에
    #   개수로 적으면 큰 도면에서 거짓말이 된다 — 사람이 「이 레이어는 500개
    #   짜리」라고 읽는데 실제로는 3,000개일 수 있다. 세는 것과 그리는 것을
    #   가른다.
    # ★길이도 같이 잰다. 어느 묶음이 배관인지 고를 때 «몇 개» 보다 «총 몇 m»
    #   가 쓸모 있다. 판정은 하지 않는다 — 판정하려다 틀린 규칙을 만들 뻔했다
    #   (실측: 평면도의 진짜 배관은 긴 선분이 10~18% 뿐이고, 층 구획선이
    #   100% 다. 「길면 배관」은 정반대로 작동한다).
    import math as _math
    from statistics import median as _median
    stat: dict[tuple, dict] = {}

    def bump(ly, c, *, seg_mm=None, circle=False, arc=False):
        k = (ly, int(c) if isinstance(c, int) else c)
        st = stat.get(k)
        if st is None:
            st = stat[k] = {"n": 0, "mm": 0.0, "lens": [], "cir": 0, "arc": 0}
        if seg_mm is not None:
            st["n"] += 1
            st["mm"] += seg_mm
            if len(st["lens"]) < 20000:      # 중앙값 표본 — 메모리 상한
                st["lens"].append(seg_mm)
        if circle:
            st["cir"] += 1
        if arc:
            st["arc"] += 1

    for ly, c, a, b in world.segs:
        n_seg += 1
        grow(a[0], a[1])
        grow(b[0], b[1])
        bump(ly, c, seg_mm=_math.hypot(b[0] - a[0], b[1] - a[1]))
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
        bump(ly, c, circle=True)
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
        bump(ly, c, arc=True)
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

    # 잰 것을 묶음에 붙인다. `n_seg` 는 «그린 수», `n_all` 은 «도면에 있는 수» —
    # 둘이 다르면 화면이 그 사실을 말한다.
    for k, b in bundles.items():
        st = stat.get(k) or {"n": 0, "mm": 0.0, "lens": [], "cir": 0, "arc": 0}
        b["n_all"] = st["n"]
        b["len_m"] = round(st["mm"] / 1000.0, 1)
        b["len_mid"] = int(round(_median(st["lens"]))) if st["lens"] else 0
        b["n_circle_all"] = st["cir"]
        b["n_arc_all"] = st["arc"]

    # 순서는 종전 뜻(«덩치 큰 것부터»)을 지키되 **잘리지 않은 수**로 센다.
    #
    # ★종전에는 `n_seg`(그려 보낸 수)로 정렬했다. 상한에 걸린 큰 도면에서는
    #   그 값이 «먼저 그려진 순» 이라 사실상 임의였다.
    # ★«총 연장» 순도 재 봤지만 **더 낫다고 못 보였다** — 계통도에서는 층
    #   구획선(4,028 m · 1,108선분 · 중앙 4,350mm)이, 기계실에서는 잡선
    #   `l4`(1,336 m)가 여전히 1등이다. 도면 테두리 4선분이 430 m 를 내기도
    #   한다. 그래서 순서를 바꾸지 않고 **수치를 옆에 적어** 사람이 보게 한다.
    ordered = sorted(bundles.items(),
                     key=lambda kv: -(kv[1]["n_all"] + kv[1]["n_circle_all"]))
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


def pick_store_dir() -> str:
    """찍은스펙이 쌓이는 폴더 — «어디에 저장본이 있나» 의 유일한 답.

    ★엔진(G)을 올리지 않고도 답할 수 있어야 한다. 업로드 청소부가 «이 도면을
      쓰는 저장본이 있나» 를 물어야 하는데, 그 물음 하나 때문에 G 트리를 통째로
      import 시킬 수는 없다(부팅 전이면 `services` 가 sys.path 에 없어 실패한다).
      `_boot()` 가 `handoff.import_write_root` 를 여기로 갈아 끼우므로 값은 같다.
    """
    return str(IMPORT_WORK_ROOT / "0단계_새찍기")


def referenced_sources() -> set[str]:
    """저장본들이 가리키는 원본 도면 (normcase 절대경로).

    업로드 청소부가 «지워도 되는 파일» 을 가리는 데 쓴다. 목록이 아니라 집합인
    이유는 부르는 쪽이 «들었나» 만 묻기 때문이다.
    """
    return {os.path.normcase(os.path.abspath(it["source_dxf"]))
            for it in _saved_keys() if it.get("source_dxf")}


def _saved_keys() -> list[dict]:
    """이미 찍어 둔 도면들 — 데스크톱 E 로 찍은 것도 여기 그대로 보인다."""
    out = []
    d = pick_store_dir()
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
