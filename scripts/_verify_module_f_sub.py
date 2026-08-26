# -*- coding: utf-8 -*-
"""[H-2 · H-3] 계통도·기계실 라우트를 실제 앱으로 확인한다.

단위 테스트는 어댑터만 본다. 여기서는 `대조 서버.py` 를 올려 HTTP 계층까지
태우고, **실도면 한 장으로 계통도 경로를 실제로 뽑아** 본다 — 두 점을 찍는
좌표까지 포함해서. 그래야 «라우트는 있는데 엔진이 안 붙었다» 를 잡는다.

    python scripts/_verify_module_f_sub.py
"""
from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FAILS: list[str] = []
# 계통도 실도면 — 저장소에 있는 것 중 하나를 골라 쓴다.
SYSTEM_DXF_CANDIDATES = [
    ROOT / "data" / "sample_problem" / "대명동201동 계통도.dxf",
    ROOT / "samples" / "dxf" / "계통도_LH_306_배관망추출.dxf",
]


def check(label: str, cond: bool, detail: str = "") -> bool:
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


def _load_app():
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    spec = importlib.util.spec_from_file_location(
        "server_app_subverify", str(ROOT / "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.app


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    app = _load_app()
    app.config["TESTING"] = True
    from routes.module_f.jobs import _new_session
    from routes.module_f.slots import _slot_switch

    print("[H-2 · H-3] 계통도·기계실 라우트 실측")

    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True

        # ── 라우트 등록
        rules = {r.rule for r in app.url_map.iter_rules()}
        for path in ("/api/module-f/system/extract",
                     "/api/module-f/machineroom/extract",
                     "/api/module-f/sub/state"):
            check(f"라우트 존재 {path}", path in rules)

        # ── 슬롯 가드 — 평면도 슬롯에서 계통도를 뽑으려 하면 막는다
        sess = _new_session()
        sid = sess["id"]
        r = c.post("/api/module-f/system/extract",
                   json={"sid": sid, "pump_x": 0, "pump_y": 0,
                         "av_x": 1, "av_y": 1})
        check("평면도 슬롯에서 계통도 추출은 409", r.status_code == 409,
              f"HTTP {r.status_code}")

        _slot_switch(sess, "system")
        r = c.post("/api/module-f/system/extract",
                   json={"sid": sid, "pump_x": 0, "pump_y": 0,
                         "av_x": 1, "av_y": 1})
        check("도면 없이 추출은 409", r.status_code == 409, f"HTTP {r.status_code}")

        # ── 입력 검사
        sess["entities"] = [{"t": "L", "l": "P", "p": [0, 0, 100, 0]}]
        r = c.post("/api/module-f/system/extract", json={"sid": sid})
        check("두 점을 안 찍으면 400", r.status_code == 400, f"HTTP {r.status_code}")

        r = c.get(f"/api/module-f/sub/state?sid={sid}")
        d = r.get_json() or {}
        check("sub/state 200", r.status_code == 200, f"HTTP {r.status_code}")
        check("아직 추출 전", d.get("extracted") is False, str(d.get("extracted")))
        check("종류가 계통도", d.get("kind") == "system", str(d.get("kind")))

        # ── 실도면으로 진짜 뽑아 본다
        dxf = next((p for p in SYSTEM_DXF_CANDIDATES if p.is_file()), None)
        if dxf is None:
            check("계통도 실도면", False,
                  "후보를 못 찾음 — 실측 생략됨(라우트만 확인)")
        else:
            print(f"\n  실도면: {dxf.name}")
            from routes.module_f.subdrawing import (entities_to_world,
                                                    parse_subdrawing)
            ents, parsed = parse_subdrawing(dxf)
            check("도면이 읽힌다", bool(ents), f"entity {len(ents):,}")
            w = entities_to_world(ents)
            check("선분이 나온다", len(w.segs) > 0, f"선분 {len(w.segs):,}")
            from routes.module_f.world import _world_payload
            from routes.module_f.common import _boot
            _boot()
            pay = _world_payload(w)
            b = pay["bounds"]
            check("경계가 유효하다", b["maxx"] > b["minx"],
                  f"x {b['minx']:.0f}~{b['maxx']:.0f} · y {b['miny']:.0f}~{b['maxy']:.0f}")

            # 두 점: 배관 선분의 양 끝에서 가장 멀리 떨어진 두 점을 쓴다.
            xs = [(s[2], s[3]) for s in w.segs]
            pts = [p for pair in xs for p in pair]
            lo = min(pts, key=lambda p: (p[0], p[1]))
            hi = max(pts, key=lambda p: (p[0], p[1]))
            sess["entities"] = ents
            r = c.post("/api/module-f/system/extract",
                       json={"sid": sid,
                             "pump_x": lo[0], "pump_y": lo[1],
                             "av_x": hi[0], "av_y": hi[1],
                             "snap_tolerance_mm": 5000})
            d = r.get_json() or {}
            if r.status_code == 200:
                s2 = d.get("summary") or {}
                check("계통도 경로 추출", True,
                      f"절점 {s2.get('nodes')} · 배관 {s2.get('pipes')}"
                      f" · 연장 {s2.get('total_m')} m")
                check("AV 절점이 10", str(s2.get("av_node_label")) == "10",
                      str(s2.get("av_node_label")))
                r = c.get(f"/api/module-f/sub/state?sid={sid}")
                check("추출 결과가 세션에 남는다",
                      (r.get_json() or {}).get("extracted") is True)
            else:
                # 미도달을 성공으로 위장하지 않는 것이 계약이다 — 그것도 확인.
                check("미도달이면 400 과 사유", r.status_code == 400,
                      f"HTTP {r.status_code} · {d.get('message', '')[:70]}")
                check("깨끗한 배관망을 권한다", d.get("suggest_clean") is True)

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  - " + f)
        return 1
    print("PASS — 계통도·기계실 라우트 전 항목 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
