# -*- coding: utf-8 -*-
"""계통도·기계실을 «화면이 보는 대로» 잰다 — 모듈 F 를 통과시켜.

앞선 진단(`_probe_riser_mr_diag.py` · `_probe_riser_layer_role.py`)은 파서를
직접 불러 잰 것이다. 사용자가 겪는 것은 그게 아니라 **화면**이므로, 같은 도면을
모듈 F 의 그 경로로 넣어 무엇이 나오는지 본다:

    · 정찰이 내려주는 띠·규칙 (자동 채택이 무엇을 고르나)
    · 찍기 화면에 서는 «재료 묶음» 목록 — 사람이 손으로 고칠 길이 있나
    · 그중 자동이 추천하는 것(cat == PIPE)과 실제 배관의 차이

★요점은 «자동이 틀렸나» 가 아니라 «틀렸을 때 사람이 고칠 수 있나» 다. 묶음
  목록에 진짜 배관 레이어가 서 있으면 이 도면은 수동으로 완주할 수 있다.
  안 서 있으면 그건 막다른 길이고, 그때는 화면을 고쳐야 한다.

    python scripts/_probe_riser_mf_screen.py [도면.dxf ...]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = [
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 계통도.dxf",
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 기계실.dxf",
]


def wait(c, sid, limit=12000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        for dxf in [Path(x) for x in sys.argv[1:]] or DEF:
            if not dxf.is_file():
                print(f"\n■ {dxf.name} — 파일 없음")
                continue
            print(f"\n{'=' * 88}")
            print(f"■ {dxf.name}")
            print("=" * 88)
            with open(dxf, "rb") as fh:
                r = c.post("/api/module-f/slot/open",
                           data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                           content_type="multipart/form-data")
            jr = r.get_json() or {}
            if "sid" not in jr:
                print(f"   ★열기 실패 — {jr.get('message') or r.status_code}")
                continue
            sid = jr["sid"]
            j = wait(c, sid)
            if j.get("state") == "error":
                print(f"   ★열기 잡 실패 — {j.get('error')}")
                continue

            rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
                   .get("recon") or {})
            ad = rec.get("adopt") or {}
            print(f"\n  정찰  후보 {rec.get('n')} · 띠 {rec.get('bands')}")
            print(f"        규칙 {ad.get('rule')} · 임계 {ad.get('conf_min')} "
                  f"· 채택 예정 {ad.get('n')}")
            print(f"        «{ad.get('why')}»")
            print(f"        추천 묶음 {rec.get('bundles')}")

            # 재료 묶음은 «읽기» 단계를 지나야 선다 — 화면이 하는 그대로.
            c.post("/api/module-f/slot/read",
                   json={"sid": sid, "method": "manual"})
            wait(c, sid)
            w = (c.get(f"/api/module-f/world?sid={sid}").get_json() or {})
            bundles = (w.get("bundles") or (w.get("world") or {}).get("bundles")
                       or [])
            print(f"\n  찍기 화면의 재료 묶음 — {len(bundles)}개 "
                  f"(레이어×색 단위, 사람이 눌러 고르는 그 목록)")
            print(f"    {'레이어':<18}{'색':<9}{'분류':<7}{'연장(m)':>10}"
                  f"{'선분':>8}{'중앙mm':>8}{'원':>6}{'호':>6}")
            # 서버가 이미 «총 연장» 순으로 세워 보낸다 — 그 순서 그대로 본다.
            for b in bundles[:14]:
                print(f"    {str(b.get('layer'))[:17]:<18}"
                      f"{str(b.get('name'))[:8]:<9}{str(b.get('cat')):<7}"
                      f"{float(b.get('len_m') or 0):>10,.1f}"
                      f"{int(b.get('n_all') or 0):>8,}"
                      f"{int(b.get('len_mid') or 0):>8,}"
                      f"{int(b.get('n_circle_all') or 0):>6,}"
                      f"{int(b.get('n_arc_all') or 0):>6,}")
            npipe = sum(1 for b in bundles if b.get("cat") == "PIPE")
            print(f"\n  → 자동이 «배관» 으로 추천하는 묶음 {npipe}개 / 전체 "
                  f"{len(bundles)}개")
            print("     나머지도 목록에 서 있으면 사람이 눌러 고칠 수 있다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
