# -*- coding: utf-8 -*-
"""[§18] 미해결 «목록» 이 «개수» 와 맞는가 — 이 변경의 유일한 불변식.

엔진이 「부속 판정 불가 n건」·「등가길이 미해결 n건」을 세기만 하고 어느
배관인지 버려서, 사람이 손으로 채울 수가 없었다. 세는 자리에서 목록도 함께
남기도록 고쳤다(순수 추가).

★고친 방식의 값어치는 «목록과 개수가 어긋날 수 없다» 는 데 있다 — 둘이 같은
  자리에서 나오기 때문이다. 그것을 실제 도면으로 확인한다. 어긋나면 F 가
  엉뚱한 배관을 미해결로 찍게 되고, 사람은 멀쩡한 값을 손으로 덮어쓴다.

    python scripts/_probe_unresolved_items.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=6000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEF
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with srv.app.test_client() as c:
        with c.session_transaction() as s_:
            s_["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True, "heads": {"conf_min": 0.9}})
        wait(c, sid)
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        s = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        pts = []
        for g in s["body_groups"]:
            a = g["segs"]
            for i in range(0, len(a) - 3, 4):
                pts.append((float(a[i]), float(a[i + 1])))
        hx, hy = float(s["heads"][0][0]), float(s["heads"][0][1])
        px, py = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": px, "y": py})
        wait(c, sid)
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
        j = wait(c, sid)
        if j["state"] != "done":
            print("표 확정 실패:", j.get("error"))
            return 2

        pv = c.get(f"/api/module-f/design/preview?sid={sid}").get_json()
        tbl = pv["tables"]
        meta = {k: v for k, v in tbl["meta"]}
        un = tbl.get("unresolved") or {}
        n_kind = int(meta.get("부속 판정 불가") or 0)
        n_len = int(meta.get("등가길이 미해결") or 0)
        ki = un.get("kind_items") or []
        li = un.get("length_items") or []
        pairs = un.get("pairs") or []

        print(f"\n■ {dxf.name}")
        print(f"    부속 판정 불가   개수 {n_kind:>4} · 목록 "
              f"{sum(int(x.get('n') or 0) for x in ki):>4}건 "
              f"({len(ki)}자리)")
        print(f"    등가길이 미해결   개수 {n_len:>4} · 목록 {len(li):>4}건 "
              f"· (종류,호칭경) 쌍 {len(pairs)}")
        ok = (sum(int(x.get("n") or 0) for x in ki) == n_kind
              and len(li) == n_len)
        print(f"    {'[OK] 목록과 개수가 같다' if ok else '★어긋난다'}")
        for x in ki[:4]:
            print(f"      부속 · 배관 {x.get('pipe')} · 노드 {x.get('node')}"
                  f" · {x.get('where')}"
                  + (f" · {x.get('angle_deg')}°" if x.get("angle_deg") is not None else ""))
        for x in li[:4]:
            print(f"      등가길이 · 배관 {x.get('pipe')} · {x.get('kind')}"
                  f" · {x.get('dia')}A")
        for p in pairs[:4]:
            print(f"      쌍 · {p.get('kind')} {p.get('dia')}A — {p.get('n')}건")
        if not (ki or li):
            print("      (이 도면엔 미해결이 없다 — 다른 도면으로도 재 볼 것)")
        return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
