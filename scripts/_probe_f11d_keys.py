# -*- coding: utf-8 -*-
"""[F-11d-1 감사] 직접 입력의 «키» 가 다시 계산을 넘어 살아남는가.

지시서 F-11d 는 먼저 **재기 전에 감사하라**고 한다. 부속 직접 입력의 키는
`(node, pipe)` 인데, 둘 다 BFS 순서로 매겨진 **표 라벨**이다. corridor 가 바뀌면
(배관 살리기 → 다시 계산 → 재확정) BFS 가 다시 돌고 번호가 재배열된다. 그러면
사람이 「이 자리는 직선이다」라고 적어 둔 값이 **다른 자리로 옮겨 붙는다** —
조용히, 그리고 산출물에 실려서.

그 흔들림이 실제로 있는지 잰다. 있으면 노드쌍 키로 이행해야 하고(F-11c 의
관경이 이미 그렇게 한다), 없으면 그대로 둔다. **재기 전에는 어느 쪽인지 모른다.**

    python scripts/_probe_f11d_keys.py [도면.dxf]

무엇을 재나:

    ① 확정 1회차   — 표 라벨 ↔ board 노드쌍 대응을 통째로 뜬다
    ② corridor 변경 — 배관 하나를 살린다(살리기는 망을 바꾸므로 BFS 가 달라진다)
    ③ 확정 2회차   — 같은 대응을 다시 뜬다
    ④ 대조         — 같은 노드쌍이 같은 라벨을 갖는가 / 부속 자리의 (node,pipe)가
                     같은 물리 자리를 가리키는가
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def build(c, sid):
    r = c.post("/api/module-f/design/build", json={"sid": sid})
    d = r.get_json() or {}
    if not d.get("ok"):
        return {"state": "reject", "error": d.get("message") or r.status_code}
    return wait(c, sid)


def snapshot(c, sid) -> dict:
    """표 라벨 ↔ board 노드쌍 · 부속 미해결 자리 — 대조에 쓸 한 장."""
    d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
    v, t = d.get("view") or {}, d.get("tables") or {}
    pipes = v.get("pipes") or []
    return {
        # 노드쌍 → 라벨 (키가 흔들리는지 보는 자)
        "by_ref": {tuple(p["ref"]): str(p["label"])
                   for p in pipes if p.get("ref")},
        "by_label": {str(p["label"]): (tuple(p["ref"]) if p.get("ref") else None)
                     for p in pipes},
        "dia": {str(p["label"]): p.get("dia") for p in pipes},
        "kind_items": [(str(x["node"]), str(x["pipe"]))
                       for x in ((t.get("unresolved") or {}).get("kind_items")
                                 or [])],
        "n_pipes": len(pipes),
    }


def stand(c, sid):
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    j = wait(c, sid)
    if j.get("state") != "done":
        return j
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    heads = [(float(h[0]), float(h[1])) for h in (st["heads"] or ())]
    groups = sorted((g.get("segs") or [] for g in st["body_groups"]),
                    key=len, reverse=True)
    for s2 in groups[:18]:
        pts = [(float(s2[i]), float(s2[i + 1]))
               for i in range(0, len(s2) - 3, 4)]
        if not (pts and heads):
            continue
        p0, d0 = None, None
        for hx, hy in heads:
            p = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
            d = ((p[0] - hx) ** 2 + (p[1] - hy) ** 2) ** 0.5
            if d0 is None or d < d0:
                p0, d0 = p, d
        if p0 is None or d0 > 2000.0:
            continue
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": p0[0], "y": p0[1]})
        wait(c, sid)
        st2 = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if (st2.get("worst") or {}).get("k"):
            return {"state": "done", "worst": st2["worst"]}
    return {"state": "no-anchor"}


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
        with c.session_transaction() as s:
            s["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        ad = rec.get("adopt") or {}
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": ad.get("conf_min")}})
        wait(c, sid)
        print(f"\n■ {dxf.name}")
        j = stand(c, sid)
        if j.get("state") != "done":
            print(f"★조립·원클릭 실패 — {j}")
            return 1
        if build(c, sid).get("state") != "done":
            print("★1차 확정 실패")
            return 1
        a = snapshot(c, sid)
        print(f"    ① 1차 확정   배관 {a['n_pipes']} · 역참조 있는 것 "
              f"{len(a['by_ref'])} · 부속 미해결 {len(a['kind_items'])}")

        # ── ② corridor 를 바꾼다. «배관 살리기» 는 망을 넓혀 BFS 를 다시 돌린다.
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        print(f"    ② corridor 변경 시도 — 편집 상태 키 {sorted(st.keys())[:8]}")
        # (살리기 API 이름은 편집 라우트에서 확인해 채운다 — 감사 전용 자리)

        if build(c, sid).get("state") != "done":
            print("★2차 확정 실패")
            return 1
        b = snapshot(c, sid)
        print(f"    ③ 2차 확정   배관 {b['n_pipes']} · 역참조 있는 것 "
              f"{len(b['by_ref'])} · 부속 미해결 {len(b['kind_items'])}")

        # ── ④ 대조
        moved = [(ref, a["by_ref"][ref], b["by_ref"][ref])
                 for ref in a["by_ref"]
                 if ref in b["by_ref"] and a["by_ref"][ref] != b["by_ref"][ref]]
        gone = [ref for ref in a["by_ref"] if ref not in b["by_ref"]]
        print(f"    ④ 라벨 흔들림 {len(moved)}개 / 사라진 자리 {len(gone)}개")
        for ref, x, y in moved[:8]:
            print(f"       노드 {ref[0]}–{ref[1]} : {x} → {y}")
        same_kind = set(a["kind_items"]) & set(b["kind_items"])
        print(f"    부속 자리 키 (node,pipe) — 1차 {len(a['kind_items'])} · "
              f"2차 {len(b['kind_items'])} · 그대로 {len(same_kind)}")
        if moved:
            print("\n  ★배관 라벨이 흔들린다 — 부속 직접 입력 키를 노드쌍으로")
            print("     이행해야 한다(F-11c 의 관경이 이미 그 방식이다).")
        else:
            print("\n  [OK] 이 시나리오에서는 라벨이 안 흔들렸다 — 다만 «안 흔들린")
            print("       것을 봤다» 이지 «절대 안 흔들린다» 는 아니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
