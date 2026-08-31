# -*- coding: utf-8 -*-
"""[F-11d-1 감사] 직접 입력의 «키» 가 다시 계산을 넘어 살아남는가.

지시서 F-11d 는 **재기 전에 감사하라**고 한다. 부속·등가길이 직접 입력의 키는
`(node, pipe)` 인데 둘 다 **BFS 순서로 매겨진 표 라벨**이다:

    label_of[nid] = str(i)   for i, nid in enumerate(bfs_order(net, root))
    pipe_row(pid) → "label": pid

corridor 가 바뀌면(기준개수 변경·배관 살리기·앵커 이동) 제한 전개 결과가 달라져
BFS 가 다시 돌고 번호가 재배열된다. 그러면 사람이 「이 자리는 직선이다」라고
적어 둔 값이 **다른 자리로 옮겨 붙는다** — 조용히, 그리고 산출물에 실려서.

★F-11c 의 관경은 이미 board 노드쌍을 키로 쓴다(D-F11-4). 부속도 그래야 하는지를
  여기서 «재고» 정한다. 흔들리지 않으면 그대로 두는 것이 맞다 — 안 흔들리는 것을
  고치면 구키 호환이라는 빚만 새로 진다.

무엇을 하나:

    ① K=30 으로 확정 → 표 라벨 ↔ board 노드쌍 대응을 통째로 뜬다
    ② K=20 으로 다시 확정 (corridor 가 실제로 바뀐다)
    ③ 같은 노드쌍이 같은 라벨을 갖는가 / 부속 자리 (node,pipe) 가 같은 곳인가
    ④ K=30 으로 되돌려 ① 과 같아지는가 (되돌아오면 «순서 의존» 은 아니다)

    python scripts/_probe_f11d_keys.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# 대명동 — 확정이 몇 초라 여러 번 돌릴 수 있다. B1F 는 1회 23분이라(BLOCKED §21)
# 감사에는 안 쓴다. 감사가 묻는 것은 «라벨이 흔들리나» 이지 도면 규모가 아니다.
DEF = ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"

_T0 = time.time()


def say(msg: str) -> None:
    print(f"  [{time.time() - _T0:6.1f}s] {msg}", flush=True)


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def build(c, sid, k):
    r = c.post("/api/module-f/design/build", json={"sid": sid, "k": k})
    d = r.get_json() or {}
    if not d.get("ok"):
        return {"state": "reject", "error": d.get("message") or r.status_code}
    j = wait(c, sid)
    if j.get("state") == "done":
        res = (c.get(f"/api/module-f/convert/result?sid={sid}").get_json()
               or {}).get("result") or {}
        if res.get("error"):
            j = dict(j, state="result-error", error=res["error"])
    return j


def snapshot(c, sid) -> dict:
    """표 라벨 ↔ board 노드쌍 · 부속 미해결 자리 — 대조에 쓸 한 장."""
    d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
    v, t = (d.get("view") or {}), (d.get("tables") or {})
    pipes = v.get("pipes") or []
    un = t.get("unresolved") or {}
    ref_of = {str(p["label"]): tuple(p["ref"])
              for p in pipes if p.get("ref")}
    items = un.get("kind_items") or []
    return {
        # 노드쌍 → 표 라벨. 이 대응이 흔들리는지가 감사의 요점이다.
        "by_ref": {tuple(p["ref"]): str(p["label"])
                   for p in pipes if p.get("ref")},
        "dia": {str(p["label"]): p.get("dia") for p in pipes},
        # ① 부속 직접 입력이 «지금» 쓰는 키.
        "kind_keys": {(str(x["node"]), str(x["pipe"])) for x in items},
        # ② 같은 자리를 **board 노드쌍 + 어느 끝** 으로 가리키면 어떤가.
        #    F-11c 의 관경이 이미 이 방식이다(D-F11-4). 이행할지 말지는
        #    «이 키가 안 흔들리는가» 로 정해야 한다 — 안 재고 옮기면 빚만 진다.
        "kind_by_ref": {(ref_of.get(str(x["pipe"])), str(x.get("where") or ""))
                        for x in items if ref_of.get(str(x["pipe"]))},
        # 역참조가 없는 자리는 이 키로 못 가리킨다 — 조용히 빼지 않고 센다.
        "kind_no_ref": sum(1 for x in items if not ref_of.get(str(x["pipe"]))),
        "pairs": {(str(p["kind"]), int(p["dia"]))
                  for p in (un.get("pairs") or [])},
        "n_pipes": len(pipes),
        "k": len(((d.get("view") or {}).get("nodes") or [])),
    }


def stand(c, sid):
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    j = wait(c, sid)
    if j.get("state") != "done":
        return j
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    heads = [(float(h[0]), float(h[1])) for h in (st["heads"] or ())]
    hs = heads[::max(1, len(heads) // 60)][:60]
    groups = sorted((g.get("segs") or [] for g in st["body_groups"]),
                    key=len, reverse=True)
    for s2 in groups[:18]:
        pts = [(float(s2[i]), float(s2[i + 1]))
               for i in range(0, len(s2) - 3, 4)]
        if len(pts) > 4000:
            pts = pts[::len(pts) // 4000]
        if not (pts and hs):
            continue
        p0, d0 = None, None
        for hx, hy in hs:
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


def compare(a, b, tag):
    """두 장을 맞대 «같은 자리가 같은 이름을 갖는가» 를 센다."""
    common = set(a["by_ref"]) & set(b["by_ref"])
    moved = [(r, a["by_ref"][r], b["by_ref"][r]) for r in common
             if a["by_ref"][r] != b["by_ref"][r]]
    gone = set(a["by_ref"]) - set(b["by_ref"])
    new = set(b["by_ref"]) - set(a["by_ref"])
    print(f"\n  ── {tag}")
    print(f"     배관 {a['n_pipes']} → {b['n_pipes']} · 공통 자리 {len(common)}")
    print(f"     ★라벨이 옮겨간 자리 {len(moved)} / 사라진 {len(gone)} · "
          f"새로 생긴 {len(new)}")
    for r, x, y in moved[:8]:
        print(f"       노드 {r[0]}–{r[1]} : {x} → {y}")
    # 부속 키는 (node, pipe) 쌍이다 — 그 쌍이 그대로 살아남나.
    ka, kb = a["kind_keys"], b["kind_keys"]
    print(f"     부속 자리 — 지금 키 (node,pipe) {len(ka)} → {len(kb)} · "
          f"그대로 {len(ka & kb)}")
    for k in sorted(ka - kb)[:4]:
        print(f"       사라진 키 (노드 {k[0]}, 배관 {k[1]})")
    ra, rb = a["kind_by_ref"], b["kind_by_ref"]
    print(f"     부속 자리 — 노드쌍 키 (ref,where) {len(ra)} → {len(rb)} · "
          f"그대로 {len(ra & rb)}"
          + (f"  [역참조 없어 못 가리킴 {a['kind_no_ref']}→{b['kind_no_ref']}]"
             if (a["kind_no_ref"] or b["kind_no_ref"]) else ""))
    print(f"     등가길이 쌍 {len(a['pairs'])} → {len(b['pairs'])} · "
          f"그대로 {len(a['pairs'] & b['pairs'])}")
    return moved, gone


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
        print(f"\n■ {dxf.name} ({dxf.stat().st_size / 1048576:.0f} MB)")
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        ad = rec.get("adopt") or {}
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": ad.get("conf_min")}})
        wait(c, sid)
        j = stand(c, sid)
        if j.get("state") != "done":
            print(f"★조립·원클릭 실패 — {j}")
            return 1
        say(f"조립·원클릭 — 최불리 {j['worst']['k']}개 · "
            f"최원 {j['worst']['far_m']} m")

        snaps = {}
        for k in (30, 20, 30):
            jb = build(c, sid, k)
            if jb.get("state") != "done":
                print(f"★K={k} 확정 실패 — {jb.get('state')} {jb.get('error')}")
                return 1
            s = snapshot(c, sid)
            say(f"K={k} 확정 — 배관 {s['n_pipes']} · 노드 {s['k']} · "
                f"부속 미해결 {len(s['kind_keys'])}")
            snaps.setdefault(k, []).append(s)

        a, b = snaps[30][0], snaps[20][0]
        moved1, _ = compare(a, b, "K 30 → 20 (corridor 가 바뀐다)")
        moved2, _ = compare(a, snaps[30][1], "K 30 → 20 → 30 (되돌린다)")

        print("\n" + "=" * 72)
        if moved1:
            print("  ★배관 라벨이 흔들린다 — 부속 직접 입력 키 (node, pipe) 는")
            print("    corridor 가 바뀌면 다른 자리를 가리킨다. 노드쌍 키로")
            print("    이행해야 한다(F-11c 의 관경이 이미 그 방식이다).")
        else:
            print("  [OK] 이 시나리오에서 라벨은 안 흔들렸다.")
        if moved2:
            print("  ★되돌려도 원래 이름으로 안 돌아온다 — 순서 의존이 있다.")
        else:
            print("  [OK] 되돌리면 원래 이름으로 돌아온다.")
        print("  («안 흔들린 것을 봤다» 이지 «절대 안 흔들린다» 는 아니다 —")
        print("   배관 살리기·앵커 이동은 이 프로브가 아직 안 건드렸다.)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
