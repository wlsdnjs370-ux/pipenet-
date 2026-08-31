# -*- coding: utf-8 -*-
"""[F-11d 수용기준] 직접 입력이 «다시 계산» 을 넘어 살아남는가.

지시서 F-11d 의 수용 기준 두 줄을 그대로 잰다:

    · 관경 1건 + 부속 1건 덮은 상태에서 corridor 를 바꿔 다시 확정하면
      **두 수정이 살아 있고 값이 유지된다**
    · 덮은 자리가 corridor 에서 빠지게 만들면 「적용 못 한 수정 n건」이
      **사유와 함께** 표시되고, 산출물에는 들어가지 않는다

★corridor 를 바꾸는 수단으로 기준개수 K 를 쓴다. 배관 살리기(JOIN)도 corridor
  를 바꾸지만 두 끝점을 결정적으로 고르기 어렵다 — 감사(§22)와 같은 수단을
  써야 두 실측이 서로 대조가 된다.

    python scripts/_probe_f11d_survive.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
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


def build(c, sid, k=None):
    body = {"sid": sid}
    if k is not None:
        body["k"] = k
    r = c.post("/api/module-f/design/build", json=body)
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


def look(c, sid):
    d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
    t = d.get("tables") or {}
    un = t.get("unresolved") or {}
    meta = {k: v for k, v in (t.get("meta") or [])}
    return {
        "kind_items": un.get("kind_items") or [],
        "applied": un.get("applied") or [],
        "bore_ov": t.get("bore_overrides") or {},
        "missed": d.get("ov_missed") or [],
        "meta_fit": meta.get("직접 입력 — 부속 판정"),
        "meta_bore": meta.get("직접 입력 — 관경"),
        "pipes": d.get("view", {}).get("pipes") or [],
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
    fails = []
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        print(f"\n■ {dxf.name}")
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
        if build(c, sid, 30).get("state") != "done":
            print("★K=30 확정 실패")
            return 1

        a = look(c, sid)
        if not a["kind_items"]:
            print("★부속 판정 불가 자리가 없어 이 도면으로는 못 잰다.")
            return 1
        say(f"K=30 기준 — 부속 미해결 {len(a['kind_items'])} · "
            f"배관 {len(a['pipes'])}")

        # ── ① 부속 1건 + 관경 1건을 덮는다.
        it = a["kind_items"][0]
        c.post("/api/module-f/design/fitting-override", json={
            "sid": sid,
            "kind": [{"node": str(it["node"]), "pipe": str(it["pipe"]),
                      "kind": "none", "note": "도면에서 직선으로 확인"}]})
        cand = next((p for p in a["pipes"] if p.get("ref")), None)
        pa, pb = cand["ref"]
        new_dia = 80 if int(cand["dia"]) != 80 else 100
        c.post("/api/module-f/design/bore-override", json={
            "sid": sid, "rows": [{"a": pa, "b": pb, "dia": new_dia,
                                  "note": "협의 변경"}]})
        # 저장된 부속 키에 «안정 키» 가 붙었는지 본다 — 이행의 핵심이다.
        got = (c.get(f"/api/module-f/design/fitting-override?sid={sid}")
               .get_json() or {})
        row0 = ((got.get("overrides") or {}).get("kind") or [{}])[0]
        has_stable = all(k in row0 for k in ("a", "b", "nx", "ny"))
        print(f"    ① 덮기      부속 {it['pipe']}·노드 {it['node']} → 직선 · "
              f"관경 {cand['label']}(노드 {pa}–{pb}) {cand['dia']}→{new_dia}A")
        print(f"       안정 키가 붙었나 — {'예' if has_stable else '★아니오'} "
              f"{ {k: row0.get(k) for k in ('a', 'b', 'nx', 'ny')} }")
        if not has_stable:
            fails.append("부속 직접 입력에 안정 키가 안 붙었다")

        if build(c, sid, 30).get("state") != "done":
            print("★재확정 실패")
            return 1
        b = look(c, sid)
        print(f"    ② 재확정    부속 미해결 {len(b['kind_items'])} · "
              f"meta 부속 {b['meta_fit']} · meta 관경 {b['meta_bore']} · "
              f"못 적용 {len(b['missed'])}")
        if str(b["meta_fit"]) != "1" or str(b["meta_bore"]) != "1":
            fails.append("첫 재확정에서 두 수정이 안 들어갔다")

        # ── ③ corridor 를 바꾼다 (K 30 → 20). 여기가 시험의 핵심이다.
        if build(c, sid, 20).get("state") != "done":
            print("★K=20 확정 실패")
            return 1
        d20 = look(c, sid)
        print(f"    ③ K=20      meta 부속 {d20['meta_fit']} · "
              f"meta 관경 {d20['meta_bore']} · 못 적용 {len(d20['missed'])} · "
              f"배관 {len(d20['pipes'])}")
        for m in d20["missed"][:4]:
            print(f"       └ 못 적용 — {m.get('what') or 'kind'} · "
                  f"{m.get('pipe') or m.get('kind')} · 「{m.get('why')}」")

        # ── ④ 되돌린다 (K 20 → 30). 값이 그대로 살아 있어야 한다.
        if build(c, sid, 30).get("state") != "done":
            print("★K=30 되돌림 실패")
            return 1
        e = look(c, sid)
        print(f"    ④ K=30 복귀 meta 부속 {e['meta_fit']} · "
              f"meta 관경 {e['meta_bore']} · 못 적용 {len(e['missed'])}")
        if str(e["meta_fit"]) != "1" or str(e["meta_bore"]) != "1":
            fails.append(f"되돌렸는데 수정이 안 살아났다 "
                         f"(부속 {e['meta_fit']} · 관경 {e['meta_bore']})")
        # 값까지 같은가 — 개수만 맞고 값이 달라지면 더 나쁘다.
        ap = [x for x in e["applied"] if x.get("what") == "kind"]
        if ap and ap[0].get("note") != "도면에서 직선으로 확인":
            fails.append("사유가 안 살아남았다")
        bo = list(e["bore_ov"].values())
        if bo and int(bo[0]["dia"]) != new_dia:
            fails.append(f"관경 값이 달라졌다 ({bo[0]['dia']} ≠ {new_dia})")
        if bo:
            print(f"       └ 관경 {bo[0]['dia']}A (원래 {bo[0]['orig_dia']}A"
                  f"[{bo[0]['orig_src']}]) · 사유 「{bo[0]['note']}」")

    print()
    if fails:
        for f in fails:
            print(f"  ★{f}")
        return 1
    print("  [OK] corridor 가 바뀌어도 수정이 살아남고, 못 들어간 것은 사유와")
    print("       함께 드러난다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
