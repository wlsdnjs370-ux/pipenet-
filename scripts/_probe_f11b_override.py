# -*- coding: utf-8 -*-
"""[F-11b 수용기준] 직접 입력의 한 바퀴 — 채움 → 재확정 → 지움 → 재확정.

지시서 F-11b 의 수용 기준을 그대로 잰다:

    부속 판정 불가 3  → 화면에서 1건 채움 → 재확정 → 2
                       산출물 meta 에 「직접 입력 — 부속 판정 1」
    지우기 → 재확정 → 미해결 3 복귀

★서버 코드는 이 항목에서 불변이다. 그래서 이 프로브는 화면이 부르는 그 API 를
  화면이 보내는 그 몸통으로 부른다 — 서버에 새 길을 내지 않는다. 화면(JS)이
  실제로 그 몸통을 만드는지는 브라우저 검증이 따로 본다.

    python scripts/_probe_f11b_override.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def wait(c, sid, limit=9000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def build(c, sid):
    """표 확정 — 값이 바뀌는 일이므로 매번 여기를 지난다.

    ★POST 가 거절되면(급수원 미지정 등) 잡이 안 서고 «앞 잡의 done» 이 그대로
      보인다. 그 done 을 성공으로 읽으면 조용히 헛돈다 — 응답부터 본다.
    """
    r = c.post("/api/module-f/design/build", json={"sid": sid})
    d = r.get_json() or {}
    if not d.get("ok"):
        return {"state": "reject", "error": d.get("message") or r.status_code}
    return wait(c, sid)


def stand(c, sid):
    """조립 → 급수 시작 위치 원클릭. 표 확정의 전제다.

    앵커를 아무 데나 찍으면 작은 조각에 걸려 최불리가 «1개» 로 나온다. 사람도
    주배관을 보고 찍으므로 큰 덩이부터 후보를 만든다(F-10g·F-11a 와 같은 방식).
    """
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
        best, bd = None, None
        for hx, hy in heads:
            p = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
            d = ((p[0] - hx) ** 2 + (p[1] - hy) ** 2) ** 0.5
            if bd is None or d < bd:
                best, bd = p, d
        if best is None or bd > 2000.0:
            continue
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": best[0], "y": best[1]})
        wait(c, sid)
        st2 = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if (st2.get("worst") or {}).get("k"):
            return {"state": "done", "worst": st2["worst"]}
    return {"state": "no-anchor"}


def look(c, sid):
    """미해결 개수 · meta 의 「직접 입력」 줄 · 채운 자리 목록."""
    d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
    t = d.get("tables") or {}
    un = t.get("unresolved") or {}
    meta = {k: v for k, v in (t.get("meta") or [])}
    return {
        "kind_items": un.get("kind_items") or [],
        "pairs": un.get("pairs") or [],
        "applied": un.get("applied") or [],
        "meta_kind": meta.get("부속 판정 불가"),
        "meta_ov_kind": meta.get("직접 입력 — 부속 판정"),
        "meta_ov_eq": meta.get("직접 입력 — 등가길이"),
    }


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
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True, "heads": {"conf_min": 0.9}})
        wait(c, sid)
        print(f"\n■ {dxf.name}")
        j = stand(c, sid)
        if j.get("state") != "done":
            print(f"★조립·원클릭 실패 — {j}")
            return 1
        print(f"    최불리 {j['worst']['k']}개 · 최원 {j['worst']['far_m']} m")
        j = build(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {j}")
            return 1

        a = look(c, sid)
        n0 = len(a["kind_items"])
        print(f"    ① 처음        부속 판정 불가 {n0}건 (meta {a['meta_kind']})"
              f" · 직접 입력 {a['meta_ov_kind']}건")
        if not n0:
            print("★채울 자리가 없어 이 도면으로는 못 잰다.")
            return 1

        # ── 고를 수 있는 종류는 서버가 준다(자유 입력 금지).
        kinds = (c.get(f"/api/module-f/design/fitting-override?sid={sid}")
                 .get_json() or {}).get("kinds") or []
        print(f"    고를 수 있는 종류 {len(kinds)}가지 — "
              + " · ".join(k["label"] for k in kinds))

        # ── ② 1건 채운다. 화면이 보내는 그 몸통 그대로.
        it = a["kind_items"][0]
        pick = next((k["value"] for k in kinds if k["value"] == "none"),
                    kinds[0]["value"])
        r = c.post("/api/module-f/design/fitting-override", json={
            "sid": sid,
            "kind": [{"node": str(it["node"]), "pipe": str(it["pipe"]),
                      "kind": pick, "note": "도면에서 직선으로 확인"}],
        }).get_json() or {}
        print(f"    ② 저장        {r.get('counts')} · 재확정 필요 "
              f"{r.get('needs_rebuild')}")
        if not r.get("needs_rebuild"):
            fails.append("저장 응답이 「재확정 필요」를 안 말한다")

        # 재확정 전에는 산출이 아직 옛 값이어야 한다 — 배지의 근거.
        mid = look(c, sid)
        if len(mid["kind_items"]) != n0:
            fails.append("재확정도 안 했는데 산출이 바뀌었다")
        print(f"    ③ 재확정 전   부속 판정 불가 {len(mid['kind_items'])}건"
              f" — 아직 안 들어감 {'[OK]' if len(mid['kind_items']) == n0 else '★'}")

        build(c, sid)
        b = look(c, sid)
        n1 = len(b["kind_items"])
        print(f"    ④ 재확정 후   부속 판정 불가 {n1}건 (meta {b['meta_kind']})"
              f" · 직접 입력 {b['meta_ov_kind']}건 · applied {len(b['applied'])}")
        if n1 != n0 - 1:
            fails.append(f"채운 1건이 안 줄었다 ({n0} → {n1}, {n0 - 1} 이어야)")
        if str(b["meta_ov_kind"]) != "1":
            fails.append(f"meta 「직접 입력 — 부속 판정」이 1 이 아니다 "
                         f"({b['meta_ov_kind']})")
        if len(b["applied"]) != 1:
            fails.append(f"applied 가 1건이 아니다 ({len(b['applied'])})")
        else:
            ap = b["applied"][0]
            if ap.get("note") != "도면에서 확인" and not ap.get("note"):
                fails.append("사유가 안 남았다")
            print(f"       └ 직접 입력 — {ap.get('pipe')} · "
                  f"{ap.get('kind')} · 사유 「{ap.get('note')}」")

        # ── ⑤ 지운다 — 빈 배열이 «지운다» 는 규약.
        r = c.post("/api/module-f/design/fitting-override",
                   json={"sid": sid, "kind": []}).get_json() or {}
        print(f"    ⑤ 지움        {r.get('counts')}")
        build(c, sid)
        d = look(c, sid)
        n2 = len(d["kind_items"])
        print(f"    ⑥ 재확정 후   부속 판정 불가 {n2}건 (meta {d['meta_kind']})"
              f" · 직접 입력 {d['meta_ov_kind']}건")
        if n2 != n0:
            fails.append(f"지운 자리가 안 돌아왔다 ({n2}, {n0} 이어야)")
        if str(d["meta_ov_kind"]) != "0":
            fails.append("지웠는데 meta 에 직접 입력이 남았다")

    print()
    if fails:
        for f in fails:
            print(f"  ★{f}")
        return 1
    print("  [OK] 채움 → 재확정 → 지움 → 재확정 한 바퀴가 돈다 "
          "(막다른 길 없음 · 값은 재확정 때만 바뀐다)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
