# -*- coding: utf-8 -*-
"""[C-0단계] 신축배관을 재료에서 빼면 헤드가 떨어지는가 — 지시서 §2 C-6.

「빼면 된다」로 바로 가지 않는다. 신축배관이 **헤드를 잇던 유일한 경로**였다면
그 헤드는 고립된다. 그러니 빼기 전에 잰다 — 코드 수정 없이 숫자만 낸다.

    · 지금 재료로 attachable_heads → wet / dropped
    · 신축배관 레이어를 재료에서 뺀 뒤 한 번 더
    ·   0 이면 그냥 빼면 된다 · 줄면 그 헤드 목록을 낸다

★레이어는 **찍기 단계에만** 산다. 손질판(EditBoard)은 pts/edges 만 들고
  레이어를 안 들고 있다 — 지시서가 「찍기 화면에서 레이어를 빼라」고 한 이유다.

    python scripts/_probe_flex_layer.py
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
# 자동 사전 — **추천만** 한다. 도면마다 관례가 다르므로 반드시 놓친다(지시서 C-6).
FLEX_WORDS = ("후렉시블", "후렉", "플렉", "flex", "fx")


def is_flex(layer: str) -> bool:
    s = str(layer or "").lower()
    return any(w.lower() in s for w in FLEX_WORDS)


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def main() -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    if not PLAN.is_file():
        print(f"표본 없음: {PLAN}")
        return 0
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(PLAN, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
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

        from routes.module_f.jobs import _sess
        from services.cad_import.design.restrict import attachable_heads
        ps = _sess(sid)["pick"]

        print(f"\n■ {PLAN.name}")
        print(f"  찍힌 재료 묶음 {len(ps.board.mat)}")
        by_layer: dict = {}
        for ly, col in ps.board.mat:
            by_layer.setdefault(str(ly), []).append(col)
        for ly in sorted(by_layer):
            n = sum(len(ps.board.by_bundle.get((ly, col), ()))
                    for col in by_layer[ly])
            mark = "  ★신축배관 추천" if is_flex(ly) else ""
            print(f"    {ly:24} 색 {sorted(by_layer[ly])} · 선분 {n}{mark}")
        flex = [ly for ly in by_layer if is_flex(ly)]
        if not flex:
            print("  ★자동 사전이 신축배관을 못 찾았다 — 이 도면에 없거나 이름"
                  " 관례가 다르다(사전은 반드시 놓친다 · S340).")

        anchor = []          # ★두 번 다 **같은 자리**에 찍는다(아래 주석 참조)

        def measure(tag):
            c.post("/api/module-f/pick/commit", json={"sid": sid})
            wait(c, sid)
            # ★급수원이 없으면 물길 판정 자체가 안 선다 — 앵커를 찍고 잰다.
            #   (처음에 이걸 빼고 재서 「헤드 0 · 물닿음 0」이 나왔다.)
            #
            # ★★그리고 앵커는 **두 번 다 같은 좌표**여야 한다. 「가장 큰 묶음의
            #   첫 점」으로 고르면 재료를 뺀 뒤 묶음 순서가 바뀌어 앵커가 옮겨
            #   간다 — 그러면 줄어든 헤드가 «신축배관 탓» 인지 «앵커가 옮겨간
            #   탓» 인지 갈리지 않는다. 실제로 첫 측정에서 그렇게 섞였다
            #   (Z1 248153,-239612 → 261437,-239750).
            st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
            if not (st.get("sources") or ()):
                if not anchor:
                    seg = sorted((g.get("segs") or []
                                  for g in st["body_groups"]),
                                 key=len, reverse=True)[0]
                    anchor.extend([seg[0], seg[1]])
                c.post("/api/module-f/edit/anchor-click",
                       json={"sid": sid, "x": anchor[0], "y": anchor[1]})
                wait(c, sid)
            pay = _sess(sid)["edit"].convert_payload()
            got = attachable_heads(pay)
            print(f"  [{tag}] 간선 {len(pay['edges'])} · 헤드 {len(pay['hcov'])}"
                  f" · 전개가 본 헤드 {got['total']}"
                  f" · 물닿음 {len(got['wet'])} · 떨어짐 {got['dropped']}")
            return got

        base = measure("지금 ")
        if not flex:
            return 0

        drop = [k for k in ps.board.mat if is_flex(k[0])]
        segs = sum(len(ps.board.by_bundle.get(k, ())) for k in drop)
        ps.board.mat = [k for k in ps.board.mat if not is_flex(k[0])]
        print(f"  [빼기] 제외 묶음 {len(drop)} {drop} · 그 선분 {segs}")
        after = measure("뺀 뒤")

        lost = sorted(set(base["wet"]) - set(after["wet"]))
        print(f"  ★물닿음 헤드 {len(base['wet'])} → {len(after['wet'])}"
              f" (줄어든 것 {len(lost)})")
        if lost:
            print(f"    떨어진 헤드 번호: {lost[:20]}"
                  + (" …" if len(lost) > 20 else ""))
            print("    → 2단계에서 최근접 가지관 노드에 붙여야 한다.")
        else:
            print("    → 0 이다. 그냥 빼면 된다(1단계로).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
