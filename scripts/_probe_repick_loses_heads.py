# -*- coding: utf-8 -*-
"""[찍기] 배관을 «하나 더» 찍으면 헤드 픽이 어떻게 되는가.

사용자 지적: 「02.찍기에서 수동으로 정의한 배관이랑 헤드를 04.수리계산에서
반영을 못한다」.

    board.py:243   cleared = self.clear_heads() if self.heads else 0

배관 클릭 한 번이 **찍은 헤드를 전부 해제한다.** 엔진은 그 수를 보고로 돌려
주는데(`헤드해제`), 화면 JS 에 그 이름은 **한 번도 나오지 않는다.**

여기서는 사람이 하는 그대로 밟아 «몇 개가 사라지는지» 센다::

    ⑴ 배관 9묶음 · 헤드 3묶음 찍기          → 스펙 헤드 N
    ⑵ 배관선택으로 돌아가 **한 묶음 더** 찍기 → 스펙 헤드 ?
    ⑶ 그대로 다음(commit)                    → 손질 헤드 ?

    python scripts/_probe_repick_loses_heads.py
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=40000):
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
        wait(c, sid)

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        ps = sess["pick"]
        bd = ps.board

        # ⑴ 추천대로 배관·헤드를 찍는다 (사람이 「자동 채택」을 누른 그 상태)
        c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
        c.post("/api/module-f/pick/mode", json={"sid": sid,
                                                "action": "complete"})
        c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "HEAD"})
        n1 = len(ps.spec().get("heads") or ())
        m1 = len(ps.spec().get("material_picks") or ())
        print(f"\n■ ⑴ 찍은 뒤 — 재료 {m1}묶음 · 헤드 {n1}묶음")

        # ⑵ 「배관선택」으로 돌아가 **배관 한 묶음을 더** 찍는다.
        #    (자동이 놓친 레이어를 사람이 손으로 보태는, 가장 흔한 동작이다)
        c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
        have = set(bd.mat)
        extra = None
        # ★사람이 실제로 보태는 것은 «자동이 놓친 배관 한 줄» 이다 — 벽체
        #   레이어(2만 선분)를 배관이라 찍는 일은 없다. 작은 묶음부터 고른다.
        for key, segs in sorted(bd.by_bundle.items(), key=lambda kv: len(kv[1])):
            if key not in have and 5 <= len(segs) <= 400:
                extra = (key, segs)
                break
        if extra is None:
            print("★더 찍을 묶음이 없습니다 — 이 도면으로는 못 잽니다.")
            return 0
        (ly, col), segs = extra
        a, b = segs[len(segs) // 2]
        rep = (c.post("/api/module-f/pick/click",
                      json={"sid": sid, "x": (a[0] + b[0]) / 2,
                            "y": (a[1] + b[1]) / 2}).get_json()
               or {}).get("report") or {}
        n2 = len(ps.spec().get("heads") or ())
        m2 = len(ps.spec().get("material_picks") or ())
        print(f"\n■ ⑵ 배관 «{ly}×{col}» 을 하나 더 찍었다")
        print(f"      보고 — {rep}")
        print(f"      재료 {m1} → {m2}묶음 · 헤드 {n1} → **{n2}**묶음")
        if n2 < n1:
            print(f"      ★★헤드 픽 {n1 - n2}묶음이 **조용히** 사라졌다"
                  f" — 보고의 «헤드해제»={rep.get('헤드해제')} 를"
                  f" 화면이 읽지 않는다.")

        # ⑶ 그대로 다음으로 넘어간다 — 사람은 사라진 줄 모른다.
        st = (c.get(f"/api/module-f/pick/state?sid={sid}").get_json()
              or {}).get("state") or {}
        print(f"\n■ ⑶ 화면 상태 — mat_done={st.get('mat_done')}"
              f" · 화면이 말하는 헤드 {st.get('heads')}")
        c.post("/api/module-f/pick/mode", json={"sid": sid,
                                                "action": "complete"})
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★손질 실패 — {str(j)[:200]}")
            return 1
        b2 = sess["edit"].board
        print(f"      손질판 — 절점 {len(b2.pts)} · 간선 {len(b2.edges)}"
              f" · **헤드 {len(b2.disks)}**")

        # ⑷ 끝까지 — 수리계산 표에 노즐이 오는가
        st2 = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st2.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st2["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 12})
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 12})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"\n■ ⑷ 표 — ★확정 실패 {str(j)[:200]}")
            return 1
        if "design" not in sess:
            print("\n■ ⑷ 표 — ★안 만들어졌습니다."
                  " 보탠 재료가 «문양 지문» 을 바꿔 헤드가 다른 원들에"
                  " 붙었습니다(물닿음 0) — 재료 픽과 헤드 판정은 붙어 있다.")
            return 1
        tbl = sess["design"]["tables"]
        print(f"\n■ ⑷ 표 — 절점 {len(tbl.nodes)} · 배관 {len(tbl.pipes)}"
              f" · **노즐 {len(tbl.nozzles)}**")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
