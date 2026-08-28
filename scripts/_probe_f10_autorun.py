# -*- coding: utf-8 -*-
"""[F-10a] 자동 차선의 «최불리 추출» 이 왜 안 끝나는가 — 잡 오류를 그대로 읽는다.

브라우저 검증에서 배관망 검출까지는 다 통과하는데 `au-run` 뒤 60초 안에
`au-to-design` 이 안 열린다. 화면만 보고는 «느린 것» 인지 «죽은 것» 인지
못 가르므로 서버에서 같은 순서를 태우고 잡 상태를 찍는다.

기본 흐름이 먼저 `/slot/read(manual)` 을 돌린 뒤 고급에서 자동으로 넘어가는
것이 F-10a 의 새 경로다 — 그 순서까지 그대로 재현한다(세션을 갈아타지 않고
방식만 바꾸는 것이 처음 생긴 길이라, 여기가 의심 지점이다).

    python scripts/_probe_f10_autorun.py [도면.dxf] [--fresh]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "samples" / "dxf" / "LH306동_평면도.dxf"


def wait(c, sid, limit=1800, tag=""):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            if j.get("error"):
                print(f"    ★{tag} 오류 — {j['error']}")
                for ln in (j.get("lines") or [])[-8:]:
                    print(f"      {ln}")
            return j
        time.sleep(0.2)
    print(f"    ★{tag} 시간 초과")
    return {"state": "timeout"}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    fresh = "--fresh" in sys.argv
    dxf = Path(args[0]) if args else DEF
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    app = srv.app
    app.config["TESTING"] = True
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    print(f"{dxf.name} · {'자동만' if fresh else '기본 흐름 뒤 자동으로 전환'}\n")
    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {}).get("sid")
        if not sid:
            print("열기 실패:", r.status_code, r.get_json())
            return 1
        wait(c, sid, tag="열기")

        if not fresh:
            # F-10a 의 새 순서 — 기본 흐름이 먼저 manual 로 읽는다.
            rr = c.post("/api/module-f/slot/read",
                        json={"sid": sid, "method": "manual"})
            print(f"  read(manual) {rr.status_code} {rr.get_json()}")
            wait(c, sid, tag="read manual")

        rr = c.post("/api/module-f/slot/read", json={"sid": sid, "method": "auto"})
        print(f"  read(auto)   {rr.status_code} {rr.get_json()}")
        wait(c, sid, tag="read auto")

        st = c.get(f"/api/module-f/auto/state?sid={sid}").get_json()
        print(f"  자동 상태 — method={st.get('method')} opened={st.get('opened')}")

        hs = c.post("/api/module-f/auto/heads", json={"sid": sid}).get_json()
        print(f"  헤드 검출 {hs.get('n')}개")
        if not hs.get("n"):
            print("  ★헤드 0 — 여기서 이미 끝난다")
            return 2
        xs = [h["x"] for h in hs["heads"]]
        ys = [h["y"] for h in hs["heads"]]
        cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
        best = min(hs["heads"], key=lambda h: (h["x"] - cx) ** 2 + (h["y"] - cy) ** 2)
        c.post("/api/module-f/auto/anchor",
               json={"sid": sid, "x": best["x"], "y": best["y"]})

        rn = c.post("/api/module-f/auto/network", json={"sid": sid})
        print(f"  망 검출 요청 {rn.status_code} {rn.get_json()}")
        wait(c, sid, tag="망 검출")

        t0 = time.perf_counter()
        rr = c.post("/api/module-f/auto/run", json={"sid": sid})
        print(f"  최불리 요청  {rr.status_code} {rr.get_json()}")
        j = wait(c, sid, tag="최불리 추출")
        print(f"  최불리 {j['state']} · {time.perf_counter() - t0:.1f}s")
        pv = c.get(f"/api/module-f/auto/preview?sid={sid}")
        print(f"  미리보기 {pv.status_code} "
              f"{str((pv.get_json() or {}).get('summary'))[:120]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
