# -*- coding: utf-8 -*-
"""최원 유하거리 «경로» 가 살아 있는 서버에서 실제로 나오는가 (BLOCKED §30 해소).

시험(test_module_f_routes)은 test_client 로 확인한다. 여기서는 **띄워 둔 서버**
에 그대로 물어, 배포된 것이 그 값을 내는지 본다.

저장본 하나를 열어 최불리까지 돌린 뒤 설계 미리보기를 청한다.
"""
from __future__ import annotations

import json
import os
import sys
import time
import urllib.parse
import urllib.request
import http.cookiejar

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(f"{name} — {detail}")


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def main() -> int:
    cj = http.cookiejar.CookieJar()
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    def get(path):
        return json.loads(op.open(BASE + path, timeout=600).read())

    def post(path, body):
        req = urllib.request.Request(
            BASE + path, data=json.dumps(body).encode(),
            headers={"Content-Type": "application/json"})
        try:
            return json.loads(op.open(req, timeout=600).read())
        except urllib.error.HTTPError as exc:
            # 실패도 자료다 — 사유를 삼키면 무엇이 막았는지 못 본다.
            try:
                return json.loads(exc.read())
            except Exception:
                return {"ok": False, "message": f"HTTP {exc.code}"}

    op.open(BASE + "/login",
            urllib.parse.urlencode({"password": _password()}).encode()).read()

    print("[1] 저장본 열기")
    items = (get("/api/module-f/saved") or {}).get("items") or []
    live = [i for i in items if i.get("source_exists")]
    if not live:
        print("  원본이 남은 저장본이 없다 — 검사 불가")
        return 0
    # 찍기만 하고 손질(알람밸브)을 안 한 저장본은 최불리가 안 선다 — 이름이
    # 비슷하다고 아무거나 고르면 «기능이 없다» 가 아니라 «준비가 덜 됐다» 로
    # 실패한다. 알람밸브가 찍혀 있는 저장본을 고른다.
    import glob
    from routes.module_f.common import IMPORT_WORK_ROOT
    ready = set()
    for f in glob.glob(str(IMPORT_WORK_ROOT / "DWG" / "*_유저손질.json")):
        try:
            j = json.load(open(f, encoding="utf-8"))
        except Exception:
            continue
        if (j.get("sources") or j.get("valve_picks")):
            ready.add(os.path.basename(f)[: -len("_유저손질.json")])
    pool = [i for i in live if i["key"] in ready] or live
    key = pool[0]["key"]
    sid = (post("/api/module-f/reopen", {"key": key}) or {}).get("sid")
    check("세션이 섰다", bool(sid), f"{key} · {sid}")
    if not sid:
        return 1

    def wait(limit=1800):
        for _ in range(int(limit / 0.5)):
            j = get(f"/api/module-f/job?sid={sid}")
            if j.get("state") in ("done", "error", "idle"):
                return j
            time.sleep(0.5)
        return {"state": "timeout"}

    st = wait()
    check("손질까지 열렸다", st.get("state") in ("done", "idle"), str(st.get("state")))

    print("[2] 최불리")
    r = post("/api/module-f/edit/worst", {"sid": sid, "k": 30})
    s = (r or {}).get("summary") or {}
    check("최불리가 섰다", bool(s),
          f"K {s.get('k')} · 최원 {s.get('far_m')} m"
          + ("" if s else f" · 서버 응답 {r}"))

    print("[3] 설계 미리보기 — 여기가 오랫동안 죽어 있던 자리")
    d = get(f"/api/module-f/design/preview?sid={sid}")
    if not (d.get("tables") or {}).get("meta"):
        print(f"      (미리보기에 표가 없다 {sorted(d)} — 확정을 먼저 돌린다)")
        rb = post("/api/module-f/design/build", {"sid": sid, "k": 30})
        j = wait()
        print(f"      확정 잡 {j.get('state')} {j.get('error') or ''} · 응답 {str(rb)[:120]}")
        d = get(f"/api/module-f/design/preview?sid={sid}")
        if not (d.get("tables") or {}).get("meta"):
            print(f"      ★여전히 표가 없다 — 응답 {str(d)[:300]}")
    meta = dict((k, v) for k, v in ((d.get("tables") or {}).get("meta") or []))
    lab = meta.get("기준 헤드 노드")
    check("기준 헤드 노드가 '?' 가 아니다", bool(lab) and lab != "?", str(lab))

    v = d.get("view") or {}
    path = v.get("worst_path") or []
    check("경로가 두 절점 이상", len(path) >= 2, f"{len(path)}절점")
    check("경로가 기준 헤드에서 끝난다",
          bool(path) and str(path[-1]) == str(lab),
          f"{path[-1] if path else None} vs {lab}")
    nodes = (d.get("tables") or {}).get("nodes") or []
    root = next((str(n.get("label")) for n in nodes
                 if str(n.get("io_node")) == "Input"), None)
    check("경로가 접속점에서 시작한다",
          bool(path) and str(path[0]) == str(root),
          f"{path[0] if path else None} vs {root}")
    m = float(v.get("worst_path_m") or 0.0)
    far = float(s.get("far_m") or 0.0)
    check("경로 길이가 최원 유하거리와 맞는다",
          m > 0 and abs(m - far) <= max(1.0, far * 0.02),
          f"경로 {m} m · 최원 {far} m")

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("최원 유하거리 경로 — 살아 있는 서버에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
