# -*- coding: utf-8 -*-
"""[§27 후속] 두 점 미리보기가 «조용히» 안 그려지는 자리를 짚는다.

화면의 `subPath()` 는 `forced_penalty_mm` 이 없으면 **null 을 돌려주고 미리보기를
접는다**(어긋난 미리보기를 그리느니 안 그리는 편이 낫다는 규약). 그래서 그 값이
빠지면 증상은 「선이 안 따라온다」 하나로만 보이고 오류는 안 난다.

살아 있는 서버에 실제로 물어, 그 값이 오는지·경로가 풀리는지 본다.
"""
from __future__ import annotations

import io
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import http.cookiejar

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in io.open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def main() -> int:
    cj = http.cookiejar.CookieJar()
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    def post(p, body):
        req = urllib.request.Request(
            BASE + p, data=json.dumps(body).encode(),
            headers={"Content-Type": "application/json"})
        try:
            return json.loads(op.open(req, timeout=900).read())
        except urllib.error.HTTPError as exc:
            try:
                return json.loads(exc.read())
            except Exception:
                return {"ok": False, "message": f"HTTP {exc.code}"}

    def get(p):
        return json.loads(op.open(BASE + p, timeout=900).read())

    op.open(BASE + "/login",
            urllib.parse.urlencode({"password": _password()}).encode()).read()

    dxf = os.path.join(_ROOT, "routes", "제출용[최종]",
                       "1. 입력도면 대명동 단위세대 계통도.dxf")
    if not os.path.isfile(dxf):
        print("표본 도면 없음")
        return 0

    # 계통도 슬롯으로 올린다 — 멀티파트를 직접 만든다.
    bound = "----mfprobe"
    with open(dxf, "rb") as f:
        blob = f.read()
    parts = []
    for k, v in (("kind", "system"),):
        parts.append(f"--{bound}\r\nContent-Disposition: form-data; "
                     f'name="{k}"\r\n\r\n{v}\r\n'.encode())
    parts.append((f"--{bound}\r\nContent-Disposition: form-data; "
                  f'name="dxf_file"; filename="{os.path.basename(dxf)}"\r\n'
                  "Content-Type: application/octet-stream\r\n\r\n").encode())
    parts.append(blob + b"\r\n")
    parts.append(f"--{bound}--\r\n".encode())
    req = urllib.request.Request(
        BASE + "/api/module-f/slot/open", data=b"".join(parts),
        headers={"Content-Type": f"multipart/form-data; boundary={bound}"})
    try:
        r = json.loads(op.open(req, timeout=900).read())
    except urllib.error.HTTPError as exc:
        r = {"ok": False, "message": exc.read()[:200].decode("utf-8", "replace")}
    sid = r.get("sid")
    print(f"[1] 계통도 슬롯 열기 — sid={sid} · {r.get('message') or ''}")
    if not sid:
        print(f"    응답: {str(r)[:300]}")
        return 1
    for _ in range(3000):
        j = get(f"/api/module-f/job?sid={sid}")
        if j.get("state") in ("done", "error", "idle"):
            break
        time.sleep(0.5)

    g = post("/api/module-f/sub/graph", {"sid": sid})
    print(f"[2] 경로 그래프 — ok={g.get('ok')} · 절점 {len(g.get('nodes') or [])}"
          f" · 간선 {len(g.get('edges') or [])}"
          f" · 조각 {g.get('components')} · 추측연결 {g.get('forced')}")
    pen = g.get("forced_penalty_mm")
    print(f"[3] ★미리보기에 필요한 forced_penalty_mm = {pen!r}"
          + ("   ← 없으면 화면이 선을 «조용히» 안 그린다" if pen is None else ""))
    keys = sorted(k for k in g if k not in ("nodes", "edges", "layers"))
    print(f"    그 밖의 키: {keys}")
    return 0 if pen is not None else 1


if __name__ == "__main__":
    raise SystemExit(main())
