# -*- coding: utf-8 -*-
"""한글 토큰이 하류 가드를 통과하는지 실측."""
import importlib.util
import io
import json
import os
import sys

sys.path.insert(0, os.getcwd())
spec = importlib.util.spec_from_file_location("daejo", "대조 서버.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

src = "samples/dwg/계통도_LH_306_배관망추출.dwg"
raw = open(src, "rb").read()

c = mod.app.test_client()
with c.session_transaction() as s:
    s["authed"] = True

r = c.post("/api/remote30/upload",
           data={"dxf_file": (io.BytesIO(raw), "평면도.dwg")},
           content_type="multipart/form-data")
print("upload ->", r.status_code, r.data.decode("utf-8")[:300])
token = json.loads(r.data.decode("utf-8"))["dxf_token"]
print("token:", token)

r2 = c.post("/api/remote30/inspect", data={"dxf_token": token})
body = r2.data.decode("utf-8")
last = [l for l in body.splitlines() if l.strip()][-1] if body.strip() else ""
print("inspect ->", r2.status_code, last[:200])

from core.upload_names import sanitize_upload_name
print("guard roundtrip:", sanitize_upload_name(token) == token)
for bad in ["../../secret.dxf", "..\\win.dxf", "a/b.dxf"]:
    print(f"  traversal {bad!r} -> pass={sanitize_upload_name(bad) == bad}")
