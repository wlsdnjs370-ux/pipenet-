# -*- coding: utf-8 -*-
"""한글 이름 도면 3장을 모듈 A 순서대로 올렸을 때 저장 경로가 겹치는지 확인한다."""
import importlib.util
import io
import os
import pathlib
import sys

os.environ.setdefault("LOGIN_PASSWORD", "probe-only-0000")
sys.path.insert(0, os.getcwd())

spec = importlib.util.spec_from_file_location("daejo", "대조 서버.py")
m = importlib.util.module_from_spec(spec)
sys.modules["daejo"] = m
spec.loader.exec_module(m)

SRC = pathlib.Path("samples/dwg/계통도_LH_306_배관망추출.dwg")
raw = SRC.read_bytes()

client = m.app.test_client()
with client.session_transaction() as s:
    s["authed"] = True

for name in ("평면도.dwg", "계통도.dwg", "기계실.dwg"):
    resp = client.post(
        "/api/remote30/upload",
        data={"dxf_file": (io.BytesIO(raw), name)},
        content_type="multipart/form-data",
    )
    print(f"{name:12s} -> {resp.status_code}  {resp.get_data(as_text=True)[:120]}")
