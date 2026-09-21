# -*- coding: utf-8 -*-
"""모듈 D 출력 워크벤치 — 테스트 클라이언트 전 경로 프로브."""
from __future__ import annotations

import importlib.util
import io
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("daejo", ROOT / "대조 서버.py")
srv = importlib.util.module_from_spec(spec)
spec.loader.exec_module(srv)

SUB = ROOT / "routes" / "제출용[최종]"
SDF, XML = SUB / "2. Pipenet_auto.sdf", SUB / "2. Pipenet_auto.xml"

client = srv.app.test_client()
with client.session_transaction() as sess:
    sess["authed"] = True

page = client.get("/module-d-output")
print("page", page.status_code, len(page.data), "bytes",
      "cache=", page.headers.get("Cache-Control"))

res = client.post("/api/module-d/load", data={
    "sdf_file": (io.BytesIO(SDF.read_bytes()), "auto.sdf"),
    "xml_file": (io.BytesIO(XML.read_bytes()), "auto.xml"),
}, content_type="multipart/form-data")
load = res.get_json()
print("load", res.status_code, load.get("ok"), load.get("message", ""))
print("  stem", load["stem"], "counts", load["counts"], "has_results", load["has_results"])
print("  link_items", len(load["link_items"]), "node_items", len(load["node_items"]),
      "presets", list(load["presets"]))
print("  saved", load["saved"])
print("  join matched", load["join"]["matched"],
      {k: len(v) for k, v in load["join"].items() if isinstance(v, list) and v})
print("  notes", load["notes"])
job = load["job_id"]

for link, node in (("Pipe volumetric flow", "None"),
                   ("Pipe pressure difference", "Node pressure"),
                   ("Pipe bore", "Node elevation")):
    res = client.post("/api/module-d/draw", json={
        "job_id": job, "link_item": link, "node_item": node,
        "show_labels": True, "show_arrows": True, "section": "프로브"})
    d = res.get_json()
    print(f"draw {link[:26]:<26}", res.status_code, d.get("ok"), d.get("message", ""),
          d.get("seconds"), d.get("orientation"), d.get("drawn"),
          "blank", len(d.get("blank_links", [])) + len(d.get("blank_nodes", [])),
          "png", d.get("png"))
    last_iso, last_png = d.get("pdf"), d.get("png")

bad_item = client.post("/api/module-d/draw", json={
    "job_id": job, "link_item": "Pipe bore; DROP", "node_item": "None"})
print("모르는 표시 항목 →", bad_item.status_code, bad_item.get_json())

res = client.post("/api/module-d/report", json={"job_id": job})
r = res.get_json()
print("report", res.status_code, r.get("ok"), r.get("message", ""),
      r.get("pages"), "쪽", r.get("rows"), "행", r.get("font_pt"), "pt",
      "vel", len(r.get("velocity_flags", [])), "noz", len(r.get("nozzle_flags", [])))

res = client.post("/api/module-d/book", json={"job_id": job})
b = res.get_json()
print("book", res.status_code, b.get("ok"), b.get("message", ""))
print("  pdf", b.get("pdf"))
print("  line", b.get("line"))
for p in b.get("parts", []):
    print("   ", p["ok"], p["name"], p["pages"], "쪽", round(p["seconds"], 2), "s", p["error"])

for name in (last_iso, last_png, r.get("pdf"), b.get("pdf")):
    got = client.get(f"/api/module-d/result/{job}/{name}")
    print("fetch", name, got.status_code, len(got.data), "bytes",
          got.headers.get("Content-Type"))

bad = client.post("/api/module-d/draw", json={"job_id": "deadbeef", "link_item": "Pipe bore"})
print("expired job →", bad.status_code, bad.get_json())

# 결과 XML 없이 — 계산서는 거절, 도면은 나온다.
res = client.post("/api/module-d/load", data={
    "sdf_file": (io.BytesIO(SDF.read_bytes()), "auto.sdf"),
}, content_type="multipart/form-data")
noxml = res.get_json()
print("load(no xml)", res.status_code, noxml.get("ok"), noxml.get("has_results"),
      noxml.get("notes"))
res = client.post("/api/module-d/report", json={"job_id": noxml["job_id"]})
print("report(no xml) →", res.status_code, res.get_json())
res = client.post("/api/module-d/draw", json={
    "job_id": noxml["job_id"], "link_item": "Pipe bore", "node_item": "None"})
d = res.get_json()
print("draw(no xml)", res.status_code, d.get("ok"), d.get("drawn"))

traverse = client.get(f"/api/module-d/result/{job}/../../../%EB%8C%80%EC%A1%B0%20%EC%84%9C%EB%B2%84.py")
print("traversal →", traverse.status_code)
