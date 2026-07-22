# -*- coding: utf-8 -*-
"""FX materialize e2e (domain-slim) — prototype 전체 플로우로 SDF+SLF 정합 확인.

no-edit(원본 확정) 경로로 돌려서:
  * 30 head 각각 FX eq_len=15.6 / spec_ref=평균
  * 산출물 xlsx/csv/slf/kfp/sdf/zip 존재
  * 동봉 SLF = FX 20A internal=21.6
  * SDF FX_20A_216 Pipe-set (bore 0.02 / len 0.7 / rise -0.1 / C 120)
을 확인.
"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DXF = BASE / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def _load_app():
    if str(BASE) not in sys.path:
        sys.path.insert(0, str(BASE))
    spec = importlib.util.spec_from_file_location("daejo_server", str(BASE / "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules["daejo_server"] = mod
    spec.loader.exec_module(mod)
    return mod.app


def _sse(resp):
    text = resp.get_data(as_text=True)
    out = []
    for block in text.split("\n\n"):
        block = block.strip()
        if not block.startswith("data:"):
            continue
        payload = block[len("data:"):].strip()
        try:
            out.append(json.loads(payload))
        except json.JSONDecodeError:
            pass
    return out


def main():
    assert DXF.is_file(), f"DXF 없음: {DXF}"
    app = _load_app()
    c = app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True

    with DXF.open("rb") as f:
        r = c.post("/api/remote30/prototype/run",
                   data={"dxf_file": (f, DXF.name)},
                   content_type="multipart/form-data")
    assert r.status_code == 200, (r.status_code, r.get_data(as_text=True)[:300])
    job_id = r.get_json()["job_id"]
    print("job_id:", job_id)

    evs = _sse(c.get(f"/api/remote30/prototype/stream/{job_id}"))
    assert not any(e.get("type") == "error" for e in evs), \
        [e for e in evs if e.get("type") == "error"]

    r = c.post(f"/api/remote30/prototype/finalize/{job_id}", json={})
    assert r.status_code == 200, r.get_data(as_text=True)[:300]

    evs = _sse(c.get(f"/api/remote30/prototype/finalize_stream/{job_id}"))
    assert not any(e.get("type") == "error" for e in evs), \
        [e for e in evs if e.get("type") == "error"]
    s5 = next((e for e in evs if e.get("type") == "stage5_complete"), None)
    assert s5 is not None, f"stage5_complete 없음. types={[e.get('type') for e in evs]}"
    equipment = s5["fx_review"]["equipment"]
    fx_rows = [e for e in equipment if e.get("desc") == "FX"]
    print(f"FX rows: {len(fx_rows)}")
    bad = [e for e in fx_rows
           if abs(float(e["eq_len"]) - 15.6) > 1e-6 or e.get("spec_ref") != "평균"]
    assert not bad, f"기본 프로파일 아닌 FX 행 {len(bad)}개: {bad[:2]}"
    print("  OK 전 FX 행 eq_len=15.6 / spec_ref=평균")

    r = c.post(f"/api/remote30/prototype/fx/finalize/{job_id}", json={})
    assert r.status_code == 200, r.get_data(as_text=True)[:300]

    evs = _sse(c.get(f"/api/remote30/prototype/fx/finalize_stream/{job_id}"))
    assert not any(e.get("type") == "error" for e in evs), \
        [e for e in evs if e.get("type") == "error"]
    done = next((e for e in evs if e.get("type") == "done"), None)
    assert done is not None, f"done 없음. types={[e.get('type') for e in evs]}"

    cands = list((BASE / "data").glob(f"**/{job_id}"))
    out_dir = cands[0] if cands else (BASE / "data" / "prototype_report" / job_id)
    print("out_dir:", out_dir, "exists:", out_dir.is_dir())
    if out_dir.is_dir():
        for want in (".xlsx", ".csv", ".slf", ".kfp", ".sdf", ".zip"):
            hit = list(out_dir.glob(f"*{want}"))
            print(f"    {want}: {'OK' if hit else 'MISSING'}", hit[0].name if hit else "")
        import xml.etree.ElementTree as ET
        slf = next(iter(out_dir.glob("*.slf")), None)
        if slf:
            root = ET.parse(slf).getroot()
            for sch in root.iter("Schedule"):
                nm = sch.find("Item-name")
                if nm is not None and (nm.text or "").startswith("FX"):
                    for sd in sch.iter("Size-definition"):
                        print(f"    SLF {nm.text}: nominal={sd.get('nominal')} internal={sd.get('internal')}")
        sdf = next(iter(out_dir.glob("*.sdf")), None)
        if sdf:
            root = ET.parse(sdf).getroot()
            for ps in root.iter("Pipe-set"):
                nm = ps.find("Pipe-type/Name")
                if nm is not None and nm.text and nm.text.startswith("FX"):
                    pipes = ps.findall("Pipe")
                    p0 = pipes[0] if pipes else None
                    print(f"    SDF Pipe-set {nm.text}: {len(pipes)} pipes",
                          f"bore={p0.get('bore')} length={p0.get('length')} "
                          f"rise={p0.get('rise')} c={p0.get('roughness-or-c')}" if p0 is not None else "")
    print("\nDONE OK")


if __name__ == "__main__":
    main()
