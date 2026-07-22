# -*- coding: utf-8 -*-
"""FX Task7 수용검증 — prototype 전체 플로우(0-2 → 3-5 → fx emit) 를 test client 로 구동.

no-edit(원본 확정) 경로로 돌려서:
  * 30 head 각각 FX eq_len=15.6 / spec_ref=평균
  * 산출물 xlsx/csv/slf/kfp/zip 존재
  * 동봉 SLF = FX_20A_216 (FX 20A internal=21.6)
을 확인. 편집(override) 반영은 별도 검증.
"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DXF = BASE / "static" / "대명동201동 단위세대_layer정리.dxf"


def _load_app():
    spec = importlib.util.spec_from_file_location("daejo_server", str(BASE / "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules["daejo_server"] = mod
    spec.loader.exec_module(mod)
    return mod.app


def _sse(resp):
    """finite SSE 응답 → event dict 리스트."""
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

    # 1) run
    with DXF.open("rb") as f:
        r = c.post("/api/remote30/prototype/run",
                   data={"dxf_file": (f, DXF.name)},
                   content_type="multipart/form-data")
    assert r.status_code == 200, (r.status_code, r.get_data(as_text=True)[:300])
    job_id = r.get_json()["job_id"]
    print("job_id:", job_id)

    # 2) stream 0-2
    evs = _sse(c.get(f"/api/remote30/prototype/stream/{job_id}"))
    types = [e.get("type") for e in evs]
    print("stage0-2 event types:", types[:4], "...", types[-3:], "count", len(evs))
    assert not any(e.get("type") == "error" for e in evs), \
        [e for e in evs if e.get("type") == "error"]

    # 3) finalize (no edits)
    r = c.post(f"/api/remote30/prototype/finalize/{job_id}", json={})
    assert r.status_code == 200, r.get_data(as_text=True)[:300]

    # 4) finalize_stream 3-5 → stage5_complete
    evs = _sse(c.get(f"/api/remote30/prototype/finalize_stream/{job_id}"))
    assert not any(e.get("type") == "error" for e in evs), \
        [e for e in evs if e.get("type") == "error"]
    s5 = next((e for e in evs if e.get("type") == "stage5_complete"), None)
    assert s5 is not None, f"stage5_complete 없음. types={[e.get('type') for e in evs]}"
    equipment = s5["fx_review"]["equipment"]
    fx_rows = [e for e in equipment if e.get("desc") == "FX"]
    av_rows = [e for e in equipment if e.get("desc") in ("AV", "A/V")]
    print(f"FX rows: {len(fx_rows)}  AV rows: {len(av_rows)}")
    bad = [e for e in fx_rows
           if abs(float(e["eq_len"]) - 15.6) > 1e-6 or e.get("spec_ref") != "평균"]
    assert not bad, f"기본 프로파일 아닌 FX 행 {len(bad)}개: {bad[:2]}"
    print(f"  ✓ 전 FX 행 eq_len=15.6 / spec_ref=평균")

    # 5) fx/finalize (원본 확정 = 편집 없음)
    r = c.post(f"/api/remote30/prototype/fx/finalize/{job_id}", json={})
    assert r.status_code == 200, r.get_data(as_text=True)[:300]
    print("fx/finalize mode:", r.get_json().get("mode"))

    # 6) fx/finalize_stream stage6 → done
    evs = _sse(c.get(f"/api/remote30/prototype/fx/finalize_stream/{job_id}"))
    assert not any(e.get("type") == "error" for e in evs), \
        [e for e in evs if e.get("type") == "error"]
    done = next((e for e in evs if e.get("type") == "done"), None)
    assert done is not None, f"done 없음. types={[e.get('type') for e in evs]}"
    outputs = done.get("outputs", {})
    print("outputs keys:", sorted(outputs.keys()))

    # 산출물 파일 존재 확인
    out_dir = BASE / "data" / "prototype_report" / job_id
    if not out_dir.is_dir():
        # fallback: PROTOTYPE_OUTPUT_DIR 는 앱 상수 — job 폴더 탐색
        cands = list((BASE / "data").glob(f"**/{job_id}"))
        out_dir = cands[0] if cands else out_dir
    print("out_dir:", out_dir, "exists:", out_dir.is_dir())
    if out_dir.is_dir():
        exts = sorted({p.suffix.lower() for p in out_dir.iterdir() if p.is_file()})
        print("  files ext:", exts)
        for want in (".xlsx", ".csv", ".slf", ".kfp", ".sdf", ".zip"):
            hit = list(out_dir.glob(f"*{want}"))
            print(f"    {want}: {'OK' if hit else 'MISSING'}", hit[0].name if hit else "")
        # 동봉 SLF 가 FX_20A_216 인지 (FX 20A internal=21.6)
        slf = next(iter(out_dir.glob("*.slf")), None)
        if slf:
            import xml.etree.ElementTree as ET
            root = ET.parse(slf).getroot()
            for sch in root.iter("Schedule"):
                nm = sch.find("Item-name")
                if nm is not None and (nm.text or "").startswith("FX"):
                    for sd in sch.iter("Size-definition"):
                        if sd.get("nominal") == "20":
                            print(f"    bundled SLF {nm.text} 20A internal = {sd.get('internal')} "
                                  f"({'FX216 OK' if sd.get('internal') in ('21.6',) else 'NOT 21.6!'})")
    print("\nDONE ✓")


if __name__ == "__main__":
    main()
