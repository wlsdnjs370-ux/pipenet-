# -*- coding: utf-8 -*-
"""FX Task7 — override 편집이 SDF 에 반영되는지 + ±50% 경고 경로 검증."""
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
    out = []
    for block in resp.get_data(as_text=True).split("\n\n"):
        block = block.strip()
        if block.startswith("data:"):
            try:
                out.append(json.loads(block[len("data:"):].strip()))
            except json.JSONDecodeError:
                pass
    return out


def main():
    app = _load_app()
    c = app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    with DXF.open("rb") as f:
        r = c.post("/api/remote30/prototype/run",
                   data={"dxf_file": (f, DXF.name)}, content_type="multipart/form-data")
    job_id = r.get_json()["job_id"]
    _sse(c.get(f"/api/remote30/prototype/stream/{job_id}"))
    c.post(f"/api/remote30/prototype/finalize/{job_id}", json={})
    evs = _sse(c.get(f"/api/remote30/prototype/finalize_stream/{job_id}"))
    s5 = next(e for e in evs if e.get("type") == "stage5_complete")
    equipment = [dict(e) for e in s5["fx_review"]["equipment"]]

    # 첫 FX 행을 40.0 으로 override — ±50% 경고를 유발한다.
    # 기준값은 하드코딩하지 않는다. 규격 프로파일 기본값이 바뀌면(22.4 → 15.6 처럼)
    # 이 하네스가 제품 회귀 없이 혼자 빨개진다.
    target = next(e for e in equipment if e["desc"] == "FX")
    tlabel = target["label"]
    base_eq = float(target["eq_len"])
    base_others = sum(1 for e in equipment
                      if e["desc"] == "FX" and e["label"] != tlabel
                      and float(e["eq_len"]) == base_eq)
    target["eq_len"] = 40.0
    target["override_flag"] = True
    target["override_note"] = "task7 override test"
    print(f"override FX label={tlabel}: {base_eq} -> 40.0 (동일 기준값 나머지 {base_others}행)")

    r = c.post(f"/api/remote30/prototype/fx/finalize/{job_id}",
               json={"equipment": equipment})
    assert r.status_code == 200, r.get_data(as_text=True)[:300]
    print("fx/finalize mode:", r.get_json().get("mode"), "edited:", r.get_json().get("edited"))

    evs = _sse(c.get(f"/api/remote30/prototype/fx/finalize_stream/{job_id}"))
    errs = [e for e in evs if e.get("type") == "error"]
    assert not errs, errs
    warns = [e for e in evs if e.get("type") == "warning"]
    print(f"warnings emitted: {len(warns)}")
    for w in warns:
        print("   ⚠", w.get("message", "")[:120])
    done = next(e for e in evs if e.get("type") == "done")

    # SDF 확인: 40.0 이 정확히 1개, 손대지 않은 FX 는 기준값 그대로
    sdf = next((BASE / "data" / "prototype_runs" / job_id).glob("*.sdf"))
    txt = sdf.read_text(encoding="utf-8", errors="replace")
    import re
    vals = re.findall(r'equivalent-length="([^"]*)"', txt)
    from collections import Counter
    cnt = Counter(vals)
    print("SDF equivalent-length counts:", dict(cnt))
    assert cnt.get("40", 0) + cnt.get("40.0", 0) == 1, "override 40.0 이 SDF 에 1개여야 함"
    kept = sum(cnt.get(k, 0) for k in {f"{base_eq:g}", str(base_eq)})
    assert kept == base_others, f"나머지 FX 는 {base_eq:g} 로 {base_others}개여야 함 (실제 {kept})"
    assert len(warns) >= 1, "±50% 편차 경고가 최소 1개 나와야 함"
    print("\nOVERRIDE VERIFY ✓ (40.0 SDF 반영 + 경고 발생)")


if __name__ == "__main__":
    main()
