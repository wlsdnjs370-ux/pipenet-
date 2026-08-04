# -*- coding: utf-8 -*-
"""기계실 천장고 배선 스모크 — 실제 DXF 를 HTTP 로 올려 표고 출처를 확인한다.

천장고를 넣은 요청과 비운 요청을 같은 도면으로 각각 보내, 첫 구간이
사람 확정 낙차 / 미확정 으로 갈리는지 본다.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))
sys.stdout.reconfigure(encoding="utf-8")

from characterization.golden_cases import _server_mod, _rp  # noqa: E402

MROOM = REPO / "data" / "sample_problem" / "옥상수조.dxf"

rp = _rp()
ents = rp.parse_dxf_for_view(MROOM, include_hidden_layers=True)["entities"]
xs, ys = [], []


def _flat(v, out):
    if isinstance(v, (list, tuple)):
        for it in v:
            _flat(it, out)
    elif isinstance(v, (int, float)):
        out.append(float(v))


for e in ents:
    if e.get("t") not in ("L", "P"):
        continue
    flat: list[float] = []
    _flat(e.get("p") or [], flat)
    xs.extend(flat[0::2]); ys.extend(flat[1::2])

mod = _server_mod()
c = mod.app.test_client()
with c.session_transaction() as s:
    s["authed"] = True

raw = MROOM.read_bytes()
for label, extra in (("천장고 4.5m", {"machine_room_ceiling_m": "4.5"}),
                     ("천장고 비움", {})):
    form = {
        "source_x": str(min(xs)), "source_y": str(max(ys)),
        "conn_x": str(max(xs)),   "conn_y": str(min(ys)),
        "snap_tolerance_mm": "1e9",
        "machineroom_dxf_file": (io.BytesIO(raw), "옥상수조.dxf"),
        **extra,
    }
    r = c.post("/api/remote30/machineroom/extract", data=form,
               content_type="multipart/form-data")
    body = r.json or {}
    mr = body.get("machine_room") or {}
    if not mr:
        print(f"{label:10s} HTTP {r.status_code} ok={body.get('ok')} "
              f"{str(body.get('message'))[:200]}")
        continue
    p0 = mr["pipes"][0]
    srcs = (mr.get("elev_sources") or {}).get("pipes") or {}
    print(f"{label:10s} HTTP {r.status_code}  천장고={mr.get('machine_room_ceiling_m')}  "
          f"첫구간 elev={p0['elev']} len={p0['length']} src={p0['elev_source']}")
    print(f"{'':10s} 노드 {len(mr['nodes'])} 배관 {len(mr['pipes'])}  "
          f"관로 출처 분포 {[(k, v) for k, v in srcs.items() if v]}")
