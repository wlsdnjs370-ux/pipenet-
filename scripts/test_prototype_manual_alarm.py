"""Manual alarm_xy 옵션 검증 — auto vs manual 두 번 실행해 비교."""

from __future__ import annotations

import io
import json
import sys
import time
from pathlib import Path

import requests

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = "http://127.0.0.1:5050"
DXF = Path(r"C:\Users\admin\PycharmProjects\JupyterProject\대명동201동 단위세대_layer정리.dxf")


def run_once(label: str, *, alarm_x: float | None = None, alarm_y: float | None = None) -> dict:
    print(f"\n=== {label} ===")
    with DXF.open("rb") as f:
        files = {"dxf_file": (DXF.name, f, "application/octet-stream")}
        data = {}
        if alarm_x is not None and alarm_y is not None:
            data["alarm_x"] = str(alarm_x); data["alarm_y"] = str(alarm_y)
        r = requests.post(f"{BASE}/api/remote30/prototype/run", files=files, data=data, timeout=60)
    j = r.json()
    if not j.get("ok"):
        print(f"ERR: {j}")
        return {}
    print(f"  job_id={j['job_id']} alarm_xy={j.get('alarm_xy')}")
    job_id = j["job_id"]
    summary = {}
    t0 = time.time()
    with requests.get(f"{BASE}/api/remote30/prototype/stream/{job_id}", stream=True, timeout=300) as resp:
        for raw in resp.iter_lines():
            if not raw or not raw.startswith(b"data: "):
                continue
            evt = json.loads(raw[6:])
            t = evt.get("type")
            if t == "entities" and evt.get("stage") == 2:
                summary = evt.get("summary", {})
                print(f"  [{(time.time()-t0):4.1f}s] stage2 summary: {summary}")
            elif t == "tables_preview":
                summary["table_counts"] = evt.get("counts", {})
                print(f"  table counts: {evt.get('counts')}")
            elif t == "done":
                summary["job_id"] = job_id
                break
            elif t == "error":
                print(f"  ERR {evt['message']}")
                break
    return summary


# auto
auto_summary = run_once("AUTO (배관-SP 2차 LINE endpoint 자동 식별)")

# 참조 input.xlsx 의 알람밸브 좌표 (input node label 10): (-11400, -3233) 이지만 이건 ref 좌표계
# 우리 도면의 헤드 무게중심 근처 임의 좌표로 시험. 입력 도면의 알람밸브가 거기 있다고 가정.
manual_summary = run_once("MANUAL (수동 좌표 264762, -243111)", alarm_x=264762.6, alarm_y=-243111.2)

print("\n=== 비교 ===")
print(f"  auto  : heads={auto_summary.get('selected_heads')} dist_m={auto_summary.get('max_distance_m')} src={auto_summary.get('source_kind')}")
print(f"  manual: heads={manual_summary.get('selected_heads')} dist_m={manual_summary.get('max_distance_m')} src={manual_summary.get('source_kind')}")
print(f"  auto tables  : {auto_summary.get('table_counts')}")
print(f"  manual tables: {manual_summary.get('table_counts')}")
