# -*- coding: utf-8 -*-
"""C3 서버 경로를 실도면 세션으로 굴린다 (지시서 §16 규칙 8).

합성 fixture 는 좌표가 정확히 맞아떨어져 문 인식의 진짜 난이도를 재지 못한다.
가장 최근에 인식한 실도면 세션의 `building.json` 을 임시 세션으로 복사해
C2 → C3 를 그대로 태우고, 나온 수치를 사람이 읽을 수 있게 적는다.
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

from flask import Flask  # noqa: E402

import routes.r30_design as r30_design  # noqa: E402
from core.design import session as S  # noqa: E402
from core.design.deterministic import zoning as Z  # noqa: E402
from core.design.schema import BuildingDraft  # noqa: E402

_REQ_OK = {k: True for k in Z.MANUAL_REQUIREMENTS}


def _pick_source() -> Path:
    found = sorted((_ROOT / "data" / "design_sessions").glob("*/building.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    for path in found:
        draft = BuildingDraft.from_dict(json.loads(path.read_text("utf-8")))
        if len(draft.rooms) >= 10 and Z.valve_candidates(draft):
            return path
    raise SystemExit("실이 10개 넘고 확정 코어가 있는 세션이 없다 — 먼저 C1 을 돌려라")


def main() -> int:
    src = _pick_source()
    raw = json.loads(src.read_text("utf-8"))
    raw["gate"] = {"passed": True, "operator": "verify", "unresolved": []}
    draft = BuildingDraft.from_dict(raw)
    print(f"원본 세션 {src.parent.name} — 실 {len(draft.rooms)} / "
          f"가상 간선 {len(draft.virtual_edges)} / 코어 {len(draft.cores)}")

    root = Path(tempfile.mkdtemp(prefix="c3verify_"))
    try:
        app = Flask(__name__)
        r30_design.register(app, DESIGN_SESSION_DIR=root, enabled=True)
        client = app.test_client()
        sid = client.post("/api/design/session",
                          json={"operator": "verify"}).get_json()["session_id"]
        S.DesignSession.open(root, sid).write("building.json", raw)

        res = client.post("/api/design/c2/constraints", json={"session_id": sid})
        if res.status_code != 200:
            print(f"  C2 실패 {res.status_code}: {res.get_json()}")
            return 1
        print(f"  C2 기준 {res.get_json()['artifact']}")

        cands = client.get(f"/api/design/c3/candidates/{sid}").get_json()["candidates"]
        print(f"  밸브 후보 {len(cands)}개 — {', '.join(c['core_id'] for c in cands[:5])}"
              + (" …" if len(cands) > 5 else ""))

        # 후보 코어의 중심을 찍는다. 사람이 화면에서 하는 일과 같은 입력이다.
        placed = [{"core_id": c["core_id"], "point": c["center"],
                   "system_type": "습식", "requirements_confirmed": dict(_REQ_OK)}
                  for c in cands if c["center"]]
        res = client.post("/api/design/c3/valves",
                          json={"session_id": sid, "operator": "verify",
                                "valves": placed})
        if res.status_code != 200:
            print(f"  C3 밸브 실패 {res.status_code}: {res.get_json()}")
            return 1
        valves = res.get_json()["valves"]
        print(f"  밸브 {len(valves)}개 확정 — {valves[0]['id']} … {valves[-1]['id']}")

        res = client.post("/api/design/c3/zones", json={"session_id": sid})
        if res.status_code != 200:
            print(f"  C3 구역 실패 {res.status_code}: {res.get_json()}")
            return 1
        body = res.get_json()
        covered = sum(len(z["rooms"]) for z in body["zones"])
        print(f"  구역 {len(body['zones'])}개 / 담당 실 {covered}개 / "
              f"미도달 {len(body['unreached'])}개 / 문 없는 실 "
              f"{len(body['isolated_rooms'])}개")
        for code in sorted({f['code'] for f in body["flags"]}):
            n = sum(1 for f in body["flags"] if f["code"] == code)
            print(f"    flag {code} × {n}")

        stage = S.DesignSession.open(root, sid).status()["meta"]["stage"]
        print(f"  단계 {stage} — 남은 문제가 있으면 c3 에 머문다")

        # 같은 입력을 다시 태워도 같은 구역이 나와야 감사에서 재현된다.
        again = client.post("/api/design/c3/zones",
                            json={"session_id": sid}).get_json()
        same = again["zones"] == body["zones"]
        print(f"  재실행 동일: {'예' if same else '아니오'}")
        return 0 if same else 1
    finally:
        shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
