# -*- coding: utf-8 -*-
"""[F-8e] 골든 — adopt 를 안 쓰는 수동 경로의 전체망 .kfp 가 F-8 전후로 같은가.

F-8 은 «새 길을 하나 더 낸 것» 이지 기존 길을 고친 것이 아니다. 그 주장을
말이 아니라 파일 해시로 증명한다: 같은 도면을 같은 손으로 찍어(레이어 추천
일괄 + 후보 전체 반영 + 급수원 1점) 전체망 .kfp 를 뽑고, 그 바이트가 F-8
이전 리비전에서 뽑은 것과 같은지 본다.

    python scripts/_golden_module_f_kfp.py --emit  data/_kfp_after.json
    (다른 리비전 워크트리에서) --emit data/_kfp_before.json
    python scripts/_golden_module_f_kfp.py --compare before.json after.json

★사용자의 저장본을 건드리지 않는다 — 쓰기 루트를 임시 폴더로 돌린다.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import io
import json
import os
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DXF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def _wait(c, sid, limit=8000):
    jb = {}
    for _ in range(limit):
        jb = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if jb.get("state") in ("done", "error"):
            return jb
        time.sleep(0.3)
    return jb


def _result(c, sid):
    return (c.get(f"/api/module-f/convert/result?sid={sid}").get_json()
            or {}).get("result")


def _app(work: str):
    spec = importlib.util.spec_from_file_location(
        "daejo", os.path.join(str(ROOT), "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.app.config["TESTING"] = True
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import disp_cache, handoff
    handoff.import_write_root = lambda: work
    handoff.OUT_DIR = handoff.pick_out_dir()
    disp_cache._DISP_CACHE_DIR = work
    os.makedirs(handoff.pick_out_dir(), exist_ok=True)
    os.makedirs(handoff.default_edits_dir(), exist_ok=True)
    return mod.app


def _candidates(c, sid):
    """후보 목록 — F-8 이후는 /recon, 이전 리비전은 /pick/suggest 로 얻는다."""
    r = c.get(f"/api/module-f/recon?sid={sid}&heads=1")
    if r.status_code == 200 and (r.get_json() or {}).get("heads") is not None:
        return r.get_json()["heads"]
    c.post("/api/module-f/pick/suggest", json={"sid": sid})
    _wait(c, sid)
    return ((_result(c, sid) or {}).get("candidates")) or []


def emit(out_json: str, dxf=None) -> int:
    # ★워크트리에서 돌릴 때를 위해 도면 경로를 받는다 — 이 샘플은 추적 대상이
    #   아니라 다른 리비전 워크트리에는 없다.
    global DXF
    if dxf:
        DXF = Path(dxf)
    if not DXF.is_file():
        print(f"샘플 DXF 없음: {DXF}")
        return 1
    with tempfile.TemporaryDirectory(prefix="f8golden_") as work:
        app = _app(work)
        with app.test_client() as c:
            with c.session_transaction() as s:
                s["authed"] = True

            with open(DXF, "rb") as f:
                raw = f.read()
            sid = c.post("/api/module-f/open", data={
                "dxf_file": (io.BytesIO(raw), os.path.basename(str(DXF)))},
                content_type="multipart/form-data").get_json()["sid"]
            if _wait(c, sid).get("state") != "done":
                print("열기 실패")
                return 1
            # F-8 이후에만 있는 문 — 이전 리비전에서는 404 이고 그래도 된다.
            c.post("/api/module-f/slot/read",
                   json={"sid": sid, "method": "manual"})

            # ── 사람 손 그대로 (adopt 없이)
            c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
            c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
            c.post("/api/module-f/pick/mode",
                   json={"sid": sid, "action": "complete"})
            c.post("/api/module-f/pick/mode",
                   json={"sid": sid, "action": "slot", "slot": "상향하향"})
            cands = _candidates(c, sid)
            for cd in cands:
                d = c.post("/api/module-f/pick/click",
                           json={"sid": sid, "x": cd["x"], "y": cd["y"],
                                 "max_d": 300}).get_json()
                if ((d.get("report") or {}).get("동작")) == "취소":
                    c.post("/api/module-f/pick/click",
                           json={"sid": sid, "x": cd["x"], "y": cd["y"],
                                 "max_d": 300})
            c.post("/api/module-f/pick/commit", json={"sid": sid})
            if _wait(c, sid).get("state") != "done":
                print("배관망 구성 실패")
                return 1

            est = (c.get(f"/api/module-f/edit/state?sid={sid}").get_json()
                   or {}).get("state") or {}
            sg = (est.get("body_groups") or [{}])[0].get("segs")
            if not sg:
                print("배관이 없다")
                return 1
            c.post("/api/module-f/edit/mode",
                   json={"sid": sid, "mode": "급수시작위치"})
            c.post("/api/module-f/edit/click",
                   json={"sid": sid, "x": (sg[0] + sg[2]) / 2,
                         "y": (sg[1] + sg[3]) / 2, "max_d": 5000})

            c.post("/api/module-f/convert/run",
                   json={"sid": sid, "outputs": {"full_kfp": True}})
            jb = _wait(c, sid)
            res = _result(c, sid) or {}
            if jb.get("state") != "done" or not res.get("ok"):
                print("변환 실패:", str(res)[:200])
                return 1

            from routes.module_f.jobs import _sess
            path = Path(_sess(sid)["kfp_path"])
            body = json.loads(path.read_text(encoding="utf-8"))
            # 세션 id·경로·시각처럼 «실행마다 다른 것» 은 뺀다. 남는 것이 망이다.
            for k in ("meta", "created", "source_path", "session"):
                body.pop(k, None)
            canon = json.dumps(body, ensure_ascii=False, sort_keys=True)
            out = {
                "cands": len(cands),
                "nodes": len(body.get("nodes_meta_runtime") or {}),
                "pipes": len(body.get("pipe_data") or {}),
                "bytes": path.stat().st_size,
                "sha256": hashlib.sha256(canon.encode("utf-8")).hexdigest(),
            }
    Path(out_json).write_text(json.dumps(out, ensure_ascii=False, indent=1),
                              encoding="utf-8")
    print(json.dumps(out, ensure_ascii=False, indent=1))
    print(f"기록: {out_json}")
    return 0


def compare(a_path: str, b_path: str) -> int:
    a = json.loads(Path(a_path).read_text(encoding="utf-8"))
    b = json.loads(Path(b_path).read_text(encoding="utf-8"))
    print(f"  전 {Path(a_path).name}: 노드 {a['nodes']} · 배관 {a['pipes']} · "
          f"{a['bytes']:,}B · {a['sha256'][:16]}…")
    print(f"  후 {Path(b_path).name}: 노드 {b['nodes']} · 배관 {b['pipes']} · "
          f"{b['bytes']:,}B · {b['sha256'][:16]}…")
    same = a["sha256"] == b["sha256"]
    print("\n" + ("PASS — 전체망 .kfp 가 비트 동일하다"
                  if same else "FAIL — 달라졌다"))
    if not same:
        for k in ("cands", "nodes", "pipes", "bytes"):
            if a.get(k) != b.get(k):
                print(f"    {k}: {a.get(k)} → {b.get(k)}")
    return 0 if same else 1


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--emit")
    ap.add_argument("--dxf")
    ap.add_argument("--compare", nargs=2)
    a = ap.parse_args()
    if a.emit:
        return emit(a.emit, a.dxf)
    if a.compare:
        return compare(*a.compare)
    ap.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
