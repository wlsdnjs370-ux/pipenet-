# -*- coding: utf-8 -*-
"""[F-8e] 세 차선을 같은 도면으로 끝까지 돌려 «값어치» 를 숫자로 만든다.

    자동(A)      알람밸브·영역만 사람이 정하고 나머지는 A
    혼합(A+E)    인식 결과를 채택해 찍은 뒤 E 의 물길로
    수동(E)      레이어 추천만 받고 사람이 찍는다

재는 것: 사람 조작 수 · 소요 시간 · 물닿는 헤드 수 · 최불리 결과.
혼합의 존재 이유(조작 수 ↓ vs 수동, 도달 헤드 ↑ vs 자동)가 여기서 보여야 한다.

★«사람 조작» 과 «클릭» 을 가른다. 수동 차선의 「전체 반영」은 단추 한 번이지만
  화면이 후보마다 /pick/click 을 태운다 — 그 수천 번을 사람이 눌렀다고 세면
  수동이 실제보다 훨씬 나빠 보인다. 반대로 그것을 안 세면 서버 부담이 안 보인다.
  그래서 둘 다 적는다: `human`(단추) · `clicks`(서버가 태운 클릭).

★사용자의 저장본을 절대 건드리지 않는다. `handoff` 의 쓰기 루트를 임시 폴더로
  돌려 이 실행이 만드는 스펙·손질·캐시가 전부 그 안에만 생기게 한다 — 백업 후
  복원은 도중에 죽으면 복구가 안 되지만, 애초에 다른 폴더에 쓰면 그럴 일이 없다.

    python scripts/_measure_module_f_lanes.py [도면.dxf] [--k 30]
"""
from __future__ import annotations

import argparse
import importlib.util
import io
import json
import os
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


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
    """서버를 올리되 쓰기 루트를 임시 폴더로 돌린다."""
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
    print(f"  쓰기 루트를 임시 폴더로 돌림 — {work}")
    return mod.app


def _open(c, path):
    with open(path, "rb") as f:
        raw = f.read()
    return c.post("/api/module-f/open", data={
        "dxf_file": (io.BytesIO(raw), os.path.basename(str(path)))},
        content_type="multipart/form-data").get_json()["sid"]


def _alarm_of(heads):
    xs = [h["x"] for h in heads]
    ys = [h["y"] for h in heads]
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    b = min(heads, key=lambda h: (h["x"] - cx) ** 2 + (h["y"] - cy) ** 2)
    return b["x"], b["y"]


def _design(c, sid, k):
    """손질망 → 최불리 표. 반환: (summary, 제외사유, 물닿음).

    ★물닿는 헤드 수는 «board 헤드 총수 − 물길 미도달» 이다. 손질 상태의 heads
      길이를 세면 그냥 board 전체라 물길 판정을 안 탄 값이 된다 — 설계 표시용
      marks 가 유일하게 갈라 놓은 자리다(`_classify_excluded`).
    """
    c.post("/api/module-f/design/build", json={"sid": sid, "k": k})
    jb = _wait(c, sid)
    r = _result(c, sid) or {}
    if jb.get("state") != "done" or not r.get("ok"):
        return None, {"error": str(r.get("error"))[:120]}, None
    s = r.get("summary") or {}
    marks = (c.get(f"/api/module-f/design/preview?sid={sid}").get_json()
             or {}).get("marks") or {}
    total = marks.get("total")
    dry = ((marks.get("dry") or {}).get("n"))
    wet = (total - dry) if (total is not None and dry is not None) else None
    return s, (s.get("excluded_detail") or {}), {"wet": wet, "board": total}


# ─────────────────────────────────────────── 차선 셋
def lane_auto(c, path, k):
    """자동 — 사람 클릭: 알람밸브 1 (영역은 안 그린다)."""
    t0 = time.perf_counter()
    sid = _open(c, path)
    if _wait(c, sid).get("state") != "done":
        return {"error": "열기 실패"}
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "auto"})
    _wait(c, sid)
    hs = c.post("/api/module-f/auto/heads", json={"sid": sid}).get_json()
    if not hs.get("n"):
        return {"error": "헤드 0"}
    x, y = _alarm_of(hs["heads"])
    c.post("/api/module-f/auto/anchor", json={"sid": sid, "x": x, "y": y})
    c.post("/api/module-f/auto/run", json={"sid": sid, "k": k})
    jb = _wait(c, sid)
    if jb.get("state") != "done":
        return {"error": str(jb.get("error"))[:120]}
    pv = c.get(f"/api/module-f/auto/preview?sid={sid}").get_json()
    s = pv.get("summary") or {}
    # 사람 조작: 「자동으로 계속」 · 알람밸브 찍기 · 「배관망 추출」
    return {"sid": sid, "human": 3, "clicks": 1,
            "sec": round(time.perf_counter() - t0, 1),
            "detected": hs["n"], "wet": None, "k": s.get("k"),
            "far_m": s.get("far_m"), "nodes": s.get("nodes"),
            "pipes": s.get("pipes"), "excluded": None,
            "note": "물닿음은 A 경로에 없다(E 의 물길 판정을 안 탄다)"}


def lane_mixed(c, path, k, conf_min):
    """혼합 — 사람 클릭: 「인식 결과로 찍기 시작」 1 + 급수 시작 1."""
    t0 = time.perf_counter()
    sid = _open(c, path)
    if _wait(c, sid).get("state") != "done":
        return {"error": "열기 실패"}
    rec = c.get(f"/api/module-f/recon?sid={sid}").get_json()["recon"]
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    c.post("/api/module-f/pick/adopt",
           json={"sid": sid, "materials": True,
                 "heads": {"conf_min": conf_min}})
    jb = _wait(c, sid)
    ad = _result(c, sid) or {}
    if jb.get("state") != "done" or not ad.get("ok"):
        return {"error": str(ad.get("error") or jb.get("error"))[:120]}
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    if _wait(c, sid).get("state") != "done":
        return {"error": "배관망 구성 실패"}
    _seed_source(c, sid)
    s, det, w = _design(c, sid, k)
    if s is None:
        return {"error": det.get("error")}
    # 사람 조작: 「인식 결과로 찍기 시작」 · 「배관망 구성」 · 급수 시작 찍기
    return {"sid": sid, "human": 3,
            "clicks": 1 + int(ad.get("head_applied") or 0)
            + 2 * int(ad.get("head_already") or 0),
            "sec": round(time.perf_counter() - t0, 1),
            "detected": rec["n"], "wet": w, "k": s.get("k"),
            "far_m": s.get("far_m"), "nodes": s.get("nodes"),
            "pipes": s.get("pipes"), "excluded": det,
            "adopt": {"mat": len(ad.get("mat_applied") or []),
                      "head": ad.get("head_applied"),
                      "already": ad.get("head_already"),
                      "ghost": ad.get("head_skipped")}}


def lane_manual(c, path, k):
    """수동 — 사람 클릭: 배관 추천 일괄 1 + 완료 1 + 후보 반영 N + 급수 1.

    「후보 반영」은 화면이 후보마다 /pick/click 을 태운다 — 사람이 그만큼
    누르는 것과 같은 값이므로 클릭으로 센다(F-5 의 그 경로 그대로).
    """
    t0 = time.perf_counter()
    sid = _open(c, path)
    if _wait(c, sid).get("state") != "done":
        return {"error": "열기 실패"}
    rec = c.get(f"/api/module-f/recon?sid={sid}").get_json()["recon"]
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    clicks = 0
    c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
    c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
    clicks += 1
    c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "complete"})
    clicks += 1
    c.post("/api/module-f/pick/mode",
           json={"sid": sid, "action": "slot", "slot": "상향하향"})
    # 후보를 사람 손으로 하나씩 (화면의 「전체 반영」과 같은 경로·같은 규칙)
    cands = (c.get(f"/api/module-f/recon?sid={sid}&heads=1").get_json()
             or {}).get("heads") or []
    for cd in cands:
        d = c.post("/api/module-f/pick/click",
                   json={"sid": sid, "x": cd["x"], "y": cd["y"],
                         "max_d": 300}).get_json()
        clicks += 1
        if ((d.get("report") or {}).get("동작")) == "취소":
            c.post("/api/module-f/pick/click",
                   json={"sid": sid, "x": cd["x"], "y": cd["y"],
                         "max_d": 300})
            clicks += 1
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    if _wait(c, sid).get("state") != "done":
        return {"error": "배관망 구성 실패"}
    _seed_source(c, sid)
    s, det, w = _design(c, sid, k)
    if s is None:
        return {"error": det.get("error")}
    # 사람 조작: 배관 추천 일괄 · 선택 완료 · 헤드 칸 · 후보 제안 · 전체 반영 ·
    #            배관망 구성 · 급수 시작 = 7. 「전체 반영」 한 번이 아래 clicks 를
    #            전부 태운다 — 사람이 그만큼 «누른» 것은 아니다.
    return {"sid": sid, "human": 7, "clicks": clicks,
            "sec": round(time.perf_counter() - t0, 1),
            "detected": rec["n"], "wet": w, "k": s.get("k"),
            "far_m": s.get("far_m"), "nodes": s.get("nodes"),
            "pipes": s.get("pipes"), "excluded": det}


def _seed_source(c, sid):
    """급수 시작을 배관 위 한 점에 찍는다 — 세 차선 공통 조건. 찍었나 여부만."""
    est = (c.get(f"/api/module-f/edit/state?sid={sid}").get_json()
           or {}).get("state") or {}
    bg = est.get("body_groups") or []
    if not bg or not bg[0].get("segs"):
        return False
    sg = bg[0]["segs"]
    c.post("/api/module-f/edit/mode", json={"sid": sid, "mode": "급수시작위치"})
    d = c.post("/api/module-f/edit/click",
               json={"sid": sid, "x": (sg[0] + sg[2]) / 2,
                     "y": (sg[1] + sg[3]) / 2, "max_d": 5000}).get_json()
    return len(((d.get("state") or {}).get("sources") or [])) == 1


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--conf", type=float, default=0.75)
    # 문턱을 바꿔 다시 잴 때 수동 차선까지 또 돌릴 이유가 없다 — 수동은 문턱과
    # 무관하다(후보 전체를 찍는다). 혼합만 골라 돌릴 수 있게 한다.
    ap.add_argument("--lanes", default="자동,혼합,수동")
    ap.add_argument("--out", default=str(ROOT / "data" / "_lanes.json"))
    a = ap.parse_args()
    want = {s.strip() for s in a.lanes.split(",") if s.strip()}

    path = Path(a.dxf)
    if not path.is_file():
        print("도면 없음:", path)
        return 1

    with tempfile.TemporaryDirectory(prefix="f8e_") as work:
        app = _app(work)
        out = {"dxf": path.name,
               "mb": round(path.stat().st_size / 1024 / 1024, 1),
               "k": a.k, "conf_min": a.conf}
        with app.test_client() as c:
            with c.session_transaction() as s:
                s["authed"] = True
            for name, fn in (("자동", lambda: lane_auto(c, path, a.k)),
                             ("혼합", lambda: lane_mixed(c, path, a.k, a.conf)),
                             ("수동", lambda: lane_manual(c, path, a.k))):
                if name not in want:
                    continue
                print(f"\n{'=' * 20} {name} 차선 {'=' * 20}")
                try:
                    out[name] = fn()
                except Exception as exc:  # noqa: BLE001
                    out[name] = {"error": f"{type(exc).__name__}: {exc}"}
                print(f"  → {json.dumps(out[name], ensure_ascii=False)}")

    Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1),
                           encoding="utf-8")
    print(f"\n기록: {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
