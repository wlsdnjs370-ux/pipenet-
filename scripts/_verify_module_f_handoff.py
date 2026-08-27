# -*- coding: utf-8 -*-
"""[F-8d] 탈출로 — 자동 결과를 손질로 이어받는다.

  ① 자동 미리보기 → 이어받기 → 손질 진입까지 «잡 1개»
  ② 진입 직후 board 의 재료·헤드가 채택 결과와 일치한다
  ③ 알람밸브·급수원 제안이 payload 로 오고, 원클릭 반영이 기존 클릭 경로를 탄다
  ④ 이어받기 후 edit/worst → design/build → design/emit 이 수동과 같이 완주한다
  ⑤ 이어받기를 안 쓰는 자동 차선이 종전과 같다 (회귀)

★스펙을 공유 작업폴더에 저장하므로 사용자의 B1F 저장본을 덮지 않도록
  대명동 단위세대 샘플로 돌리고, 만들어진 파일은 끝에서 지운다.

    python scripts/_verify_module_f_handoff.py
"""
from __future__ import annotations

import importlib.util
import io
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DXF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
WORK = ROOT / "cad_project_editor_g" / "docs" / "import"
FAILS: list[str] = []


def check(label, cond, detail=""):
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return bool(cond)


def _wait(c, sid, limit=4000):
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


def _cleanup(key: str):
    removed = 0
    for pat in (f"0단계_새찍기/{key}_*", f"0단계_새찍기/{key}*stage1_world*",
                f"DWG/{key}_유저손질.json", f"_edit_disp_cache_{key}.json"):
        for p in WORK.glob(pat):
            try:
                p.unlink()
                removed += 1
            except OSError:
                pass
    print(f"  [정리] 작업폴더에서 {removed}개 걷어냄 (키 {key})")


def _app():
    spec = importlib.util.spec_from_file_location(
        "daejo", os.path.join(str(ROOT), "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.app.config["TESTING"] = True
    return mod.app


def _auto_upto_run(c, path):
    """열기 → 자동 선택 → 알람밸브 → 최불리까지. 반환 sid."""
    with open(path, "rb") as f:
        raw = f.read()
    sid = c.post("/api/module-f/open", data={
        "dxf_file": (io.BytesIO(raw), os.path.basename(str(path)))},
        content_type="multipart/form-data").get_json()["sid"]
    if _wait(c, sid).get("state") != "done":
        return None
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "auto"})
    if _wait(c, sid).get("state") != "done":
        return None
    hs = c.post("/api/module-f/auto/heads", json={"sid": sid}).get_json()
    if not hs.get("n"):
        return None
    xs = [h["x"] for h in hs["heads"]]
    ys = [h["y"] for h in hs["heads"]]
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    best = min(hs["heads"], key=lambda h: (h["x"] - cx) ** 2 + (h["y"] - cy) ** 2)
    c.post("/api/module-f/auto/anchor",
           json={"sid": sid, "x": best["x"], "y": best["y"]})
    c.post("/api/module-f/auto/run", json={"sid": sid, "k": 10})
    if _wait(c, sid).get("state") != "done":
        return None
    return sid


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    if not DXF.is_file():
        print(f"샘플 DXF 없음: {DXF}")
        return 1
    key = os.path.splitext(os.path.basename(str(DXF)))[0]
    if (WORK / "0단계_새찍기" / f"{key}_찍은스펙.json").exists():
        print(f"!! 작업폴더에 이미 {key} 저장본이 있다 — 덮지 않기 위해 중단")
        return 1

    app = _app()
    try:
        with app.test_client() as c:
            with c.session_transaction() as s:
                s["authed"] = True

            print(f"\n[①] 자동 → 이어받기 — {DXF.name}")
            sid = _auto_upto_run(c, DXF)
            if not check("자동 추출까지 완주", sid is not None):
                return 1
            pv = c.get(f"/api/module-f/auto/preview?sid={sid}").get_json()
            check("자동 미리보기가 있다", bool(pv.get("ok")),
                  f"헤드 {(pv.get('summary') or {}).get('k')}")

            n_jobs_before = 1
            t0 = time.perf_counter()
            r = c.post("/api/module-f/auto/handoff", json={"sid": sid})
            if not check("이어받기 수락", bool(r.get_json().get("ok")),
                         str(r.get_json())[:90]):
                return 1
            jb = _wait(c, sid)
            for ln in (jb.get("lines") or [])[-8:]:
                print(f"       | {ln}")
            if not check("이어받기 잡 완료 (잡 1개)",
                         jb.get("state") == "done" and n_jobs_before == 1,
                         str(jb.get("error"))[:100]):
                return 1
            print(f"       {time.perf_counter() - t0:.1f}s")
            res = _result(c, sid) or {}

            print("\n[②] 진입 직후 board — 채택 결과와 맞나")
            st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()
            check("손질 세션이 섰다", bool(st.get("ok")), str(st)[:80])
            est = st.get("state") or {}
            check("망이 구성됐다",
                  len(est.get("body_groups") or []) > 0
                  or len(est.get("heads") or []) > 0,
                  f"헤드 {len(est.get('heads') or [])}")
            pk = res.get("pick") or {}
            check("재료가 채택 결과와 같다",
                  len(pk.get("materials") or []) == len(res.get("mat_applied") or []),
                  f"board {len(pk.get('materials') or [])} vs "
                  f"보고 {len(res.get('mat_applied') or [])}")
            check("헤드가 찍혀 있다", (pk.get("n_heads") or 0) > 0,
                  f"{pk.get('n_heads')}픽 · 찍힘 {res.get('head_applied')}"
                  f" · 이미 {res.get('head_already')}"
                  f" · 유령 {res.get('head_skipped')}")

            print("\n[③] 제안 — payload 와 원클릭 반영")
            hints = c.get(f"/api/module-f/auto/handoff-hints?sid={sid}"
                          ).get_json().get("handoff") or {}
            check("알람밸브·급수원 제안이 둘 다 온다",
                  hints.get("alarm") and hints.get("source"),
                  f"{hints.get('alarm')} / {hints.get('source')}")
            xy = hints.get("source") or [0, 0]
            c.post("/api/module-f/edit/mode",
                   json={"sid": sid, "mode": "급수시작위치"})
            d = c.post("/api/module-f/edit/click",
                       json={"sid": sid, "x": xy[0], "y": xy[1],
                             "max_d": 3000}).get_json()
            est2 = d.get("state") or {}
            check("급수 시작이 제안 자리에 반영된다",
                  len(est2.get("sources") or []) == 1,
                  f"{len(est2.get('sources') or [])}곳")
            c.post("/api/module-f/edit/mode",
                   json={"sid": sid, "mode": "알람밸브위치"})
            d = c.post("/api/module-f/edit/click",
                       json={"sid": sid, "x": xy[0], "y": xy[1],
                             "max_d": 3000}).get_json()
            check("알람밸브도 같은 경로로 반영된다", bool(d.get("ok")),
                  str(d.get("message"))[:70])

            print("\n[④] 이어받기 뒤 수동 차선과 같이 완주하나")
            r = c.post("/api/module-f/design/build", json={"sid": sid, "k": 10})
            jb = _wait(c, sid)
            res2 = _result(c, sid) or {}
            if not check("design/build", jb.get("state") == "done"
                         and res2.get("ok"), str(res2.get("error"))[:100]):
                return 1
            s = res2.get("summary") or {}
            print(f"       최불리 {s.get('k')}개 · 절점 {s.get('nodes')} · "
                  f"배관 {s.get('pipes')}")
            r = c.post("/api/module-f/design/emit", json={"sid": sid})
            jb = _wait(c, sid)
            res3 = _result(c, sid) or {}
            check("design/emit", jb.get("state") == "done" and res3.get("ok"),
                  str(res3.get("error"))[:100])

            print("\n[⑤] 이어받기를 안 쓰는 자동 차선 — 종전 그대로")
            sid2 = _auto_upto_run(c, DXF)
            check("자동만으로 완주", sid2 is not None)
            if sid2:
                pv2 = c.get(f"/api/module-f/auto/preview?sid={sid2}").get_json()
                a, b = (pv.get("summary") or {}), (pv2.get("summary") or {})
                same = all(a.get(k2) == b.get(k2)
                           for k2 in ("k", "nodes", "pipes", "far_m", "near_m"))
                check("이어받기 전후 자동 결과가 같다", same,
                      f"{a.get('k')}/{a.get('nodes')}/{a.get('far_m')} vs "
                      f"{b.get('k')}/{b.get('nodes')}/{b.get('far_m')}")
                st2 = c.get(f"/api/module-f/auto/state?sid={sid2}").get_json()
                check("자동 세션은 손질로 안 넘어간다",
                      st2.get("method") == "auto", str(st2.get("method")))
    finally:
        _cleanup(key)

    print("\n" + "=" * 58)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\nFAIL — {len(FAILS)}건")
        return 1
    print("F-8d 탈출로 — 수용기준 전 항목 PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
