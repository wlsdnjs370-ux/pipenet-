# -*- coding: utf-8 -*-
"""[F-8a] 정찰을 실도면으로 확인한다 — 수용기준 다섯 줄.

  ① B1F 업로드 한 번(추가 클릭 0)으로 sess["recon"] 이 채워진다
  ② 도면이 정찰 «전에» 화면에 앉는다 (잡 로그 순서로 증명)
  ③ remote30_prototype import 를 깨뜨려도 열기·찍기는 살고 recon.error 만 남는다
  ④ 계통도 슬롯 열기에서는 정찰이 돌지 않는다 (로그 부재로 증명)
  ⑤ /pick/suggest 가 종전과 같은 모양으로 답한다

    python scripts/_verify_module_f_recon.py [평면도.dxf]
"""
from __future__ import annotations

import importlib.util
import io
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
SYSTEM = ROOT / "data" / "sample_problem" / "대명동201동 계통도_최소.dxf"
FAILS: list[str] = []


def check(label, cond, detail=""):
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return bool(cond)


def _wait(c, sid, limit=2400):
    jb = {}
    for _ in range(limit):
        jb = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if jb.get("state") in ("done", "error"):
            return jb
        time.sleep(0.3)
    return jb


def _open(c, path, kind="plan", sid=None):
    with open(path, "rb") as f:
        raw = f.read()
    data = {"dxf_file": (io.BytesIO(raw), os.path.basename(str(path)))}
    if kind == "plan" and sid is None:
        r = c.post("/api/module-f/open", data=data,
                   content_type="multipart/form-data")
    else:
        data["kind"] = kind
        if sid:
            data["sid"] = sid
        r = c.post("/api/module-f/slot/open", data=data,
                   content_type="multipart/form-data")
    return r.get_json()


def _app():
    spec = importlib.util.spec_from_file_location(
        "daejo", os.path.join(str(ROOT), "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.app.config["TESTING"] = True
    return mod.app


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    plan = Path(sys.argv[1]) if len(sys.argv) > 1 else PLAN
    if not plan.is_file():
        print(f"평면도 없음: {plan}")
        return 1

    app = _app()
    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True

        # ─────────────────────────── ①② 업로드 한 번으로 정찰까지
        print(f"\n[①②] 업로드 한 번 — {plan.name} "
              f"({plan.stat().st_size / 1024 / 1024:.1f} MB)")
        t0 = time.perf_counter()
        sid = _open(c, plan)["sid"]
        jb = _wait(c, sid)
        if not check("열기 잡 완료", jb.get("state") == "done",
                     str(jb.get("error"))[:90]):
            return 1
        print(f"       총 {time.perf_counter() - t0:.1f}s")

        lines = jb.get("lines") or []
        for ln in lines:
            print(f"       | {ln}")
        i_world = next((i for i, l in enumerate(lines)
                        if l.startswith("[찍기] 완료")), -1)
        i_recon = next((i for i, l in enumerate(lines)
                        if l.startswith("[정찰]")), -1)
        check("도면이 정찰보다 먼저 앉는다",
              i_world >= 0 and i_recon > i_world,
              f"찍기완료 #{i_world} < 정찰 #{i_recon}")
        check("정찰 진행이 잡 로그(SSE 원천)에 보인다", i_recon >= 0)

        rv = c.get(f"/api/module-f/recon?sid={sid}").get_json()
        rec = (rv or {}).get("recon") or {}
        if not check("정찰이 채워졌다(추가 클릭 0)", rec.get("state") == "ok",
                     str(rec)[:120]):
            return 1
        check("배관 묶음을 셌다", rec["bundles"].get("PIPE", 0) > 0,
              str(rec["bundles"]))
        check("헤드 후보가 나왔다", rec["n"] > 0,
              f"{rec['n']}개 · {rec['bands']}")
        check("띠 합 == 후보 수", sum(rec["bands"].values()) == rec["n"],
              f"{sum(rec['bands'].values())} vs {rec['n']}")

        full = c.get(f"/api/module-f/recon?sid={sid}&heads=1").get_json()
        check("좌표는 청할 때만 실린다",
              "heads" not in rv and len(full.get("heads") or []) == rec["n"],
              f"기본 없음 · heads=1 로 {len(full.get('heads') or [])}개")

        # 화면이 정찰 도는 중에도 도면을 그릴 수 있었나 — world 가 있다.
        w = c.get(f"/api/module-f/world?sid={sid}").get_json()
        check("도면이 내려간다", bool((w.get("world") or {}).get("bundles")),
              f"묶음 {len((w.get('world') or {}).get('bundles') or [])}")

        # ─────────────────────────── ⑤ suggest 종전 그대로
        print("\n[⑤] /pick/suggest — 종전 모양")
        r = c.post("/api/module-f/pick/suggest", json={"sid": sid})
        if check("수락", bool(r.get_json().get("ok")), str(r.get_json())[:80]):
            _wait(c, sid)
            res = (c.get(f"/api/module-f/convert/result?sid={sid}")
                   .get_json() or {}).get("result") or {}
            check("n · bands · candidates 가 그대로 온다",
                  {"n", "bands", "candidates"} <= set(res),
                  f"키 {sorted(res)}")
            check("suggest 결과가 정찰과 같은 수",
                  res.get("n") == rec["n"],
                  f"{res.get('n')} vs {rec['n']}")

        # ─────────────────────────── ④ 계통도는 정찰하지 않는다
        print("\n[④] 계통도 슬롯 — 정찰 없음")
        if not SYSTEM.is_file():
            print(f"  [건너뜀] 계통도 샘플 없음: {SYSTEM}")
        else:
            _open(c, SYSTEM, kind="system", sid=sid)
            jb2 = _wait(c, sid)
            lines2 = jb2.get("lines") or []
            for ln in lines2[:6]:
                print(f"       | {ln}")
            check("계통도 열기 완료", jb2.get("state") == "done",
                  str(jb2.get("error"))[:80])
            check("정찰 로그가 없다",
                  not any(l.startswith("[정찰]") for l in lines2),
                  f"{sum(1 for l in lines2 if l.startswith('[정찰]'))}줄")
            rv2 = c.get(f"/api/module-f/recon?sid={sid}").get_json()
            check("계통도 슬롯의 정찰은 비어 있다",
                  (rv2.get("recon") or {}).get("state") == "none",
                  str(rv2.get("recon"))[:80])

        # ─────────────────────────── ③ A 를 깨뜨려도 열기는 산다
        print("\n[③] remote30_prototype import 를 깨뜨린 채 열기")
        saved = sys.modules.get("remote30_prototype", "<없음>")
        sys.modules["remote30_prototype"] = None   # import 하면 ImportError
        try:
            sid3 = _open(c, SYSTEM if SYSTEM.is_file() else plan)["sid"]
            jb3 = _wait(c, sid3)
            check("열기는 그대로 성공", jb3.get("state") == "done",
                  str(jb3.get("error"))[:90])
            rv3 = c.get(f"/api/module-f/recon?sid={sid3}").get_json()
            r3 = rv3.get("recon") or {}
            check("정찰에는 사유만 남는다", r3.get("state") == "error",
                  str(r3.get("error"))[:90])
            w3 = c.get(f"/api/module-f/world?sid={sid3}").get_json()
            check("찍기는 종전대로 쓸 수 있다",
                  bool((w3.get("state") or {}).get("armed") is not None),
                  str(w3.get("state"))[:70])
        finally:
            if saved == "<없음>":
                sys.modules.pop("remote30_prototype", None)
            else:
                sys.modules["remote30_prototype"] = saved

    print("\n" + "=" * 58)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\nFAIL — {len(FAILS)}건")
        return 1
    print("F-8a 정찰 — 수용기준 전 항목 PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
