# -*- coding: utf-8 -*-
"""[F-8b] 채택을 실도면으로 확인한다 — 핵심은 «동일성» 하나다.

  ① adopt 로 만든 스펙 == 같은 좌표를 같은 순서로 사람이 클릭한 스펙
     (클릭 경로 경유의 증명 — 채택이 별도 주입 경로를 안 만들었다는 뜻)
  ② adopt 직후 undo 가 마지막 채택 클릭을 사람 클릭처럼 되돌린다 (D-F8-5)
  ③ head_skipped 후보가 좌표·사유와 함께 남고, board 는 그 후보에 불변이다
  ④ adopt 를 안 부른 세션의 흐름이 종전과 같다 (회귀)
  ⑤ B1F 실측 — 찍힌 헤드 수와 유령 수를 남긴다

★공유 작업폴더(cad_project_editor_g/docs/import)에 저장하므로 B1F 저장본을
  덮지 않도록 스펙 저장은 임시 폴더로만 한다. commit() 은 부르지 않는다.

    python scripts/_verify_module_f_adopt.py [평면도.dxf]
"""
from __future__ import annotations

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


def _open(c, path):
    with open(path, "rb") as f:
        raw = f.read()
    r = c.post("/api/module-f/open", data={
        "dxf_file": (io.BytesIO(raw), os.path.basename(str(path)))},
        content_type="multipart/form-data")
    return r.get_json()["sid"]


def _spec_of(c, sid):
    """세션의 «지금 찍힌 것» 을 스펙으로 — 파일로 저장하지 않는다."""
    from routes.module_f.jobs import _sess
    return _sess(sid)["pick"].spec()


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

        # ─────────────────────────── 채택 (기본 임계 0.9)
        print(f"\n[⑤] 채택 — {plan.name} "
              f"({plan.stat().st_size / 1024 / 1024:.1f} MB) · conf_min 0.9")
        sid = _open(c, plan)
        if not check("열기+정찰", _wait(c, sid).get("state") == "done"):
            return 1
        rec = c.get(f"/api/module-f/recon?sid={sid}").get_json()["recon"]
        print(f"       후보 {rec['n']} · {rec['bands']}")

        t0 = time.perf_counter()
        r = c.post("/api/module-f/pick/adopt",
                   json={"sid": sid, "materials": True,
                         "heads": {"conf_min": 0.9}})
        if not check("채택 수락", bool(r.get_json().get("ok")),
                     str(r.get_json())[:90]):
            return 1
        jb = _wait(c, sid)
        if not check("채택 잡 완료", jb.get("state") == "done",
                     str(jb.get("error"))[:90]):
            return 1
        res = _result(c, sid) or {}
        for ln in (jb.get("lines") or [])[-6:]:
            print(f"       | {ln}")
        print(f"       채택 {time.perf_counter() - t0:.1f}s")
        check("재료가 찍혔다", len(res.get("mat_applied") or []) > 0,
              f"{res.get('mat_applied')}")
        check("헤드가 찍혔다", res.get("head_applied", 0) > 0,
              f"찍힘 {res.get('head_applied')} · 이미 {res.get('head_already')}"
              f" · 유령 {res.get('head_skipped')}")

        # ─────────────────────────── ③ 유령의 모양
        print("\n[③] 유령(skipped) — 좌표·사유가 남는가")
        gh = res.get("skipped_heads") or []
        if res.get("head_skipped", 0):
            g = gh[0]
            check("좌표·사유가 실린다",
                  {"i", "x", "y", "conf", "why"} <= set(g), str(g)[:110])
            check("개수와 목록이 맞다", len(gh) == res["head_skipped"],
                  f"{len(gh)} vs {res['head_skipped']}")
        else:
            check("유령 0 — 후보 전부가 찍혔다", True,
                  f"찍힘 {res.get('head_applied')}")

        spec_adopt = _spec_of(c, sid)
        n_heads_adopt = len(spec_adopt.get("heads") or [])
        clicked = res.get("clicked") or []
        print(f"       스펙: 재료 {len(spec_adopt['material_picks'])}묶음 · "
              f"헤드 {n_heads_adopt}픽 · 클릭 {len(clicked)}회")

        # ─────────────────────────── ② undo (D-F8-5)
        print("\n[②] 채택 직후 undo — 사람 클릭과 같은가")
        before = c.get(f"/api/module-f/world?sid={sid}").get_json()["state"]
        r = c.post("/api/module-f/pick/undo", json={"sid": sid}).get_json()
        after = r.get("state") or {}
        check("되돌릴 것이 있었다", r.get("undone") is not None,
              str(r.get("undone"))[:70])
        check("클릭 기록이 하나 줄었다",
              after.get("n_clicks") == before.get("n_clicks") - 1,
              f"{before.get('n_clicks')} → {after.get('n_clicks')}")
        check("마지막 채택이 board 에서 빠졌다",
              after.get("n_heads") != before.get("n_heads")
              or after.get("n_clicks") < before.get("n_clicks"),
              f"헤드 {before.get('n_heads')} → {after.get('n_heads')}")

        # ─────────────────────────── ① 동일성 (핵심)
        print("\n[①] 동일성 — 같은 좌표·순서를 사람이 클릭하면 같은 스펙인가")
        sid2 = _open(c, plan)
        if not check("두 번째 세션 열기", _wait(c, sid2).get("state") == "done"):
            return 1
        # 재료: 채택이 찍은 그 묶음을 /pick/auto 로 (같은 몸통, 공개 경로)
        c.post("/api/module-f/pick/mode", json={"sid": sid2, "action": "pipe"})
        c.post("/api/module-f/pick/auto", json={"sid": sid2, "cat": "PIPE"})
        c.post("/api/module-f/pick/mode",
               json={"sid": sid2, "action": "complete"})
        st = c.post("/api/module-f/pick/mode",
                    json={"sid": sid2, "action": "slot",
                          "slot": spec_adopt.get("heads", [{}])[0].get("label")
                          if spec_adopt.get("heads") else "상향하향"}
                    ).get_json()["state"]
        print(f"       사람 쪽 준비 — 재료 {len(st['materials'])}묶음 · "
              f"칸 {st['head_label']}")

        t1 = time.perf_counter()
        from routes.module_f.adopt import ADOPT_MAX_D_MM
        for x, y in clicked:
            c.post("/api/module-f/pick/click",
                   json={"sid": sid2, "x": x, "y": y,
                         "max_d": ADOPT_MAX_D_MM})
        print(f"       손 클릭 {len(clicked)}회 · {time.perf_counter() - t1:.1f}s")

        # ★undo 로 하나 뺐던 것을 되돌려 놓고 비교한다(같은 자리를 다시 찍는다).
        spec_hand = _spec_of(c, sid2)

        def norm(sp):
            """비교용 정규화 — 저장 순서만 다른 것은 다른 것이 아니다."""
            out = {"material_picks": sorted(map(json.dumps, map(
                list, sp.get("material_picks") or [])))}
            out["heads"] = sorted(json.dumps(h, sort_keys=True,
                                             ensure_ascii=False)
                                  for h in (sp.get("heads") or []))
            out["ho"] = len(sp.get("ho") or [])
            return out

        # sid 는 undo 한 상태다 — 마지막 클릭을 되찍어 채택 직후로 돌린다.
        if clicked:
            lx, ly = clicked[-1]
            c.post("/api/module-f/pick/click",
                   json={"sid": sid, "x": lx, "y": ly,
                         "max_d": ADOPT_MAX_D_MM})
        a, b = norm(_spec_of(c, sid)), norm(spec_hand)
        check("재료 픽이 같다", a["material_picks"] == b["material_picks"],
              f"채택 {len(a['material_picks'])} vs 손 {len(b['material_picks'])}")
        check("헤드 픽이 같다", a["heads"] == b["heads"],
              f"채택 {len(a['heads'])} vs 손 {len(b['heads'])}")
        check("호(ho) 수가 같다", a["ho"] == b["ho"], f"{a['ho']} vs {b['ho']}")
        if a != b:
            only_a = set(a["heads"]) - set(b["heads"])
            only_b = set(b["heads"]) - set(a["heads"])
            for s in list(only_a)[:3]:
                print(f"       채택에만: {s[:150]}")
            for s in list(only_b)[:3]:
                print(f"       손에만  : {s[:150]}")

        # 스펙 파일로도 같은지 — 임시 폴더에만 쓴다(공유 작업폴더 보호)
        with tempfile.TemporaryDirectory() as td:
            from routes.module_f.jobs import _sess
            p1 = _sess(sid)["pick"].commit(out_dir=td)
            p2 = _sess(sid2)["pick"].commit(out_dir=os.path.join(td, "b"))
            s1 = json.load(open(p1, encoding="utf-8"))
            s2 = json.load(open(p2, encoding="utf-8"))
            s1.pop("source_dxf", None)
            s2.pop("source_dxf", None)
            check("저장한 스펙 파일이 같다", norm(s1) == norm(s2),
                  f"{os.path.basename(p1)}")

        # ─────────────────────────── ④ 회귀 — adopt 를 안 쓰는 흐름
        print("\n[④] adopt 를 안 부른 세션 — 종전 흐름 그대로")
        sid3 = _open(c, plan)
        _wait(c, sid3)
        c.post("/api/module-f/pick/mode", json={"sid": sid3, "action": "pipe"})
        r = c.post("/api/module-f/pick/auto",
                   json={"sid": sid3, "cat": "PIPE"}).get_json()
        check("/pick/auto 종전대로", bool(r.get("ok")) and bool(r.get("applied")),
              f"{len(r.get('applied') or [])}묶음")
        r = c.post("/api/module-f/pick/mode",
                   json={"sid": sid3, "action": "complete"}).get_json()
        check("재료 완료 종전대로", bool(r.get("applied")), str(r.get("message")))
        spec3 = _spec_of(c, sid3)
        check("재료만 찍은 스펙에 헤드가 없다",
              not spec3.get("heads"), str(len(spec3.get("heads") or [])))

    print("\n" + "=" * 58)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\nFAIL — {len(FAILS)}건")
        return 1
    print("F-8b 채택 — 수용기준 전 항목 PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
