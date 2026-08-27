# -*- coding: utf-8 -*-
"""모듈 F 라우트 — 1단계 찍기(재료·헤드)."""
from __future__ import annotations

import time

from flask import jsonify, request

from routes.module_f.common import _fail
from routes.module_f.jobs import _job_running, _run_job, _sess
from routes.module_f.remote30 import _sheet_frames
from routes.module_f.views import _pick_state


def register(app):
    # ─────────────────────────────────────────── 1. 찍기
    @app.post("/api/module-f/pick/mode")
    def module_f_pick_mode():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        action = str(body.get("action") or "")
        if action == "pipe":
            ok = ps.select_pipe()
            msg = "배관(재료)을 찍으세요. 레이어×색 단위로 잡힙니다."
        elif action == "complete":
            ok = ps.complete_pipe()
            msg = ("재료 선택 완료 — 이제 헤드를 찍습니다."
                   if ok else "재료를 하나 이상 찍어야 완료할 수 있습니다.")
        elif action == "slot":
            ok = ps.set_slot(body.get("slot"))
            msg = (f"헤드 칸 = {ps.head_label}" if ok
                   else "재료 선택을 먼저 완료하세요.")
        else:
            return _fail(f"모르는 동작입니다: {action}")
        return jsonify({"ok": True, "applied": bool(ok), "message": msg,
                        "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/auto")
    def module_f_pick_auto():
        """모듈 A 의 레이어 사전이 고른 묶음을 한 번에 찍는다.

        `board.mat` 에 직접 밀어넣지 않고 **그 묶음의 실제 선분 중점**으로
        정상 클릭 경로(`PickSession.click`)를 태운다. 그래야 클릭 기록·되돌리기
        ·스펙 저장이 사람이 찍은 것과 완전히 같은 상태가 된다.
        """
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        want = str(body.get("cat") or "PIPE").upper()
        if want not in {"PIPE", "HEAD", "ALARM"}:
            return _fail(f"추천 카테고리가 아닙니다: {want}")

        world = sess.get("world") or {}
        targets = [b for b in (world.get("bundles") or []) if b.get("cat") == want]
        if not targets:
            return _fail(f"{want} 로 추천된 레이어가 없습니다. 직접 찍어 주세요.")

        if want == "PIPE":
            ps.select_pipe()
        else:
            if not ps.mat_done:
                return _fail("재료 선택을 먼저 완료해야 헤드를 찍을 수 있습니다.")
            ps.set_slot(ps.head_label)

        applied, skipped = [], []
        for b in targets:
            segs = ps.board.by_bundle.get((b["layer"], b["color"])) or []
            if not segs:
                skipped.append(b["layer"])
                continue
            a, c = segs[0]
            rep = ps.click((a[0] + c[0]) / 2.0, (a[1] + c[1]) / 2.0)
            if rep is None or rep.get("동작") != "추가":
                skipped.append(b["layer"])
            else:
                applied.append(b["layer"])
        return jsonify({
            "ok": True, "applied": applied, "skipped": skipped,
            "message": (f"{want} 추천 {len(applied)}묶음을 찍었습니다."
                        + (f" ({len(skipped)}묶음은 이미 찍혀 있거나 건너뜀)"
                           if skipped else "")),
            "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/click")
    def module_f_pick_click():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        try:
            x = float(body.get("x"))
            y = float(body.get("y"))
        except (TypeError, ValueError):
            return _fail("클릭 좌표가 올바르지 않습니다.")
        max_d = body.get("max_d")
        max_d = float(max_d) if max_d is not None else None
        rep = ps.click(x, y, max_d=max_d)
        return jsonify({"ok": True, "report": rep,
                        "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/suggest")
    def module_f_pick_suggest():
        """[F-5 · D2] 찍기 후보 제안 — 모듈 A 의 detect_heads(R1~R5·신뢰도).

        후보를 «제안만» 한다. board 는 여기서 절대 바뀌지 않는다 — 반영은
        화면이 기존 찍기 경로(사람 클릭과 같은 /pick/click)로만 한다.
        E 의 「표시가 없으면 추측하지 않는다」 확정 게이트를 우회하는 별도
        주입 경로를 만들지 않는 것이 D2 의 요점이다.

        인식이 실패해도(A import 불가 포함) 찍기는 종전대로 동작한다 —
        잡이 error 로 끝날 뿐 세션은 멀쩡하다.
        """
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("pick") is None:
            return _fail("찍기 세션이 없습니다.")
        dxf = sess.get("dxf")
        if not dxf:
            return _fail("올린 DXF 가 없는 세션입니다 — 후보 제안은 "
                         "도면을 올린 찍기 세션에서만 됩니다.")
        if _job_running(sess):
            return _fail("이미 작업이 돌고 있습니다. 끝난 뒤에 다시 눌러 주세요.",
                         409)

        def job():
            # [F-8a] 인식 자체는 `recon.run_recon` 하나뿐이다 — 열기 잡의 정찰과
            # 같은 것을 쓴다. 둘이 각자 A 를 부르면 언젠가 한쪽만 고쳐져 카드와
            # 찍기 화면이 서로 다른 후보 수를 말하게 된다.
            from routes.module_f.recon import run_recon
            rec = run_recon(dxf, world=sess.get("world"), tag="제안")
            cands = rec["heads"]
            sess["suggest"] = cands
            # 정찰을 안 돌린 세션(옛 흐름)이라면 이 결과를 그대로 정찰로도 쓴다.
            sess.setdefault("recon", rec)
            return {"ok": True, "n": len(cands), "bands": rec["bands"],
                    "candidates": cands}

        _run_job(sess, "찍기 후보 제안", job)
        return jsonify({"ok": True})

    @app.post("/api/module-f/pick/undo")
    def module_f_pick_undo():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        undone = ps.undo()
        return jsonify({"ok": True, "undone": undone,
                        "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/commit")
    def module_f_pick_commit():
        """찍은 스펙을 저장하고, 그 스펙으로 1~6단계를 다시 돌려 손질망을 만든다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        if not ps.mat_done:
            return _fail("재료(배관) 선택을 완료해야 다음으로 넘어갈 수 있습니다.")
        if _job_running(sess):
            return _fail("이미 작업이 돌고 있습니다. 끝난 뒤에 다시 눌러 주세요.", 409)

        def job():
            from services.cad_import.edit.session import EditSession
            t0 = time.perf_counter()
            spec_path = ps.commit()
            print(f"[찍기] 스펙 저장 — {spec_path}")
            print("[손질] 찍은 스펙으로 배관망을 다시 구성하는 중…")
            es = EditSession.open(ps.key, out_dir=None, load_saved=False,
                                  use_cache=False)
            sess["edit"] = es
            sess["sheets"] = _sheet_frames(es.board)
            print(f"[손질] 완료 {time.perf_counter() - t0:.1f}s · "
                  f"노드 {len(es.board.pts)} · 간선 {len(es.board.edges)} · "
                  f"헤드 {len(es.board.disks)}")
            return {"spec_path": spec_path}

        _run_job(sess, "배관망 구성", job)
        return jsonify({"ok": True})
