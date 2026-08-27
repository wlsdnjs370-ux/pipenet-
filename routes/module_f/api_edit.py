# -*- coding: utf-8 -*-
"""모듈 F 라우트 — 2단계 손질(이음·삭제·급수·종류·자동 이음·최불리)."""
from __future__ import annotations

import os
import time

from flask import jsonify, request

from routes.module_f.common import REMOTE_K_DEFAULT, _fail, _r1
from routes.module_f.graph import _autojoin_apply, _autojoin_scan
from routes.module_f.jobs import _job_running, _run_job, _sess
from routes.module_f.remote30 import _worst_k_heads
from routes.module_f.views import _edit_state


def register(app):
    @app.get("/api/module-f/worst/reference-counts")
    def module_f_reference_counts():
        """NFTC 103 표 2.1.1.1 — 기준개수를 고를 수 있게 그대로 내려보낸다.

        표를 화면에 옮겨 적지 않는다. 법정 수치를 두 곳에 두면 개정이 왔을 때
        한쪽만 고쳐지고, 그 어긋남은 산출로만 드러난다 — `core/nftc_rules.py`
        가 유일한 출처다.
        """
        try:
            from nftc_rules import reference_count_options
            rows = reference_count_options()
        except Exception as exc:  # noqa: BLE001 — 표가 없어도 직접 입력은 된다
            print(f"[최불리] 기준개수 표를 읽지 못했습니다: {exc}")
            return jsonify({"ok": True, "rows": [], "default": REMOTE_K_DEFAULT,
                            "source": None,
                            "message": f"표를 읽지 못했습니다 — 직접 입력하세요: {exc}"})
        return jsonify({"ok": True, "rows": rows,
                        "default": REMOTE_K_DEFAULT,
                        "source": "NFTC 103 표 2.1.1.1"})

    # ─────────────────────────────────────────── 2. 손질
    @app.get("/api/module-f/edit/state")
    def module_f_edit_state():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("edit") is None:
            return _fail("손질 세션이 없습니다.")
        # 화면이 처음 여는 길이다 — 제 사본이 없으므로 망 도형을 반드시 싣는다.
        return jsonify({"ok": True, "key": sess["key"],
                        "state": _edit_state(sess, full=True)})

    @app.post("/api/module-f/edit/mode")
    def module_f_edit_mode():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        from services.cad_import.edit.session import (
            MODE_DELETE, MODE_JOIN, MODE_SOURCE, MODE_VALVE)
        allowed = {MODE_JOIN, MODE_DELETE, MODE_SOURCE, MODE_VALVE}
        mode = str(body.get("mode") or "")
        if mode not in allowed:
            return _fail(f"모르는 손질 모드입니다: {mode}")
        es.set_mode(mode)
        return jsonify({"ok": True, "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/click")
    def module_f_edit_click():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        try:
            x = float(body.get("x"))
            y = float(body.get("y"))
            max_d = float(body.get("max_d"))
        except (TypeError, ValueError):
            return _fail("클릭 좌표가 올바르지 않습니다.")
        rep = es.click(x, y, max_d)
        if rep and rep.get("동작") not in ("헤드선택",):
            # 망이 바뀌면 앞서 잡아 둔 물길·최불리 선정은 더 이상 사실이 아니다.
            sess["water_path"] = None
            sess["worst"] = None
        return jsonify({"ok": True, "report": rep,
                        "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/kind")
    def module_f_edit_kind():
        """고른 헤드의 종류를 덮는다. 미지정이 남으면 변환이 막힌다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        from services.cad_import.kinds import CONFIRMED_KINDS
        kind = str(body.get("kind") or "")
        if kind not in CONFIRMED_KINDS:
            return _fail(f"헤드 종류가 아닙니다: {kind}")
        applied = es.set_kind(kind)
        if applied is None:
            return _fail("먼저 헤드를 하나 고르세요.")
        return jsonify({"ok": True, "applied": applied,
                        "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/undo")
    def module_f_edit_undo():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        ok = es.undo()
        if ok:
            sess["water_path"] = None
            sess["worst"] = None
            # 자동 이음을 되돌렸을 수도 있다 — 지난 결과를 남겨두면 거짓말이 된다.
            sess["autojoin_report"] = None
        return jsonify({"ok": True, "undone": bool(ok),
                        "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/autojoin/scan")
    def module_f_edit_autojoin_scan():
        """끊긴 관 끝을 짝짓고 이을 여유를 도면에서 잰다. 아직 붙이지 않는다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        try:
            force = body.get("eps_mm")
            force = None if force in (None, "", 0) else float(force)
        except (TypeError, ValueError):
            force = None
        t0 = time.perf_counter()
        scan = _autojoin_scan(es.board, force_eps=force)
        sess["autojoin"] = scan
        sess["aj_seq"] = sess.get("aj_seq", 0) + 1
        sess["autojoin_report"] = None
        best = next((t for t in scan["trials"]
                     if t["eps_mm"] == scan["eps_mm"]), scan["trials"][-1])
        print(f"[자동이음] 관 끝 {scan['ends']} · 여유 {scan['eps_mm']}mm "
              f"({'직접 지정' if force else '도면 실측'}) · 후보 "
              f"{len(scan['cands'])}곳 {scan['by_kind']} · "
              f"{time.perf_counter() - t0:.1f}s")
        msg = (f"여유 {scan['eps_mm']:.0f}mm(도면 실측) · 이을 곳 "
               f"{len(scan['cands'])}군데 · 덩이 {scan['bodies_before']} → "
               f"{best['bodies']} 예상 · 가까운 짝 {scan['near']}쌍 중 "
               f"{scan['kept']}쌍이 관 방향과 맞음")
        if scan["dropped"]:
            msg += f" · 후보 {scan['dropped']}곳은 상한을 넘어 제외"
        if not scan["cands"]:
            msg = ("이을 만한 끊긴 자리를 찾지 못했습니다 — 이미 이어져 있거나 "
                   "틈이 800mm 보다 넓습니다.")
        return jsonify({"ok": True, "message": msg,
                        "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/autojoin/apply")
    def module_f_edit_autojoin_apply():
        """후보를 모듈 E 의 이음 판정에 태운다. 무거워서 잡으로 돌린다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        scan = sess.get("autojoin")
        if not scan or not scan.get("cands"):
            return _fail("먼저 «끊긴 곳 찾기» 를 눌러 후보를 뽑으세요.")
        if _job_running(sess):
            return _fail("이미 작업이 돌고 있습니다. 끝난 뒤에 다시 눌러 주세요.", 409)

        def job():
            rep = _autojoin_apply(es.board, scan)
            # 망이 바뀌었으니 앞서 잡아 둔 물길·최불리는 더 이상 사실이 아니다.
            sess["water_path"] = None
            sess["worst"] = None
            sess["autojoin"] = None
            sess["aj_seq"] = sess.get("aj_seq", 0) + 1
            sess["autojoin_report"] = rep
            return rep

        _run_job(sess, "자동 이음", job)
        return jsonify({"ok": True, "sid": sess["id"]})

    @app.post("/api/module-f/edit/autojoin/clear")
    def module_f_edit_autojoin_clear():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("edit") is None:
            return _fail("손질 세션이 없습니다.")
        sess["autojoin"] = None
        sess["aj_seq"] = sess.get("aj_seq", 0) + 1
        return jsonify({"ok": True, "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/flow")
    def module_f_edit_flow():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        state = es.flow()
        if state is None:
            return _fail("급수 시작 위치를 먼저 찍어야 물흐름을 볼 수 있습니다.")
        # 브라우저에는 연출 프레임을 돌리지 않는다 — 끝까지 돌려 최종 상태로 둔다.
        while es.flow_tick():
            pass
        pts = es.board.pts
        sess["water_path"] = [
            [_r1(pts[a][0]), _r1(pts[a][1]), _r1(pts[b][0]), _r1(pts[b][1])]
            for a, b in state["wet_edges"]]
        return jsonify({
            "ok": True,
            "water": {"wet_heads": len(state["wet_heads"]),
                      "total_heads": state["total_heads"],
                      "wet_edges": len(state["wet_edges"]),
                      "reach": len(state["reach"])},
            "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/worst")
    def module_f_edit_worst():
        """Remote 30 — 급수원에서 가장 불리한 K 헤드와 그 경로를 고른다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        b = es.board
        if not b.sources:
            return _fail("급수 시작 위치를 먼저 찍어야 최불리 헤드를 고를 수 있습니다.")
        try:
            k = int(body.get("k") or REMOTE_K_DEFAULT)
        except (TypeError, ValueError):
            k = REMOTE_K_DEFAULT
        k = max(1, min(k, 200))

        # 한 파일에 도면이 여러 장이면 최불리 K 가 서로 다른 도면의 헤드를 섞어
        # 뽑는다 — 모듈 A 가 실측으로 겪은 그 문제다. 장을 고르면 그 장 안에서만.
        only, sheet_no = None, None
        sheets = sess.get("sheets") or []
        try:
            want = int(body.get("sheet") or 0)
        except (TypeError, ValueError):
            want = 0
        if want and sheets:
            hit = next((f for f in sheets if int(f.get("index", 0)) == want), None)
            if hit is None:
                return _fail(f"그런 도면 장이 없습니다: {want}")
            x0, y0, x1, y1 = [float(v) for v in hit["bbox"]]
            only = {hi for hi, d in enumerate(b.disks)
                    if x0 <= float(d[0]) <= x1 and y0 <= float(d[1]) <= y1}
            sheet_no = want
            print(f"[최불리] 도면 {want} 장 안으로 범위를 좁힘 — 헤드 {len(only)}개")

        # ── 영역 지정 (모듈 A 의 zones) — 사람이 사각형으로 후보를 가둔다.
        #
        # 도면 장 나누기는 «자동으로 잰 경계» 라 실무에서 늘 맞지는 않는다.
        # 한 층에 방화구획이 여럿이거나, 계산에서 빼야 할 구역(주차장·기계실)이
        # 섞여 있으면 앵커가 그리로 튄다 — 그때는 사람이 직접 가두는 수밖에 없다.
        # A 가 같은 이유로 zones 를 갖는다.
        #
        # 여러 개면 합집합이다(∪). 장 선택과 함께 쓰면 교집합이 된다 —
        # «이 장의 이 구역» 이 자연스러운 읽기다.
        zones = body.get("zones")
        if zones:
            try:
                rects = []
                for z in zones:
                    x0, y0, x1, y1 = (float(z[0]), float(z[1]),
                                      float(z[2]), float(z[3]))
                    rects.append((min(x0, x1), min(y0, y1),
                                  max(x0, x1), max(y0, y1)))
            except (TypeError, ValueError, IndexError):
                return _fail("영역 좌표가 올바르지 않습니다 "
                             "([[x0,y0,x1,y1], …] 형식).")
            if not rects:
                return _fail("영역이 비었습니다.")
            in_zone = {hi for hi, d in enumerate(b.disks)
                       if any(x0 <= float(d[0]) <= x1 and y0 <= float(d[1]) <= y1
                              for x0, y0, x1, y1 in rects)}
            if not in_zone:
                return _fail(f"영역 {len(rects)}곳 안에 헤드가 없습니다. "
                             "영역을 다시 그리세요.")
            only = in_zone if only is None else (only & in_zone)
            if not only:
                return _fail("고른 도면 장과 영역이 겹치는 헤드가 없습니다.")
            print(f"[최불리] 영역 {len(rects)}곳으로 범위를 좁힘 — 헤드 {len(only)}개")

        # [F-1 · D4] 급수원이 여럿이면 «어느 하나 기준» 인지 사람이 정한다 —
        # 전체망 .kfp 변환의 source_selection_required 와 같은 규약(태그 Z1…
        # 또는 1부터 번호). 어느 급수원에서든 가장 먼 헤드는 앵커가 못 된다:
        # 급수원마다 최원 유하거리가 다르기 때문이다(G BLOCKED B2 — 이것으로 해소).
        src_index = None
        picked_tag = None
        cands = [{"tag": f"Z{i + 1}",
                  "x": _r1(b.pts[n][0]), "y": _r1(b.pts[n][1])}
                 for i, n in enumerate(b.sources)
                 if isinstance(n, int) and 0 <= n < len(b.pts)]
        want_src = str(body.get("source") or "").strip()
        # ★명시가 우선이다 — 1곳뿐이어도 틀린 태그를 조용히 무시하면, 사용자는
        #   Z2 기준을 골랐다고 믿은 채 Z1 결과를 읽게 된다.
        if want_src:
            for i in range(len(b.sources)):
                if want_src.upper() == f"Z{i + 1}" or want_src == str(i + 1):
                    src_index, picked_tag = i, f"Z{i + 1}"
                    break
            if src_index is None:
                return jsonify({"ok": False, "code": "source_selection_required",
                                "message": f"급수원 '{want_src}'를 찾지 못했습니다.",
                                "sources": cands}), 400
        elif len(b.sources) == 1:
            src_index, picked_tag = 0, "Z1"      # 1곳이면 자동으로 그것
        else:
            return jsonify({"ok": False, "code": "source_selection_required",
                            "message": "급수원이 여러 곳입니다. 어느 급수원 기준의 "
                                       "최불리인지 하나를 지정하세요.",
                            "sources": cands}), 400

        w = _worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=k,
                           only_heads=only, source_index=src_index)
        if not w["heads"]:
            sess["worst"] = None
            return _fail("급수원에서 닿는 헤드가 없습니다. 이음·급수 위치를 확인하세요.")
        w["sheet"] = sheet_no
        w["source_tag"] = picked_tag          # 화면이 «어느 급수원 기준» 인지 안다
        w["source_index"] = src_index
        w["zones"] = [list(r) for r in rects] if zones else []
        w["candidates"] = len(only) if only is not None else w["reachable"]
        sess["worst"] = w
        sess["worst_zones"] = w["zones"]      # 다시 누를 때 같은 영역을 쓴다
        return jsonify({
            "ok": True,
            "summary": {"k": len(w["heads"]), "reachable": w["reachable"],
                        "far_m": w["far_m"], "near_m": w["near_m"],
                        "span_m": w.get("span_m", 0.0),
                        "total_m": w.get("total_m", 0.0),
                        "max_load": w.get("max_load", 0),
                        "sheet": sheet_no,
                        "source": picked_tag,
                        "zones": len(w["zones"]),
                        "candidates": w["candidates"],
                        # 최원 유하거리 «경로» — 그 거리가 어느 줄인지.
                        "anchor_path_m": w.get("anchor_path_m", 0.0),
                        "anchor_path_nodes": len(w.get("anchor_path") or ()),
                        "path_edges": len(w["edges"])},
            "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/worst-clear")
    def module_f_edit_worst_clear():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("edit") is None:
            return _fail("손질 세션이 없습니다.")
        sess["worst"] = None
        return jsonify({"ok": True, "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/save")
    def module_f_edit_save():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        path = es.commit()
        # 파일 이름만 돌려준다 — 서버 폴더 구조는 밖으로 나갈 이유가 없다.
        return jsonify({"ok": True, "file": os.path.basename(path),
                        "message": f"유저손질을 저장했습니다: {os.path.basename(path)}"})
