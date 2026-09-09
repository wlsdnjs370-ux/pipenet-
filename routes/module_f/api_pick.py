# -*- coding: utf-8 -*-
"""모듈 F 라우트 — 1단계 찍기(재료·헤드)."""
from __future__ import annotations

import time

from flask import jsonify

from routes.module_f.common import (
    COORD_LIMIT_MM, _check_num, _check_xy, _fail)
from routes.module_f.jobs import _job_running, _run_job, _sess, route_session
from routes.module_f.remote30 import _sheet_frames
from routes.module_f.views import _pick_state


def _need_pick(body):
    """찍기판이 선 세션인가 — 이 파일의 라우트 일곱 곳이 같은 검사를 했다.

    찍기 라우트는 예외 없이 `sess["pick"]` 을 만진다. 없는데 들어오면 하나같이
    같은 문장으로 400 을 냈으므로, 그 판단을 한 자리로 모은다.
    """
    sess = _sess(body.get("sid"))
    if sess.get("pick") is None:
        return sess, "찍기 세션이 없습니다."
    return sess, None


def _apply_fold(sess, ps, es) -> dict:
    """[신축배관 접기] 지정된 레이어의 가닥을 손질판에서 직선으로 편다.

    지정이 없으면 **아무것도 안 한다** — 그러면 산출이 한 바이트도 안 바뀐다
    (D-F11-1: 사람의 명시적 수정 외에는 산출을 바꾸지 않는다).

    접은 결과는 세션에 남겨 화면·산출 기록이 그대로 읽는다. 못 접은 가닥이
    있으면 **개수와 이유**가 함께 남는다 — 조용히 넘기지 않는다(S340).
    """
    want = [str(h.get("layer")) for h in (sess.get("pick_layer_folded") or ())
            if h.get("on")]
    if not want:
        sess.pop("fold", None)
        return {}
    from routes.module_f import flexfold as ff
    pieces = ff.flex_pieces(ps.board.w, want)
    chains, tangled, loose = ff.strands(pieces)
    got = ff.fold_board(es.board, chains)
    # 선언 길이는 손질판이 들고 간다 — `convert_payload()` 가 그대로 실어 보낸다.
    es.board.edge_len_mm = dict(got["edge_len_mm"])
    rec = {
        "layers": want, "pieces": len(pieces), "strands": len(chains),
        "tangled": len(tangled), "loose": len(loose),
        "folded_runs": got["folded"], "cuts": got.get("cuts", 0),
        "skipped": got["skipped"], "why": dict(got["why"]),
        "total_mm": round(sum(ff.run_length(p) for p in chains), 1),
        "declared_mm": round(sum(got["edge_len_mm"].values()), 1),
        "edges_before": got["before"], "edges_after": got["after"],
    }
    sess["fold"] = rec
    print(f"[접기] {want} · 가닥 {rec['strands']}"
          f" → 접힌 구간 {rec['folded_runs']}"
          f" (분기에서 자른 곳 {rec['cuts']})"
          f" · 간선 {rec['edges_before']} → {rec['edges_after']}"
          f" · 총연장 {rec['total_mm']:,.0f} mm 중 {rec['declared_mm']:,.0f} 선언")
    if rec["skipped"]:
        print(f"[접기] ★못 접은 가닥 {rec['skipped']} — {rec['why']}")
    if tangled or loose:
        print(f"[접기] ★사슬이 아닌 것 {len(tangled)} · 어디에도 안 든 조각"
              f" {len(loose)} — 그대로 둡니다")
    return rec


def register(app):
    # ─────────────────────────────────────────── 1. 찍기
    @app.post("/api/module-f/pick/mode")
    @route_session(_need_pick, post=True, why_code=400)
    def module_f_pick_mode(sess, body):
        ps = sess["pick"]
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
    @route_session(_need_pick, post=True, why_code=400)
    def module_f_pick_auto(sess, body):
        """모듈 A 의 레이어 사전이 고른 묶음을 한 번에 찍는다.

        `board.mat` 에 직접 밀어넣지 않고 **그 묶음의 실제 선분 중점**으로
        정상 클릭 경로(`PickSession.click`)를 태운다. 그래야 클릭 기록·되돌리기
        ·스펙 저장이 사람이 찍은 것과 완전히 같은 상태가 된다.
        """
        ps = sess["pick"]
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

        # [F-8b] 클릭 몸통은 `adopt.adopt_bundles` 하나다 — 채택도 이 길로 간다.
        from routes.module_f.adopt import adopt_bundles
        got = adopt_bundles(ps, world, want)
        applied, skipped = got["applied"], got["skipped"]
        return jsonify({
            "ok": True, "applied": applied, "skipped": skipped,
            "message": (f"{want} 추천 {len(applied)}묶음을 찍었습니다."
                        + (f" ({len(skipped)}묶음은 이미 찍혀 있거나 건너뜀)"
                           if skipped else "")),
            "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/click")
    @route_session(_need_pick, post=True, why_code=400)
    def module_f_pick_click(sess, body):
        ps = sess["pick"]
        try:
            x, y = _check_xy(body, what="클릭")
            # 안 주면 None — 엔진이 제 기본값을 쓴다(종전 규약).
            max_d = (None if body.get("max_d") is None else
                     _check_num(body.get("max_d"), "허용 거리",
                                lo=0.0, hi=COORD_LIMIT_MM))
        except ValueError as exc:
            return _fail(str(exc))
        rep = ps.click(x, y, max_d=max_d)
        return jsonify({"ok": True, "report": rep,
                        "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/suggest")
    @route_session(_need_pick, post=True, why_code=400)
    def module_f_pick_suggest(sess, body):
        """[F-5 · D2] 찍기 후보 제안 — 모듈 A 의 detect_heads(R1~R5·신뢰도).

        후보를 «제안만» 한다. board 는 여기서 절대 바뀌지 않는다 — 반영은
        화면이 기존 찍기 경로(사람 클릭과 같은 /pick/click)로만 한다.
        E 의 「표시가 없으면 추측하지 않는다」 확정 게이트를 우회하는 별도
        주입 경로를 만들지 않는 것이 D2 의 요점이다.

        인식이 실패해도(A import 불가 포함) 찍기는 종전대로 동작한다 —
        잡이 error 로 끝날 뿐 세션은 멀쩡하다.
        """
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

    # ── 신축배관 등 «배관 형상이 아닌» 레이어를 재료에서 뺀다 [지시서 C-1]
    #
    #   ★조용히 빼지 않는다. 자동 사전은 도면마다 다른 관례를 **반드시**
    #     놓치고(사내 도면 8장에서 5종), 조용히 빼면 배관이 끊긴 것을 아무도
    #     모른다(S340). 그래서 추천만 하고 확정은 사람이 한다.
    #
    #   ★그리고 빼면 무엇을 잃는지 **먼저 잰다.** 대명동 실측: 신축배관을 빼면
    #     물닿음 헤드가 111 → 5 로 떨어진다. 신축배관이 헤드를 가지관에 잇는
    #     유일한 경로이기 때문이다. 그 사실을 모른 채 빼면 망이 무너진다.
    FLEX_WORDS = ("후렉시블", "후렉", "플렉", "flex", "fx")

    def _is_flex(layer) -> bool:
        low = str(layer or "").lower()
        return any(w.lower() in low for w in FLEX_WORDS)

    @app.get("/api/module-f/pick/materials")
    @route_session(_need_pick)
    def module_f_pick_materials(sess, body):
        """찍힌 재료 묶음 — 레이어별로 묶어 «신축배관 추천» 을 붙인다.

        ★재료에서 «빼는» 길은 없앴다(어제의 `exclude-layer`). 빼면 물닿음 헤드가
          111 → 5 가 된다 — 신축배관이 헤드를 가지관에 잇는 유일한 경로다.
          대신 «접기» 를 지정한다: 길이와 연결은 그대로 두고 굴곡만 편다.
        """
        ps = sess["pick"]
        pool = getattr(ps.board, "by_bundle", {}) or {}
        by: dict = {}

        def _row(ly):
            return by.setdefault(str(ly), {"layer": str(ly), "colors": [],
                                           "segs": 0, "flex": _is_flex(ly),
                                           "on": False, "fold": False,
                                           "strands": 0, "fold_mm": 0.0})

        for ly, col in ps.board.mat:
            row = _row(ly)
            row["on"] = True
            row["colors"].append(col)
            row["segs"] += len(pool.get((ly, col), ()))
        # 접기로 지정된 레이어를 표시한다 — 재료에서 «빠지는» 것이 아니라
        # 그대로 있고, 굴곡만 펴진다. 그래서 `on` 은 그대로 두고 칸을 따로 둔다.
        fold = {str(h.get("layer")): h
                for h in (sess.get("pick_layer_folded") or ())
                if h.get("on")}
        for ly, h in fold.items():
            row = _row(ly)
            row["fold"] = True
            row["strands"] = int(h.get("strands") or 0)
            row["fold_mm"] = float(h.get("total_mm") or 0.0)
        rows = sorted(by.values(), key=lambda r: (-r["segs"], r["layer"]))
        return jsonify({"ok": True, "materials": rows,
                        "flex_hint": [r["layer"] for r in rows if r["flex"]]})

    @app.post("/api/module-f/pick/fold-layer")
    @route_session(_need_pick, post=True)
    def module_f_pick_fold_layer(sess, body):
        """레이어 하나를 «접기» 로 지정하거나 되돌린다(`on`: false 면 되돌리기).

        ★어제의 `exclude-layer`(재료에서 **빼기**)는 없앴다. 빼면 대명동에서
          물닿음 헤드가 111 → 5 가 된다(실측) — 신축배관이 헤드를 가지관에 잇는
          유일한 경로이기 때문이다. 남겨 두면 언젠가 누가 누른다.

        여기서는 «지정» 만 한다. 실제로 접는 것은 찍기를 확정한 뒤 손질판에서다
        (`pick/commit`) — 그때라야 접을 간선이 존재한다.

        Body: sid · layer · [on]
        """
        ps = sess["pick"]
        layer = str(body.get("layer") or "")
        if not layer:
            return _fail("어느 레이어인지 주세요.")
        want_on = body.get("on")
        want_on = True if want_on is None else bool(want_on)
        pool = getattr(ps.board, "by_bundle", {}) or {}
        keys = [k for k in pool if str(k[0]) == layer]
        if not keys:
            return _fail(f"그런 레이어가 없습니다: {layer}")

        # 접을 가닥이 실제로 몇 개인지 **미리 세어** 돌려준다 — 누르기 전에
        # 무엇이 일어나는지 알아야 사람이 정할 수 있다.
        from routes.module_f import flexfold as ff
        pieces = ff.flex_pieces(ps.board.w, [layer])
        chains, tangled, rest = ff.strands(pieces)
        total = sum(ff.run_length(p) for p in chains)

        hist = sess.setdefault("pick_layer_folded", [])
        rec = {"layer": layer, "on": want_on, "bundles": len(keys),
               "pieces": len(pieces), "strands": len(chains),
               "tangled": len(tangled), "loose": len(rest),
               "total_mm": round(total, 1)}
        hist[:] = [h for h in hist if h.get("layer") != layer] + [rec]
        # 조용히 넘기지 않는다 — 무엇을 접기로 했는지 세션과 로그에 남는다.
        #
        #   ★찍기 기록(`board.clicks`)에는 넣지 않는다. 그 목록의 각 항목은
        #     좌표를 가진 «클릭» 이고 `highlight_geom()` 이 마지막 항목에서
        #     `cl["x"]` 를 읽는다 — 좌표 없는 항목을 끼우면 500 이다(실측).
        print(f"[찍기] 레이어 {'접기 지정' if want_on else '접기 해제'} — {layer}"
              f" · 조각 {len(pieces)} → 가닥 {len(chains)}"
              f" · 총연장 {total:,.0f} mm"
              + (f" · 사슬 아님 {len(tangled)}" if tangled else ""))
        return jsonify({"ok": True, **rec, "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/adopt")
    @route_session(_need_pick, post=True, why_code=400)
    def module_f_pick_adopt(sess, body):
        """[F-8b] 정찰 결과를 찍기 스펙으로 채택한다 — 전 과정이 클릭 경로다.

        body: {sid, materials: true, heads: {conf_min: 0.9} | {indices: [...]}}

        board 에 픽을 넣는 코드는 이 라우트 어디에도 없다(D-F8-3). 재료는
        `/pick/auto` 와 같은 함수로 묶음 중점을 클릭하고, 헤드는 후보 좌표마다
        `PickSession.click` 을 태운다. 그래서 채택 직후 undo 가 마지막 채택
        클릭을 사람 클릭과 똑같이 되돌린다(D-F8-5).

        후보 수백~수천 개 클릭은 무거우므로 잡으로 돌린다(F-8b-5).
        """
        from routes.module_f.adopt import (
            ADOPT_MAX_D_MM, adopt_bundles, adopt_heads, select_heads)

        ps = sess["pick"]
        if _job_running(sess):
            return _fail("이미 작업이 돌고 있습니다. 끝난 뒤에 다시 눌러 주세요.",
                         409)

        rec = sess.get("recon") or {}
        if rec.get("error"):
            return _fail(f"정찰이 실패한 도면입니다 — 직접 찍어 주세요. "
                         f"({rec['error']})")

        want_mat = bool(body.get("materials", True))
        hspec = body.get("heads")
        if hspec is None:
            hspec = {}
        elif not isinstance(hspec, dict):
            return _fail("heads 는 {conf_min: …} 또는 {indices: [...]} 입니다.")
        try:
            conf_min = (None if hspec.get("conf_min") is None
                        else float(hspec["conf_min"]))
            indices = (None if hspec.get("indices") is None
                       else [int(i) for i in hspec["indices"]])
            max_d = float(body.get("max_d") or ADOPT_MAX_D_MM)
        except (TypeError, ValueError):
            return _fail("채택 조건이 올바르지 않습니다.")
        if not want_mat and conf_min is None and indices is None:
            return _fail("채택할 것이 없습니다 — 재료나 헤드 중 하나는 골라야 합니다.")

        cands = list(rec.get("heads") or ())
        world = sess.get("world") or {}

        def job():
            out = {"ok": True, "mat_applied": [], "mat_skipped": [],
                   "head_applied": 0, "head_already": 0, "head_skipped": 0,
                   "skipped_heads": [], "clicked": []}
            if want_mat:
                ps.select_pipe()
                got = adopt_bundles(ps, world, "PIPE")
                out["mat_applied"] = got["applied"]
                out["mat_skipped"] = got["skipped"]
                print(f"[채택] 재료 {len(got['applied'])}묶음 찍음 "
                      f"(건너뜀 {len(got['skipped'])})")
                # 재료가 1개 이상이어야 헤드 칸이 열린다 — board 의 기존 판정.
                if not ps.complete_pipe():
                    out["ok"] = False
                    out["error"] = ("재료를 하나도 못 찍었습니다 — "
                                    "배관 레이어를 직접 찍어 주세요.")
                    print("[채택] 재료 0묶음 — 헤드 채택은 건너뜁니다")
                    out["state"] = _pick_state(sess)
                    return out

            if conf_min is not None or indices is not None:
                if not ps.mat_done:
                    out["ok"] = False
                    out["error"] = "재료 선택을 먼저 완료해야 헤드를 찍습니다."
                    out["state"] = _pick_state(sess)
                    return out
                ps.set_slot(ps.head_label)
                picks = select_heads(cands, conf_min=conf_min,
                                     indices=indices)
                print(f"[채택] 헤드 후보 {len(picks)}/{len(cands)}개를 찍는 중… "
                      f"(칸 {ps.head_label} · 허용 {max_d:.0f}mm)")

                def say(n, total, ok, dup, bad):
                    print(f"[채택] {n}/{total} — 찍힘 {ok} · 이미 {dup} · "
                          f"유령 {bad}")

                got = adopt_heads(ps, picks, max_d=max_d, progress=say)
                out["head_applied"] = got["applied"]
                out["head_already"] = got["already"]
                out["head_skipped"] = len(got["skipped"])
                out["skipped_heads"] = got["skipped"]
                out["clicked"] = got["clicked"]
                print(f"[채택] 완료 — 찍힘 {got['applied']} · "
                      f"이미 반영 {got['already']} · 유령 {len(got['skipped'])}")
            out["state"] = _pick_state(sess)
            return out

        _run_job(sess, "인식 결과 채택", job)
        return jsonify({"ok": True})

    @app.post("/api/module-f/pick/undo")
    @route_session(_need_pick, post=True, why_code=400)
    def module_f_pick_undo(sess, body):
        ps = sess["pick"]
        undone = ps.undo()
        return jsonify({"ok": True, "undone": undone,
                        "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/commit")
    @route_session(_need_pick, post=True, why_code=400)
    def module_f_pick_commit(sess, body):
        """찍은 스펙을 저장하고, 그 스펙으로 1~6단계를 다시 돌려 손질망을 만든다."""
        ps = sess["pick"]
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
            # ★[신축배관 접기] 여기가 두 판이 **동시에** 있는 유일한 자리다.
            #   레이어는 찍기 판에만 살고(손질판은 pts/edges 뿐), 접을 간선은
            #   손질판에만 있다. 그래서 접기는 반드시 이 자리에서 끝난다.
            _apply_fold(sess, ps, es)
            sess["edit"] = es
            sess["sheets"] = _sheet_frames(es.board)
            print(f"[손질] 완료 {time.perf_counter() - t0:.1f}s · "
                  f"노드 {len(es.board.pts)} · 간선 {len(es.board.edges)} · "
                  f"헤드 {len(es.board.disks)}")
            return {"spec_path": spec_path}

        _run_job(sess, "배관망 구성", job)
        return jsonify({"ok": True})
