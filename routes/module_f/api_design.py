# -*- coding: utf-8 -*-
"""[F-2] 수리계산 입력 라우트 — G 의 design/ 을 그대로 HTTP 로 연다.

엔진 함수는 전부 `cad_project_editor_g/services/cad_import/design/` 에서
import 한다 — 여기서는 판단하지 않는다(F 를 위해 G 를 몰래 바꾸는 것 금지).

    POST /api/module-f/design/build     최불리 → 제한 전개 → 5표 (메모리)
    GET  /api/module-f/design/preview   정규화+베이크 좌표 (저장 좌표 그대로)
    POST /api/module-f/design/emit      .sdf + .slf 저장 → 내려받기 링크

★미리보기 전용 좌표계는 없다. preview 와 emit 은 **같은 함수**
(`display_tables`)의 사본을 쓴다 — 화면과 파일이 달라지면 미리보기가 무의미하다
(G16 원칙 그대로).
"""
from __future__ import annotations

import os
from pathlib import Path

from flask import jsonify, request, send_file

from routes.module_f.common import (
    REMOTE_K_DEFAULT, _boot, _fail, _r1)
from routes.module_f.jobs import _job_running, _run_job, _sess

# 설정 7종(지시서 F-2) — body 로 받고 세션에 기억한다.
_DEFAULT_SETTINGS = {
    "k": REMOTE_K_DEFAULT,           # 기준개수
    "schedule": "KSD 3507",          # 배관 규격 기본값
    "iso": True,                     # 아이소매트릭 보기
    "iso_z_scale": 1.0,              # 고도 펼침 배율
    "canvas_units": 3000.0,          # 캔버스 크기
    "lift_ref": "valve",             # lift 영점 (valve | mid)
    "head_stub_pct": 2.5,            # 헤드 스텁 길이 (%)
}


def _settings(sess: dict, body: dict) -> dict:
    """설정 7종 — 준 것만 덮고 세션에 기억한다."""
    cur = dict(sess.get("design_settings") or _DEFAULT_SETTINGS)
    for key, cast in (("k", int), ("schedule", str), ("iso", bool),
                      ("iso_z_scale", float), ("canvas_units", float),
                      ("lift_ref", str), ("head_stub_pct", float)):
        if key in body and body[key] is not None:
            try:
                cur[key] = cast(body[key])
            except (TypeError, ValueError):
                return_fail = f"설정 값이 올바르지 않습니다: {key}={body[key]!r}"
                raise ValueError(return_fail)
    cur["k"] = max(1, min(int(cur["k"]), 200))
    sess["design_settings"] = cur
    return cur


def _dia_texts(sess: dict) -> list:
    """치수 텍스트 — handoff 캐시에서. 못 읽으면 빈 목록(관경은 별표1 폴백).

    G 데스크톱 4번째 창(`dialog_design_input._dia_texts`)과 같은 경로다 —
    여기가 다르면 같은 도면의 관경 근거가 웹과 데스크톱에서 갈라진다.

    ★실패를 조용히 넘기지 않는다. 여기가 빈 목록을 돌려주면 관경이 **전부**
      별표1 폴백이 되는데, 화면에는 «폴백 100%» 라는 결과만 남고 원인은
      어디에도 안 나온다. 실측으로 그렇게 한 번 당했다 — 원본 DXF 의 mtime 만
      바뀌어 handoff 캐시가 기각되면서 치수 텍스트 533개가 통째로 사라졌고,
      관경표는 아무 경고 없이 규약값으로만 채워졌다.
    """
    key = sess.get("key")
    try:
        import json
        from services.cad_import.design.bore import extract_dia_text_points
        from services.cad_import.pipeline import handoff, stage1 as s1
        spec = os.path.join(handoff.pick_out_dir(), f"{key}_찍은스펙.json")
        with open(spec, encoding="utf-8") as f:
            src = json.load(f).get("source_dxf")
        w = handoff.load_world(key, src, s1.World)
        if w is None:
            print(f"[설계] ★치수 텍스트 없음 — handoff 캐시를 쓸 수 없습니다"
                  f" (원본: {src}). 관경은 전부 별표1 로 정해집니다.")
            return []
        pts = extract_dia_text_points(w.texts)
        if not pts:
            print(f"[설계] ★도면 문자 {len(w.texts):,}개 중 치수로 읽힌 것이"
                  f" 0개입니다 — 관경은 전부 별표1 로 정해집니다.")
        else:
            print(f"[설계] 치수 텍스트 {len(pts):,}개"
                  f" (도면 문자 {len(w.texts):,}개 중)")
        return pts
    except Exception as exc:  # noqa: BLE001
        print(f"[설계] ★치수 텍스트를 읽지 못했습니다 — 관경은 별표1 로만: {exc}")
        return []


def _valve_label(tables):
    """알람밸브 노드의 표 라벨 — G 대화상자 `_valve_label` 과 같은 규칙."""
    for row in (getattr(tables, "equipment", None) or ()):
        if str(row.get("desc")) == "A/V":
            return row.get("in")
    return None


def _view_opts(cfg: dict) -> dict:
    """설정 7종 → display_tables/emit_design_sdf 인자. 두 곳이 같은 값을 쓴다."""
    return {
        "iso": bool(cfg["iso"]),
        "iso_z_scale": float(cfg["iso_z_scale"]),
        "canvas_units": float(cfg["canvas_units"]),
        "head_stub_ratio": float(cfg["head_stub_pct"]) / 100.0,
    }


def _summary(got: dict, tbl) -> dict:
    """build 요약 — G 대화상자 `_render` 가 보여주는 그 수치들."""
    meta = dict(tbl.meta)
    w = got.get("worst") or {}
    return {
        "k": len(w.get("heads") or []),
        "far_m": w.get("far_m"), "near_m": w.get("near_m"),
        "span_m": w.get("span_m"), "total_m": w.get("total_m"),
        "max_load": w.get("max_load"),
        "source": w.get("source_tag"),
        "counts": {"nodes": len(tbl.nodes), "pipes": len(tbl.pipes),
                   "nozzles": len(tbl.nozzles), "fittings": len(tbl.fittings),
                   "equipment": len(tbl.equipment)},
        "bore_src": {
            "text": meta.get("관경 근거 — 도면 텍스트"),
            "nfpc_min": meta.get("관경 근거 — 별표1 보강 (text<min)"),
            "nfpc_fallback": meta.get("관경 근거 — 별표1 폴백 (text 없음)"),
        },
        "fitting_unresolved": meta.get("부속 판정 불가"),
        "eq_len_unresolved": meta.get("등가길이 미해결"),
        "loops": meta.get("루프 잔여 배관(표 꼬리)"),
        "excluded_heads": got.get("excluded_heads", 0),
        "candidate_heads": got.get("candidate_heads", 0),
        "total_heads": got.get("total_heads", 0),
    }


def _classify_excluded(sess: dict, got: dict, board) -> dict:
    """[F-5] 빠진 헤드를 세 갈래로 가른다 — «제외 2,864» 를 숫자로 쪼갠다.

        찍히지 않음     A 후보인데 board 에 없는 것 (suggest 를 돌린 세션만)
        이음 끊김       board 물길은 닿는데 전개가 못 붙인 것 (B4 부착 실패)
        물길 미도달     board 물길 자체가 안 닿는 것 (이음 끊김의 상류)

    좌표를 함께 돌려준다 — 화면이 분류별로 켜고 끌 수 있어야 어디를 이어야
    하는지 보인다. 숫자만 주면 «크다» 만 알고 «어디» 를 모른다.
    """
    disks = getattr(board, "disks", []) or []
    total = len(disks)
    # board 물길 도달 — 급수원 성분에 붙은 헤드.
    try:
        wet_board = set((board.water_state() or {}).get("wet_heads") or [])
    except Exception as exc:  # noqa: BLE001
        print(f"[설계] 물길 분류 실패 — 미도달로 못 가른다: {exc}")
        wet_board = None
    # 전개가 붙일 수 있는 헤드 — 엔진의 공개 probe 를 그대로 쓴다.
    attach = None
    try:
        from services.cad_import.design.restrict import attachable_heads
        es = sess.get("edit")
        probe = attachable_heads(es.convert_payload())
        if probe.get("ok"):
            attach = set(probe.get("wet") or ())
    except Exception as exc:  # noqa: BLE001
        print(f"[설계] 부착 probe 실패 — 이음 끊김을 못 가른다: {exc}")

    def xy(i):
        d = disks[i]
        return [round(float(d[0]), 1), round(float(d[1]), 1)]

    out = {"total": total}
    if wet_board is not None:
        dry = [i for i in range(total) if i not in wet_board]
        out["dry"] = {"n": len(dry), "xy": [xy(i) for i in dry]}
        if attach is not None:
            unatt = [i for i in wet_board if i not in attach and i < total]
            out["unattached"] = {"n": len(unatt), "xy": [xy(i) for i in unatt]}
    # 찍히지 않음 — suggest 후보 중 어느 board 헤드와도 250mm 안에 없는 것.
    cands = sess.get("suggest")
    if cands:
        import math as _m
        centers = [(float(d[0]), float(d[1])) for d in disks]
        missing = []
        for c_ in cands:
            cx, cy = float(c_["x"]), float(c_["y"])
            if not any(_m.hypot(cx - px, cy - py) <= 250.0
                       for px, py in centers):
                missing.append([round(cx, 1), round(cy, 1)])
        out["unpicked"] = {"n": len(missing), "xy": missing}
    return out


def emit_design_files(sess: dict, UPLOAD_DIR, cfg: dict | None = None):
    """[F-4 공유] 세션의 확정된 표 → .sdf+.slf 한 쌍. (경로, 오류) 를 돌려준다.

    design/emit 라우트와 변환 단계의 «최불리 .sdf» 체크가 **같은 함수**를 탄다 —
    두 경로가 각자 emit 을 들고 있으면 설정 한쪽만 고쳐지는 날이 온다.
    """
    d = sess.get("design")
    if not d:
        return None, "먼저 수리계산 입력에서 표를 확정하세요."
    cfg = cfg or dict(sess.get("design_settings") or _DEFAULT_SETTINGS)

    from services.cad_import.design.emit import AssetMissing, emit_design_sdf
    key = sess.get("key") or "design"
    # ★파일 이름이 곧 SDF 안의 User-lib 참조다(G14 — 파일명만 담는다).
    #   G 데스크톱과 같은 이름이라야 산출이 같다. 세션끼리 안 섞이게 폴더를 가른다.
    out_dir = Path(UPLOAD_DIR) / "module_f" / f"{sess['id']}_design"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{key}_수리계산입력.sdf"
    tbl = d["tables"]
    try:
        out = emit_design_sdf(
            tbl, out_path,
            project_title=f"{key} 수리계산 입력",
            **_view_opts(cfg),
            iso_ref_label=(_valve_label(tbl)
                           if cfg["lift_ref"] == "valve" else None))
    except AssetMissing as exc:
        return None, str(exc)
    except Exception as exc:  # noqa: BLE001
        return None, f"{type(exc).__name__}: {exc}"
    sess["design_sdf_path"] = str(out)
    sess["design_slf_path"] = str(out.with_suffix(".slf"))
    return out, None


def register(app, *, UPLOAD_DIR):
    # ─────────────────────────────────────── 수리계산 입력 (설계)
    @app.post("/api/module-f/design/build")
    def module_f_design_build():
        """최불리 선정 → corridor 제한 전개 → 5표. 파일은 쓰지 않는다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        if not getattr(es.board, "sources", None):
            return _fail("급수 시작 위치를 먼저 찍어야 설계면적을 고를 수 있습니다.")
        undecided = sum(1 for kk in getattr(es.board, "disk_kinds", []) or []
                        if kk == "미지정")
        if undecided:
            return _fail(f"헤드 종류가 미지정인 것이 {undecided}개 있습니다. "
                         "손질에서 종류를 정한 뒤 다시 시도하세요.")
        if _job_running(sess):
            return _fail("이미 작업이 돌고 있습니다. 끝난 뒤에 다시 눌러 주세요.",
                         409)
        try:
            _boot()
            cfg = _settings(sess, body)
        except ValueError as exc:
            return _fail(str(exc))
        source = body.get("source")

        def job():
            from services.cad_import.design.restrict import select_and_expand
            from services.cad_import.design.tables import build_design_tables
            from services.cad_import.design.sdf_post import UnknownSchedule

            payload = es.convert_payload()
            srcs = payload.get("sources") or ()
            sel = source if source is not None else (
                srcs[0].get("tag") if len(srcs) > 1 and isinstance(srcs[0], dict)
                else None)
            got = select_and_expand(payload, es.board, k=cfg["k"],
                                    selected_source=sel)
            if not got.get("ok"):
                return {"ok": False, "error": got.get("error")}
            texts = _dia_texts(sess)
            try:
                tbl = build_design_tables(
                    got["kfp"], got["worst"], got["edge_ref"], texts,
                    board_pts=es.board.pts,
                    excluded_heads=got.get("excluded_heads", 0),
                    default_schedule=cfg["schedule"],
                    tree_loads=got.get("tree_loads"))
            except UnknownSchedule as exc:
                return {"ok": False, "error": str(exc)}
            # 표는 메모리에만 — emit 을 눌러야 파일이 생긴다.
            marks = _classify_excluded(sess, got, es.board)
            sess["design"] = {"got": got, "tables": tbl, "k": cfg["k"],
                              "schedule": cfg["schedule"], "marks": marks}
            s = _summary(got, tbl)
            s["excluded_detail"] = {
                k2: v2["n"] for k2, v2 in marks.items()
                if isinstance(v2, dict)}
            det = s["excluded_detail"]
            print("[설계] 제외 사유 — "
                  + " · ".join(f"{lab} {det[k2]:,}" for k2, lab in
                               (("dry", "물길 미도달"),
                                ("unattached", "이음 끊김"),
                                ("unpicked", "찍히지 않음")) if k2 in det))
            print(f"[설계] 표 확정 · 헤드 {s['k']} · 앵커 {s['far_m']} m · "
                  f"배관 {s['counts']['pipes']} · 제외 {s['excluded_heads']:,}")
            return {"ok": True, "summary": s}

        _run_job(sess, "수리계산 입력", job)
        return jsonify({"ok": True})

    @app.get("/api/module-f/design/preview")
    def module_f_design_preview():
        """emit 에 넘길 좌표 **그대로** — 표시 전용 좌표계를 따로 두지 않는다.

        보기 설정만 바뀌면 build 를 다시 돌지 않는다 — 캐시한 표에 표시 변환만
        다시 얹는다(G16 의 «최불리 재계산 없이 다시 그리기» 그대로).
        """
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        d = sess.get("design")
        if not d:
            job = sess.get("job") or {}
            if job.get("phase") == "수리계산 입력" and job.get("state") == "run":
                return _fail("아직 계산 중입니다.", 409)
            return _fail("먼저 design/build 로 표를 확정하세요.", 404)
        # 보기 설정 — 쿼리로 덮을 수 있고, 덮은 값은 기억된다.
        # (iso 는 문자열 "false" 를 bool() 로 캐스팅하면 True 가 되므로 따로 푼다.)
        try:
            cfg = _settings(sess, {k: request.args.get(k)
                                   for k in ("iso_z_scale", "canvas_units",
                                             "lift_ref", "head_stub_pct")
                                   if request.args.get(k) is not None})
        except ValueError as exc:
            return _fail(str(exc))
        if "iso" in request.args:
            cfg["iso"] = request.args.get("iso") in ("1", "true", "True", "on")
            sess["design_settings"] = cfg

        from services.cad_import.design.emit import display_tables
        tbl = d["tables"]
        view, stood = display_tables(
            tbl, **_view_opts(cfg),
            iso_ref_label=(_valve_label(tbl)
                           if cfg["lift_ref"] == "valve" else None))

        # 담당 헤드 수 — 간선 굵기의 근거. 도면 간선은 worst.loads, 세로
        # 구간(역참조 없음)은 tree_loads 가 안다.
        got = d["got"]
        loads = ((got.get("worst") or {}).get("loads")) or {}
        ref = got.get("edge_ref") or {}
        tree = got.get("tree_loads") or {}
        load_of = {}
        for pid, edge in ref.items():
            try:
                i, j = int(edge[0]), int(edge[1])
                load_of[str(pid)] = int(loads.get((min(i, j), max(i, j)), 0))
            except (TypeError, ValueError, IndexError):
                continue
        for pid, n in tree.items():
            load_of.setdefault(str(pid), int(n))

        at = {str(n.get("label")): n for n in view.nodes}
        elev = {lab: float(n.get("elevation", 0) or 0) for lab, n in at.items()}
        parent_of = {}
        for row in view.pipes:
            a, b = str(row.get("in")), str(row.get("out"))
            parent_of.setdefault(b, a)
            parent_of.setdefault(a, b)
        heads = {str(r.get("in")) for r in view.nozzles}
        av = _valve_label(tbl)

        # ── 최원 유하거리 «경로» — 급수원 → 앵커 (표 라벨 기준)
        #
        # 손질 단계에서는 board 절점으로 그렸지만 여기는 설계 표의 좌표계다.
        # 표에 이미 «앵커 노드» 가 meta 로 적혀 있으니 그것에서 급수원까지
        # 되짚는다. `parent_of` 는 양방향 첫이웃이라 트리가 아니다 — 여기서
        # 급수원 기점 BFS 로 제대로 된 부모를 만든다(안 그러면 되짚다 맴돈다).
        anchor_lab = dict(tbl.meta).get("앵커 노드")
        anchor_lab = str(anchor_lab) if anchor_lab not in (None, "?") else None
        root = next((lab for lab, n in at.items()
                     if str(n.get("io_node")) == "Input"), None)
        anchor_path: list[str] = []
        anchor_path_m = 0.0
        if anchor_lab and root:
            adj: dict[str, list[tuple[str, str]]] = {}
            plen: dict[str, float] = {}
            for row in view.pipes:
                a, b2 = str(row.get("in")), str(row.get("out"))
                pid = str(row.get("label"))
                try:
                    plen[pid] = float(row.get("length") or 0.0)
                except (TypeError, ValueError):
                    plen[pid] = 0.0
                adj.setdefault(a, []).append((b2, pid))
                adj.setdefault(b2, []).append((a, pid))
            par: dict[str, tuple[str, str]] = {}
            seen = {root}
            queue = [root]
            while queue:
                cur = queue.pop(0)
                for nxt, pid in adj.get(cur, ()):
                    if nxt in seen:
                        continue
                    seen.add(nxt)
                    par[nxt] = (cur, pid)
                    queue.append(nxt)
            if anchor_lab in seen:
                cur = anchor_lab
                while cur != root:
                    up, pid = par[cur]
                    anchor_path.append(cur)
                    anchor_path_m += plen.get(pid, 0.0)
                    cur = up
                anchor_path.append(root)
                anchor_path.reverse()

        nodes = []
        for lab, n in at.items():
            # ★반올림하지 않는다 — writer 는 좌표를 `.6g`(유효 6자리)로 찍는데
            #   소수 3자리 반올림은 그보다 거칠어(실측: -75.9731 → -75.973)
            #   「preview == 저장 Position」 이 깨진다. 같은 double 을 그대로
            #   보내면 양쪽을 `.6g` 로 찍었을 때 정확히 같은 문자열이 된다.
            rec = {"label": lab, "x": float(n.get("x", 0)),
                   "y": float(n.get("y", 0))}
            if lab in heads:
                rec["head"] = True
                rec["up"] = (elev.get(lab, 0.0)
                             - elev.get(parent_of.get(lab, ""), 0.0)) >= 0
            if str(n.get("io_node")) == "Input":
                rec["input"] = True
            if av is not None and lab == str(av):
                rec["valve"] = True
            if anchor_lab and lab == anchor_lab:
                rec["anchor"] = True      # 기준압을 잡는 지점
            nodes.append(rec)
        pipes = [{"label": str(r.get("label")),
                  "a": str(r.get("in")), "b": str(r.get("out")),
                  "dia": r.get("dia"), "len_m": r.get("length"),
                  "src": r.get("dia_src"),
                  "load": load_of.get(str(r.get("label")), 0)}
                 for r in view.pipes]
        return jsonify({
            "ok": True, "settings": cfg,
            "stood": stood,
            "view": {"nodes": nodes, "pipes": pipes,
                     # 최원 유하거리 경로 — far_m 이 «어느 줄» 인지.
                     "anchor": anchor_lab,
                     "anchor_path": anchor_path,
                     "anchor_path_m": round(anchor_path_m, 2)},
            "tables": tbl.as_dict(),        # 저장될 값 그대로 (F-3 표 4종)
            # [F-5] 제외 사유 분류 — mm 세계좌표. 설계 캔버스(정규화 좌표)가
            # 아니라 손질 망 위에 그려야 «어디» 인지 보인다.
            "marks": d.get("marks") or {},
        })

    @app.post("/api/module-f/design/emit")
    def module_f_design_emit():
        """.sdf + .slf 한 쌍을 쓴다. 자산이 없으면 실패(G 정책 그대로)."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        d = sess.get("design")
        if not d:
            return _fail("먼저 design/build 로 표를 확정하세요.", 404)
        try:
            cfg = _settings(sess, body)
        except ValueError as exc:
            return _fail(str(exc))

        out, err = emit_design_files(sess, UPLOAD_DIR, cfg)
        if err:
            return _fail(err, 500)
        slf = out.with_suffix(".slf")
        return jsonify({
            "ok": True,
            "sdf": {"name": out.name, "bytes": out.stat().st_size},
            "slf": {"name": slf.name, "bytes": slf.stat().st_size},
            "download": f"/api/module-f/download?sid={sess['id']}&what=design",
        })
