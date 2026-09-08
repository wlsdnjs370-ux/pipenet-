# -*- coding: utf-8 -*-
"""[H-4 · H-5 · H-6] 모듈 F 라우트 — 제5국면 S700.

    /merge/mode     S710  급수방식 선택 (사람이 고른다)
    /merge/build    S720~S740  입상관 · 기계실 전단 접속 · 결합
    /merge/state    지금까지의 결합 상태
    /merge/emit     S750 · S760 · S770  입력파일 · 형식변환 · 압축

★결합은 **세션 안에서만** 일어난다. 슬롯마다 뽑아 둔 것을 모아 한 망으로
  만들 뿐, 어느 슬롯의 저장본도 건드리지 않는다.

★평면도 단독도 정상 경로다. 계통도가 없으면 결합할 입상관이 없으니 평면도의
  설계 표가 그대로 산출이 된다(지시서 H-5 — 그 경우 산출이 H-4 이전과 바이트
  동일해야 한다).
"""
from __future__ import annotations

import os

from flask import jsonify, request, send_file

from routes.module_f.common import _fail
from routes.module_f.jobs import _job_running, _run_job, route_session
from routes.module_f.merge import (
    ANCHOR_LABEL, SUPPLY_MODES, MergeError, check_supply_mode,
    combined_summary, merge_network)
from routes.module_f.slots import SLOT_KINDS, _slot_active, _slot_capture

# 결합에 쓸 재료가 어느 슬롯에 있는가 — 활성 슬롯이 아니어도 꺼내 온다.
_SLOT_PICK = {
    "plan": ("design", "설계 표"),
    "system": ("riser", "계통도 입상관"),
    "machineroom": ("machineroom", "기계실 경로"),
}


def _slot_value(sess: dict, kind: str, key: str):
    """슬롯 하나에서 값을 꺼낸다 — 활성이면 평면 dict, 아니면 저장소에서."""
    if _slot_active(sess) == kind:
        return _slot_capture(sess).get(key)
    return ((sess.get("slots") or {}).get(kind) or {}).get(key)


def _materials(sess: dict) -> dict:
    """세 슬롯이 지금 내놓을 수 있는 재료.

    평면도만 한 겹 더 들어간다 — `sess["design"]` 은 `{got, tables, k, …}` 묶음
    이고 결합이 쓰는 것은 그중 `tables` 다.
    """
    out = {}
    for kind in SLOT_KINDS:
        key, _label = _SLOT_PICK[kind]
        val = _slot_value(sess, kind, key)
        if kind == "plan" and isinstance(val, dict):
            # 어느 길로 온 표인지 함께 들고 간다 — 라벨 오프셋이 갈린다.
            out["plan_method"] = val.get("method") or "manual"
            val = val.get("tables")
        out[kind] = val
    out.setdefault("plan_method", "manual")
    return out


def register(app, *, UPLOAD_DIR):
    # ─────────────────────────────────── S710
    @app.get("/api/module-f/merge/modes")
    def module_f_merge_modes():
        """고를 수 있는 급수방식 — 화면이 이 목록으로 라디오를 그린다."""
        return jsonify({"ok": True,
                        "modes": [{"key": k, "label": v}
                                  for k, v in SUPPLY_MODES.items()]})

    @app.post("/api/module-f/merge/mode")
    @route_session(post=True)
    def module_f_merge_mode(sess, body):
        """급수방식을 고른다. 자동 추정하지 않는다 — 도면에 없는 값이다."""
        # ★결합 잡은 급수방식·낙차·펌프 제원을 «돌면서» 읽는다(merge_network 호출
        #   시점). 도는 중에 바꾸면 로그에 찍힌 방식과 실제 쓰인 값이 갈린다.
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 바꿀 수 있습니다.", 409)
        try:
            mode = check_supply_mode(body.get("mode"))
        except MergeError as exc:
            return _fail(str(exc))
        sess["supply_mode"] = mode
        # 펌프 가압에서만 뜻이 있는 값들 — 없으면 0/미지정으로 둔다.
        for key, cast in (("source_drop_m", float),):
            if body.get(key) is not None:
                try:
                    sess[key] = cast(body[key])
                except (TypeError, ValueError):
                    return _fail(f"{key} 값이 올바르지 않습니다: {body[key]!r}")
        pump = body.get("pump")
        if isinstance(pump, dict):
            sess["pump_spec"] = pump
        return jsonify({"ok": True, "mode": mode,
                        "label": SUPPLY_MODES[mode],
                        "source_drop_m": sess.get("source_drop_m", 0.0)})

    # ─────────────────────────────────── 상태
    @app.get("/api/module-f/merge/state")
    @route_session()
    def module_f_merge_state(sess, body):
        """재료가 갖춰졌나 · 무엇이 비었나 — S650 이 «남은 도면» 을 묻는 자리."""
        mats = _materials(sess)
        return jsonify({
            "ok": True,
            "mode": sess.get("supply_mode"),
            "mode_label": SUPPLY_MODES.get(sess.get("supply_mode") or ""),
            "source_drop_m": sess.get("source_drop_m", 0.0),
            "ready": {kind: bool(v) for kind, v in mats.items()},
            "labels": {kind: _SLOT_PICK[kind][1] for kind in SLOT_KINDS},
            # 평면도만 있으면 결합 없이 지나간다 — 그것도 정상이다.
            "can_build": bool(mats["plan"]) and bool(sess.get("supply_mode")),
            "merged": bool(sess.get("merged")),
            "summary": sess.get("merge_summary"),
        })

    # ─────────────────────────────────── S720~S740
    @app.post("/api/module-f/merge/build")
    @route_session(post=True)
    def module_f_merge_build(sess, body):
        """세 도면을 한 배관망으로. 무거우므로 잡으로 돌린다."""
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 결합할 수 있습니다.", 409)

        mode = sess.get("supply_mode")
        if not mode:
            return _fail("급수방식을 먼저 고르세요 (S710).", 400)

        mats = _materials(sess)
        if not mats["plan"]:
            return _fail("평면도의 설계 표를 먼저 확정하세요 "
                         "(수리계산 단계의 «표 확정»).", 400)

        def job():
            print(f"[결합] S700 시작 — 급수방식 {SUPPLY_MODES[mode]}")
            print("[결합]   평면도 경로: "
                  + ("자동(A 위상 검출)" if mats["plan_method"] == "auto"
                     else "수동(E 색 찍기)"))
            for kind in SLOT_KINDS:
                print(f"[결합]   {_SLOT_PICK[kind][1]}: "
                      + ("있음" if mats[kind] else "없음"))
            got = merge_network(
                mats["plan"], riser=mats["system"],
                machineroom=mats["machineroom"], mode=mode,
                source_drop_m=sess.get("source_drop_m", 0.0),
                pump=sess.get("pump_spec"),
                method=mats["plan_method"])
            sess["merged"] = got
            summary = combined_summary(got)
            sess["merge_summary"] = summary
            for line in summary.get("steps") or ():
                print(f"[결합]   · {line}")
            print(f"[결합] 완료 — 절점 {summary['nodes']} · 배관 {summary['pipes']}"
                  f" · 노즐 {summary['nozzles']}")
            return summary

        _run_job(sess, "배관망 결합", job)
        return jsonify({"ok": True, "sid": sess["id"]})

    # ─────────────────────────────────── 결합망 미리보기
    @app.get("/api/module-f/merge/preview")
    @route_session()
    def module_f_merge_preview(sess, body):
        """결합된 배관망을 **화면에 그릴 모양**으로.

        ★결합해 놓고 보여 주지 않으면 사람은 무엇이 합쳐졌는지 알 수 없다 —
          숫자(절점 308 · 배관 307)만으로는 세 도면이 제대로 이어졌는지 판단할
          길이 없다. 여기서 세 도면을 **색으로 갈라** 한 그림으로 준다.

        좌표는 결합망 그대로(평면)다 — `emit_merged` 가 내는 파일도 이 좌표라
        「보이는 것 = 저장되는 것」이 성립한다. `iso=1` 이면 30° 등각으로
        굽되, 그것은 **보기 전용**이라는 것을 화면이 말한다.

        Query: sid · [iso=0|1] · [iso_z_scale]
        """
        got = sess.get("merged")
        if not got:
            return jsonify({"ok": True, "view": None,
                            "message": "먼저 결합하세요 (S740)."})
        c = got.get("combined")
        if c is None:
            return jsonify({"ok": True, "view": None,
                            "message": "계통도가 없어 결합망이 없습니다 — "
                                       "평면도 단독 산출입니다."})
        parts = got.get("parts") or {}
        of = {}
        for kind in ("system", "machineroom", "plan"):
            for lab in (parts.get(kind) or ()):
                of[str(lab)] = kind
        # ★기준점(라벨 10)은 **두 망 모두에** 있다 — 특허 S740 이 평면도 라벨을
        #   +9 해서 그 한 점에서 만나게 하기 때문이다. 색으로는 한쪽에 넣되,
        #   «여기가 이음매» 라는 것을 따로 표시한다. 결합이 제대로 됐는지는
        #   결국 그 한 점을 보고 판단한다.
        shared = ({str(x) for x in (parts.get("plan") or ())}
                  & ({str(x) for x in (parts.get("system") or ())}
                     | {str(x) for x in (parts.get("machineroom") or ())}))

        nodes = [dict(n) for n in (c.nodes or ())]
        mr_edges = [list(map(float, e)) for e in
                    (getattr(c, "machine_room_plan_edges", None) or ())]
        iso = (request.args.get("iso") or "0") in ("1", "true", "True", "on")
        if iso:
            try:
                zs = float(request.args.get("iso_z_scale") or 1.0)
            except (TypeError, ValueError):
                zs = 1.0
            # ★부위마다 «맞는» 투영이 다르다 — 한 식으로 다 굽으면 깨진다.
            #
            #   · 평면도: 평면이니 30° 회전 + 표고 lift. lift 는 평면과 같은 자
            #     (1 m = 1000, §T3 좌표가 mm) — 설계 화면과 같은 규칙이다.
            #   · 계통도(라이저): schematic y 가 이미 **수직**이다. 회전을
            #     먹이면 수직 막대가 사선이 된다(실측: x 퍼짐 1,732 — 사용자
            #     지적 「계통도가 기울어져 있다」). 기준점의 아이소 위치에
            #     평면 오프셋을 그대로 얹어 수직으로 세운다.
            #   · 기계실: 평면 군집이니 회전하되, 접속점(펌프 junction)이
            #     라이저의 «새» 자리에 그대로 붙도록 평행이동한다 — 안 하면
            #     이음매가 찢어진다.
            COS30, SIN30 = 0.8660254037844387, 0.5

            def _rot(x, y):
                return ((x - y) * COS30, (x + y) * SIN30)

            at0 = {str(n.get("label")): (float(n.get("x", 0) or 0),
                                         float(n.get("y", 0) or 0))
                   for n in nodes}
            ax, ay = at0.get(ANCHOR_LABEL, (0.0, 0.0))
            a_iso = _rot(ax, ay)
            e_ref = next((float(n.get("elevation", 0) or 0) for n in nodes
                          if str(n.get("label")) == ANCHOR_LABEL), 0.0)
            lift = 1000.0 * zs

            pj = got.get("pump_junction")
            pj_xy = at0.get(str(pj)) if pj else None
            shift = (0.0, 0.0)
            if pj_xy is not None:
                # 펌프 junction 은 라이저 규칙으로 옮겨진다 — 그 새 자리와
                # 평면 회전 자리의 차가 기계실 군집의 평행이동이다.
                new_pj = (a_iso[0] + (pj_xy[0] - ax),
                          a_iso[1] + (pj_xy[1] - ay))
                rot_pj = _rot(*pj_xy)
                shift = (new_pj[0] - rot_pj[0], new_pj[1] - rot_pj[1])

            for n in nodes:
                lab = str(n.get("label"))
                x = float(n.get("x", 0) or 0)
                y = float(n.get("y", 0) or 0)
                kind = of.get(lab, "plan")
                if kind == "system":
                    n["x"] = a_iso[0] + (x - ax)
                    n["y"] = a_iso[1] + (y - ay)
                elif kind == "machineroom":
                    rx, ry = _rot(x, y)
                    n["x"] = rx + shift[0]
                    n["y"] = ry + shift[1]
                else:
                    rx, ry = _rot(x, y)
                    e = float(n.get("elevation", 0) or 0)
                    n["x"] = rx
                    n["y"] = ry + (e - e_ref) * lift
            mr_edges = []
            for e in (getattr(c, "machine_room_plan_edges", None) or ()):
                r1 = _rot(float(e[0]), float(e[1]))
                r2 = _rot(float(e[2]), float(e[3]))
                mr_edges.append([r1[0] + shift[0], r1[1] + shift[1],
                                 r2[0] + shift[0], r2[1] + shift[1]])

        heads = {str(r.get("in")) for r in (c.nozzles or ())}
        pumps = {str(r.get("in")) for r in (c.pumps or ())}
        pumps |= {str(r.get("out")) for r in (c.pumps or ())}
        valves = {str(r.get("in")) for r in (c.valves or ())}

        out_nodes = []
        for n in nodes:
            lab = str(n.get("label"))
            rec = {"label": lab, "x": float(n.get("x", 0) or 0),
                   "y": float(n.get("y", 0) or 0),
                   "e": round(float(n.get("elevation", 0) or 0), 3),
                   "part": of.get(lab, "plan")}
            if lab in heads:
                rec["head"] = True
            if lab in pumps:
                rec["pump"] = True
            if lab in valves:
                rec["valve"] = True
            if str(n.get("io_node")) == "Input":
                rec["input"] = True
            if lab in shared:
                rec["anchor"] = True
            out_nodes.append(rec)

        out_pipes = []
        for r in (c.pipes or ()):
            a, b = str(r.get("in")), str(r.get("out"))
            # 두 도면을 잇는 배관은 «이음매» 다 — 그 자리가 결합의 핵심이라
            # 화면이 따로 그릴 수 있게 표시해 둔다.
            ka, kb = of.get(a, "plan"), of.get(b, "plan")
            out_pipes.append({"label": str(r.get("label")), "a": a, "b": b,
                              "dia": r.get("dia"), "len_m": r.get("length"),
                              "part": ka if ka == kb else "seam"})

        return jsonify({
            "ok": True, "iso": iso,
            "view": {"nodes": out_nodes, "pipes": out_pipes,
                     # 기계실 평면 배관망 — SDF 에는 없고 «보기» 로만 쓴다.
                     "mr_plan_edges": mr_edges},
            "counts": {"plan": sum(1 for n in out_nodes
                                   if n["part"] == "plan"),
                       "system": sum(1 for n in out_nodes
                                     if n["part"] == "system"),
                       "machineroom": sum(1 for n in out_nodes
                                          if n["part"] == "machineroom"),
                       "seam": sum(1 for p in out_pipes
                                   if p["part"] == "seam"),
                       "anchor": sorted(shared)},
        })

    # ─────────────────────────────────── S750 · S760 · S770
    @app.post("/api/module-f/merge/emit")
    @route_session(post=True)
    def module_f_merge_emit(sess, body):
        """결합망 → 입력파일 3종 + 압축.

        S760 은 «별도 산출이 아니라 S750 의 결과 파일 자체를 원본으로» 삼는다
        (특허 도 9 주석). 그래서 SDF 를 먼저 쓰고 그 파일에서 나머지를 만든다 —
        형식마다 따로 뽑으면 같은 배관망을 가리킨다는 보장이 사라진다.
        """
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 저장할 수 있습니다.", 409)
        got = sess.get("merged")
        if not got:
            return _fail("먼저 결합하세요 (S740).", 400)
        if got.get("combined") is None:
            return _fail("계통도가 없어 결합망이 없습니다 — 평면도 산출은 "
                         "수리계산 단계의 «.sdf + .slf 저장» 을 쓰세요.", 400)

        from pathlib import Path

        out_dir = Path(UPLOAD_DIR).parent / "module_f_merged" / sess["id"]

        def job():
            from routes.module_f.emit import emit_merged
            print("[결합] S750 입력파일 생성")
            files = emit_merged(
                got["combined"], out_dir,
                title=f"모듈 F 통합 — {sess.get('key') or ''}")
            sess["merge_files"] = files
            for k, v in files.items():
                if v:
                    print(f"[결합]   {k}: {os.path.basename(v)}")
            return {k: os.path.basename(v) for k, v in files.items() if v}

        _run_job(sess, "산출물 생성", job)
        return jsonify({"ok": True, "sid": sess["id"]})

    @app.get("/api/module-f/merge/download")
    @route_session()
    def module_f_merge_download(sess, body):
        """산출물 내려받기 — 세션이 만든 것만."""
        what = str(request.args.get("what") or "zip")
        files = sess.get("merge_files") or {}
        path = files.get(what)
        if not path or not os.path.isfile(path):
            return _fail(f"그런 산출물이 없습니다: {what}", 404)
        return send_file(path, as_attachment=True,
                         download_name=os.path.basename(path))
