# -*- coding: utf-8 -*-
"""[H-2 · H-3] 모듈 F 라우트 — 계통도(S720) · 기계실(S730) 추출.

두 라우트가 하는 일은 같다: 사람이 찍은 **두 점**을 받아 그 사이의 배관 경로를
뽑는다. 계통도는 펌프↔알람밸브, 기계실은 수원↔입상관 연결점이다.

엔진은 모듈 A 것을 그대로 쓴다(`subdrawing.py`). 여기는 세션·입력 검사·실패
보고만 맡는다.

★실패를 «성공한 빈 결과» 로 바꾸지 않는다. 클릭이 배관에서 너무 멀거나 두 점이
  안 이어지면 400 으로 그 사실을 말한다 — 특허 S340 의 «임의로 메우지 아니하고
  미도달로 보고한다» 가 이 자리의 규범이다.
"""
from __future__ import annotations

from flask import jsonify, request

from routes.module_f.common import _fail
from routes.module_f.jobs import _job_running, _sess, route_session
from routes.module_f.slots import _slot_active
from routes.module_f.subdrawing import (
    extract_machineroom, extract_system, extract_system_clean, riser_summary)

# 클릭 ↔ 그래프 절점 허용 거리. A 의 기본값과 같다.
SNAP_DEFAULT_MM = 2500.0
SNAP_MAX_MM = 50_000.0


def _xy(body, kx, ky, what):
    x, y = body.get(kx), body.get(ky)
    if x is None or y is None:
        raise ValueError(f"{what} 위치를 도면에서 찍으세요.")
    try:
        return float(x), float(y)
    except (TypeError, ValueError):
        raise ValueError(f"{what} 좌표가 숫자가 아닙니다: {x!r}, {y!r}") from None


def _snap(body) -> float:
    try:
        v = float(body.get("snap_tolerance_mm") or SNAP_DEFAULT_MM)
    except (TypeError, ValueError):
        return SNAP_DEFAULT_MM
    return min(max(v, 1.0), SNAP_MAX_MM)


def _need_slot(body, kind: str):
    """그 슬롯이 활성이고 도면이 읽혀 있는가."""
    sess = _sess(body.get("sid"))
    if _job_running(sess):
        return sess, "작업이 끝난 뒤에 추출할 수 있습니다."
    if _slot_active(sess) != kind:
        return sess, f"«{kind}» 슬롯으로 먼저 바꾸세요."
    if not sess.get("entities"):
        return sess, "도면이 아직 준비되지 않았습니다."
    return sess, None


def register(app):
    # ─────────────────────────────────── 계통도 (S720)
    @app.post("/api/module-f/system/extract")
    @route_session(lambda b: _need_slot(b, "system"), post=True)
    def module_f_system_extract(sess, body):
        """펌프 → 알람밸브 경로 = 입상관.

        Body: sid · pump_x · pump_y · av_x · av_y · [snap_tolerance_mm]
              [waypoints:[[x,y],…]] · [clean:true]
        """
        # 조각난 풀 계통도용 폴백 — 두 점 없이 파일의 단일망을 통째로 읽는다.
        if bool(body.get("clean")):
            try:
                riser = extract_system_clean(sess.get("dxf"))
            except Exception as exc:  # noqa: BLE001
                return _fail(f"깨끗한 배관망으로도 읽지 못했습니다: {exc}", 400)
            sess["riser"] = riser
            sess["riser_mode"] = "clean_network"
            return jsonify({"ok": True, "mode": "clean_network",
                            "summary": riser_summary(riser)})

        try:
            pump = _xy(body, "pump_x", "pump_y", "펌프")
            av = _xy(body, "av_x", "av_y", "알람밸브")
        except ValueError as exc:
            return _fail(str(exc))
        wps = []
        for p in (body.get("waypoints") or []):
            try:
                wps.append((float(p[0]), float(p[1])))
            except (TypeError, ValueError, IndexError):
                return _fail(f"경유점 좌표가 잘못되었습니다: {p!r}")

        try:
            riser = extract_system(sess["entities"], pump, av,
                                   snap_tolerance_mm=_snap(body),
                                   waypoints=wps or None)
        except ValueError as exc:
            # 사용자 입력 문제 — 미도달을 그대로 말한다(S340).
            return jsonify({"ok": False, "message": str(exc),
                            "suggest_clean": True}), 400
        except Exception as exc:  # noqa: BLE001
            return _fail(f"계통도 추출에 실패했습니다: {exc}", 500)

        sess["riser"] = riser
        sess["riser_mode"] = "dxf_path_v1"
        return jsonify({"ok": True, "mode": "dxf_path_v1",
                        "summary": riser_summary(riser)})

    # ─────────────────────────────────── 기계실 (S730)
    @app.post("/api/module-f/machineroom/extract")
    @route_session(lambda b: _need_slot(b, "machineroom"), post=True)
    def module_f_machineroom_extract(sess, body):
        """수원(탱크) → 입상관 연결점 경로. 좌표는 평면 그대로 보존한다.

        Body: sid · source_x · source_y · conn_x · conn_y
              [snap_tolerance_mm] · [ceiling_m]
        """
        try:
            src = _xy(body, "source_x", "source_y", "수원(탱크 토출구)")
            conn = _xy(body, "conn_x", "conn_y", "입상관 연결점")
        except ValueError as exc:
            return _fail(str(exc))
        ceiling = body.get("ceiling_m")
        if ceiling is not None:
            try:
                ceiling = float(ceiling)
            except (TypeError, ValueError):
                return _fail(f"기계실 천장고가 숫자가 아닙니다: {ceiling!r}")

        try:
            mr = extract_machineroom(sess["entities"], src, conn,
                                     snap_tolerance_mm=_snap(body),
                                     ceiling_m=ceiling)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            return _fail(f"기계실 추출에 실패했습니다: {exc}", 500)

        # 사람이 찍은 연결점을 함께 남긴다 — 결합(S740) 때 기계실 평면을 어디에
        # 붙일지의 기준이다. 추출 결과 dict 에는 라벨만 있고 좌표는 없다.
        mr["conn_xy"] = [conn[0], conn[1]]
        sess["machineroom"] = mr
        summary = riser_summary(mr)
        # 실측 edge 와 추정 edge 를 갈라 보고한다 — 통합해 그리면 안 된다.
        summary["plan_edges"] = len(mr.get("plan_edges") or ())
        summary["plan_edges_estimated"] = len(mr.get("plan_edges_estimated") or ())
        # 천장고가 없으면 첫 구간 표고가 미확정으로 남는다 — 숨기지 않는다.
        summary["ceiling_m"] = ceiling
        summary["elevation_unresolved"] = ceiling is None
        return jsonify({"ok": True, "summary": summary})

    # ─────────────────────────────────── 추출 결과 되읽기
    @app.get("/api/module-f/sub/state")
    @route_session()
    def module_f_sub_state(sess, body):
        """지금 슬롯의 추출 결과 요약 — 없으면 빈 것으로 답한다."""
        kind = _slot_active(sess)
        got = sess.get("riser") if kind == "system" else sess.get("machineroom")
        return jsonify({
            "ok": True, "kind": kind,
            "opened": bool(sess.get("entities")),
            "extracted": bool(got),
            "mode": sess.get("riser_mode"),
            "summary": riser_summary(got) if got else None,
        })
