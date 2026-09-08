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

from flask import jsonify

from routes.module_f.common import _check_xy, _fail
from routes.module_f.jobs import _job_running, _sess, route_session
from routes.module_f.slots import _slot_active
from routes.module_f.sub_fix import (
    SLOT_KEY, apply_overrides, parse_rows, rows_for_view)
from routes.module_f.subdrawing import (
    extract_machineroom, extract_system, extract_system_clean, graph_payload,
    layer_options, pick_system_layer, riser_summary)

# 클릭 ↔ 그래프 절점 허용 거리. A 의 기본값과 같다.
SNAP_DEFAULT_MM = 2500.0
SNAP_MAX_MM = 50_000.0


def _xy(body, kx, ky, what):
    x, y = body.get(kx), body.get(ky)
    if x is None or y is None:
        raise ValueError(f"{what} 위치를 도면에서 찍으세요.")
    # 셀 수 있는 수인가까지 본다 — 제곱에서 터지는 자리가 엔진 안에 있다
    # (common._check_xy 주석 참조). 네 입구가 같은 자를 쓴다.
    return _check_xy(body, kx=kx, ky=ky, what=what)


def _snap(body) -> float:
    try:
        v = float(body.get("snap_tolerance_mm") or SNAP_DEFAULT_MM)
    except (TypeError, ValueError):
        return SNAP_DEFAULT_MM
    return min(max(v, 1.0), SNAP_MAX_MM)


def _layers(sess, body):
    """사람이 고른 «배관으로 볼 레이어». 안 고르면 None = 도면 전체.

    ★고른 것은 세션에 남는다. 미리보기(그래프)와 추출이 **같은 레이어**를
      써야 화면에 그려진 경로와 뽑히는 경로가 같다 — 다르면 미리보기가
      거짓말을 한다.
    """
    if "layers" in body:
        raw = body.get("layers")
        if raw in (None, "", []):
            sess["sub_layers"] = None
        elif isinstance(raw, list):
            sess["sub_layers"] = sorted({str(v) for v in raw})
        # ★사람이 직접 고른 순간부터 «자동 좁히기» 는 손을 뗀다. 사람 결정을
        #   기계가 다음 클릭에 덮으면 고를 수 있다는 말이 거짓이 된다.
        sess["sub_layers_auto"] = False
        # 목록이 아니면 조용히 무시하지 않고 그대로 둔다(옛 값 유지).
    got = sess.get("sub_layers")
    return set(got) if got else None


def _fixes(sess, kind: str) -> list:
    return list((sess.get("sub_fixes") or {}).get(kind) or ())


def _reapply(sess, kind: str) -> dict:
    """다시 뽑은 결과에 사람이 고친 값을 **되붙인다**.

    ★안 되붙이면 추출을 한 번 더 누른 순간 손질이 통째로 사라진다. 자리는
      좌표로 가리키므로 같은 구간이면 다시 붙고, 사라진 구간은 «못 붙였다» 로
      세어 화면에 말한다(조용히 버리지 않는다).
    """
    got = sess.get(SLOT_KEY[kind])
    rows = _fixes(sess, kind)
    if not got or not rows:
        return {"applied": 0, "given": len(rows), "unmatched": 0}
    return apply_overrides(got, rows)


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
                                   waypoints=wps or None,
                                   layer_filter=_layers(sess, body))
        except ValueError as exc:
            # 사용자 입력 문제 — 미도달을 그대로 말한다(S340).
            return jsonify({"ok": False, "message": str(exc),
                            "suggest_clean": True}), 400
        except Exception as exc:  # noqa: BLE001
            return _fail(f"계통도 추출에 실패했습니다: {exc}", 500)

        sess["riser"] = riser
        sess["riser_mode"] = "dxf_path_v1"
        fixed = _reapply(sess, "system")
        return jsonify({"ok": True, "mode": "dxf_path_v1",
                        "summary": riser_summary(riser), "fixed": fixed})

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
                                     ceiling_m=ceiling,
                                     layer_filter=_layers(sess, body))
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            return _fail(f"기계실 추출에 실패했습니다: {exc}", 500)

        # 사람이 찍은 연결점을 함께 남긴다 — 결합(S740) 때 기계실 평면을 어디에
        # 붙일지의 기준이다. 추출 결과 dict 에는 라벨만 있고 좌표는 없다.
        mr["conn_xy"] = [conn[0], conn[1]]
        sess["machineroom"] = mr
        # ★요약을 만들기 **전에** 되붙인다 — 뒤에 하면 화면에 뜨는 연장이
        #   손질 전 값이 되어, 표와 요약이 서로 다른 말을 한다.
        fixed = _reapply(sess, "machineroom")
        summary = riser_summary(mr)
        # 실측 edge 와 추정 edge 를 갈라 보고한다 — 통합해 그리면 안 된다.
        summary["plan_edges"] = len(mr.get("plan_edges") or ())
        summary["plan_edges_estimated"] = len(mr.get("plan_edges_estimated") or ())
        # 천장고가 없으면 첫 구간 표고가 미확정으로 남는다 — 숨기지 않는다.
        summary["ceiling_m"] = ceiling
        summary["elevation_unresolved"] = ceiling is None
        return jsonify({"ok": True, "summary": summary, "fixed": fixed})

    # ─────────────────────────────────── 경로 그래프 (실시간 미리보기용)
    @app.post("/api/module-f/sub/graph")
    @route_session(post=True)
    def module_f_sub_graph(sess, body):
        """두 점 사이의 선이 «따라오게» 하려면 화면이 그래프를 들어야 한다.

        마우스가 움직일 때마다 서버를 왕복하면 LAN·터널에서 눈에 띄게 밀린다.
        실측으로 이 그래프는 작다(노드 132~382 · 간선 131~411 · 3~8KB) —
        통째로 내려보내고 브라우저가 직접 최단경로를 푼다.

        ★추출이 쓰는 **바로 그 그래프** 다(`subdrawing.path_graph`). 미리보기와
          결과가 다른 그래프를 쓰면 화면이 거짓말을 한다.

        body: {sid, [layers: ["레이어명", …]]}  · layers 를 주면 그 값이 세션에
              남아 추출까지 같이 쓴다. 빈 배열/None 이면 «도면 전체».
        """
        if not sess.get("entities"):
            return _fail("도면이 아직 준비되지 않았습니다.", 409)
        lf = _layers(sess, body)
        # ★두 점을 다 찍었으면 «그 두 점이 있는 계통» 하나로 좁힌다.
        #   섞은 채로 두면 최단경로가 고층↔저층을 넘나든다(실측 4회). 좁히기는
        #   미리보기 그래프 자체를 바꾸므로 추출과 어긋나지 않는다 — 세션에
        #   남겨 두 곳이 같은 레이어를 쓴다.
        narrowed = None
        pa, pb = body.get("a"), body.get("b")
        # ★«아직 안 골랐을 때만» 으로 걸면 한 번 좁힌 뒤로는 다시 못 좁힌다 —
        #   점을 다른 계통에 다시 찍어도 첫 계통에 갇힌다. 기준은 «사람이
        #   골랐는가» 하나뿐이다(`sub_layers_auto`).
        auto_ok = sess.get("sub_layers_auto", True)
        if auto_ok and pa and pb:
            try:
                nm, diag = pick_system_layer(
                    sess["entities"], (float(pa[0]), float(pa[1])),
                    (float(pb[0]), float(pb[1])),
                    snap_tolerance_mm=_snap(body))
            except Exception as exc:  # noqa: BLE001 — 좁히기 실패는 치명이 아니다
                nm, diag = None, {"reason": f"계통을 못 골랐습니다: {exc}"}
            if nm:
                sess["sub_layers"] = [nm]
                lf = {nm}
            else:
                # 못 좁혔으면 **지난번 자동 선택도 푼다** — 점을 옮겼는데 옛
                # 계통에 갇힌 채로 뽑으면 그게 더 나쁜 거짓말이다.
                sess["sub_layers"] = None
                lf = None
            sess["sub_layers_auto"] = True
            narrowed = {"layer": nm, **diag}
        try:
            got = graph_payload(sess["entities"], layer_filter=lf)
        except Exception as exc:  # noqa: BLE001 — 못 만들면 사유를 말한다
            return _fail(f"경로 그래프를 만들지 못했습니다: {exc}", 400)
        got.update({
            "ok": True,
            "kind": _slot_active(sess),
            "layers": layer_options(sess["entities"],
                                    sess.get("layer_colors")),
            "chosen": sorted(lf) if lf else None,
            "chosen_auto": bool(narrowed and narrowed.get("layer")),
            "narrowed": narrowed,
            "snap_default_mm": SNAP_DEFAULT_MM,
        })
        return jsonify(got)

    # ─────────────────────────────────── 뽑힌 배관 손보기 [§27 후속]
    @app.get("/api/module-f/sub/pipes")
    @route_session()
    def module_f_sub_pipes(sess, body):
        """뽑힌 배관표 — 볼 자리가 없으면 고칠 수도 없다.

        ★실측이 이 자리를 만들게 했다: 대명동 계통도는 뽑힌 배관 53개의 관경이
          **전부 «추측 150A»** 다(도면 치수 텍스트 매치 0건). 그 값이 그대로
          최종 SDF 의 입상관이 되는데 사람이 볼 길이 없었다.
        """
        kind = _slot_active(sess)
        if kind not in SLOT_KEY:
            return _fail("계통도·기계실 슬롯에서만 볼 수 있습니다.")
        got = sess.get(SLOT_KEY[kind])
        return jsonify({
            "ok": True, "kind": kind, "extracted": bool(got),
            "rows": rows_for_view(got) if got else [],
            "fixes": _fixes(sess, kind),
        })

    @app.post("/api/module-f/sub/pipe-fix")
    @route_session(post=True)
    def module_f_sub_pipe_fix(sess, body):
        """뽑힌 배관의 관경·길이를 사람이 덮는다.

        추출은 조각난 도면을 «다리» 로 이어 세우고 관경은 못 읽으면 150A 로
        둔다. 그 판정을 늘리는 대신 **고칠 자리**를 준다 — §27 이 두 번
        확인한 방향이다(추측 규칙을 얹으면 «틀린 확신» 만 는다).

        body: {sid, rows: [{a:[x,y], b:[x,y], dia?, length?, note?}]}
              빈 배열이면 전부 지우고 원래 값으로 돌아간다.
        """
        kind = _slot_active(sess)
        if kind not in SLOT_KEY:
            return _fail("계통도·기계실 슬롯에서만 고칠 수 있습니다.")
        got = sess.get(SLOT_KEY[kind])
        if not got:
            return _fail("먼저 경로를 추출하세요.", 409)
        try:
            rows = parse_rows(body.get("rows"))
        except ValueError as exc:
            return _fail(str(exc))
        stat = apply_overrides(got, rows)
        if stat["unmatched"]:
            # 조용히 버리지 않는다 — 어느 자리가 사라졌는지는 사람만 안다.
            return _fail(
                f"{stat['unmatched']}개는 지금 뽑힌 경로에 없는 구간입니다 — "
                f"다시 뽑으면서 그 구간이 빠졌는지 확인하세요.")
        sess.setdefault("sub_fixes", {})[kind] = rows
        print(f"[{kind}] 배관 손질 {len(rows)}개 적용")
        return jsonify({"ok": True, "rows": rows_for_view(got),
                        "fixes": rows, "counts": stat,
                        "summary": riser_summary(got)})

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
