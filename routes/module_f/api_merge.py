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
    SUPPLY_MODES, MergeError, check_supply_mode, combined_summary,
    merge_network)
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
