# -*- coding: utf-8 -*-
"""[H-0] 모듈 F 라우트 — 도면 슬롯(특허 S650).

세 라우트뿐이다. 슬롯을 **열고 · 바꾸고 · 들여다본다.** 도면을 실제로 여는 일은
`api_open._open_job` 을 그대로 쓴다 — 특허 S650 이 «같은 절차를 반복 적용» 하라고
하므로, 계통도·기계실이 평면도와 다른 제1국면을 밟으면 그 자체가 오구현이다.

계통도·기계실의 **추출**(A 엔진 접합)은 H-2 · H-3 의 일이다. 여기서는 슬롯이
서로를 덮지 않는다는 계약까지만 세운다.
"""
from __future__ import annotations

import os

from flask import jsonify, request

from routes.module_f.api_open import _open_job
from routes.module_f.common import _boot, _check_key, _fail
from routes.module_f.jobs import _job_running, _new_session, _run_job, _sess
from routes.module_f.slots import (
    SLOT_LABELS, _check_slot_kind, _slot_active, _slot_state, _slot_switch)
from routes.module_f.world import _world_payload


def _auto_open_job(sess: dict, dxf):
    """[A 방식] 평면도를 «자동 추출» 로 연다 — 모듈 A 의 파서로 읽는다.

    수동(E)은 찍기판(PickSession)을 만들지만 여기는 그럴 이유가 없다. A 의
    위상 검출은 entity 목록과 레이어 분류 위에서 바로 돈다 — 사람이 정할 것은
    알람밸브 한 점과 헤드 영역뿐이다.
    """
    import os
    import time

    from routes.module_f.auto import parse_plan
    from routes.module_f.subdrawing import entities_to_world

    def job():
        t0 = time.perf_counter()
        print(f"[자동] DXF 읽는 중 — {os.path.basename(str(dxf))}")
        ents, layer_cat, diag = parse_plan(dxf)
        sess["entities"] = ents
        sess["layer_cat"] = layer_cat
        sess["auto_diag"] = diag
        # 키는 뒤에서 산출 파일 이름이 된다 — 수동 경로와 같은 자를 통과시킨다.
        sess["key"] = _check_key(os.path.splitext(os.path.basename(str(dxf)))[0])
        sess["world"] = _world_payload(entities_to_world(ents))
        print(f"[자동] 완료 {time.perf_counter() - t0:.1f}s · "
              f"도형 {diag['entities']:,} · 레이어 {diag['layers']}")
        cats = diag.get("cats") or {}
        print("[자동] 레이어 용도: "
              + " · ".join(f"{k} {v}" for k, v in sorted(cats.items())))
        # 외부참조 시트면 도면 내용이 딴 파일에 있다 — 헤드 0개로 끝난다.
        xr = diag.get("xref") or {}
        if xr.get("is_xref_shell"):
            print("[자동] ★이 파일은 외부참조(XREF) 껍데기입니다 — 도면 내용이 "
                  "딴 파일에 있어 헤드를 찾지 못할 수 있습니다.")
        return {"key": sess["key"], "entities": diag["entities"]}
    return job


def _sub_open_job(sess: dict, dxf, kind: str):
    """[H-2 · H-3] 계통도·기계실을 여는 잡 — 평면도와 다른 제1국면.

    평면도는 사람이 재료를 찍어야 하므로 E 의 `PickSession` 으로 간다. 계통도·
    기계실은 찍을 재료가 없다 — 두 점(펌프↔알람밸브 / 수원↔연결점)을 잇는
    경로가 전부라, 도면을 그대로 띄워 놓고 사람이 그 두 점을 찍는다.

    그래서 여기서는 A 의 시각화 파서로 읽어 캔버스에만 올린다. 추출은 두 점이
    정해진 뒤 `/api/module-f/<kind>/extract` 에서 한다.
    """
    import os
    import time

    from routes.module_f.subdrawing import entities_to_world, parse_subdrawing

    def job():
        t0 = time.perf_counter()
        label = SLOT_LABELS[kind]
        print(f"[{label}] DXF 읽는 중 — {os.path.basename(str(dxf))}")
        entities, parsed = parse_subdrawing(dxf)
        sess["entities"] = entities
        sess["key"] = os.path.splitext(os.path.basename(str(dxf)))[0]
        payload = _world_payload(entities_to_world(entities))
        sess["world"] = payload
        skipped = parsed.get("skipped") or {}
        print(f"[{label}] 완료 {time.perf_counter() - t0:.1f}s · "
              f"도형 {len(entities):,} · 선분 {payload['counts']['segs']:,}"
              + (f" · 못 읽은 것 {sum(skipped.values())}" if skipped else ""))
        if skipped:
            # 조용히 넘기지 않는다 — 못 그린 것이 배관이면 경로가 끊긴다.
            print(f"[{label}] 못 읽은 종류: "
                  + ", ".join(f"{k}×{v}" for k, v in sorted(skipped.items())))
        return {"key": sess["key"], "entities": len(entities)}
    return job


def register(app, *, _save_upload):
    @app.get("/api/module-f/slot/state")
    def module_f_slot_state():
        """세 슬롯의 진행 한 장 — S650 이 «남은 도면이 있나» 를 묻는 자리."""
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        out = _slot_state(sess)
        out["ok"] = True
        return jsonify(out)

    @app.post("/api/module-f/slot/switch")
    def module_f_slot_switch():
        """활성 슬롯을 바꾼다. 작업이 도는 중에는 거절한다.

        ★잡이 도는 중에 슬롯을 바꾸면 워커가 **다른 슬롯의 평면 dict** 에 결과를
          쓴다. `_open_job` 의 클로저가 붙잡은 것은 세션이지 슬롯이 아니기
          때문이다 — 계통도를 읽던 잡이 평면도의 찍기 상태를 덮어쓴다.
        """
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 도면을 바꿀 수 있습니다.", 409)
        try:
            kind = _slot_switch(sess, body.get("kind"))
        except ValueError as exc:
            return _fail(str(exc))
        out = _slot_state(sess)
        out["ok"] = True
        out["switched"] = kind
        return jsonify(out)

    @app.post("/api/module-f/slot/open")
    def module_f_slot_open():
        """도면을 이 슬롯에 **놓는다** — 올리기까지. 읽지는 않는다.

        ★읽기를 여기서 시작하지 않는 이유: 평면도는 방식마다 파서가 다르다
          (자동 = A 의 `parse_dxf_bundle`, 수동 = E 의 `PickSession.open`).
          방식을 모르는 채로 읽으려면 둘 다 돌리거나 하나를 찍어야 하는데,
          둘 다 돌리면 큰 도면에서 파싱을 두 번 하고(B1F 는 한 번이 9초가
          넘는다) 하나를 찍으면 사람이 고르기도 전에 길이 정해진다.

          그래서 «놓기» 와 «읽기» 를 가른다. 순서도 그것이 맞다 — 도면을
          올린 다음에 어떻게 읽을지 고른다. 읽기는 `/slot/read` 다.

        `sid` 가 있으면 그 세션의 해당 슬롯으로, 없으면 새 세션을 그 슬롯으로
        시작한다.
        """
        try:
            kind = _check_slot_kind(request.form.get("kind"))
        except ValueError as exc:
            return _fail(str(exc))

        sid = (request.form.get("sid") or "").strip()
        sess = None
        if sid:
            try:
                sess = _sess(sid)
            except ValueError as exc:
                return _fail(str(exc), 410)
            if _job_running(sess):
                return _fail("작업이 끝난 뒤에 도면을 열 수 있습니다.", 409)

        try:
            _boot()
            dxf = _save_upload("dxf_file", {".dxf"}, required=True)
        except ValueError as exc:
            return _fail(str(exc))
        except Exception as exc:  # noqa: BLE001
            return _fail(f"도면을 저장하지 못했습니다: {exc}", 500)

        if sess is None:
            sess = _new_session(slot=kind, dxf=str(dxf))
        else:
            _slot_switch(sess, kind)
            sess["dxf"] = str(dxf)
        # 올린 것뿐 — 아직 읽지 않았다. 앞서 이 슬롯에 있던 것은 지운다.
        sess["method"] = None
        for k in ("world", "pick", "edit", "entities", "layer_cat", "auto",
                  "auto_diag", "auto_heads", "auto_alarm", "auto_zones"):
            sess.pop(k, None)

        return jsonify({"ok": True, "sid": sess["id"], "kind": kind,
                        "filename": os.path.basename(str(dxf)),
                        # 평면도만 방식을 물어야 한다 — 계통도·기계실은 두 점
                        # 찍기 하나뿐이라 곧바로 읽으면 된다.
                        "needs_method": kind == "plan"})

    @app.post("/api/module-f/slot/read")
    def module_f_slot_read():
        """올려 둔 도면을 **읽기 시작한다** — 평면도는 방식을 여기서 정한다.

        올리기(`/slot/open`)와 가른 이유는 그쪽 머리말에 있다. 방식이 정해져야
        어느 파서로 읽을지가 정해지므로, 이 호출이 곧 «길을 고르는» 순간이다.
        """
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 읽을 수 있습니다.", 409)
        dxf = sess.get("dxf")
        if not dxf or not os.path.isfile(str(dxf)):
            return _fail("먼저 도면을 올리세요.", 400)
        kind = _slot_active(sess)

        method = str(body.get("method") or "").strip().lower()
        if kind == "plan":
            if method not in ("manual", "auto"):
                return _fail("추출 방식을 고르세요 — 자동(auto) 또는 수동(manual).")
        else:
            method = "manual"          # 계통도·기계실은 갈릴 것이 없다
        sess["method"] = method

        # 계통도·기계실은 찍을 재료가 없다(S650 의 «같은 절차» 는 같은 구현을
        # 뜻하지 않는다. subdrawing.py 머리말 참조).
        if kind != "plan":
            job = _sub_open_job(sess, dxf, kind)
            phase = f"{SLOT_LABELS[kind]} 읽기"
        elif method == "auto":
            job = _auto_open_job(sess, dxf)
            phase = "평면도 읽기 (자동)"
        else:
            job = _open_job(sess, dxf)
            phase = "평면도 읽기"
        _run_job(sess, phase, job)
        return jsonify({"ok": True, "sid": sess["id"], "kind": kind,
                        "method": method,
                        "filename": os.path.basename(str(dxf))})
