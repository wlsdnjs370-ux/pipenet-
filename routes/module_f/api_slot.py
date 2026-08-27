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


def _auto_augment_job(sess: dict, dxf):
    """[A 방식] 자동 추출이 쓸 것을 **보태는** 잡 — 도면은 이미 화면에 있다.

    ★열기는 이미 E 의 `PickSession` 으로 끝났다(그쪽이 싸다 — 실측 2.5s vs
      A 의 파서 5.3s). 그래서 화면에 그릴 것(`world`)은 그대로 두고, A 의
      위상 검출이 필요한 것만 얹는다: entity 목록과 레이어 분류.

      「불러오기 → 도면이 보인다」 는 방식과 무관해야 한다. 방식을 고르기
      전에 화면이 비어 있으면 무엇을 고르는지 모른 채 고르게 된다.
    """
    import os
    import time

    from routes.module_f.auto import parse_plan

    def job():
        t0 = time.perf_counter()
        print(f"[자동] 위상 검출용으로 다시 읽는 중 — "
              f"{os.path.basename(str(dxf))}")
        print("[자동]   (자동은 모듈 A 의 파서를 따로 씁니다. 처음 한 번만 "
              "오래 걸리고, 같은 도면은 다음부터 즉시입니다.)")
        ents, layer_cat, diag = parse_plan(dxf)
        sess["entities"] = ents
        sess["layer_cat"] = layer_cat
        sess["auto_diag"] = diag
        # world 는 덮지 않는다 — 찍기판이 만든 것이 색까지 살아 있어 더 낫다.
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
        return {"key": sess.get("key"), "entities": diag["entities"]}
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
        """도면을 올려 **읽고 화면에 띄운다** — 방식과 무관한 공통 단계.

        ★「불러오기 → 도면이 보인다」 는 방식이 무엇이든 같아야 한다. 방식을
          고르기 전에 화면이 비어 있으면 무엇을 고르는지 모른 채 고르게 된다.

        ★어느 파서로 여느냐 — 실측으로 정했다(scripts/_probe_parse_cost.py,
          LH306 16MB):

              PickSession.open    2.46s   ← 여기서 쓴다
              parse_dxf_bundle    5.29s
              parse_dxf_for_view  5.42s

          찍기판이 가장 싸고 화면에 그릴 것도 충분하다(선분 26,377 · 원 807).
          그리고 수동을 고르면 **추가 파싱이 0** 이다 — 이미 그것이 수동이
          쓰는 바로 그 판이다. 자동을 고른 경우에만 A 의 파서를 한 번 더
          돌려 위상 검출용 entity·레이어 분류를 보탠다(`/slot/read`).

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
        # 새 도면이다 — 앞서 이 슬롯에 있던 것은 지운다. 남겨 두면 새 도면을
        # 올렸는데 옛 결과가 그대로 뜬다.
        sess["method"] = None
        for k in ("world", "pick", "edit", "entities", "layer_cat", "auto",
                  "auto_diag", "auto_heads", "auto_alarm", "auto_zones",
                  # [S270·S310] 검출한 망도 «그 도면» 의 것이다.
                  "auto_net",
                  "design", "worst", "water_path",
                  # [F-8a] 정찰·제안은 «그 도면» 의 것이다. 남겨 두면 새 도면을
                  # 올렸는데 카드가 앞 도면의 후보 수를 그린다.
                  "recon", "suggest"):
            sess.pop(k, None)

        # 읽어서 화면에 띄우는 것까지가 공통이다.
        # [D-F8-2] 정찰은 평면도만 — `_open_job` 이 종류를 보고 가른다.
        job = (_open_job(sess, dxf, kind=kind) if kind == "plan"
               else _sub_open_job(sess, dxf, kind))
        _run_job(sess, f"{SLOT_LABELS[kind]} 읽기", job)
        return jsonify({"ok": True, "sid": sess["id"], "kind": kind,
                        "filename": os.path.basename(str(dxf)),
                        # 평면도만 방식을 물어야 한다 — 계통도·기계실은 두 점
                        # 찍기 하나뿐이라 갈릴 것이 없다.
                        "needs_method": kind == "plan"})

    @app.post("/api/module-f/slot/read")
    def module_f_slot_read():
        """읽어 놓은 도면을 **어느 길로 갈지 정한다**.

        도면은 `/slot/open` 이 이미 읽어 화면에 띄웠다. 여기서 갈리는 것은
        «그 다음» 이다:

            수동  더 읽을 것이 없다 — 이미 찍기판이 서 있다 (추가 0초)
            자동  A 의 파서로 한 번 더 읽어 위상 검출용을 보탠다 (실측 +5.3s)

        `started` 로 잡을 돌렸는지 알린다 — 화면이 기다릴지 바로 넘어갈지를
        그것으로 가른다.
        """
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 고를 수 있습니다.", 409)
        dxf = sess.get("dxf")
        if not dxf or not os.path.isfile(str(dxf)):
            return _fail("먼저 도면을 올리세요.", 400)
        kind = _slot_active(sess)
        if not sess.get("world"):
            return _fail("도면을 아직 다 읽지 못했습니다.", 409)

        method = str(body.get("method") or "").strip().lower()
        if kind == "plan":
            if method not in ("manual", "auto"):
                return _fail("추출 방식을 고르세요 — 자동(auto) 또는 수동(manual).")
        else:
            method = "manual"          # 계통도·기계실은 갈릴 것이 없다
        sess["method"] = method

        started = False
        if kind == "plan" and method == "auto":
            _run_job(sess, "자동 추출 준비", _auto_augment_job(sess, dxf))
            started = True
        return jsonify({"ok": True, "sid": sess["id"], "kind": kind,
                        "method": method, "started": started,
                        "filename": os.path.basename(str(dxf))})
