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
from routes.module_f.common import _boot, _fail
from routes.module_f.jobs import (_job_running, _new_session, _run_job, _sess,
                                  route_session)
from routes.module_f.slots import (
    SLOT_LABELS, _check_slot_kind, _slot_active, _slot_state, _slot_switch)
# [오너 2026-09-22 · 그림 38] 계통도·기계실 읽기 — 서버가 뜰 때 불러 둔다. 그래야
#   «같은 도면 기억» 의 코드 판번호가 지금 돌고 있는 코드를 가리킨다.
from routes.module_f import sub_fastread  # noqa: F401
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

    from routes.module_f.sub_fastread import describe, read_view
    from routes.module_f.subdrawing import entities_to_world, layer_colors

    def job():
        t0 = time.perf_counter()
        label = SLOT_LABELS[kind]
        print(f"[{label}] DXF 읽는 중 — {os.path.basename(str(dxf))}")
        # [오너 2026-09-22 · 그림 38 ①③] 도면에 안 쓰이는 블록 정의는 속을 비운
        # 사본으로 읽고, 같은 도면은 기억해 둔 것을 쓴다. 결과는 A 의 파서로
        # 원본을 읽은 것과 같다(`parse_subdrawing` 과 같은 길 · sub_fastread).
        view = read_view(dxf)
        entities, parsed = view["entities"], view["parsed"]
        note = describe(view)
        if note:
            print(f"[{label}]   {note}")
        sess["entities"] = entities
        sess["key"] = os.path.splitext(os.path.basename(str(dxf)))[0]
        # 도면 색 그대로 그린다 — 계통도·기계실도 평면도와 같은 규칙이다.
        # 종전에는 한 색으로 눌러 그려서 배관·기호·건축선이 구별되지 않았다.
        colors = layer_colors(parsed)
        # 레이어 색은 «배관 레이어 고르기» 목록도 쓴다 — 목록과 도면이 같은
        # 색이라야 사람이 둘을 맞대 볼 수 있다.
        sess["layer_colors"] = colors
        payload = _world_payload(entities_to_world(entities, colors))
        if kind == "system":
            # ②: 도면을 먼저 띄우고 ★추적은 이어서 — world 는 그 안에서 앉는다.
            _system_layers(sess, entities, parsed, payload, label,
                           key=view.get("key"), t_open=t0)
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


# ★추적을 아직 고르는 중인 동안 표가 받는 자리 — 끝나면 서버가 정한 것으로 바뀐다
#   (화면은 `/sub/graph` 응답의 `trace` 로 표를 다시 그린다).
_TRACE_PENDING = {"mode": "pending", "layers": None, "junk": [], "candidates": [],
                  "forced": None, "split": None,
                  "reason": "경로 추적 레이어(★)를 고르는 중입니다 — 잠시 뒤 채워집니다."}


def _system_layers(sess: dict, entities, parsed, payload: dict, label: str, *,
                   key: str | None = None, t_open: float | None = None) -> None:
    """[오너 2026-09-22] 계통도 칸 — 레이어 세 묶음과 경로 추적 레이어(★).

    화면(`payload["sub_layers"]`)은 묶음 표로 레이어를 켜고 끄고, 기본값은
    배관망 + 건축만 보인다. 추적은 `sess["sub_trace"]` 를 쓴다(api_sub `_trace`).
    ★실패해도 도면 열기는 막지 않는다 — 종전 화면·종전 추적으로 떨어지고,
      그 사실을 로그에 남긴다(조용히 넘기지 않는다).

    [오너 2026-09-22 · 그림 38 ②③]
    ★도면을 먼저 띄운다. 표(묶음)까지 만든 뒤 `sess["world"]` 를 앉히고, ★추적
      그래프(B1F 6~7초)는 그 뒤에 만든다 — 두 점을 찍을 때 쓰는 것이지 도면을
      보는 데는 필요 없다. 그 사이 표의 ★칸은 «고르는 중» 이다.
    ★같은 도면이면 전에 고른 표 · ★추적을 그대로 쓴다(`sub_fastread` 기억 —
      열쇠는 파일 내용 + 코드 판번호라 코드가 바뀌면 다시 고른다).
    """
    import time

    from routes.module_f.sub_fastread import load_trace, store_trace
    from routes.module_f.sub_trace import (
        GROUP_LABELS, layer_rows, plan_trace, public_rows, public_trace)
    t0 = time.perf_counter()

    def say_groups(rows) -> None:
        n = {g: sum(1 for r in rows if r["group"] == g) for g in GROUP_LABELS}
        print(f"[{label}] 레이어 묶음 — 배관망 {n['net']} · 건축 {n['arch']} · "
              f"숨김 {n['etc']} ({time.perf_counter() - t0:.1f}s)")

    got = load_trace(key)
    if got is not None:
        rows, tr = got
        sess["sub_trace"] = tr
        payload["sub_layers"] = {"rows": public_rows(rows),
                                 "trace": public_trace(tr), "groups": GROUP_LABELS}
        sess["world"] = payload
        say_groups(rows)
        print(f"[{label}] 경로 추적 — {tr['reason']} (전에 골라 둔 것)")
        return
    try:
        rows = layer_rows(entities, parsed)
    except Exception as exc:  # noqa: BLE001
        sess["sub_trace"] = None
        print(f"[{label}] 레이어 묶음을 못 만들었습니다 — 종전 화면으로 엽니다: {exc}")
        return
    payload["sub_layers"] = {"rows": public_rows(rows), "trace": dict(_TRACE_PENDING),
                             "groups": GROUP_LABELS}
    sess["world"] = payload
    say_groups(rows)
    if t_open is not None:
        print(f"[{label}] 도면을 먼저 띄웁니다 ({time.perf_counter() - t_open:.1f}s) "
              f"— 경로 추적 레이어(★)는 이어서 고릅니다")
    t1 = time.perf_counter()
    try:
        tr = plan_trace(entities, rows)
    except Exception as exc:  # noqa: BLE001
        sess["sub_trace"] = None
        payload["sub_layers"]["trace"] = None
        print(f"[{label}] 경로 추적 레이어를 못 골랐습니다 — 자동 레이어로 추적합니다: {exc}")
        return
    dt = time.perf_counter() - t1
    sess["sub_trace"] = tr
    payload["sub_layers"]["trace"] = public_trace(tr)
    store_trace(key, rows, tr, dt)
    print(f"[{label}] 경로 추적 — {tr['reason']} ({dt:.1f}s)")


def register(app, *, _save_upload):
    @app.get("/api/module-f/slot/state")
    @route_session()
    def module_f_slot_state(sess, body):
        """세 슬롯의 진행 한 장 — S650 이 «남은 도면이 있나» 를 묻는 자리."""
        out = _slot_state(sess)
        out["ok"] = True
        return jsonify(out)

    @app.post("/api/module-f/slot/switch")
    @route_session(post=True)
    def module_f_slot_switch(sess, body):
        """활성 슬롯을 바꾼다. 작업이 도는 중에는 거절한다.

        ★잡이 도는 중에 슬롯을 바꾸면 워커가 **다른 슬롯의 평면 dict** 에 결과를
          쓴다. `_open_job` 의 클로저가 붙잡은 것은 세션이지 슬롯이 아니기
          때문이다 — 계통도를 읽던 잡이 평면도의 찍기 상태를 덮어쓴다.
        """
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
                  # 「배관으로 취급」 지정도 그 도면의 레이어 이야기다.
                  "auto_pipe_layers",
                  "design", "worst", "water_path",
                  # [F-8a] 정찰·제안은 «그 도면» 의 것이다. 남겨 두면 새 도면을
                  # 올렸는데 카드가 앞 도면의 후보 수를 그린다.
                  "recon", "suggest",
                  # [오너 2026-09-22] 계통도 칸의 추적 레이어·고른 배관 레이어도
                  # 그 도면의 것이다 — 남기면 새 도면을 옛 레이어로 추적한다.
                  "sub_trace", "sub_layers", "sub_layers_auto"):
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
    @route_session(post=True)
    def module_f_slot_read(sess, body):
        """읽어 놓은 도면을 **어느 길로 갈지 정한다**.

        도면은 `/slot/open` 이 이미 읽어 화면에 띄웠다. 여기서 갈리는 것은
        «그 다음» 이다:

            수동  더 읽을 것이 없다 — 이미 찍기판이 서 있다 (추가 0초)
            자동  A 의 파서로 한 번 더 읽어 위상 검출용을 보탠다 (실측 +5.3s)

        `started` 로 잡을 돌렸는지 알린다 — 화면이 기다릴지 바로 넘어갈지를
        그것으로 가른다.
        """
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
