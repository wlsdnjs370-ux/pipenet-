# -*- coding: utf-8 -*-
"""[A 방식] 모듈 F 라우트 — 평면도 자동 추출(모듈 A 위상 검출).

수동(E)이 「색으로 배관을 찍고 헤드를 찍고 급수원을 찍는」 길이라면, 여기는
「알람밸브 한 점 + 헤드 영역」만 사람이 정하고 나머지를 A 가 하는 길이다.

    /auto/state      지금 무엇이 준비됐나 (도면 · 알람밸브 · 영역)
    /auto/heads      헤드 후보만 먼저 (영역을 정하기 전에 어디 있나 보려고)
    /auto/anchor     알람밸브 위치 지정
    /auto/run        선정 실행 → 5종 입력표
    /auto/preview    결과 캔버스 payload

★결과는 수동 경로와 **같은 자리**에 놓인다(`sess["design"]`) — 그래야 수리계산
  표·통합·산출이 두 길을 구분하지 않고 받는다. 다른 것은 오는 길뿐이다.
"""
from __future__ import annotations

import os

from flask import jsonify, request

from routes.module_f.auto import AutoError, detect_head_candidates, preview_view, run_auto
from routes.module_f.common import REMOTE_K_DEFAULT, _fail
from routes.module_f.jobs import _job_running, _run_job, _sess
from routes.module_f.slots import _slot_active


# 영역 개수 상한 — `HeadRegion.contains` 는 사각형마다 훑으므로 헤드 × 영역으로
# 늘어난다. 손으로 그리는 것이라 실제로는 몇 개면 충분하다.
MAX_ZONES = 64
# 화면에 그리는 헤드 후보 상한. 넘으면 «조용히» 자르지 않고 몇 개를 뺐는지 싣는다.
HEAD_PREVIEW_CAP = 4000


def _need_auto(body_or_args):
    """자동 슬롯이고 도면이 읽혀 있는가 — 작업이 도는 중이면 거절한다.

    ★잡이 도는 동안 이 세션의 값을 바꾸면, 워커가 읽는 것과 화면이 보는 것이
      갈린다. 잡 자체는 시작할 때 값을 읽으므로 계산이 틀어지진 않지만, 끝난
      뒤 「내가 넣은 영역으로 뽑힌 것」이라고 오해하게 된다.
    """
    sess = _sess(body_or_args.get("sid"))
    if _job_running(sess):
        return sess, "작업이 끝난 뒤에 바꿀 수 있습니다."
    if _slot_active(sess) != "plan":
        return sess, "평면도 슬롯에서만 자동 추출을 씁니다."
    if sess.get("method") != "auto":
        return sess, "이 도면은 수동(색 찍기)으로 열렸습니다."
    if not sess.get("entities"):
        return sess, "도면이 아직 준비되지 않았습니다."
    return sess, None


def register(app):
    @app.get("/api/module-f/auto/state")
    def module_f_auto_state():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        return jsonify({
            "ok": True,
            # ★기본값을 채우지 않는다. 아직 안 고른 슬롯을 «수동» 이라고
            #   답하면 화면이 고르지도 않은 길의 단계를 펼친다.
            "method": sess.get("method"),
            # 올려는 뒀는데 아직 안 읽은 슬롯으로 돌아왔을 때 화면이 무엇을
            # 올렸는지 말할 수 있게.
            "dxf_name": (os.path.basename(str(sess["dxf"]))
                         if sess.get("dxf") else None),
            "opened": bool(sess.get("entities")),
            "alarm": sess.get("auto_alarm"),
            # ★개수가 아니라 사각형 그대로 돌려준다. 슬롯을 오갔다 오면 화면은
            #   제 상태를 잃는데, 서버는 그대로 들고 있다 — 개수만 주면 «영역
            #   3곳» 이라 적히면서 캔버스에는 아무것도 안 그려져, 지워진 줄 알고
            #   다시 그리게 된다.
            "zones": [list(z) for z in (sess.get("auto_zones") or ())],
            "k": int(sess.get("auto_k") or REMOTE_K_DEFAULT),
            "diag": sess.get("auto_diag"),
            "done": bool(sess.get("auto")),
            "summary": (sess.get("auto") or {}).get("summary"),
        })

    @app.post("/api/module-f/auto/anchor")
    def module_f_auto_anchor():
        """알람밸브(기준점) 위치 — 특허 S210·S220 의 «사용자 지정»."""
        body = request.get_json(silent=True) or {}
        try:
            sess, why = _need_auto(body)
        except ValueError as exc:
            return _fail(str(exc), 410)
        if why:
            return _fail(why, 409)
        x, y = body.get("x"), body.get("y")
        if x is None or y is None:
            sess["auto_alarm"] = None            # 지우기
            return jsonify({"ok": True, "alarm": None})
        try:
            sess["auto_alarm"] = [float(x), float(y)]
        except (TypeError, ValueError):
            return _fail(f"좌표가 숫자가 아닙니다: {x!r}, {y!r}")
        return jsonify({"ok": True, "alarm": sess["auto_alarm"]})

    @app.post("/api/module-f/auto/zones")
    def module_f_auto_zones():
        """헤드 영역 — anchored 선정의 필수 입력(`head_region`)."""
        body = request.get_json(silent=True) or {}
        try:
            sess, why = _need_auto(body)
        except ValueError as exc:
            return _fail(str(exc), 410)
        if why:
            return _fail(why, 409)
        raw = body.get("zones") or []
        if len(raw) > MAX_ZONES:
            return _fail(f"영역이 너무 많습니다: {len(raw)}곳 "
                         f"(최대 {MAX_ZONES}). 넓은 사각형 하나로 묶으세요.")
        try:
            rects = [[float(v) for v in r[:4]] for r in raw]
        except (TypeError, ValueError, IndexError):
            return _fail("영역 좌표가 올바르지 않습니다 ([[x0,y0,x1,y1], …]).")
        if any(len(r) != 4 for r in rects):
            return _fail("영역은 [x0,y0,x1,y1] 네 값이어야 합니다.")
        sess["auto_zones"] = rects
        return jsonify({"ok": True, "zones": len(rects)})

    @app.post("/api/module-f/auto/heads")
    def module_f_auto_heads():
        """② 헤드 검출 — 도면에서 헤드를 **전부** 찾는다.

        모듈 A 의 `detect_heads`(R1~R5·신뢰도) 그대로다. 이것이 「알람밸브 찍고
        나면 자동으로 헤드가 다 나온다」의 그 단계이고, 범위(영역)의 기본값도
        여기서 나온다 — 영역을 그리는 것은 그것을 «좁히는» 선택일 뿐이다.
        """
        body = request.get_json(silent=True) or {}
        try:
            sess, why = _need_auto(body)
        except ValueError as exc:
            return _fail(str(exc), 410)
        if why:
            return _fail(why, 409)
        try:
            heads = detect_head_candidates(
                sess["entities"], sess["layer_cat"],
                rects=sess.get("auto_zones") or None)
        except Exception as exc:  # noqa: BLE001
            return _fail(f"헤드 후보를 찾지 못했습니다: {exc}", 500)
        sess["auto_heads"] = heads
        print(f"[자동] 헤드 검출 {len(heads):,}개"
              + (f" (영역 {len(sess['auto_zones'])}곳 안)"
                 if sess.get("auto_zones") else " (도면 전체)"))
        # 조용히 자르지 않는다 — 몇 개를 뺐는지 응답에 실어 화면이 그대로
        # 말하게 한다(이 저장소의 표시 상한 규약).
        shown = heads[:HEAD_PREVIEW_CAP]
        return jsonify({"ok": True, "n": len(heads), "heads": shown,
                        "shown": len(shown),
                        "dropped": len(heads) - len(shown)})

    @app.post("/api/module-f/auto/run")
    def module_f_auto_run():
        """선정 실행. 무거우므로 잡으로 돌린다."""
        body = request.get_json(silent=True) or {}
        try:
            sess, why = _need_auto(body)
        except ValueError as exc:
            return _fail(str(exc), 410)
        if why:
            return _fail(why, 409)
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 실행할 수 있습니다.", 409)
        if not sess.get("auto_alarm"):
            return _fail("알람밸브 위치를 먼저 찍으세요.", 400)
        # 영역은 «좁히는» 선택이다 — 안 그렸으면 검출한 헤드 전부를 범위로.
        try:
            k = max(1, min(int(body.get("k") or REMOTE_K_DEFAULT), 200))
        except (TypeError, ValueError):
            k = REMOTE_K_DEFAULT
        sess["auto_k"] = k

        def job():
            zones = sess.get("auto_zones") or []
            print(f"[자동] 최불리 추리기 — 기준개수 {k} · 범위 "
                  + (f"영역 {len(zones)}곳" if zones else "도면 전체"))
            got = run_auto(sess["entities"], sess["layer_cat"],
                           alarm_xy=sess["auto_alarm"],
                           rects=zones, k=k,
                           project_title=f"모듈 F 자동 — {sess.get('key') or ''}",
                           progress_cb=lambda f, m: print(f"[자동] {m}"))
            sess["auto"] = got
            # ★수동 경로와 같은 자리에 놓는다 — 하류가 두 길을 구분하지 않는다.
            #   `got` 은 G 의 제한전개 산출이라 자동 경로엔 없다. 빈 dict 를 두면
            #   미리보기의 담당 헤드 수가 0 으로 떨어질 뿐 나머지는 그대로 돈다.
            sess["design"] = {"got": {}, "tables": got["tables"], "k": k,
                              "schedule": None, "marks": {}, "method": "auto"}
            s = got["summary"]
            print(f"[자동] 완료 — 헤드 {s['k']} · 절점 {s['nodes']} · "
                  f"배관 {s['pipes']} · 최원 {s['far_m']} m")
            if s["source_fallback"]:
                print("[자동] ★급수원이 그래프에서 멀어 최근접 절점으로 대체됐습니다 "
                      "— 알람밸브 위치를 확인하세요.")
            return s

        _run_job(sess, "자동 추출", job)
        return jsonify({"ok": True, "sid": sess["id"]})

    @app.get("/api/module-f/auto/preview")
    def module_f_auto_preview():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        got = sess.get("auto")
        if not got:
            return _fail("아직 자동 추출을 실행하지 않았습니다.", 404)
        return jsonify({"ok": True, "view": preview_view(got["tables"]),
                        "summary": got["summary"],
                        "tables": got["tables"].as_dict()})
