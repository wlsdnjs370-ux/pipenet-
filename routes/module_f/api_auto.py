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


def _pipe_ents(sess):
    """자동이 볼 도형 — 사람이 「배관으로 취급」이라 지정한 묶음을 올려서 준다.

    자동 경로의 «모든» 입구가 이걸 거쳐야 한다. 한 군데라도 빠뜨리면 헤드 검출과
    망 검출이 서로 다른 도면을 보게 된다.
    """
    from routes.module_f.auto import apply_pipe_overrides
    return apply_pipe_overrides(sess["entities"], sess["layer_cat"],
                                sess.get("auto_pipe_layers"))


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
            # 사람이 「배관으로 취급」이라 찍은 묶음 — 슬롯을 오갔다 와도 되살린다.
            "pipe_layers": list(sess.get("auto_pipe_layers") or ()),
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

    @app.post("/api/module-f/auto/pipe-layers")
    def module_f_auto_pipe_layers():
        """「이 레이어를 배관으로 취급」 — 사람이 찍은 레이어×색 묶음.

        자동 차선에는 이 길이 없었다. 레이어 이름 사전이 OTHER 로 떨어뜨리면
        그것으로 끝이라, 사람이 보기에 명백한 배관도 손댈 수가 없었다.

        ★추측 규칙을 늘리지 않는 이유: 「선을 따라 헤드가 일정 간격·일정 거리로
          정렬」을 지문으로 재 봤더니 건축선(A-B1)에 28줄이 걸렸다. 벽이 배관과
          나란히 지나가기 때문이다. 그런 규칙은 다른 현장에서 벽을 배관으로 먹는다.
          수동 차선은 색으로 찍어 확정하는 길이 이미 있다 — 자동에도 그 결정을 준다.

        body: {sid, layers: [{layer, color}, …]}
        """
        body = request.get_json(silent=True) or {}
        try:
            sess, why = _need_auto(body)
        except ValueError as exc:
            return _fail(str(exc), 410)
        if why:
            return _fail(why, 409)
        raw = body.get("layers")
        if raw is None:
            raw = []
        if not isinstance(raw, list) or len(raw) > 64:
            return _fail("배관으로 취급할 묶음 목록이 올바르지 않습니다 "
                         "(최대 64묶음).")
        picks = []
        for it in raw:
            if not isinstance(it, dict) or not str(it.get("layer") or "").strip():
                return _fail("묶음은 {layer, color} 형식이어야 합니다.")
            picks.append({"layer": str(it["layer"]), "color": it.get("color")})
        sess["auto_pipe_layers"] = picks
        # 지정이 바뀌면 앞서 뽑은 것은 «다른 도면» 의 결과다 — 남겨 두면 사람이
        # 새 지정으로 나온 줄 안다.
        for k in ("auto_net", "auto", "design", "auto_heads"):
            sess.pop(k, None)
        n = sum(1 for e in sess["entities"]
                if any(str(e.get("l") or "0") == p["layer"] for p in picks))
        print(f"[자동] 배관으로 취급 — {len(picks)}묶음 · entity {n:,}개")
        return jsonify({"ok": True, "layers": picks, "entities": n})

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
            ents, cat = _pipe_ents(sess)
            heads = detect_head_candidates(
                ents, cat, rects=sess.get("auto_zones") or None)
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

    @app.post("/api/module-f/auto/network")
    def module_f_auto_network():
        """[S270 · S310] 배관망 검출 — 최불리를 고르기 «전» 의 단계.

        `scripts/평면도 배관망 추출논리.pdf` 의 순서에서 S320(내림차순 정렬) 앞
        까지다. 담당 헤드 수로 물 안 가는 관로를 자르고(S270), 밸브에서 각 헤드
        까지 거리를 잰다(S310). 최불리는 그 목록을 자르는 일이라 다음 단계다.

        이 단계가 없으면 사람은 «거리를 어디서 재는지» 를 못 보고 결과만 받는다.
        """
        from routes.module_f.auto import network_view, run_network

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
        # S270 가지치기 — A 는 기본 off 지만 논리 문서는 켜는 것을 전제한다.
        prune = bool(body.get("prune", True))

        def job():
            print(f"[망검출] S270 담당 헤드 수 · S310 거리 측정 "
                  + ("(물 안 가는 관로 잘라냄)" if prune else "(가지치기 끔)"))
            ents, cat = _pipe_ents(sess)
            got = run_network(ents, cat,
                              alarm_xy=sess["auto_alarm"],
                              rects=sess.get("auto_zones") or None,
                              prune=prune,
                              progress_cb=lambda f, m: print(f"[망검출] {m}"))
            sess["auto_net"] = got
            s = got["summary"]
            print(f"[망검출] 완료 — 절점 {s['nodes']:,} · 배관 {s['pipes']:,} · "
                  f"연장 {s['len_m']:,.1f} m")
            print(f"[망검출] 도달 헤드 {s['reached']:,}/{s['detected']:,} · "
                  f"거리 최근 {s['near_m']} m · 중앙 {s['mid_m']} m · "
                  f"최원 {s['far_m']} m")
            if s["cut_pipes"]:
                print(f"[망검출] 물 안 가는 관로 {s['cut_pipes']:,}개 · "
                      f"{s['cut_m']:,.1f} m 잘라냄 (S270)")
            return s

        _run_job(sess, "배관망 검출", job)
        return jsonify({"ok": True})

    @app.get("/api/module-f/auto/network-view")
    def module_f_auto_network_view():
        """검출한 망의 도형 — 새로고침·슬롯 왕복에도 다시 그릴 수 있게."""
        from routes.module_f.auto import network_view
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        got = sess.get("auto_net")
        # ★«아직 안 돌렸다» 는 오류가 아니다. 404 로 답하면 화면이 단계에 들어올
        #   때마다 콘솔에 붉은 줄을 남긴다 — 진짜 오류가 그 사이에 묻힌다.
        if not got:
            return jsonify({"ok": True, "summary": None, "view": None})
        return jsonify({"ok": True, "summary": got["summary"],
                        "view": network_view(got["selection"])})

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
            print(f"[자동] 최불리 추출 — 기준개수 {k} · 범위 "
                  + (f"영역 {len(zones)}곳" if zones else "도면 전체"))
            ents, cat = _pipe_ents(sess)
            got = run_auto(ents, cat,
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

    @app.post("/api/module-f/auto/handoff")
    def module_f_auto_handoff():
        """[F-8d] 탈출로 — 자동 결과가 이상할 때 손질로 이어받는다.

        자동이 마음에 안 든다고 처음부터 다시 시작하게 두지 않는다. 같은
        세션의 찍기판(`sess["pick"]`)은 살아 있다 — 자동 차선도 열기는 같은
        `/open` 을 탄다. 그것을 그대로 써서 채택 → 스펙 저장 → 손질 진입까지
        **잡 하나**로 잇는다.

        채택 범위는 «정찰 후보 전체» 다. 자동이 영역으로 좁혔더라도 이어받기는
        넓게 준다 — 좁히는 것은 손질에서 사람이 할 일이고, 여기서 미리 잘라
        두면 그 결정을 되돌릴 방법이 없다.

        자동이 알던 것(알람밸브)은 **제안** 으로만 넘긴다. 반영은 손질의 기존
        클릭 경로(`edit/mode` + `edit/click`)로만 한다 — D-F8-3 은 여기서도
        같다. 알람밸브 자리와 급수 시작 자리를 하나로 합칠지는 미결이라
        (BLOCKED §5) 제안 두 개로 나눠 둔다.
        """
        from routes.module_f.adopt import adopt_bundles, adopt_heads, select_heads
        from routes.module_f.remote30 import _sheet_frames
        from routes.module_f.views import _pick_state

        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 이어받을 수 있습니다.", 409)
        if _slot_active(sess) != "plan":
            return _fail("평면도 슬롯에서만 이어받을 수 있습니다.", 409)
        if not sess.get("auto"):
            return _fail("아직 자동 추출을 실행하지 않았습니다.", 409)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기판이 없는 세션입니다 — 도면을 다시 여세요.", 409)
        rec = sess.get("recon") or {}
        if rec.get("error"):
            return _fail(f"정찰이 실패한 도면이라 이어받을 것이 없습니다. "
                         f"({rec['error']})", 409)

        cands = list(rec.get("heads") or ())
        world = sess.get("world") or {}
        alarm = sess.get("auto_alarm")

        def job():
            from services.cad_import.edit.session import EditSession
            print(f"[이어받기] 자동 결과를 손질로 — 후보 {len(cands)}개 전부")
            ps.select_pipe()
            mat = adopt_bundles(ps, world, "PIPE")
            print(f"[이어받기] 재료 {len(mat['applied'])}묶음 "
                  f"(건너뜀 {len(mat['skipped'])})")
            if not ps.complete_pipe():
                raise RuntimeError(
                    "재료를 하나도 못 찍었습니다 — 배관 레이어를 직접 찍어 주세요.")
            ps.set_slot(ps.head_label)
            got = adopt_heads(ps, select_heads(cands),
                              progress=lambda n, t, a, d, b: print(
                                  f"[이어받기] {n}/{t} — 찍힘 {a} · 이미 {d} · 유령 {b}"))
            print(f"[이어받기] 헤드 — 찍힘 {got['applied']} · "
                  f"이미 반영 {got['already']} · 유령 {len(got['skipped'])}")

            spec_path = ps.commit()
            print(f"[찍기] 스펙 저장 — {spec_path}")
            print("[손질] 찍은 스펙으로 배관망을 다시 구성하는 중…")
            es = EditSession.open(ps.key, out_dir=None, load_saved=False,
                                  use_cache=False)
            sess["edit"] = es
            sess["sheets"] = _sheet_frames(es.board)
            # ★자동 흐름을 떠난다 — 단계바·슬롯 진행이 손질 쪽을 가리켜야 한다.
            #   `auto`·`design` 은 지우지 않는다: 손질 뒤 사람이 다시 최불리를
            #   고르면 그때 덮인다(기존 규약).
            sess["method"] = "manual"
            # 자동이 알던 것 — «제안» 이다. 반영은 손질 클릭으로만 한다.
            sess["handoff"] = {
                "alarm": list(alarm) if alarm else None,
                "source": list(alarm) if alarm else None,
                "mat_applied": mat["applied"],
                "head_applied": got["applied"],
                "head_already": got["already"],
                "head_skipped": len(got["skipped"]),
                "skipped_heads": got["skipped"],
            }
            print(f"[손질] 완료 · 노드 {len(es.board.pts)} · "
                  f"간선 {len(es.board.edges)} · 헤드 {len(es.board.disks)}")
            return {"ok": True, "spec_path": spec_path,
                    **sess["handoff"], "pick": _pick_state(sess)}

        _run_job(sess, "손질로 이어받기", job)
        return jsonify({"ok": True})

    @app.get("/api/module-f/auto/handoff-hints")
    def module_f_auto_handoff_hints():
        """이어받기가 넘긴 제안 — 새로고침해도 오버레이가 다시 뜨게."""
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        return jsonify({"ok": True, "handoff": sess.get("handoff")})
