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

from routes.module_f import overrides as ov
from routes.module_f.common import _fail
from routes.module_f.jobs import _job_running, _run_job, route_session
from routes.module_f.merge import (
    ANCHOR_LABEL, SUPPLY_MODES, MergeError, bake_combined_iso, bake_combined_plan,
    check_supply_mode, combined_summary, merge_network)
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


def _split_note(sess: dict):
    """[오너 2026-09-21] 편집으로 통합망이 끊겼으면 그 사실을 한 문장으로 준다.

    결합 직후(편집 전)보다 덩어리가 늘었을 때만 말한다 — 결합 때부터 따로
    떨어진 조각이 있던 도면은 종전과 똑같이 산출된다.
    """
    from routes.module_f.merge import SPLIT_MESSAGE, network_pieces
    got = sess.get("merged") or {}
    editor = sess.get("merge_editor")
    if got.get("combined") is None or not editor or editor.get("base") is None:
        return None
    if network_pieces(got["combined"]) > network_pieces(editor["base"].tables):
        return SPLIT_MESSAGE
    return None


def _merge_underlay(sess: dict, got: dict, nodes, iso: bool, zs: float):
    """[§3-4] 통합 밑그림 변환 — board mm 한 점을 **이 화면** 의 자리로.

    04 의 `_underlay_xf` 와 **같은 이름·같은 모양**의 숫자 일곱 개다. 화면이
    한 함수(`drawUnderlay`)로 두 화면을 그리므로 모양이 갈리면 안 된다.

    ■ 두 화면의 다른 점은 «배율» 하나다

      04 는 표 좌표를 `norm.scale` 로 정규화해 그리지만, 통합은 **표 mm 를
      그대로** 쓴다(`merge_network` 가 좌표를 안 건드린다 — §2 실측으로
      기준점 차 0.000 mm). 그러니 여기서 k 는 1 이고, 남는 것은 board 원점을
      표 원점으로 옮기는 평행이동뿐이다:

        표 mm = board mm − origin + 1000        (`main_walk.xf_mm_to_m` 의 mm 판)

    ■ 표고

      통합의 평면 식은 `bake_combined_iso` 에서 `(e − e_ref)·lift` 이고
      e_ref 는 **기준점(라벨 10)** 의 표고다. board 평면이 놓인 높이는 그
      기준점 자리, 즉 평면도가 라이저와 만나는 접속점이므로 두 값이 같다 —
      dz 는 0 이 된다. 그래도 0 을 박아 두지 않고 **재서** 보낸다. 기준점이
      옮겨지면 그 사실이 여기 그대로 드러나야 한다.

    ★못 만들면 None 이다 — 화면은 그때 밑그림을 그리지 않는다(F-10e 규약).
    """
    origin = ((_slot_value(sess,"plan","design") or {}).get("got") or {}).get("origin_mm")
    if not origin:
        return None
    elev = {str(n.get("label")): float(n.get("elevation", 0) or 0)
            for n in (nodes or ())}
    e_ref = elev.get(ANCHOR_LABEL)
    if e_ref is None:
        return None
    return {
        "k": 1.0,
        "tx": 1000.0 - float(origin[0]),
        "ty": 1000.0 - float(origin[1]),
        "cos30": 0.8660254037844387, "sin30": 0.5,
        "iso": bool(iso),
        "lift": 1000.0 * float(zs or 1.0),
        "e_ref": e_ref,
        "e": e_ref,
    }


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



def rebuild_merged(sess: dict, *, persist_overrides: bool = True) -> dict:
    """Rebuild from current tables, sharing all existing override handling."""
    mode = sess.get("supply_mode")
    mats = _materials(sess)
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
    # Library choices made in the live editor must survive the merge's default
    # bore normalization. Otherwise the displayed DN and saved inner bore split.
    chosen = {str(p['label']):p for p in mats['plan'].pipes if p.get('inner_mm') is not None}
    if got.get('combined') is not None:
        for row in got['combined'].pipes:
            spec = chosen.get(str(row.get('label')))
            if spec:
                for field in ('type','dia','inner_mm','c','roughness_mm'):
                    row[field] = spec[field]
    sess["merged"] = got
    # ★[요소속성 수정카드] 적용 ④ — 계통도·기계실 요소(§4).
    #
    #   회랑 요소는 여기 없다. 그것은 설계 표에서 이미 덮여 결합으로
    #   흘러든다 — 두 자리가 같은 값을 덮으면 한쪽만 고치는 날 두
    #   화면이 갈린다. 여기서는 **통합에서만 사는 것**만 덮는다.
    #   여러 번 눌러도 같은 결과다(결합망을 매번 새로 만든다).
    el_rows = ov.ensure_loaded(sess)
    mg_missed = []
    if el_rows:
        _n, mg_missed = ov.apply_to_merge(got, el_rows)
        if _n and persist_overrides:
            ov.save(sess, el_rows)     # 원값이 채워졌다 — 카드가 본다
            try:
                ov.write_file(sess.get("key") or "design", el_rows)
            except OSError as exc:
                print(f"[수정] ★원값을 파일에 못 썼습니다 — {exc}")
    # ★[§3-3-1] 적용 ★ — 계통도·기계실 대상 **위상** 수정.
    #
    #   회랑 요소의 위상 수정은 여기 없다. 그것은 04 의 「표 확정」
    #   에서 이미 망에 먹었고, 그 표가 결합으로 흘러들었다(D8) —
    #   그래서 04 의 `.sdf` 와 통합 `.sdf` 가 같은 망을 말한다(⑨).
    mg_ops = ov.ensure_ops_loaded(sess)
    if mg_ops:
        _no, _mo, _rep = ov.apply_ops_to_merge(got, mg_ops)
        mg_missed = list(mg_missed) + list(_mo)
        # ★[§3-3-1 ⑤] 「사람이 **만든** 요소의 값 수정」은 여기서야
        #   돌 수 있다 — 위 ④ 에서 막 만들어졌기 때문이다. 위의 값
        #   수정(④ 전)은 그 요소를 아직 못 가리킨다.
        _add = [r for r in el_rows
                if str((r.get("key") or [None])[0]) == "add"]
        if _add:
            _n5, _m5 = ov.apply_to_merge(got, _add)
            mg_missed = list(mg_missed) + list(_m5)
            if _n5 and persist_overrides:
                ov.save(sess, el_rows)
    sess["merge_missed"] = mg_missed
    if mg_missed:
        print(f"[결합] ★적용 못 한 요소 수정 {len(mg_missed)}건 — "
              "조용히 버리지 않고 화면에 올린다")
    summary = combined_summary(got)
    sess["merge_summary"] = summary
    for line in summary.get("steps") or ():
        print(f"[결합]   · {line}")
    print(f"[결합] 완료 — 절점 {summary['nodes']} · 배관 {summary['pipes']}"
          f" · 노즐 {summary['nozzles']}")
    return summary

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
            # [D5] 결합 뒤 검사 — 화면이 «성립했는가» 를 볼 수 있게 그대로.
            "checks": ((sess.get("merged") or {}).get("checks")
                       if isinstance(sess.get("merged"), dict) else None),
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
            summary = rebuild_merged(sess)
            if (sess.get('merged') or {}).get('combined'):
                from routes.module_f.network_edit import accept_rebuilt
                editor = accept_rebuilt(sess,'merge')
                summary = combined_summary(sess['merged'])
                summary['editor_notice'] = editor.get('notice')
                sess['merge_summary'] = summary
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
        if got.get("attached") and got.get("pump_junction"):
            shared.add(str(got["pump_junction"]))

        nodes = [dict(n) for n in (c.nodes or ())]
        mr_edges = [list(map(float, e)) for e in
                    (getattr(c, "machine_room_plan_edges", None) or ())]
        iso = (request.args.get("iso") or "0") in ("1", "true", "True", "on")
        try:
            zs = float(request.args.get("iso_z_scale") or 1.0)
        except (TypeError, ValueError):
            zs = 1.0
        # ★표고 기준은 **굽기 전** 좌표에서 읽는다 — 구운 뒤 nodes 의 x·y 는
        #   화면 자리로 바뀌지만 elevation 은 그대로라 어느 쪽이든 같다. 그래도
        #   굽기 전에 잡아 두면 「무엇으로 쟀나」가 한 줄로 읽힌다(§3-4).
        under = _merge_underlay(sess, got, nodes, iso, zs)
        if iso:
            # 굽는 식은 `merge.bake_combined_iso` 하나뿐이다 — 산출(.sdf)도
            # 같은 함수를 쓴다. 두 자리가 각자 셈하면 화면과 파일이 갈린다.
            nodes, mr_edges = bake_combined_iso(got, iso_z_scale=zs)
        else:
            # [오너 2026-09-21] 평면 보기 — 계통도 세로관을 세로로 세운다(높이차
            #   없는 배관은 그대로). 식은 `merge.bake_combined_plan` 하나뿐이고
            #   평면 .sdf 도 같은 함수를 쓴다(보이는 것 = 저장되는 것).
            nodes, mr_edges = bake_combined_plan(got)

        heads = {str(r.get("in")) for r in (c.nozzles or ())}
        pumps = {str(r.get("in")) for r in (c.pumps or ())}
        pumps |= {str(r.get("out")) for r in (c.pumps or ())}
        valves = {str(r.get("in")) for r in (c.valves or ())}

        # ★[요소속성 수정카드 §4] **안정 키**를 함께 싣는다 — 통합 화면에서도
        #   그 자리에서 값을 고칠 수 있어야 한다(오너 2026-09-14 「계통도,
        #   기계실도 별도로 수정 만들면 좋지」).
        #
        #   회랑(평면도) 요소는 수리계산 화면과 **같은 주소**를 쓴다. 라벨만
        #   S740 이 +9 옮겼을 뿐 같은 자리다(배관 이름은 안 옮긴다 —
        #   `to_head_tables` 가 노드 라벨만 민다). 그러니 라벨을 되밀어 설계
        #   주소록에서 찾는다. 계통도·기계실은 board 가 없어 라벨이 곧 주소다.
        from routes.module_f.merge import label_offset_for
        _dk = (_slot_value(sess,"plan","design") or {}).get("keys") or {}
        _off = label_offset_for(_materials(sess)["plan_method"])

        def _plan_node_key(lab):
            try:
                back = str(int(lab) - _off)
            except (TypeError, ValueError):
                return None
            return (_dk.get("node") or {}).get(back)

        def _key_of_node(lab, part):
            if part == "system":
                return ["sys", lab]
            if part == "machineroom":
                return ["mr", lab]
            return _plan_node_key(lab)

        def _key_of_pipe(lab, part):
            if part == "system":
                return ["sys", lab]
            if part == "machineroom":
                return ["mr", lab]
            if part == "plan":
                return (_dk.get("pipe") or {}).get(str(lab))
            return None                 # 이음매 — 어느 도면 것도 아니다

        out_nodes = []
        for n in nodes:
            lab = str(n.get("label"))
            rec = {"label": lab, "x": float(n.get("x", 0) or 0),
                   "y": float(n.get("y", 0) or 0),
                   "e": round(float(n.get("elevation", 0) or 0), 3),
                   "part": of.get(lab, "plan")}
            rec["key"] = _key_of_node(lab, rec["part"])
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
            part = (got.get("pipe_parts") or {}).get(str(r.get("label")))
            if part is None:  # Legacy snapshots and explicitly edited new edges.
                part = ka if ka == kb else "seam"
            out_pipes.append({"label": str(r.get("label")), "a": a, "b": b,
                              "dia": r.get("dia"), "len_m": r.get("length"),
                              "c": r.get("c"), "elev": r.get("elev"),
                              "type": r.get("type"), "inner_mm": r.get("inner_mm"),
                              "eq_len": r.get("eq_len"),
                              "part": part,
                              "boundary": a in shared or b in shared,
                              "key": _key_of_pipe(str(r.get("label")), part)})

        from routes.module_f.underlays import reference_layers
        layers = reference_layers(sess,got,iso=iso,geometry=request.args.get('underlays')=='1')
        from routes.module_f.fitting_inspection import build_merged_inspection
        inspection = build_merged_inspection(
            c, _slot_value(sess, "plan", "design") or {}, out_nodes,
            offset=_off, transform=under,
            board=getattr(_slot_value(sess, "plan", "edit"), "board", None),
            plan_editor=_slot_value(sess, "plan", "network_editor"),
            merge_editor=sess.get("merge_editor"))
        return jsonify({
            "ok": True, "iso": iso,
            "view": {"nodes": out_nodes, "pipes": out_pipes,
                     "inspection": inspection,
                     # 기계실 평면 배관망 — SDF 에는 없고 «보기» 로만 쓴다.
                     "mr_plan_edges": mr_edges,
                     # [§3-4] 평면도 밑그림 변환 — 04 와 같은 이름·같은 모양.
                     "underlay": under, "underlays": layers,
                     # [오너 2026-09-21] 세로관을 「아니오」로 지워 끊겼으면 화면 하단 경고 문구.
                     "split": _split_note(sess)},
            # [요소속성 수정카드] 카드가 원값·사유·시각을 나란히 보인다(규칙 5)
            #   와 「적용 못 한 수정」(규칙 4).
            "overrides": ov.ensure_loaded(sess),
            # [§3-3] 위상 수정 — 두 화면이 **같은 목록**을 본다(⑨).
            "ops": ov.ensure_ops_loaded(sess),
            "ov_missed": sess.get("merge_missed") or [],
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

        # [오너 2026-09-21] 세로관을 「아니오」로 지워 끊긴 통합망은 산출하지 않는다.
        split = _split_note(sess)
        if split:
            return _fail(split, 409)

        from pathlib import Path

        out_dir = Path(UPLOAD_DIR).parent / "module_f_merged" / sess["id"]

        def job():
            from routes.module_f.emit import emit_merged
            print("[결합] S750 입력파일 생성")
            # 아이소 좌표 한 벌을 함께 낸다 — 화면 미리보기가 쓰는 **그 함수**로
            # 구운 절점을 넘긴다(두 자리가 각자 셈하면 화면과 파일이 갈린다).
            iso_nodes, _iso_edges = bake_combined_iso(got)
            files = emit_merged(
                got["combined"], out_dir,
                title=f"모듈 F 통합 — {sess.get('key') or ''}",
                iso_nodes=iso_nodes,
                # 평면 .sdf 의 그림 자리 — 평면 보기 화면과 같은 함수(bake_combined_plan).
                plan_nodes=bake_combined_plan(got)[0],
                display_reference_labels=list((got.get("parts") or {}).get("plan") or ()))
            sess["merge_files"] = files
            for k, v in files.items():
                if isinstance(v, str) and v:
                    print(f"[결합]   {k}: {os.path.basename(v)}")
            for warning in files.get("warnings", []):
                print(f"[결합] {warning}")
            return {**{k: os.path.basename(v) for k, v in files.items()
                       if isinstance(v, str) and v}, "warnings": files.get("warnings", [])}

        _run_job(sess, "산출물 생성", job)
        return jsonify({"ok": True, "sid": sess["id"]})

    @app.get("/api/module-f/merge/download")
    @route_session()
    def module_f_merge_download(sess, body):
        """산출물 내려받기 — 세션이 만든 것만, **형태별로 하나씩**.

        ★[오너 2026-09-14] 「zip 한 벌」 단추를 없앴다. `what` 은 `emit_merged`
          가 낸 이름 그대로다 — `sdf · slf · kfp · has` 와 아이소본
          `sdf_iso · slf_iso · kfp_iso · has_iso`. SDF 는 `.slf` 와 한 쌍이라
          화면이 두 번 부른다(묶어 주면 압축을 푸는 손이 한 번 더 든다).
          `zip` 도 여전히 받긴 하지만 화면에 단추는 없다.
        """
        what = str(request.args.get("what") or "sdf")
        files = sess.get("merge_files") or {}
        path = files.get(what)
        if not isinstance(path, str) or not path or not os.path.isfile(path):
            return _fail(f"그런 산출물이 없습니다: {what}", 404)
        if what in ("kfp", "kfp_iso"):
            from routes.module_f.kfp_export import send_kfp
            return send_kfp(path, os.path.basename(path))
        return send_file(path, as_attachment=True,
                         download_name=os.path.basename(path))
