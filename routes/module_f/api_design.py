# -*- coding: utf-8 -*-
"""[F-2] 수리계산 입력 라우트 — G 의 design/ 을 그대로 HTTP 로 연다.

엔진 함수는 전부 `cad_project_editor_g/services/cad_import/design/` 에서
import 한다 — 여기서는 판단하지 않는다(F 를 위해 G 를 몰래 바꾸는 것 금지).

    POST /api/module-f/design/build     최불리 → 제한 전개 → 5표 (메모리)
    GET  /api/module-f/design/preview   정규화+베이크 좌표 (저장 좌표 그대로)
    POST /api/module-f/design/emit      .sdf + .slf 저장 → 내려받기 링크

★미리보기 전용 좌표계는 없다. preview 와 emit 은 **같은 함수**
(`display_tables`)의 사본을 쓴다 — 화면과 파일이 달라지면 미리보기가 무의미하다
(G16 원칙 그대로).
"""
from __future__ import annotations

import os
from pathlib import Path

from flask import jsonify, request, send_file

from routes.module_f.common import (
    REMOTE_K_DEFAULT, _boot, _fail, _r1)
from routes.module_f.jobs import _job_running, _run_job, _sess, route_session

# 설정 7종(지시서 F-2) — body 로 받고 세션에 기억한다.
_DEFAULT_SETTINGS = {
    "k": REMOTE_K_DEFAULT,           # 기준개수
    "schedule": "KSD 3507",          # 배관 규격 기본값
    "iso": True,                     # 아이소매트릭 보기
    "iso_z_scale": 1.0,              # 고도 펼침 배율
    "canvas_units": 3000.0,          # 캔버스 크기
    "lift_ref": "valve",             # lift 영점 (valve | mid)
    "head_stub_pct": 2.5,            # 헤드 스텁 길이 (%)
}


def _settings(sess: dict, body: dict) -> dict:
    """설정 7종 — 준 것만 덮고 세션에 기억한다."""
    cur = dict(sess.get("design_settings") or _DEFAULT_SETTINGS)
    for key, cast in (("k", int), ("schedule", str), ("iso", bool),
                      ("iso_z_scale", float), ("canvas_units", float),
                      ("lift_ref", str), ("head_stub_pct", float)):
        if key in body and body[key] is not None:
            try:
                cur[key] = cast(body[key])
            except (TypeError, ValueError):
                return_fail = f"설정 값이 올바르지 않습니다: {key}={body[key]!r}"
                raise ValueError(return_fail)
    cur["k"] = max(1, min(int(cur["k"]), 200))
    sess["design_settings"] = cur
    return cur


def _dia_texts(sess: dict) -> list:
    """치수 텍스트 — handoff 캐시에서. 못 읽으면 빈 목록(관경은 별표1 폴백).

    G 데스크톱 4번째 창(`dialog_design_input._dia_texts`)과 같은 경로다 —
    여기가 다르면 같은 도면의 관경 근거가 웹과 데스크톱에서 갈라진다.

    ★실패를 조용히 넘기지 않는다. 여기가 빈 목록을 돌려주면 관경이 **전부**
      별표1 폴백이 되는데, 화면에는 «폴백 100%» 라는 결과만 남고 원인은
      어디에도 안 나온다. 실측으로 그렇게 한 번 당했다 — 원본 DXF 의 mtime 만
      바뀌어 handoff 캐시가 기각되면서 치수 텍스트 533개가 통째로 사라졌고,
      관경표는 아무 경고 없이 규약값으로만 채워졌다.
    """
    key = sess.get("key")
    try:
        import json
        from services.cad_import.design.bore import extract_dia_text_points
        from services.cad_import.pipeline import handoff, stage1 as s1
        spec = os.path.join(handoff.pick_out_dir(), f"{key}_찍은스펙.json")
        with open(spec, encoding="utf-8") as f:
            src = json.load(f).get("source_dxf")
        w = handoff.load_world(key, src, s1.World)
        if w is None:
            print(f"[설계] ★치수 텍스트 없음 — handoff 캐시를 쓸 수 없습니다"
                  f" (원본: {src}). 관경은 전부 별표1 로 정해집니다.")
            return []
        pts = extract_dia_text_points(w.texts)
        if not pts:
            print(f"[설계] ★도면 문자 {len(w.texts):,}개 중 치수로 읽힌 것이"
                  f" 0개입니다 — 관경은 전부 별표1 로 정해집니다.")
        else:
            print(f"[설계] 치수 텍스트 {len(pts):,}개"
                  f" (도면 문자 {len(w.texts):,}개 중)")
        return pts
    except Exception as exc:  # noqa: BLE001
        print(f"[설계] ★치수 텍스트를 읽지 못했습니다 — 관경은 별표1 로만: {exc}")
        return []


def _valve_label(tables):
    """알람밸브 노드의 표 라벨 — G 대화상자 `_valve_label` 과 같은 규칙."""
    for row in (getattr(tables, "equipment", None) or ()):
        if str(row.get("desc")) == "A/V":
            return row.get("in")
    return None


def _view_opts(cfg: dict) -> dict:
    """설정 7종 → display_tables/emit_design_sdf 인자. 두 곳이 같은 값을 쓴다."""
    return {
        "iso": bool(cfg["iso"]),
        "iso_z_scale": float(cfg["iso_z_scale"]),
        "canvas_units": float(cfg["canvas_units"]),
        "head_stub_ratio": float(cfg["head_stub_pct"]) / 100.0,
    }


def _summary(got: dict, tbl) -> dict:
    """build 요약 — G 대화상자 `_render` 가 보여주는 그 수치들."""
    meta = dict(tbl.meta)
    w = got.get("worst") or {}
    return {
        "k": len(w.get("heads") or []),
        "far_m": w.get("far_m"), "near_m": w.get("near_m"),
        "span_m": w.get("span_m"), "total_m": w.get("total_m"),
        "max_load": w.get("max_load"),
        "source": w.get("source_tag"),
        "counts": {"nodes": len(tbl.nodes), "pipes": len(tbl.pipes),
                   "nozzles": len(tbl.nozzles), "fittings": len(tbl.fittings),
                   "equipment": len(tbl.equipment)},
        "bore_src": {
            "text": meta.get("관경 근거 — 도면 텍스트"),
            "nfpc_min": meta.get("관경 근거 — 별표1 보강 (text<min)"),
            "nfpc_fallback": meta.get("관경 근거 — 별표1 폴백 (text 없음)"),
        },
        "fitting_unresolved": meta.get("부속 판정 불가"),
        "eq_len_unresolved": meta.get("등가길이 미해결"),
        "loops": meta.get("루프 잔여 배관(표 꼬리)"),
        "excluded_heads": got.get("excluded_heads", 0),
        "candidate_heads": got.get("candidate_heads", 0),
        "total_heads": got.get("total_heads", 0),
    }


def _classify_excluded(sess: dict, got: dict, board) -> dict:
    """[F-5] 빠진 헤드를 세 갈래로 가른다 — «제외 2,864» 를 숫자로 쪼갠다.

        찍히지 않음     A 후보인데 board 에 없는 것 (suggest 를 돌린 세션만)
        이음 끊김       board 물길은 닿는데 전개가 못 붙인 것 (B4 부착 실패)
        물길 미도달     board 물길 자체가 안 닿는 것 (이음 끊김의 상류)

    좌표를 함께 돌려준다 — 화면이 분류별로 켜고 끌 수 있어야 어디를 이어야
    하는지 보인다. 숫자만 주면 «크다» 만 알고 «어디» 를 모른다.
    """
    disks = getattr(board, "disks", []) or []
    total = len(disks)
    # board 물길 도달 — 급수원 성분에 붙은 헤드.
    try:
        wet_board = set((board.water_state() or {}).get("wet_heads") or [])
    except Exception as exc:  # noqa: BLE001
        print(f"[설계] 물길 분류 실패 — 미도달로 못 가른다: {exc}")
        wet_board = None
    # 전개가 붙일 수 있는 헤드 — 엔진의 공개 probe 를 그대로 쓴다.
    attach = None
    try:
        from services.cad_import.design.restrict import attachable_heads
        es = sess.get("edit")
        probe = attachable_heads(es.convert_payload())
        if probe.get("ok"):
            attach = set(probe.get("wet") or ())
    except Exception as exc:  # noqa: BLE001
        print(f"[설계] 부착 probe 실패 — 이음 끊김을 못 가른다: {exc}")

    def xy(i):
        d = disks[i]
        return [round(float(d[0]), 1), round(float(d[1]), 1)]

    out = {"total": total}
    if wet_board is not None:
        dry = [i for i in range(total) if i not in wet_board]
        out["dry"] = {"n": len(dry), "xy": [xy(i) for i in dry]}
        if attach is not None:
            unatt = [i for i in wet_board if i not in attach and i < total]
            out["unattached"] = {"n": len(unatt), "xy": [xy(i) for i in unatt]}
    # 찍히지 않음 — suggest 후보 중 어느 board 헤드와도 250mm 안에 없는 것.
    cands = sess.get("suggest")
    if cands:
        import math as _m
        centers = [(float(d[0]), float(d[1])) for d in disks]
        missing = []
        for c_ in cands:
            cx, cy = float(c_["x"]), float(c_["y"])
            if not any(_m.hypot(cx - px, cy - py) <= 250.0
                       for px, py in centers):
                missing.append([round(cx, 1), round(cy, 1)])
        out["unpicked"] = {"n": len(missing), "xy": missing}
    return out


def emit_design_files(sess: dict, UPLOAD_DIR, cfg: dict | None = None):
    """[F-4 공유] 세션의 확정된 표 → .sdf+.slf 한 쌍. (경로, 오류) 를 돌려준다.

    design/emit 라우트와 변환 단계의 «최불리 .sdf» 체크가 **같은 함수**를 탄다 —
    두 경로가 각자 emit 을 들고 있으면 설정 한쪽만 고쳐지는 날이 온다.
    """
    d = sess.get("design")
    if not d:
        return None, "먼저 수리계산 입력에서 표를 확정하세요."
    cfg = cfg or dict(sess.get("design_settings") or _DEFAULT_SETTINGS)

    from services.cad_import.design.emit import AssetMissing, emit_design_sdf
    key = sess.get("key") or "design"
    # ★파일 이름이 곧 SDF 안의 User-lib 참조다(G14 — 파일명만 담는다).
    #   G 데스크톱과 같은 이름이라야 산출이 같다. 세션끼리 안 섞이게 폴더를 가른다.
    out_dir = Path(UPLOAD_DIR) / "module_f" / f"{sess['id']}_design"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{key}_수리계산입력.sdf"
    tbl = d["tables"]
    try:
        out = emit_design_sdf(
            tbl, out_path,
            project_title=f"{key} 수리계산 입력",
            **_view_opts(cfg),
            iso_ref_label=(_valve_label(tbl)
                           if cfg["lift_ref"] == "valve" else None))
    except AssetMissing as exc:
        return None, str(exc)
    except Exception as exc:  # noqa: BLE001
        return None, f"{type(exc).__name__}: {exc}"
    sess["design_sdf_path"] = str(out)
    sess["design_slf_path"] = str(out.with_suffix(".slf"))
    return out, None


def schedule_bores_mm(name: str) -> set:
    """그 관종의 SLF 스케줄에 실제로 있는 호칭경(mm).

    ★「엘베」 교훈의 관경판이다. 자유 숫자를 받으면 저장은 되지만 SLF 에 그
      호칭경이 없어 PIPENET 이 그 배관을 못 푼다 — 문제를 뒤로 미룰 뿐이다.
      그래서 **그 자리에서 거절**한다. 목록은 엔진의 `SCHEDULE_DEFS` 한 곳에서
      읽는다(값은 m 단위로 적혀 있다 — 0.065 == 65A).
    """
    _boot()
    from services.cad_import.design.sdf_post import SCHEDULE_DEFS
    for nm, _c, sizes in SCHEDULE_DEFS:
        if nm == str(name or "").strip():
            return {int(round(float(d) * 1000)) for d, _sch in sizes}
    return set()


def _bore_ov_map(sess) -> dict:
    """세션의 관경 덮기 → `decide_bores(overrides=)` 가 받는 모양.

    저장은 «사람이 읽는 모양»(a·b·dia·note 목록)으로 하고, 엔진에는 그 엔진이
    받는 모양으로 번역해 넘긴다. 세션에 엔진 모양을 그대로 담아 두면 나중에
    엔진 시그니처가 바뀔 때 옛 세션이 조용히 깨진다.
    """
    out = {}
    for r in (sess.get("bore_overrides") or ()):
        try:
            a, b = int(r["a"]), int(r["b"])
        except (KeyError, TypeError, ValueError):
            continue
        out[(min(a, b), max(a, b))] = (int(r["dia"]), str(r.get("note") or ""))
    return out


# ── [F-11d-1] 직접 입력의 «안정 키» ─────────────────────────────────
#
# 부속 직접 입력의 키 `(node, pipe)` 는 둘 다 제한 전개(corridor)에서 새로
# 매겨지는 이름이다. corridor 가 바뀌면 BFS 가 다시 돌아 번호가 재배열된다.
#
# 실측(BLOCKED §22 · 대명동, 기준개수 K 30→20):
#     배관 라벨   공통 자리 105개 중 **105개가 옮겨감**
#     (node,pipe) 키  4 → 2 · **그대로 0**       ← 하나도 안 남는다
#     board 노드쌍 키 4 → 2 · **그대로 2**       ← 살아남은 자리를 다 지킨다
#
# 그래서 **세션에는 안정 키로 담고, 엔진에 넘길 때 번역한다.**
# `build_fittings(overrides=)` 는 `(node, pipe)` 를 받고 그 파일은 이 지시서에서
# 읽기 전용이므로(§3) 엔진을 바꾸지 않는다.
#
# ★이 번역은 «판정» 이 아니라 «이름 바꾸기» 다. 어느 자리가 미해결인지는 엔진이
#   정하고(`build_fittings`), 여기서는 그 자리를 corridor 가 바뀌어도 가리킬 수
#   있는 이름으로 옮겨 적을 뿐이다. F 가 판정을 다시 하면 규칙이 두 벌이 된다.

_SPOT_MM = 1          # 좌표 반올림 자리 — 같은 board 점이면 같은 값이 나온다


def _kfp_ends(net, pid):
    """배관의 두 끝 노드. kfp 는 start/end 와 from/to 두 이름을 다 쓴다."""
    pr = ((net or {}).get("pipe_data") or {}).get(str(pid)) or {}
    return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")


def _node_xyz(net, nid):
    c = (((net or {}).get("nodes_meta_runtime") or {}).get(nid) or {}
         ).get("coords") or None
    if not c:
        return None
    return tuple(round(float(v), _SPOT_MM) for v in (list(c) + [0.0, 0.0])[:3])


def spot_key(got, node, pipe):
    """`(kfp node, kfp pipe)` → 안정 키 `(board 노드쌍, 그 자리 좌표)`.

    배관은 board 간선으로, 자리는 좌표로 가리킨다 — 둘 다 corridor 가 바뀌어도
    같은 값이다. 역참조가 없는 배관(헤드 접속관·가지 상승)은 board 간선이 없어
    안정 키를 못 만든다 — None 을 돌려주고 부르는 쪽이 «구키로 남긴다».
    """
    net = (got or {}).get("kfp") or {}
    ref = ((got or {}).get("edge_ref") or {}).get(str(pipe))
    if ref is None:
        return None
    xyz = _node_xyz(net, node)
    if xyz is None:
        return None
    try:
        a, b = int(ref[0]), int(ref[1])
    except (TypeError, ValueError, IndexError):
        return None
    return (min(a, b), max(a, b)), xyz


def spot_index(got) -> dict:
    """이번 빌드의 `안정 키 → (node, pipe)` — 번역의 반대 방향.

    배관마다 두 끝을 다 넣는다. 어느 끝인지는 좌표가 가른다.
    """
    net = (got or {}).get("kfp") or {}
    out = {}
    for pid, ref in ((got or {}).get("edge_ref") or {}).items():
        try:
            a, b = int(ref[0]), int(ref[1])
        except (TypeError, ValueError, IndexError):
            continue
        pair = (min(a, b), max(a, b))
        for nid in _kfp_ends(net, pid):
            if nid is None:
                continue
            xyz = _node_xyz(net, nid)
            if xyz is not None:
                out[(pair, xyz)] = (str(nid), str(pid))
    return out


def fitting_ov_for_engine(sess, got) -> tuple:
    """세션의 직접 입력 → 엔진이 받는 모양 + **적용 못 한 것**.

    ★못 적용한 것을 조용히 버리지 않는다(지시서 F-11d-2 「조용한 소실 금지」).
      그 자리가 corridor 에서 빠졌거나, 미해결이 아니게 됐거나, 안정 키를 못
      만드는 배관이면 값이 안 들어간다 — 사람은 그 사실을 알아야 한다.
    """
    cur = dict(sess.get("fitting_overrides") or {})
    idx = spot_index(got)
    out_kind, missed = [], []
    for r in (cur.get("kind") or ()):
        node, pipe = r.get("node"), r.get("pipe")
        key = _row_key(r)
        if key is not None and key in idx:
            node, pipe = idx[key]           # 안정 키가 이번 빌드의 이름을 준다
        elif key is not None:
            missed.append({**r, "why": "그 자리가 이번 계산 범위에 없습니다"})
            continue
        # 안정 키가 없는 구(舊) 항목은 적어 둔 이름 그대로 시도한다(읽기 호환).
        out_kind.append({"node": node, "pipe": pipe,
                         "kind": r.get("kind"), "note": r.get("note")})
    return ({"kind": out_kind, "eq_len": list(cur.get("eq_len") or ())},
            missed)


def _row_key(r):
    """저장된 한 줄에서 안정 키를 꺼낸다. 구(舊) 항목이면 None."""
    try:
        a, b = int(r["a"]), int(r["b"])
        xyz = tuple(round(float(r[k]), _SPOT_MM) for k in ("nx", "ny", "nz"))
    except (KeyError, TypeError, ValueError):
        return None
    return (min(a, b), max(a, b)), xyz


def register(app, *, UPLOAD_DIR):
    # ─────────────────────────────────────── 수리계산 입력 (설계)
    @app.post("/api/module-f/design/build")
    @route_session(post=True)
    def module_f_design_build(sess, body):
        """최불리 선정 → corridor 제한 전개 → 5표. 파일은 쓰지 않는다."""
        es = sess.get("edit")
        if es is None:
            # 자동(A) 경로는 표가 이미 나와 있다 — 여기서 다시 만들 것이 없다.
            # 화면은 이 단추를 감추지만, 옛 클라이언트가 부를 수 있으므로
            # «손질 세션이 없다» 대신 무엇이 잘못됐는지 말한다.
            if sess.get("method") == "auto":
                return _fail("자동 경로는 «자동 추출» 이 이미 표를 냈습니다 — "
                             "여기서 다시 확정하지 않습니다.", 409)
            return _fail("손질 세션이 없습니다.")
        if not getattr(es.board, "sources", None):
            return _fail("급수 시작 위치를 먼저 찍어야 설계면적을 고를 수 있습니다.")
        undecided = sum(1 for kk in getattr(es.board, "disk_kinds", []) or []
                        if kk == "미지정")
        if undecided:
            return _fail(f"헤드 종류가 미지정인 것이 {undecided}개 있습니다. "
                         "손질에서 종류를 정한 뒤 다시 시도하세요.")
        if _job_running(sess):
            return _fail("이미 작업이 돌고 있습니다. 끝난 뒤에 다시 눌러 주세요.",
                         409)
        try:
            _boot()
            cfg = _settings(sess, body)
        except ValueError as exc:
            return _fail(str(exc))
        source = body.get("source")

        def job():
            from services.cad_import.design.restrict import select_and_expand
            from services.cad_import.design.tables import build_design_tables
            from services.cad_import.design.sdf_post import UnknownSchedule

            payload = es.convert_payload()
            srcs = payload.get("sources") or ()
            sel = source if source is not None else (
                srcs[0].get("tag") if len(srcs) > 1 and isinstance(srcs[0], dict)
                else None)
            got = select_and_expand(payload, es.board, k=cfg["k"],
                                    selected_source=sel)
            if not got.get("ok"):
                return {"ok": False, "error": got.get("error")}
            texts = _dia_texts(sess)
            # [F-11d] 직접 입력을 **이번 계산의 이름으로 번역**한다. 세션에는
            #   corridor 가 바뀌어도 같은 자리를 가리키는 안정 키로 담겨 있다
            #   (BLOCKED §22). 못 옮긴 것은 버리지 않고 세어 화면에 올린다.
            fit_ov, fit_missed = fitting_ov_for_engine(sess, got)
            try:
                tbl = build_design_tables(
                    got["kfp"], got["worst"], got["edge_ref"], texts,
                    board_pts=es.board.pts,
                    excluded_heads=got.get("excluded_heads", 0),
                    default_schedule=cfg["schedule"],
                    tree_loads=got.get("tree_loads"),
                    # [§18] 사람이 손으로 채운 값 — 규칙이 못 가린 자리에만
                    #   쓰인다(엔진이 그렇게 판단한다). 세션에 남아 있으므로
                    #   「다시 계산」·「표 확정」을 다시 눌러도 그대로 간다.
                    fitting_overrides=fit_ov,
                    # [F-11c] 관경 덮기 — 부속과 달리 «규칙 값도» 덮는다
                    #   (D-F11-3). 키는 board 노드쌍이라 corridor 가 다시
                    #   계산돼도 같은 자리를 가리킨다(D-F11-4).
                    bore_overrides=_bore_ov_map(sess))
            except UnknownSchedule as exc:
                return {"ok": False, "error": str(exc)}
            # ★[F-11d-2] 넘긴 것 중 «엔진이 실제로 쓴 것» 을 맞대 본다.
            #   자리가 corridor 에 남아 있어도 그 사이에 미해결이 아니게 됐으면
            #   값은 안 들어간다 — 그것도 «적용 못 한 수정» 이다. 개수만 세면
            #   사람은 들어간 줄 안다.
            used = {(str(a.get("node")), str(a.get("pipe")))
                    for a in ((getattr(tbl, "unresolved", None) or {})
                              .get("applied") or ())
                    if a.get("what") == "kind"}
            sent = {(str(r.get("node")), str(r.get("pipe"))): r
                    for r in (fit_ov.get("kind") or ())}
            for k2, r in sent.items():
                if k2 not in used:
                    fit_missed.append(
                        {**r, "why": "그 자리는 이제 «판정 불가» 가 아닙니다 "
                                     "— 자동이 답을 냈습니다"})
            # 등가길이 쌍도 같은 방식으로 맞댄다((종류,호칭경) 이 단위다).
            used_eq = {(str(a.get("kind")), int(a.get("dia")))
                       for a in ((getattr(tbl, "unresolved", None) or {})
                                 .get("applied") or ())
                       if a.get("what") == "eq_len" and a.get("dia") is not None}
            for r in (fit_ov.get("eq_len") or ()):
                try:
                    kk = (str(r.get("kind")), int(r.get("dia")))
                except (TypeError, ValueError):
                    continue
                if kk not in used_eq:
                    fit_missed.append(
                        {**r, "what": "eq_len",
                         "why": "그 (종류, 호칭경) 쌍이 이번 계산에 없습니다"})
            if fit_missed:
                print(f"[설계] ★적용 못 한 직접 입력 {len(fit_missed)}건 — "
                      "조용히 버리지 않고 화면에 올린다")
            sess["ov_missed"] = fit_missed
            # 표는 메모리에만 — emit 을 눌러야 파일이 생긴다.
            marks = _classify_excluded(sess, got, es.board)
            sess["design"] = {"got": got, "tables": tbl, "k": cfg["k"],
                              "schedule": cfg["schedule"], "marks": marks}
            s = _summary(got, tbl)
            s["excluded_detail"] = {
                k2: v2["n"] for k2, v2 in marks.items()
                if isinstance(v2, dict)}
            det = s["excluded_detail"]
            print("[설계] 제외 사유 — "
                  + " · ".join(f"{lab} {det[k2]:,}" for k2, lab in
                               (("dry", "물길 미도달"),
                                ("unattached", "이음 끊김"),
                                ("unpicked", "찍히지 않음")) if k2 in det))
            print(f"[설계] 표 확정 · 헤드 {s['k']} · 앵커 {s['far_m']} m · "
                  f"배관 {s['counts']['pipes']} · 제외 {s['excluded_heads']:,}")
            return {"ok": True, "summary": s}

        _run_job(sess, "수리계산 입력", job)
        return jsonify({"ok": True})

    @app.post("/api/module-f/design/fitting-override")
    @route_session(post=True)
    def module_f_design_fitting_override(sess, body):
        """[§18] 규칙이 못 가린 자리를 사람이 채운다 — 「직접 입력」.

        모듈 A 의 방식을 그대로 따른다(`override_flag`·`override_note`):
        ⑴ 값을 덮고, ⑵ **덮었다는 사실과 사유를 함께 남기고**, ⑶ 화면과
        산출물이 「직접 입력」이라고 밝힌다. 자동이 낸 값과 사람이 넣은 값을
        같은 얼굴로 두지 않는 것이 이 방식의 요점이다.

        ★단위가 둘인 이유는 문제의 성격이 다르기 때문이다:
          · 부속 판정 — **자리(노드·배관)** 단위. 기하가 자리마다 달라 못 묶는다.
          · 등가길이 — **(종류, 호칭경) 쌍** 단위. 라이브러리 구멍이라 한 번
            채우면 그 쌍을 쓰는 배관이 한꺼번에 풀린다.

        ★여기서 «판정» 을 하지 않는다. 어느 자리가 미해결인지는 엔진이 정하고
          (`build_fittings`), 이 값은 그 자리에만 쓰인다. F 가 다시 판정하면
          규칙이 두 벌이 되어 언젠가 갈린다.

        body: {sid, kind: [{node, pipe, kind, note}], eq_len: [{kind, dia, m, note}]}
              칸을 안 보내면 그 갈래는 그대로 둔다. 빈 배열을 보내면 지운다.
        표를 다시 확정해야 값이 산출에 들어간다 — 그 사실을 응답에 실어 준다.
        """
        cur = dict(sess.get("fitting_overrides") or {})
        for key, need in (("kind", ("node", "pipe", "kind")),
                          ("eq_len", ("kind", "dia", "m"))):
            if key not in body:
                continue
            rows = body.get(key)
            if rows is None:
                rows = []
            if not isinstance(rows, list) or len(rows) > 500:
                return _fail(f"{key} 목록이 올바르지 않습니다 (최대 500).")
            clean = []
            for r in rows:
                if not isinstance(r, dict):
                    return _fail(f"{key} 항목은 객체여야 합니다.")
                if any(str(r.get(n) or "").strip() == "" for n in need):
                    return _fail(f"{key} 항목에 {', '.join(need)} 가 다 있어야 합니다.")
                row = {n: r.get(n) for n in need}
                if key == "eq_len":
                    try:
                        row["dia"] = int(r["dia"])
                        row["m"] = float(r["m"])
                    except (TypeError, ValueError):
                        return _fail("호칭경은 정수, 등가길이는 숫자여야 합니다.")
                    if row["m"] < 0:
                        return _fail("등가길이는 음수일 수 없습니다.")
                note = str(r.get("note") or "").strip()
                if len(note) > 200:
                    return _fail("사유가 너무 깁니다 (200자).")
                row["note"] = note
                if key == "kind":
                    # [F-11d-1] 받은 자리를 «안정 키» 로도 적어 둔다. 화면은
                    #   지금처럼 (node, pipe) 로 가리키면 되고, 그 이름이 다음
                    #   계산에서 다른 자리를 뜻하게 되는 것을 여기서 막는다.
                    #   이미 안정 키를 실어 보냈으면 그것을 그대로 믿는다.
                    st = _row_key(r)
                    if st is None:
                        # ★`sess["design"]` 이 아니라 그 안의 `got` 이다 —
                        #   `kfp`·`edge_ref` 는 거기 있다. 한 겹 위를 넘기면
                        #   조용히 None 이 되어 «안정 키가 안 붙은 채» 저장된다.
                        st = spot_key((sess.get("design") or {}).get("got"),
                                      row.get("node"), row.get("pipe"))
                    if st is not None:
                        (row["a"], row["b"]), xyz = st
                        row["nx"], row["ny"], row["nz"] = xyz
                clean.append(row)
            cur[key] = clean
        sess["fitting_overrides"] = cur
        n_k, n_e = len(cur.get("kind") or ()), len(cur.get("eq_len") or ())
        n_st = sum(1 for r in (cur.get("kind") or ()) if _row_key(r))
        if n_k:
            print(f"[설계] 직접 입력 부속 {n_k}자리 중 안정 키 {n_st}자리 "
                  f"(나머지는 board 역참조가 없어 구키로 남는다)")
        print(f"[설계] 직접 입력 — 부속 {n_k}자리 · 등가길이 {n_e}쌍 "
              f"(표를 다시 확정해야 산출에 들어간다)")
        return jsonify({
            "ok": True, "overrides": cur,
            "counts": {"kind": n_k, "eq_len": n_e},
            # 표시가 아니라 «값» 이 바뀌는 일이다 — 다시 확정하라고 분명히 말한다.
            "needs_rebuild": bool(sess.get("design")),
            "message": ("직접 입력을 저장했습니다 — "
                        "「표 확정」을 다시 눌러야 산출에 반영됩니다."),
        })

    @app.get("/api/module-f/design/fitting-override")
    @route_session()
    def module_f_design_fitting_override_get(sess, body):
        """지금 저장된 직접 입력 + **고를 수 있는 부속 종류**.

        ★종류를 자유 입력으로 두면 안 된다. 라이브러리에 없는 이름을 넣으면
          그 부속의 등가길이가 다시 «미해결» 이 된다(실측: 「엘베」라고 적었더니
          부속 판정 불가 3→2 로 줄면서 등가길이 미해결이 0→1 로 늘었다).
          그래서 엔진이 아는 이름만 내려보내고 화면은 그중에서 고르게 한다.

        「직선 — 부속 없음」도 정답의 하나다. 22.5° 미만은 45° 엘보보다 직선에
        가깝지만 collinear merge 가 흡수를 거부한 각이라 프로그램이 단정할 수
        없어 판정 불가로 셌다 — 사람은 도면을 보고 단정할 수 있다.
        """
        kinds = [{"value": "none", "label": "직선 — 부속 없음"}]
        try:
            import sys as _s
            core = str(Path(__file__).resolve().parents[2] / "core")
            if core not in _s.path:
                _s.path.append(core)
            import fitting_rules as fr
            kinds += [
                {"value": fr.ELBOW_45, "label": "45° 엘보"},
                {"value": fr.ELBOW_90, "label": "90° 엘보"},
                {"value": fr.TEE, "label": "티"},
            ]
        except Exception as exc:  # noqa: BLE001 — 목록을 못 읽어도 조회는 된다
            print(f"[설계] 부속 종류 목록을 못 읽었습니다: {exc}")
        cur = dict(sess.get("fitting_overrides") or {})
        applied = ((sess.get("design") or {}).get("tables"))
        applied = (getattr(applied, "unresolved", None) or {}).get("applied") or []
        return jsonify({"ok": True, "overrides": cur, "applied": applied,
                        "kinds": kinds})

    @app.post("/api/module-f/design/bore-override")
    @route_session(post=True)
    def module_f_design_bore_override(sess, body):
        """[F-11c · D-F11-3] 관경 «직접 입력» — 규칙 값도 덮는다.

        부속·등가길이(§18)와 문법은 같지만 **범위가 다르다**. 저 둘은 「규칙이
        못 가린 자리에만」 쓰는데, 관경은 규칙이 낸 값도 덮는다 — 도면 치수가
        틀렸거나 설계 협의로 바뀌는 일이 실제로 있기 때문이다. 대신 덮은 자리는
        **원값·원출처를 반드시 남긴다**(`tables.bore_overrides`).

        ★키는 «정렬된 board 노드쌍» 이다(D-F11-4). 배관 라벨(P12)은 BFS 순서로
          매겨지므로 corridor 가 바뀌면 같은 이름이 다른 배관을 가리킨다 —
          사람이 65A 라고 적어 둔 자리가 조용히 옆 배관으로 옮겨간다.

        body: {sid, rows: [{a, b, dia, note}]}  · 빈 배열을 보내면 지운다.
        """
        rows = body.get("rows")
        if rows is None:
            rows = []
        if not isinstance(rows, list) or len(rows) > 500:
            return _fail("덮기 목록이 올바르지 않습니다 (최대 500).")
        sched = (sess.get("design_settings") or _DEFAULT_SETTINGS).get(
            "schedule") or (sess.get("design") or {}).get("schedule")
        try:
            allow = schedule_bores_mm(sched)
        except Exception as exc:  # noqa: BLE001 — 부팅 실패도 사유로 말한다
            return _fail(f"규격표를 못 읽었습니다: {exc}")
        clean, seen = [], set()
        for r in rows:
            if not isinstance(r, dict):
                return _fail("덮기 항목은 객체여야 합니다.")
            try:
                a, b, dia = int(r["a"]), int(r["b"]), int(r["dia"])
            except (KeyError, TypeError, ValueError):
                return _fail("덮기 항목에 a, b, dia 가 다 있어야 합니다 "
                             "(a·b 는 노드 번호, dia 는 호칭경 mm).")
            if allow and dia not in allow:
                # ★여기서 거절해야 한다. 저장해 두면 SLF 에 그 호칭경이 없어
                #   PIPENET 이 그 배관을 못 푼다 — 문제를 뒤로 미룰 뿐이다.
                return _fail(
                    f"{dia}A 는 «{sched}» 규격표에 없는 호칭경입니다. "
                    f"쓸 수 있는 것: {' · '.join(str(v) for v in sorted(allow))}")
            note = str(r.get("note") or "").strip()
            if len(note) > 200:
                return _fail("사유가 너무 깁니다 (200자).")
            key = (min(a, b), max(a, b))
            if key in seen:
                return _fail(f"같은 배관({key[0]}–{key[1]})을 두 번 덮었습니다.")
            seen.add(key)
            clean.append({"a": key[0], "b": key[1], "dia": dia, "note": note})
        sess["bore_overrides"] = clean
        print(f"[설계] 관경 직접 입력 — {len(clean)}개 "
              f"(표를 다시 확정해야 산출에 들어간다)")
        return jsonify({
            "ok": True, "rows": clean, "counts": {"bore": len(clean)},
            "schedule": sched, "allowed": sorted(allow),
            "needs_rebuild": bool(sess.get("design")),
            "message": ("관경 직접 입력을 저장했습니다 — "
                        "「표 확정」을 다시 눌러야 산출에 반영됩니다."),
        })

    @app.get("/api/module-f/design/bore-override")
    @route_session()
    def module_f_design_bore_override_get(sess, body):
        """지금 저장된 관경 덮기 + **고를 수 있는 호칭경**.

        고를 수 있는 값을 서버가 주는 이유는 §18 의 부속 종류와 같다 — 화면이
        따로 목록을 들고 있으면 규격표가 바뀔 때 둘이 갈린다.
        """
        sched = (sess.get("design_settings") or _DEFAULT_SETTINGS).get(
            "schedule") or (sess.get("design") or {}).get("schedule")
        try:
            allow = sorted(schedule_bores_mm(sched))
        except Exception as exc:  # noqa: BLE001
            print(f"[설계] 규격표를 못 읽었습니다: {exc}")
            allow = []
        return jsonify({"ok": True, "rows": sess.get("bore_overrides") or [],
                        "schedule": sched, "allowed": allow})

    @app.get("/api/module-f/design/preview")
    @route_session()
    def module_f_design_preview(sess, body):
        """emit 에 넘길 좌표 **그대로** — 표시 전용 좌표계를 따로 두지 않는다.

        보기 설정만 바뀌면 build 를 다시 돌지 않는다 — 캐시한 표에 표시 변환만
        다시 얹는다(G16 의 «최불리 재계산 없이 다시 그리기» 그대로).
        """
        d = sess.get("design")
        if not d:
            job = sess.get("job") or {}
            if job.get("phase") == "수리계산 입력" and job.get("state") == "run":
                return _fail("아직 계산 중입니다.", 409)
            # ★«아직 확정 안 함» 은 오류가 아니다. 404 로 답하면 사람이 수리계산
            #   화면에 들어올 때마다 브라우저 콘솔에 붉은 줄이 남고, 진짜 오류가
            #   그 사이에 묻힌다. 같은 이유로 `auto/network-view` 는 이미
            #   «없음» 을 200 으로 답한다 — 그 규약을 여기에도 맞춘다.
            return jsonify({"ok": True, "view": None, "tables": None,
                            "marks": {}, "stood": None,
                            "settings": dict(sess.get("design_settings")
                                             or _DEFAULT_SETTINGS),
                            "message": "먼저 「표 확정」을 눌러 주세요."})
        # 보기 설정 — 쿼리로 덮을 수 있고, 덮은 값은 기억된다.
        # (iso 는 문자열 "false" 를 bool() 로 캐스팅하면 True 가 되므로 따로 푼다.)
        try:
            cfg = _settings(sess, {k: request.args.get(k)
                                   for k in ("iso_z_scale", "canvas_units",
                                             "lift_ref", "head_stub_pct")
                                   if request.args.get(k) is not None})
        except ValueError as exc:
            return _fail(str(exc))
        if "iso" in request.args:
            cfg["iso"] = request.args.get("iso") in ("1", "true", "True", "on")
            sess["design_settings"] = cfg

        from services.cad_import.design.emit import display_tables
        tbl = d["tables"]
        view, stood = display_tables(
            tbl, **_view_opts(cfg),
            iso_ref_label=(_valve_label(tbl)
                           if cfg["lift_ref"] == "valve" else None))

        # 담당 헤드 수 — 간선 굵기의 근거. 도면 간선은 worst.loads, 세로
        # 구간(역참조 없음)은 tree_loads 가 안다.
        got = d["got"]
        loads = ((got.get("worst") or {}).get("loads")) or {}
        ref = got.get("edge_ref") or {}
        tree = got.get("tree_loads") or {}
        load_of = {}
        for pid, edge in ref.items():
            try:
                i, j = int(edge[0]), int(edge[1])
                load_of[str(pid)] = int(loads.get((min(i, j), max(i, j)), 0))
            except (TypeError, ValueError, IndexError):
                continue
        for pid, n in tree.items():
            load_of.setdefault(str(pid), int(n))

        at = {str(n.get("label")): n for n in view.nodes}
        elev = {lab: float(n.get("elevation", 0) or 0) for lab, n in at.items()}
        parent_of = {}
        for row in view.pipes:
            a, b = str(row.get("in")), str(row.get("out"))
            parent_of.setdefault(b, a)
            parent_of.setdefault(a, b)
        heads = {str(r.get("in")) for r in view.nozzles}
        av = _valve_label(tbl)

        # ── 최원 유하거리 «경로» — 급수원 → 앵커 (표 라벨 기준)
        #
        # 손질 단계에서는 board 절점으로 그렸지만 여기는 설계 표의 좌표계다.
        # 표에 이미 «앵커 노드» 가 meta 로 적혀 있으니 그것에서 급수원까지
        # 되짚는다. `parent_of` 는 양방향 첫이웃이라 트리가 아니다 — 여기서
        # 급수원 기점 BFS 로 제대로 된 부모를 만든다(안 그러면 되짚다 맴돈다).
        anchor_lab = dict(tbl.meta).get("앵커 노드")
        anchor_lab = str(anchor_lab) if anchor_lab not in (None, "?") else None
        root = next((lab for lab, n in at.items()
                     if str(n.get("io_node")) == "Input"), None)
        anchor_path: list[str] = []
        anchor_path_m = 0.0
        if anchor_lab and root:
            adj: dict[str, list[tuple[str, str]]] = {}
            plen: dict[str, float] = {}
            for row in view.pipes:
                a, b2 = str(row.get("in")), str(row.get("out"))
                pid = str(row.get("label"))
                try:
                    plen[pid] = float(row.get("length") or 0.0)
                except (TypeError, ValueError):
                    plen[pid] = 0.0
                adj.setdefault(a, []).append((b2, pid))
                adj.setdefault(b2, []).append((a, pid))
            par: dict[str, tuple[str, str]] = {}
            seen = {root}
            queue = [root]
            while queue:
                cur = queue.pop(0)
                for nxt, pid in adj.get(cur, ()):
                    if nxt in seen:
                        continue
                    seen.add(nxt)
                    par[nxt] = (cur, pid)
                    queue.append(nxt)
            if anchor_lab in seen:
                cur = anchor_lab
                while cur != root:
                    up, pid = par[cur]
                    anchor_path.append(cur)
                    anchor_path_m += plen.get(pid, 0.0)
                    cur = up
                anchor_path.append(root)
                anchor_path.reverse()

        nodes = []
        for lab, n in at.items():
            # ★반올림하지 않는다 — writer 는 좌표를 `.6g`(유효 6자리)로 찍는데
            #   소수 3자리 반올림은 그보다 거칠어(실측: -75.9731 → -75.973)
            #   「preview == 저장 Position」 이 깨진다. 같은 double 을 그대로
            #   보내면 양쪽을 `.6g` 로 찍었을 때 정확히 같은 문자열이 된다.
            rec = {"label": lab, "x": float(n.get("x", 0)),
                   "y": float(n.get("y", 0))}
            if lab in heads:
                rec["head"] = True
                rec["up"] = (elev.get(lab, 0.0)
                             - elev.get(parent_of.get(lab, ""), 0.0)) >= 0
            if str(n.get("io_node")) == "Input":
                rec["input"] = True
            if av is not None and lab == str(av):
                rec["valve"] = True
            if anchor_lab and lab == anchor_lab:
                rec["anchor"] = True      # 기준압을 잡는 지점
            nodes.append(rec)
        # [F-11c · D-F11-4] 관경 덮기의 «키» — board 노드쌍. 표 라벨(P12·노드 3)은
        #   BFS 순서로 매겨져 corridor 가 바뀌면 다른 자리를 가리키므로, 화면이
        #   덮을 자리를 가리킬 때는 이 쌍을 쓴다. 역참조가 없는 배관(헤드
        #   접속관·가지 상승)은 board 간선이 없어 null 이다 — 못 덮는다.
        ref_of = {}
        for pid, edge in ref.items():
            try:
                i, j = int(edge[0]), int(edge[1])
            except (TypeError, ValueError, IndexError):
                continue
            ref_of[str(pid)] = [min(i, j), max(i, j)]
        pipes = [{"label": str(r.get("label")),
                  "a": str(r.get("in")), "b": str(r.get("out")),
                  "dia": r.get("dia"), "len_m": r.get("length"),
                  "src": r.get("dia_src"),
                  "ref": ref_of.get(str(r.get("label"))),
                  "load": load_of.get(str(r.get("label")), 0)}
                 for r in view.pipes]
        return jsonify({
            "ok": True, "settings": cfg,
            "stood": stood,
            "view": {"nodes": nodes, "pipes": pipes,
                     # 최원 유하거리 경로 — far_m 이 «어느 줄» 인지.
                     "anchor": anchor_lab,
                     "anchor_path": anchor_path,
                     "anchor_path_m": round(anchor_path_m, 2)},
            "tables": tbl.as_dict(),        # 저장될 값 그대로 (F-3 표 4종)
            # [F-5] 제외 사유 분류 — mm 세계좌표. 설계 캔버스(정규화 좌표)가
            # 아니라 손질 망 위에 그려야 «어디» 인지 보인다.
            "marks": d.get("marks") or {},
            # [F-11d-2] 직접 입력 중 «이번 계산에 못 들어간 것». 조용한 소실
            #   금지 — 개수만 세면 사람은 들어간 줄 안다. 사유를 함께 싣는다.
            "ov_missed": sess.get("ov_missed") or [],
        })

    @app.post("/api/module-f/design/emit")
    @route_session(post=True)
    def module_f_design_emit(sess, body):
        """.sdf + .slf 한 쌍을 쓴다. 자산이 없으면 실패(G 정책 그대로)."""
        # ★「표 확정」 잡이 도는 동안 저장하면 — 새 표가 나오기 직전의 «옛 표» 로
        #   파일이 써진다. 사용자는 방금 누른 확정이 반영됐다고 읽는다.
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 저장할 수 있습니다.", 409)
        d = sess.get("design")
        if not d:
            return _fail("먼저 design/build 로 표를 확정하세요.", 404)
        try:
            cfg = _settings(sess, body)
        except ValueError as exc:
            return _fail(str(exc))

        out, err = emit_design_files(sess, UPLOAD_DIR, cfg)
        if err:
            return _fail(err, 500)
        slf = out.with_suffix(".slf")
        return jsonify({
            "ok": True,
            "sdf": {"name": out.name, "bytes": out.stat().st_size},
            "slf": {"name": slf.name, "bytes": slf.stat().st_size},
            "download": f"/api/module-f/download?sid={sess['id']}&what=design",
        })
