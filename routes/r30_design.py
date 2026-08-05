# -*- coding: utf-8 -*-
"""모듈 C 설계 자동화 API — 지시서 §11.

`register(app, ...)` 패턴. 기존 `/api/remote30/*` 는 건드리지 않으며, 이 모듈은
`DESIGN_WORKBENCH_ENABLED` 가 꺼져 있으면 라우트를 아예 등록하지 않는다(§C.5).

C2 이후 엔드포인트는 **게이트 강제만** 걸어 자리를 잡아 두고, 게이트를 통과한
뒤에는 아직 구현되지 않았음을 501 로 밝힌다. 통과한 것처럼 200 을 돌려주면 화면이
진행된 것으로 오해한다.
"""
from __future__ import annotations

import contextlib
import gzip
import hashlib
import json
import os
import uuid
from pathlib import Path

from flask import Response, jsonify, make_response, render_template, request
from werkzeug.utils import secure_filename

from core.design import gate as G
from core.design import session as S
from core.design.recognize import params as RP
from core.design.recognize import pipeline as PL
from core.design.schema import BuildingDraft

# [문서정합] §11.2 는 이 상수를 쓰라고만 하고 둘 곳을 지정하지 않았다. 앱 엔트리
# (`대조 서버.py`)는 §2 가 register 호출 한 줄만 허용하므로 여기 둔다. C1 사슬의
# 판정이 바뀌면 이 값을 올려야 옛 결과가 캐시에서 되살아나지 않는다.
DESIGN_CACHE_VERSION = "c1-v1"

# 게이트 뒤에 서는 단계들. 지시서 §4.1 — `/api/design/c2/*` 이후 전부.
_GATED_STAGES = (
    ("/api/design/c2/constraints", "POST", "C2"),
    ("/api/design/c2b/esfr", "POST", "C2B"),
    ("/api/design/c3/valves", "POST", "C3"),
    ("/api/design/c3/zones", "POST", "C3-zones"),
    ("/api/design/c4/heads", "POST", "C4"),
    ("/api/design/c5/route", "POST", "C5"),
    ("/api/design/emit", "POST", "EMIT"),
    ("/api/design/checks/<sid>", "GET", "CHECKS"),
)


def _fail(code: str, message: str, status: int, **extra):
    return jsonify({"ok": False, "code": code, "message": message, **extra}), status


def _content_hash(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fp:
        for block in iter(lambda: fp.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def register(app, *, DESIGN_SESSION_DIR, UPLOAD_DIR=None, INSPECT_CACHE_DIR=None,
             INSPECT_CACHE_VERSION: str = "", enabled: bool = False):
    """[문서정합] §11 의 서명은 `(app, *, UPLOAD_DIR, _save_upload, DESIGN_SESSION_DIR)`.

    `_save_upload` 는 받지 않는다. C1 은 화면이 이미 그린 도면을 다시 보는 단계라
    업로드는 `/api/remote30/upload` 가 이미 끝냈고, 여기서 또 받으면 같은 도면이
    두 벌 쌓인다. 대신 inspect 렌더 캐시를 읽어야 해서 그 위치와 버전을 받는다.
    """
    if not enabled:
        return

    DESIGN_SESSION_DIR.mkdir(parents=True, exist_ok=True)
    DESIGN_CACHE_DIR = (UPLOAD_DIR / "_design_cache") if UPLOAD_DIR else None
    if DESIGN_CACHE_DIR is not None:
        DESIGN_CACHE_DIR.mkdir(parents=True, exist_ok=True)

    def _open(sid: str):
        """반환 `(세션, None)` 또는 `(None, 오류응답)`."""
        try:
            return S.DesignSession.open(DESIGN_SESSION_DIR, sid), None
        except S.SessionNotFound:
            return None, _fail("SESSION_NOT_FOUND", "세션을 찾을 수 없습니다.", 404)

    def _load_draft(sess):
        """반환 `(draft, version, None)` 또는 `(None, 0, 오류응답)`."""
        try:
            raw, version = sess.read("building.json")
        except S.ArtifactNotFound:
            return None, 0, _fail(
                "C1_NOT_DONE", "C1 인식이 끝나지 않아 확정할 대상이 없습니다.", 409)
        return BuildingDraft.from_dict(raw), version, None

    # ── 페이지 ──────────────────────────────────────────────────────────
    # [문서정합] §11 은 이 라우트를 `routes/pages.py` 에 두라고 한다. 여기 둔 이유는
    # 플래그가 하나이기 때문이다. 페이지만 열리고 /api/design/* 이 404 면 화면은
    # 이유를 알 수 없는 고장으로 보인다. 페이지와 API 는 같이 있거나 같이 없어야 한다.
    @app.get("/design-workbench")
    def design_workbench_page():
        response = make_response(render_template("design_workbench.html"))
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response

    # ── 세션 ────────────────────────────────────────────────────────────
    @app.post("/api/design/session")
    def design_session_create():
        body = request.get_json(silent=True) or {}
        operator = (body.get("operator") or "").strip() or None
        sess = S.DesignSession.create(DESIGN_SESSION_DIR, operator=operator)
        return jsonify({"ok": True, "session_id": sess.sid}), 201

    @app.get("/api/design/session/<sid>")
    def design_session_status(sid):
        sess, err = _open(sid)
        if err:
            return err
        return jsonify({"ok": True, **sess.status()})

    # ── C1 인식 (§11.1) ─────────────────────────────────────────────────
    def _dxf_path(token: str):
        """업로드된 도면 토큰 → 경로. 업로드 폴더 밖이면 None.

        `r30_inspect` 에 같은 함수가 있으나 그쪽 register 클로저 안이고 §2 가 그
        파일 수정을 막는다. 경로 조작을 막는 유일한 방어선이라 우회하지 않고 다시
        쓴다.
        """
        token = (token or "").strip()
        if not token or UPLOAD_DIR is None or secure_filename(token) != token:
            return None
        candidate = UPLOAD_DIR / token
        return candidate if candidate.is_file() and candidate.suffix.lower() == ".dxf" else None

    def _inspect_cache(content_hash: str):
        """inspect 가 남긴 (엔티티 gz, 메타 json). 도면 내용 해시가 키다."""
        if INSPECT_CACHE_DIR is None:
            return None, None
        key = f"{INSPECT_CACHE_VERSION}_{content_hash}"
        return (INSPECT_CACHE_DIR / f"{key}.entities.ndjson.gz",
                INSPECT_CACHE_DIR / f"{key}.meta.json")

    def _cache_path(content_hash: str, floor: str, wall_layers: list):
        """§11.2 캐시 키 — 도면 내용 + 인식 파라미터 전체."""
        if DESIGN_CACHE_DIR is None:
            return None
        signature = json.dumps({"params": RP.recognize_params(), "floor": floor,
                                "wall_layers": wall_layers},
                               sort_keys=True, ensure_ascii=False)
        digest = hashlib.sha256(signature.encode("utf-8")).hexdigest()[:12]
        return DESIGN_CACHE_DIR / f"{DESIGN_CACHE_VERSION}_{content_hash}_{digest}.json.gz"

    def _cache_get(path):
        if path is None or not path.is_file():
            return None
        try:
            with gzip.open(path, "rt", encoding="utf-8") as gz:
                return json.load(gz)
        except (OSError, ValueError):
            return None

    def _cache_put(path, building):
        if path is None:
            return
        tmp = path.with_suffix(f".{uuid.uuid4().hex[:12]}.tmp")
        try:
            with gzip.open(tmp, "wt", encoding="utf-8") as gz:
                json.dump(building, gz, ensure_ascii=False)
            os.replace(tmp, path)
        except OSError:
            tmp.unlink(missing_ok=True)

    @app.post("/api/design/c1/recognize")
    def design_c1_recognize():
        """DXF(토큰) → C130~C190 → `building.json`. NDJSON 스트림(§11.1)."""
        body = request.get_json(silent=True) or {}
        sess, err = _open(str(body.get("session_id") or ""))
        if err:
            return err
        dxf_path = _dxf_path(str(body.get("dxf_token") or ""))
        if dxf_path is None:
            return _fail("DXF_NOT_FOUND", "업로드된 도면을 찾을 수 없습니다.", 400)
        try:
            content_hash = _content_hash(dxf_path)
        except OSError as exc:
            return _fail("DXF_UNREADABLE", f"도면을 읽지 못했습니다: {exc}", 400)

        ent_path, meta_path = _inspect_cache(content_hash)
        if ent_path is None or not (ent_path.is_file() and meta_path.is_file()):
            # 화면이 그린 엔티티를 다시 쓰는 구조라, 캐시가 없으면 도면을 아직
            # 안 올렸거나 스트림이 중간에 끊긴 것이다. 여기서 DXF 를 다시 파싱하면
            # 392k 도면 기준 40초를 말없이 더 쓴다.
            return _fail("DXF_NOT_INSPECTED",
                         "도면을 화면에 먼저 올려야 인식을 시작할 수 있습니다.", 409)

        floor = str(body.get("floor") or "1F").strip() or "1F"
        wall_layers = sorted({str(name) for name in (body.get("wall_layers") or [])})
        cache_path = _cache_path(content_hash, floor, wall_layers)

        # 같은 세션에 두 번 걸리면 building.json 을 서로 덮어쓴다.
        stack = contextlib.ExitStack()
        try:
            stack.enter_context(sess.stage_lock("c1"))
        except S.StageBusy as exc:
            return _fail("STAGE_BUSY", str(exc), 409)

        def _replay(building):
            """캐시 hit — 화면이 받는 메시지를 draft 에서 되살린다.

            hit 과 miss 가 다른 메시지를 내면 두 번째 실행에서 화면이 빈다. 지문과
            중심선은 draft 에 없어 되살리지 않는다 — 그 둘은 진단용이다.
            """
            draft = building["draft"]
            yield {"type": "phase", "phase": "recognize", "cached": True,
                   "unit": building["unit"], "wall_layers": building["wall_layers"],
                   "wall_source": building["wall_source"]}
            yield {"type": "virtual_edges", "edges": draft["virtual_edges"]}
            yield {"type": "rooms", "rooms": draft["rooms"]}
            yield {"type": "cores", "cores": draft["cores"]}

        def _recognized():
            """캐시 hit 이면 재사용, 아니면 사슬을 돌린다. 반환은 `building` 메시지."""
            cached = _cache_get(cache_path)
            if cached is not None:
                yield from _replay(cached)
                return cached

            meta = json.loads(meta_path.read_text(encoding="utf-8"))
            total = int((meta.get("counts") or {}).get("total_entities") or 0)
            entities = []
            with gzip.open(ent_path, "rt", encoding="utf-8") as gz:
                for line in gz:
                    chunk = json.loads(line).get("entities")
                    if not chunk:
                        continue
                    entities.extend(chunk)
                    yield {"type": "parse", "stage": "cache", "done": len(entities),
                           "total": max(total, len(entities))}

            building = None
            for msg in PL.recognize(entities, meta.get("bbox"), floor=floor,
                                    session_id=sess.sid, wall_layers=wall_layers,
                                    source={"dxf_filename": dxf_path.name,
                                            "dxf_token": dxf_path.name}):
                if msg["type"] == "building":
                    building = msg
                else:
                    yield msg
            # 막힌 결과는 캐시하지 않는다 — 되살릴 때 사람이 고를 차선(candidates)이
            # draft 에 없어, 두 번째 실행이 이유 없이 멈춘 것처럼 보인다.
            if not building["blocked"]:
                _cache_put(cache_path, building)
            return building

        def _messages():
            building = yield from _recognized()
            result = {"type": "result", "ok": True, "session_id": sess.sid,
                      "blocked": building["blocked"], "unit": building["unit"],
                      "wall_layers": building["wall_layers"],
                      "wall_source": building["wall_source"],
                      "counts": building["counts"], "stages": building["stages"],
                      "seconds": building["seconds"]}
            if building["blocked"]:
                # 막힌 채로 building.json 을 쓰면 GATE 가 빈 도면을 확정 대상으로
                # 삼는다. 안 써야 /gate_items 가 C1_NOT_DONE 으로 남는다.
                sess.audit("system", "C1", "blocked", {"reason": building["blocked"]})
                yield {**result, "ok": False, "code": building["blocked"].upper(),
                       "message": "WALL 로 볼 레이어를 고르면 인식을 이어갑니다."}
                return
            version = sess.write("building.json", building["draft"])
            sess.audit("system", "C1", "recognized",
                       {"version": version, **building["counts"]})
            yield {**result, "version": version,
                   "gate_items_url": f"/api/design/c1/gate_items/{sess.sid}"}

        def _stream():
            with stack:
                try:
                    for msg in _messages():
                        yield json.dumps(msg, ensure_ascii=False) + "\n"
                except Exception as exc:  # noqa: BLE001 — 스트림이 열린 뒤라 상태코드가 없다
                    yield json.dumps({"type": "error", "ok": False, "code": "C1_FAILED",
                                      "message": f"C1 인식 실패: {exc}"},
                                     ensure_ascii=False) + "\n"

        return Response(_stream(), mimetype="application/x-ndjson")

    # ── GATE ────────────────────────────────────────────────────────────
    @app.get("/api/design/c1/gate_items/<sid>")
    def design_gate_items(sid):
        sess, err = _open(sid)
        if err:
            return err
        draft, _version, err = _load_draft(sess)
        if err:
            return err
        return jsonify(G.gate_items(draft))

    # [문서정합] §12.6 은 편집을 `gate/confirm` 의 `edits` 로만 보내라고 적었다.
    # 그러면 화면이 합친 뒤의 폴리곤을 스스로 계산해 그려야 하고, 서버가 "맞닿지
    # 않아 못 합친다" 고 거절하는 것을 확정 순간에야 알게 된다. 편집은 한 건씩
    # 바로 반영해 결과 폴리곤을 돌려주고, `confirm` 의 `edits` 는 계약대로 남긴다.
    @app.post("/api/design/gate/edit")
    def design_gate_edit():
        body = request.get_json(silent=True) or {}
        sess, err = _open(str(body.get("session_id") or ""))
        if err:
            return err
        draft, version, err = _load_draft(sess)
        if err:
            return err
        if draft.gate.passed:
            return _fail("GATE_ALREADY_PASSED",
                         "이미 확정된 세션입니다. 실을 다시 고칠 수 없습니다.", 409)

        try:
            edits = G.apply_edits(draft, [body.get("edit") or {}])
        except (ValueError, TypeError) as exc:
            return _fail("INVALID_EDIT", str(exc), 400)

        draft.gate.edits.extend(edits)
        draft.gate.unresolved = G.unresolved(draft)
        try:
            new_version = sess.write("building.json", draft.to_dict(),
                                     if_version=version)
        except S.VersionConflict as conflict:
            return _fail("VERSION_CONFLICT",
                         "다른 곳에서 먼저 저장했습니다. 현재 내용을 확인하세요.", 409,
                         current_version=conflict.current, current=conflict.data)

        actor = (body.get("operator") or "").strip() or "unknown"
        sess.audit(actor, "GATE", "room_edit", edits[0])
        return jsonify({"ok": True, "version": new_version, "edit": edits[0],
                        "rooms": [r.to_dict() for r in draft.rooms]})

    @app.post("/api/design/gate/confirm")
    def design_gate_confirm():
        body = request.get_json(silent=True) or {}
        sess, err = _open(str(body.get("session_id") or ""))
        if err:
            return err
        draft, version, err = _load_draft(sess)
        if err:
            return err

        operator = (body.get("operator") or "").strip() or None
        if_version = body.get("if_version")
        # 편집이 먼저다. 사람이 실을 자른 뒤 그 자식 실의 용도를 같은 요청에
        # 실어 보내므로, 값을 먼저 반영하면 아직 없는 실을 가리킨다.
        try:
            edits = G.apply_edits(draft, body.get("edits") or [])
            changes = G.apply_values(draft, body.get("values") or {})
        except (ValueError, TypeError) as exc:
            return _fail("INVALID_VALUE", str(exc), 400)

        defaults = G.apply_defaults(draft)
        draft.gate.edits.extend(edits)
        draft.gate.operator = operator
        draft.gate.unresolved = G.unresolved(draft)
        draft.gate.passed = not draft.gate.unresolved
        draft.gate.passed_at = S.now_iso() if draft.gate.passed else None

        try:
            new_version = sess.write(
                "building.json", draft.to_dict(),
                if_version=int(if_version) if if_version is not None else version)
        except S.VersionConflict as conflict:
            return _fail("VERSION_CONFLICT",
                         "다른 곳에서 먼저 저장했습니다. 현재 내용을 확인하세요.", 409,
                         current_version=conflict.current, current=conflict.data)

        actor = operator or "unknown"
        for change in changes:
            event = ("use_override"
                     if change["field"] == "use" and change.get("suggested") else "value_confirmed")
            sess.audit(actor, "GATE", event, change)
        for applied in defaults:
            sess.audit("system", "GATE", "default_applied", applied)
        for edit in edits:
            sess.audit(actor, "GATE", "room_edit", edit)

        if not draft.gate.passed:
            sess.audit(actor, "GATE", "confirm_incomplete",
                       {"unresolved": len(draft.gate.unresolved)})
            return jsonify({
                "ok": False, "code": "GATE_INCOMPLETE",
                "message": "확정되지 않은 필수 항목이 남아 있습니다.",
                "unresolved": draft.gate.unresolved, "version": new_version,
                "defaults": defaults,
            }), 422

        sess.update_meta(stage="c2", gate_passed=True, operator=operator)
        sess.audit(actor, "GATE", "passed",
                   {"rooms": len(draft.rooms), "edits": len(draft.gate.edits)})
        # `defaults` 를 함께 낸다 — 서버가 근거를 들어 채운 값을 화면이 모르면
        # 결손이 아닌데 비어 있는 칸이 생긴다.
        return jsonify({"ok": True, "passed": True, "version": new_version,
                        "passed_at": draft.gate.passed_at, "defaults": defaults})

    # ── 게이트 뒤 단계 (자리만) ─────────────────────────────────────────
    def _make_gated(stage: str):
        def view(sid=None):
            session_id = sid or str((request.get_json(silent=True) or {}).get("session_id") or "")
            sess, err = _open(session_id)
            if err:
                return err
            draft, _version, err = _load_draft(sess)
            if err:
                return err
            try:
                G.require_gate(draft)
            except G.GateNotPassed as exc:
                return _fail("GATE_NOT_PASSED", str(exc), 409, unresolved=exc.unresolved)
            return _fail("STAGE_NOT_IMPLEMENTED",
                         f"{stage} 단계는 아직 구현되지 않았습니다.", 501)
        return view

    for rule, method, stage in _GATED_STAGES:
        app.add_url_rule(rule, endpoint=f"design_stage_{stage.lower().replace('-', '_')}",
                         view_func=_make_gated(stage), methods=[method])
