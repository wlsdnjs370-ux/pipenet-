# -*- coding: utf-8 -*-
"""모듈 C 설계 자동화 API — 지시서 §11.

`register(app, ...)` 패턴. 기존 `/api/remote30/*` 는 건드리지 않으며, 이 모듈은
`DESIGN_WORKBENCH_ENABLED` 가 꺼져 있으면 라우트를 아예 등록하지 않는다(§C.5).

아직 구현되지 않은 C3 이후 엔드포인트는 **게이트 강제만** 걸어 자리를 잡아 두고,
게이트를 통과한 뒤에는 구현이 없음을 501 로 밝힌다. 통과한 것처럼 200 을 돌려주면
화면이 진행된 것으로 오해한다.
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
from core.design.deterministic import constraints as CN
from core.design.deterministic import esfr as ESFR
from core.design.deterministic import zoning as Z
from core.design.recognize import params as RP
from core.design.recognize import pipeline as PL
from core.design.schema import BuildingDraft

# [문서정합] §11.2 는 이 상수를 쓰라고만 하고 둘 곳을 지정하지 않았다. 앱 엔트리
# (`대조 서버.py`)는 §2 가 register 호출 한 줄만 허용하므로 여기 둔다. C1 사슬의
# 판정이 바뀌면 이 값을 올려야 옛 결과가 캐시에서 되살아나지 않는다.
DESIGN_CACHE_VERSION = "c1-v1"

# 게이트 뒤에 서지만 아직 구현이 없는 단계들. C2·C2B 는 실제 라우트로 빠졌다.
_GATED_STAGES = (
    ("/api/design/c4/heads", "POST", "C4"),
    ("/api/design/c5/route", "POST", "C5"),
    ("/api/design/emit", "POST", "EMIT"),
    ("/api/design/checks/<sid>", "GET", "CHECKS"),
)


# meta.stage 는 앞으로만 간다. C2B 로 기준을 다시 구우면서 stage 를 되감으면 이미
# C4 까지 간 세션이 화면에서 C2 로 돌아간 것처럼 보인다.
_STAGE_ORDER = ("c1", "c2", "c3", "c4", "c5", "emit")

# 세션 파일이 관리하는 값. 내용 비교에서는 빼야 같은 기준을 다시 구웠는지 알 수 있다.
_VOLATILE_KEYS = ("version", "updated_at")


def _fail(code: str, message: str, status: int, **extra):
    return jsonify({"ok": False, "code": code, "message": message, **extra}), status


def _stage_fields(meta: dict, stage: str) -> dict:
    """`update_meta` 에 실을 stage 갱신. 뒤로 가는 값은 실지 않는다."""
    current = meta.get("stage")
    if current in _STAGE_ORDER and _STAGE_ORDER.index(current) >= _STAGE_ORDER.index(stage):
        return {}
    return {"stage": stage}


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
        sess.audit_many(actor, "GATE", [
            (("use_override" if change["field"] == "use" and change.get("suggested")
              else "value_confirmed"), change) for change in changes
        ] + [("room_edit", edit) for edit in edits])
        # 기본값은 사람이 아니라 서버가 채웠다. 행위자를 섞으면 감리에서 누가
        # 확정했는지 가려낼 수 없다.
        sess.audit_many("system", "GATE", [("default_applied", a) for a in defaults])

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

    # ── 게이트 뒤 공통 ──────────────────────────────────────────────────
    def _gated(session_id: str):
        """반환 `(세션, draft, None)` 또는 `(None, None, 오류응답)`."""
        sess, err = _open(session_id)
        if err:
            return None, None, err
        draft, _version, err = _load_draft(sess)
        if err:
            return None, None, err
        try:
            G.require_gate(draft)
        except G.GateNotPassed as exc:
            return None, None, _fail("GATE_NOT_PASSED", str(exc), 409,
                                     unresolved=exc.unresolved)
        return sess, draft, None

    # ── C2 결정론 기준 (§5, §11) ────────────────────────────────────────
    def _current_constraints(sess, meta, data):
        """내용이 같은 현행 기준이 있으면 `(이름, version, 내용)`, 없으면 None.

        같은 값을 다시 구울 때마다 `constraints.v2.json` 을 만들면 화면 새로고침
        한 번에 세대가 하나씩 늘어 감사 기록에서 진짜 재검토가 묻힌다.
        """
        name = meta.get("constraints")
        if not name:
            return None
        current, version = sess.read_or_none(name)
        if current is None:
            return None
        stripped = {k: v for k, v in current.items() if k not in _VOLATILE_KEYS}
        return (name, version, current) if stripped == data else None

    def _bake_constraints(sess, draft, actor, decision=None):
        """반환 `(payload, None)` 또는 `(None, 오류응답)`.

        기준은 덮어쓰지 않는다(§16.3). 값이 실제로 달라졌을 때만
        `constraints.v2.json` 을 새로 쓰고 meta 가 현행을 가리킨다 — 어떤 기준으로
        설계했는지가 감사 대상이라 옛 기준이 남아 있어야 한다.
        """
        try:
            data = CN.build_constraints(draft.to_dict()).to_dict()
        except CN.MissingBuildingFact as exc:
            # GATE 로 돌려보낸다. 결손은 코드가 추측할 항목이 아니다.
            return None, _fail("MISSING_BUILDING_FACT", str(exc), 422,
                               field=exc.field, detail=exc.detail)
        if decision is None:
            stored, _v = sess.read_or_none("esfr.json")
            if stored is not None:
                decision = ESFR.EsfrDecision.from_dict(stored)
        if decision is not None:
            data = ESFR.apply(data, decision)

        meta, _ = sess.read("meta.json")
        same = _current_constraints(sess, meta, data)
        if same is not None:
            name, version, stored_data = same
            return {"artifact": name, "version": version, "constraints": stored_data,
                    "rebaked": False}, None

        name = sess.next_constraints_name()
        version = sess.write(name, data)
        sess.update_meta(constraints=name, **_stage_fields(meta, "c3"))
        sess.audit(actor, "C2", "constraints_built", {
            "artifact": name,
            "scenario_head_count": data.get("scenario_head_count"),
            "horizontal_distance_m": data.get("horizontal_distance_m"),
            "esfr_active": bool((data.get("esfr") or {}).get("active")),
        })
        return {"artifact": name, "version": version, "constraints": data,
                "rebaked": True}, None

    @app.post("/api/design/c2/constraints")
    def design_c2_constraints():
        body = request.get_json(silent=True) or {}
        sess, draft, err = _gated(str(body.get("session_id") or ""))
        if err:
            return err
        actor = (body.get("operator") or "").strip() or draft.gate.operator or "unknown"
        try:
            with sess.stage_lock("c2"):
                payload, err = _bake_constraints(sess, draft, actor)
        except S.StageBusy as exc:
            return _fail("STAGE_BUSY", str(exc), 409)
        if err:
            return err
        return jsonify({"ok": True, **payload})

    # ── C2B 103B 활성 (§6) ──────────────────────────────────────────────
    @app.post("/api/design/c2b/esfr")
    def design_c2b_esfr():
        body = request.get_json(silent=True) or {}
        sess, draft, err = _gated(str(body.get("session_id") or ""))
        if err:
            return err
        # 빠진 값을 False 로 읽으면 "묻지 않았다"가 "아니라고 답했다"가 되고 그 답이
        # 설비 종류를 가른다 — 세 조건 전부 참/거짓으로 명시돼야 받는다.
        missing = [key for key in ESFR.INPUTS if not isinstance(body.get(key), bool)]
        if missing:
            return _fail("ESFR_INPUT_REQUIRED",
                         "103B 세 조건을 참/거짓으로 명시해야 합니다: " + ", ".join(missing),
                         400, fields=missing)
        try:
            decision = ESFR.decide(
                operator=(body.get("operator") or "").strip(),
                decided_at=S.now_iso(), note=str(body.get("note") or ""),
                **{key: body[key] for key in ESFR.INPUTS})
        except ValueError as exc:
            return _fail("OPERATOR_REQUIRED", str(exc), 400)

        try:
            with sess.stage_lock("c2"):
                payload = None
                # 기준을 이미 구웠다면 그 자리에서 다시 굽는다. 두 산출물이 어긋난
                # 채 남으면 하류가 어느 쪽을 믿느냐에 따라 설계가 갈린다. 굽지
                # 못하면 판단도 남기지 않는다 — 한쪽만 남으면 다음 C2 가 조용히
                # 다른 기준을 만든다.
                if sess.path(S.CONSTRAINTS).is_file():
                    payload, err = _bake_constraints(sess, draft, decision.operator,
                                                     decision=decision)
                    if err:
                        return err
                sess.write("esfr.json", decision.to_dict())
                sess.audit(decision.operator, "C2B", "esfr_decided", decision.to_dict())
        except S.StageBusy as exc:
            return _fail("STAGE_BUSY", str(exc), 409)
        return jsonify({"ok": True, "esfr": decision.to_dict(), **(payload or {})})

    # ── C3 유수검지장치·방호구역 (§7) ───────────────────────────────────
    def _load_constraints(sess):
        """현행 기준. C2 를 굽기 전이면 오류 — 기준 없이 구역을 나눌 수 없다."""
        meta, _ = sess.read("meta.json")
        data, _v = sess.read_or_none(meta.get("constraints") or S.CONSTRAINTS)
        if data is None:
            return None, _fail("CONSTRAINTS_REQUIRED",
                               "C2 규범 조건을 먼저 실행하세요.", 409)
        return data, None

    @app.post("/api/design/c3/valves")
    def design_c3_valves():
        body = request.get_json(silent=True) or {}
        sess, draft, err = _gated(str(body.get("session_id") or ""))
        if err:
            return err
        operator = (body.get("operator") or "").strip()
        if not operator:
            return _fail("OPERATOR_REQUIRED", "밸브를 확정한 사람이 필요합니다.", 400)
        # 기준 없이 확정한 밸브는 어떤 규범으로 놓았는지 말할 수 없다.
        _constraints, err = _load_constraints(sess)
        if err:
            return err

        candidates = {c["core_id"]: c for c in Z.valve_candidates(draft)}
        placed = body.get("valves")
        if not isinstance(placed, list) or not placed:
            return _fail("VALVE_REQUIRED",
                         "유수검지장치를 하나 이상 찍어야 합니다.", 400,
                         candidates=list(candidates.values()))

        rooms_by_id = {r.id: r for r in draft.rooms}
        valves = []
        for i, item in enumerate(placed, start=1):
            core_id = str((item or {}).get("core_id") or "")
            if core_id not in candidates:
                return _fail("VALVE_CORE_UNCONFIRMED",
                             f"{core_id or '(빈 값)'} 은(는) 확정된 샤프트·계단 코어가 "
                             "아닙니다. 후보 중에서 고르세요.", 422,
                             candidates=list(candidates.values()))
            system_type = str(item.get("system_type") or "")
            if system_type not in Z.SYSTEM_TYPES:
                # 동결·수손 우려는 도면에 없다. 기본값을 넣으면 습식이 아닌 현장이
                # 조용히 습식으로 확정된다.
                return _fail("SYSTEM_TYPE_REQUIRED",
                             f"{core_id} 의 설비 종류를 골라야 합니다: "
                             + ", ".join(Z.SYSTEM_TYPES), 422, core_id=core_id)
            point = item.get("point")
            if not (isinstance(point, (list, tuple)) and len(point) == 2):
                return _fail("VALVE_POINT_REQUIRED",
                             f"{core_id} 의 설치 좌표가 없습니다.", 422, core_id=core_id)
            room_id = Z.seed_room(draft, point)
            if room_id is None:
                return _fail("VALVE_OUTSIDE_ROOMS",
                             f"{core_id} 에 찍은 자리가 어느 실 안에도 없습니다. "
                             "코어 폴리곤 안을 찍으세요.", 422, core_id=core_id)
            try:
                confirmed = Z.check_requirements(item.get("requirements_confirmed") or {})
            except Z.ZoningError as exc:
                return _fail(exc.code, f"{core_id}: {exc}", 422,
                             core_id=core_id, **exc.detail)
            # 층은 코어가 아니라 밸브를 찍은 실이 안다. `Core` 에는 층이 없다.
            floor = rooms_by_id[room_id].floor
            valves.append({
                "id": f"AV-{floor or '?'}-{i:02d}",
                "core_id": core_id, "floor": floor,
                "point": [float(point[0]), float(point[1])],
                "room_id": room_id, "system_type": system_type,
                "manual": True, "operator": operator,
                "requirements_confirmed": confirmed,
            })

        unmet = [{"valve_id": v["id"], "fields": [k for k in Z.MANUAL_REQUIREMENTS
                                                  if not v["requirements_confirmed"][k]]}
                 for v in valves]
        unmet = [u for u in unmet if u["fields"]]

        try:
            with sess.stage_lock("c3"):
                design, _v = sess.read_or_none("design.json")
                design = dict(design or {"schema": "fncadnet.design/1"})
                design["valves"] = valves
                # 밸브가 바뀌면 옛 구역은 근거를 잃는다. 남겨 두면 화면이 이전
                # 구역을 새 밸브의 것으로 읽는다.
                design.pop("zones", None)
                sess.write("design.json", design)
                sess.audit(operator, "C3", "valves_placed", {
                    "count": len(valves),
                    "valves": [{k: v[k] for k in ("id", "core_id", "system_type",
                                                  "room_id", "requirements_confirmed")}
                               for v in valves],
                })
        except S.StageBusy as exc:
            return _fail("STAGE_BUSY", str(exc), 409)
        return jsonify({"ok": True, "valves": valves, "requirements_unmet": unmet})

    @app.get("/api/design/c3/candidates/<sid>")
    def design_c3_candidates(sid):
        sess, draft, err = _gated(sid)
        if err:
            return err
        return jsonify({"ok": True, "candidates": Z.valve_candidates(draft),
                        "system_types": list(Z.SYSTEM_TYPES),
                        "requirements": [{"key": k, "label": Z.REQUIREMENT_LABELS[k]}
                                         for k in Z.MANUAL_REQUIREMENTS]})

    @app.post("/api/design/c3/zones")
    def design_c3_zones():
        body = request.get_json(silent=True) or {}
        sess, draft, err = _gated(str(body.get("session_id") or ""))
        if err:
            return err
        constraints, err = _load_constraints(sess)
        if err:
            return err
        design, _v = sess.read_or_none("design.json")
        valves = (design or {}).get("valves") or []
        if not valves:
            return _fail("VALVE_REQUIRED",
                         "유수검지장치를 먼저 확정하세요.", 409)

        result = Z.split_zones(draft, valves, constraints)
        actor = (body.get("operator") or "").strip() or valves[0].get("operator") or "unknown"
        try:
            with sess.stage_lock("c3"):
                stored = dict(design)
                stored["zones"] = result["zones"]
                stored["flags"] = result["flags"]
                sess.write("design.json", stored)
                # 플래그가 남아 있으면 구역이 아직 사람 손을 기다린다. 여기서 stage 를
                # 밀면 닿지 않는 실을 안은 채로 C4 헤드 배치가 열린다.
                if not result["flags"]:
                    meta, _ = sess.read("meta.json")
                    fields = _stage_fields(meta, "c4")
                    if fields:
                        sess.update_meta(**fields)
                sess.audit(actor, "C3", "zones_split", {
                    "zones": len(result["zones"]),
                    "unreached": len(result["unreached"]),
                    "flags": [f["code"] for f in result["flags"]],
                })
        except S.StageBusy as exc:
            return _fail("STAGE_BUSY", str(exc), 409)
        return jsonify({"ok": True, **result})

    # ── 게이트 뒤 단계 (자리만) ─────────────────────────────────────────
    def _make_gated(stage: str):
        def view(sid=None):
            session_id = sid or str((request.get_json(silent=True) or {}).get("session_id") or "")
            _sess, _draft, err = _gated(session_id)
            if err:
                return err
            return _fail("STAGE_NOT_IMPLEMENTED",
                         f"{stage} 단계는 아직 구현되지 않았습니다.", 501)
        return view

    for rule, method, stage in _GATED_STAGES:
        app.add_url_rule(rule, endpoint=f"design_stage_{stage.lower().replace('-', '_')}",
                         view_func=_make_gated(stage), methods=[method])
