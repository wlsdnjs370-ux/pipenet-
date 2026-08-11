# -*- coding: utf-8 -*-
"""모듈 C 설계 자동화 API — 지시서 §11.

`register(app, ...)` 패턴. 기존 `/api/remote30/*` 는 건드리지 않으며, 이 모듈은
`DESIGN_WORKBENCH_ENABLED` 가 꺼져 있으면 라우트를 아예 등록하지 않는다(§C.5).

C2 이후는 전부 게이트 뒤에 선다. 단계를 넘길지는 서버가 정하고(`meta.stage`),
방출은 그 값이 `emit` 일 때만 돈다 — 막는 플래그를 안은 망이 솔버 파일이 되어
나가지 않게 하는 마지막 잠금이다.
"""
from __future__ import annotations

import contextlib
import gzip
import hashlib
import json
import math
import os
import shutil
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path

from flask import (Response, jsonify, make_response, render_template, request,
                   send_from_directory)
from werkzeug.utils import secure_filename

from core.upload_names import sanitize_upload_name
from core.design import gate as G
from core.design import session as S
from core.design.checks import dimensional as DIM
from core.design.deterministic import build_network as BN
from core.design.deterministic import constraints as CN
from core.design.deterministic import esfr as ESFR
from core.design.deterministic import head_layout as HL
from core.design.deterministic import zoning as Z
from core.design.recognize import params as RP
from core.design.recognize import pipeline as PL
from core.design.schema import BuildingDraft

# [문서정합] §11.2 는 이 상수를 쓰라고만 하고 둘 곳을 지정하지 않았다. 앱 엔트리
# (`대조 서버.py`)는 §2 가 register 호출 한 줄만 허용하므로 여기 둔다. C1 사슬의
# 판정이 바뀌면 이 값을 올려야 옛 결과가 캐시에서 되살아나지 않는다.
DESIGN_CACHE_VERSION = "c1-v1"

# 방출 산출물이 놓이는 자리. 세션 안에 두어 세션이 지워질 때 함께 지워진다.
_EMIT_DIR = "emit"

# C5 플래그 중 **경고**로 볼 것. 나머지는 방출을 막는다. C4 와 같은 이유로 목록을
# 뒤집어 둔다 — 새 플래그의 기본은 "막는다" 여야 한다.
_C5_WARNINGS = frozenset({"MONOTONIC_CLAMP", "MULTIPLE_RISERS"})

# C4 배치 플래그 중 **경고**로 볼 것. 나머지는 전부 다음 단계를 막는다.
# 목록을 뒤집어 둔 이유는, 새 플래그가 생겼을 때 기본이 "막는다" 여야 하기
# 때문이다. 통과 목록을 두면 새 플래그가 조용히 통과한다.
_C4_WARNINGS = frozenset({"OBSTACLE_UNVERIFIED"})


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


class _HeadLimits:
    """`constraints.json` 중 C4 가 쓰는 값만 속성으로 읽는다.

    `Constraints` 는 frozen dataclass 라 json 에서 되살리려면 전 필드를 다 채워야
    하는데, 그 되살리는 코드가 곧 두 번째 진실 출처가 된다 — 한쪽만 고쳐지면 화면과
    구운 기준이 조용히 갈린다. C4 가 실제로 보는 값은 네 개뿐이므로 그 넷만 읽고,
    하나라도 없으면 기본값을 넣지 않고 거절한다.
    """

    FIELDS = ("horizontal_distance_m", "head_spacing_square_m",
              "head_clearance_radius_m", "head_to_wall_clearance_m")

    def __init__(self, data: dict):
        self.missing = [k for k in self.FIELDS if data.get(k) is None]
        for key in self.FIELDS:
            setattr(self, key, float(data.get(key) or 0.0))


def _emit_stem(network: dict, index: int) -> str:
    """방출 파일 이름. 코어 이름이 그대로 파일 이름이 되므로 여기서 씻는다."""
    core = secure_filename(str((network.get("riser") or {}).get("core_id") or ""))
    return f"net{index:02d}_{core}" if core else f"net{index:02d}"


def _sdf_pipe_lengths(sdf: Path) -> list[float]:
    """SDF 가 스스로 적어 둔 배관장(m).

    `parse_sdf` 의 `length_m` 을 쓰지 않는 이유가 있다. 그 함수는 `length` 와 노드
    좌표거리의 비 중앙값이 0.7~1.5 면 **좌표거리를 권위로 삼아 length 를 덮어쓴다**
    (K-solver 규약이 좌표거리＝배관장이라서다). 그런데 `emit_sdf` 의 Position 은
    최장변 3000 으로 정규화한 스키매틱 좌표라 미터가 아니다 — 그 배율이 우연히
    저 창에 들면 읽은 배관장이 일제히 밀린다(실측: 양주옥정 1망, 비 1.5,
    5.632m → 5.443m). PIPENET 은 `length` 를 읽으므로 파일 자체는 옳다. 그래서
    "낸 값이 그대로 있는가" 는 `length` 로 재고, 좌표는 여기서 보지 않는다.
    """
    root = ET.parse(str(sdf)).getroot()
    return [float(el.attrib["length"]) for el in root.iter()
            if el.tag.rpartition("}")[2] == "Pipe" and "length" in el.attrib]


def _round_trip(combined, net, sdf_lengths: list[float]) -> dict:
    """§13.6 — 방출한 SDF 를 다시 읽어 그래프와 맞대어 본다.

    `kfp_sdf_converter.round_trip_check` 는 KFP 를 기점으로 도는 함수라 여기 쓸 수
    없다(우리 기점은 그래프다). 같은 항목을 같은 방식으로 잰다.

    모듈 A 는 신축배관 등가길이를 실배관 한 토막으로 편다(`_materialize_fx_pipes`).
    그래서 노드 수·배관 수·총연장이 신축배관 개수만큼 **느는 것이 정상**이다.
    """
    flex = len(combined.equipment)
    flex_len_m = sum(float(eq["drawing_len_mm"]) / 1000.0
                     for eq in combined.equipment)
    length_before = sum(float(p["length"]) for p in combined.pipes)
    length_after = sum(sdf_lengths)
    dia_before = {f"P{p['label']}": int(p["dia"]) for p in combined.pipes}
    dia_after = {p.id: p.nominal_mm for p in net.pipes.values()}
    same_dia = sum(1 for pid, dn in dia_before.items() if dia_after.get(pid) == dn)
    heads_after = sum(1 for n in net.nodes.values() if n.kind == "head")

    out = {
        "nodes": [len(combined.nodes), len(net.nodes)],
        "pipes": [len(combined.pipes), len(net.pipes)],
        "heads": [len(combined.nozzles), heads_after],
        "length_m": [round(length_before, 3), round(length_after, 3)],
        "dia_match": [same_dia, len(dia_before)],
        "flex": flex, "flex_length_m": round(flex_len_m, 3),
    }
    out["ok"] = bool(
        out["nodes"][1] == out["nodes"][0] + flex
        and out["pipes"][1] == out["pipes"][0] + flex
        and heads_after == len(combined.nozzles)
        and math.isclose(length_before + flex_len_m, length_after, abs_tol=0.01)
        and same_dia == len(dia_before))
    return out


def _emit_bundle(graph: dict, out_dir: Path, stem: str, *, title: str) -> dict:
    """그래프 한 벌 → SDF(+동봉 SLF)·KFP·HAS + 왕복 검증.

    모듈 A 는 부르기만 하고 고치지 않는다(금지사항 G8). import 를 함수 안에 두는
    이유는 `_inner_dia_mm` 과 같다 — `remote30_prototype` 이 평면 import 로 짜여
    `core` 자체가 sys.path 에 올라야 열린다.

    세 포맷 중 하나라도 실패하면 여기서 터진다. 잡아서 넘기면 SDF 만 나온 것을
    "3종 생성" 으로 읽게 된다.
    """
    from kfp_sdf_converter import parse_sdf
    from remote30_full_network import (CombinedTables, ProjectContext, ZoneSpec,
                                       ZoneType, emit_full_sdf)
    from remote30_prototype import emit_has, emit_kfp

    combined = CombinedTables(
        nodes=list(graph["nodes"]), pipes=list(graph["pipes"]),
        nozzles=list(graph["nozzles"]), fittings=list(graph["fittings"]),
        equipment=list(graph["equipment"]),
        meta=[tuple(pair) for pair in (graph.get("meta") or ())])
    # `zone_spec` 은 `ProjectContext` 의 필수 인자지만 이 경로의 `emit_full_sdf` 는
    # `report_title()` 만 본다. 가압방식은 모듈 C 가 아직 정하지 않으므로, 여기 적은
    # 값이 수리계산에 흘러들지 않는다는 사실이 곧 이 줄이 G9 위반이 아닌 근거다.
    ctx = ProjectContext(zone_spec=ZoneSpec(zone_type=ZoneType.LSP_GRAVITY),
                         project_title=title)

    sdf = out_dir / f"{stem}.sdf"
    emit_full_sdf(combined, sdf, ctx=ctx)
    kfp = out_dir / f"{stem}.kfp"
    # display_geometry=False — K-Fire Solver 는 노드 좌표거리에서 배관장을 역산한다.
    emit_kfp(sdf, kfp, coord_scale=1.0, display_geometry=False)
    has = out_dir / f"{stem}.has"
    # HASS 는 InsertionPoint 2D 좌표를 그대로 표시한다 — 등각으로 구워야 입상관이
    # 기둥으로 보인다.
    emit_has(sdf, has, isometric=True, iso_z_scale=1.0)

    # PIPENET 은 .sdf 옆의 .slf 로 호칭경↔내경을 찾는다. `emit_sdf` 가 같은 이름으로
    # 동봉하므로 내려보낼 때도 반드시 이 쌍이어야 한다.
    files = [{"name": path.name, "bytes": path.stat().st_size}
             for path in (sdf, sdf.with_suffix(".slf"), kfp, has)]
    return {"stem": stem, "files": files,
            "round_trip": _round_trip(combined, parse_sdf(sdf),
                                      _sdf_pipe_lengths(sdf))}


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
        if not token or UPLOAD_DIR is None or sanitize_upload_name(token) != token:
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

    # ── C4 헤드 배치 (§8) ───────────────────────────────────────────────
    @app.post("/api/design/c4/heads")
    def design_c4_heads():
        body = request.get_json(silent=True) or {}
        sess, draft, err = _gated(str(body.get("session_id") or ""))
        if err:
            return err
        raw, err = _load_constraints(sess)
        if err:
            return err
        limits = _HeadLimits(raw)
        if limits.missing:
            return _fail("CONSTRAINTS_INCOMPLETE",
                         "기준에 헤드 값이 없습니다: " + ", ".join(limits.missing), 409)

        design, _v = sess.read_or_none("design.json")
        zones = (design or {}).get("zones") or []
        if not zones:
            return _fail("ZONE_REQUIRED", "C3 방호구역을 먼저 나누세요.", 409)
        # 어느 밸브도 담당하지 않는 실에 헤드를 놓으면, 그 헤드에 물을 보내는
        # 배관을 C5 가 그릴 수 없다. 빠진 실은 지우지 말고 세어서 보인다.
        zoned = {rid for z in zones for rid in (z.get("rooms") or [])}

        layouts, skipped = [], []
        for index, room in enumerate(draft.rooms):
            if room.id not in zoned:
                skipped.append(room.id)
                continue
            layout = HL.layout_heads(room, limits, room_index=index)
            HL.check_obstacles(layout, room, limits, draft.obstacles)
            layouts.append(layout)

        flags = [dict(f, room=f.get("room") or layout.room_id)
                 for layout in layouts for f in layout.flags]
        blocking = [f for f in flags if f["code"] not in _C4_WARNINGS]
        checks = DIM.run_checks(draft.rooms, layouts, limits,
                                gross_floor_area_m2=body.get("gross_floor_area_m2"))
        heads = [layout.to_dict() for layout in layouts]
        actor = (body.get("operator") or "").strip() or "unknown"

        try:
            with sess.stage_lock("c4"):
                stored = dict(design)
                stored["heads"] = heads
                stored["head_flags"] = flags
                sess.write("design.json", stored)
                sess.write("checks.json", checks)
                # C3 가 아직 문제를 안고 있으면 stage 는 "c3" 에 서 있다. 거기서 c5 로
                # 밀면 닿지 않는 실을 안은 채 C5 가 열린다. 배치는 하되(빠진 실은
                # `rooms_without_zone` 로 센다) 단계는 C3 가 풀린 뒤에 넘긴다.
                meta, _ = sess.read("meta.json")
                if not blocking and meta.get("stage") == "c4":
                    sess.update_meta(stage="c5")
                sess.audit(actor, "C4", "heads_placed", {
                    "rooms": len(layouts), "skipped": len(skipped),
                    "heads": sum(len(h["heads"]) for h in heads),
                    "flags": sorted({f["code"] for f in flags}),
                    "check_flags": [c["code"] for c in checks["flags"]],
                })
        except S.StageBusy as exc:
            return _fail("STAGE_BUSY", str(exc), 409)

        return jsonify({"ok": True, "rooms": heads, "flags": flags,
                        "blocking": [f["code"] for f in blocking],
                        "rooms_without_zone": skipped, "checks": checks})

    # ── C5 배관망 라우팅 (§9) ───────────────────────────────────────────
    def _inner_dia_mm():
        """호칭경 → KS 실내경. 모듈 A 의 표를 그대로 쓴다.

        모듈 최상단에서 import 하지 않는 이유는 `hydraulic_solver` 가 평면
        import(`from hb_rules import ...`)로 짜여 `core` 자체가 sys.path 에 올라야
        열리기 때문이다. 그 부담을 이 모듈에 지우면 C1~C4 라우트까지 함께 묶인다.
        같은 표를 설계 쪽에 한 벌 더 두면 두 값이 조용히 갈린다.
        """
        import hydraulic_solver as HS
        return HS.inner_diameter_mm

    @app.post("/api/design/c5/route")
    def design_c5_route():
        body = request.get_json(silent=True) or {}
        sess, draft, err = _gated(str(body.get("session_id") or ""))
        if err:
            return err
        constraints, err = _load_constraints(sess)
        if err:
            return err
        design, _v = sess.read_or_none("design.json")
        if not (design or {}).get("zones"):
            return _fail("ZONE_REQUIRED", "C3 방호구역을 먼저 나누세요.", 409)
        if not (design or {}).get("heads"):
            return _fail("HEADS_REQUIRED", "C4 헤드 배치를 먼저 실행하세요.", 409)

        result = BN.build_networks(
            draft, design, constraints, inner_dia_mm=_inner_dia_mm(),
            meta=[("Project", draft.source.get("dxf") or sess.sid)])
        networks = [n.to_dict() for n in result.networks]
        flags = result.flags + [f for n in result.networks
                                for f in (n.tables.flags if n.tables else ())]
        blocking = [f for f in flags if f["code"] not in _C5_WARNINGS]
        actor = (body.get("operator") or "").strip() or "unknown"

        try:
            with sess.stage_lock("c5"):
                stored = dict(design)
                stored["networks"] = networks
                stored["network_flags"] = flags
                sess.write("design.json", stored)
                sess.write("graph.json", {"schema": "fncadnet.graph/1",
                                          "networks": networks})
                meta, _ = sess.read("meta.json")
                stage = meta.get("stage")
                # 앞 단계가 문제를 안고 있으면 stage 는 거기 서 있다. 단계를 화면이
                # 따로 셈하면 서버와 갈리므로 여기서 정한 값을 그대로 돌려준다.
                if not blocking and stage == "c5":
                    sess.update_meta(stage="emit")
                    stage = "emit"
                sess.audit(actor, "C5", "network_routed", {
                    "networks": len(networks),
                    "metrics": [n["metrics"] for n in networks],
                    "flags": sorted({f["code"] for f in flags}),
                })
        except S.StageBusy as exc:
            return _fail("STAGE_BUSY", str(exc), 409)

        return jsonify({"ok": True, "networks": networks, "flags": flags,
                        "stage": stage,
                        "blocking": sorted({f["code"] for f in blocking})})

    @app.get("/api/design/checks/<sid>")
    def design_checks(sid):
        sess, _draft, err = _gated(sid)
        if err:
            return err
        data, _v = sess.read_or_none("checks.json")
        if data is None:
            return _fail("CHECKS_NOT_RUN", "C4 헤드 배치를 먼저 실행하세요.", 409)
        return jsonify({"ok": True, **data})

    # ── 방출 — 모듈 A 체이닝 (§11 · §13.6) ──────────────────────────────
    @app.post("/api/design/emit")
    def design_emit():
        body = request.get_json(silent=True) or {}
        sess, draft, err = _gated(str(body.get("session_id") or ""))
        if err:
            return err
        graph, _v = sess.read_or_none("graph.json")
        networks = (graph or {}).get("networks") or []
        if not networks:
            return _fail("NETWORK_REQUIRED", "C5 배관망 라우팅을 먼저 실행하세요.", 409)
        # 단계는 C5 가 정해 둔 것을 그대로 믿는다. 배관망 파일만 있으면 방출하도록
        # 하면 막는 플래그를 안은 망이 솔버 파일이 되어 나간다.
        meta, _ = sess.read("meta.json")
        if meta.get("stage") != "emit":
            return _fail("EMIT_NOT_READY",
                         "앞 단계에 남은 문제가 있어 방출하지 않습니다.", 409,
                         stage=meta.get("stage"))

        title = draft.source.get("dxf") or sess.sid
        actor = (body.get("operator") or "").strip() or "unknown"
        out_dir = sess.path(_EMIT_DIR)

        try:
            with sess.stage_lock("emit"):
                # 옛 방출물을 남겨 두면 배관망을 다시 그린 뒤에도 내려받기가 옛 파일을
                # 준다 — 화면이 말하는 망과 손에 쥔 파일이 갈린다.
                shutil.rmtree(out_dir, ignore_errors=True)
                out_dir.mkdir(parents=True)
                bundles = [_emit_bundle(net["graph"], out_dir,
                                        _emit_stem(net, index), title=title)
                           for index, net in enumerate(networks, start=1)]
                broken = [b["stem"] for b in bundles if not b["round_trip"]["ok"]]
                flags = [{
                    "code": "ROUND_TRIP_MISMATCH", "networks": broken,
                    "message": f"방출한 SDF 를 다시 읽었더니 그래프와 어긋난 망이 "
                               f"{len(broken)}개 있습니다 — 솔버에 그대로 넣지 마세요.",
                }] if broken else []
                sess.write("emit.json", {"schema": "fncadnet.design_emit/1",
                                         "bundles": bundles, "flags": flags})
                sess.audit(actor, "EMIT", "files_emitted", {
                    "networks": len(bundles),
                    "files": [f["name"] for b in bundles for f in b["files"]],
                    "round_trip_ok": not broken,
                })
        except S.StageBusy as exc:
            return _fail("STAGE_BUSY", str(exc), 409)

        return jsonify({"ok": True, "bundles": bundles, "flags": flags})

    # [문서정합] §11 에 내려받기 라우트가 없다. 방출은 했는데 가져갈 길이 없으면
    # 산출물이 서버 안에만 쌓이므로 여기 둔다. 이름은 `emit.json` 에 적힌 것만 준다.
    @app.get("/api/design/emit/<sid>/<name>")
    def design_emit_file(sid, name):
        sess, _draft, err = _gated(sid)
        if err:
            return err
        data, _v = sess.read_or_none("emit.json")
        allowed = {f["name"] for bundle in (data or {}).get("bundles") or []
                   for f in bundle["files"]}
        if name not in allowed:
            return _fail("EMIT_FILE_NOT_FOUND", "그런 방출 파일이 없습니다.", 404)
        return send_from_directory(sess.path(_EMIT_DIR), name, as_attachment=True)
