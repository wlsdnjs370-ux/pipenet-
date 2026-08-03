# -*- coding: utf-8 -*-
"""Remote 30 도면 검사(inspect) 도메인 — DXF 업로드 → 레이어/엔티티 뷰 캐시.

`대조 서버.py` 에서 `register(app, ...)` 로 등록. 캐시 경로/버전 전역과
_save_upload 를 주입받는다. _inspect_cache_paths 는 주입 전역을 참조하므로
register 안에 중첩. 라우트 본문·엔드포인트명 원본 그대로.
"""
from __future__ import annotations

import gzip
import hashlib
import json
import math
import os
from pathlib import Path

from flask import Response, jsonify


def _inspect_layer_visibility(doc):
    """DXF 레이어 가시성 수집 → (doc_layer_info, hidden_layers).

    hidden_layers = CAD 가 화면에 안 그리는 레이어(off/frozen/color<0). plot=0(비출력)은
    화면엔 그대로 보이므로(내진·치수·배치도 등) hidden 에 넣지 않고 no_plot 으로만 참고 보관.
    """
    doc_layer_info: dict[str, dict] = {}
    hidden_layers: set[str] = set()
    try:
        for ly in doc.layers:
            try:
                color = int(ly.dxf.color)
            except Exception:
                color = 7
            name = str(ly.dxf.name)
            is_off = bool(ly.is_off())
            is_frozen = bool(ly.is_frozen())
            try:
                no_plot = int(getattr(ly.dxf, "plot", 1)) == 0
            except Exception:
                no_plot = False
            doc_layer_info[name] = {
                "is_off": is_off,
                "is_frozen": is_frozen,
                "is_locked": bool(ly.is_locked()),
                "color": color,
                "no_plot": no_plot,
            }
            if is_off or is_frozen or color < 0:
                hidden_layers.add(name)
    except Exception:
        pass
    return doc_layer_info, hidden_layers


def register(app, *, INSPECT_CACHE_DIR, INSPECT_CACHE_VERSION, _save_upload):

    def _inspect_cache_paths(dxf_path: Path):
        """도면 내용 해시 기반 inspect 렌더 캐시 경로 → (entities_gz, meta_json).

        동일 도면 재업로드(=내용 해시 일치)면 렌더 결과를 그대로 스트리밍하기 위함.
        해시 실패 시 (None, None) — 캐시 비활성으로 동작.
        """
        try:
            h = hashlib.sha256()
            with open(dxf_path, "rb") as _f:
                for _blk in iter(lambda: _f.read(1024 * 1024), b""):
                    h.update(_blk)
            content_hash = h.hexdigest()
        except Exception:
            content_hash = ""
        if not content_hash:
            return None, None
        cache_key = f"{INSPECT_CACHE_VERSION}_{content_hash}"
        return (INSPECT_CACHE_DIR / f"{cache_key}.entities.ndjson.gz",
                INSPECT_CACHE_DIR / f"{cache_key}.meta.json")

    @app.post("/api/remote30/inspect")
    def remote30_inspect():
        """DXF 업로드 → 모든 entity JSON + 레이어 통계 + 카테고리 자동 추천."""
        try:
            dxf_path = _save_upload("dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400

        dxf_name = dxf_path.name

        # ── 바이너리 캐시 조회 — 동일 도면 재업로드면 렌더 결과를 그대로 스트리밍.
        cache_ent_path, cache_meta_path = _inspect_cache_paths(dxf_path)

        if cache_ent_path and cache_ent_path.exists() and cache_meta_path.exists():
            try:
                meta = json.loads(cache_meta_path.read_text(encoding="utf-8"))
            except Exception:
                meta = None
            if meta is not None:
                def _stream_cached():
                    with gzip.open(cache_ent_path, "rt", encoding="utf-8") as gz:
                        for line in gz:
                            if line:
                                yield line
                    yield json.dumps({
                        "type": "result",
                        "ok": True,
                        "dxf_filename": dxf_name,
                        "dxf_token": dxf_name,
                        "bbox": meta.get("bbox"),
                        "layers": meta.get("layers", []),
                        "counts": meta.get("counts", {}),
                        "dropped_types": meta.get("dropped_types", {}),
                        "bg_truncated": meta.get("bg_truncated", False),
                        "bg_entities": meta.get("bg_entities", 0),
                        "bg_budget": meta.get("bg_budget", 0),
                        "cached": True,
                    }, ensure_ascii=False) + "\n"
                return Response(_stream_cached(), mimetype="application/x-ndjson")

        try:
            import ezdxf
            from sprinkler_remote30_extractor import Remote30Settings, layer_match
        except ImportError as exc:
            return jsonify({"ok": False, "message": f"의존성 누락: {exc}"}), 500

        try:
            doc = ezdxf.readfile(str(dxf_path))
            msp = doc.modelspace()
        except Exception as exc:
            return jsonify({"ok": False, "message": f"DXF 파싱 실패: {exc}"}), 500

        # DXF 레이어 가시성 (off/frozen/color<0 만 hidden, plot=0 은 화면엔 보이므로 렌더).
        doc_layer_info, hidden_layers = _inspect_layer_visibility(doc)

        entities = []
        bbox = [float("inf"), float("inf"), float("-inf"), float("-inf")]

        def _upd(x, y):
            if x < bbox[0]:
                bbox[0] = x
            if y < bbox[1]:
                bbox[1] = y
            if x > bbox[2]:
                bbox[2] = x
            if y > bbox[3]:
                bbox[3] = y

        dropped_types: dict[str, int] = {}
        MAX_INSERT_DEPTH = 10  # cycle 방지용 — 정상 도면은 3~4 단계면 충분
        # 배경(건축/제외) 렌더 상한. 옛 구현은 예산(12만) 초과 시 배경을 **통째** 생략해
        # 141MB LH 지하층배관도(배경 28만 = 9개 XREF INSERT)의 건축 레이어가 화면에서
        # 통으로 사라졌다 — 기계실 모드에서 수원/연결점을 찍으려면 그 배경이 필요하다.
        # 지금은 상한을 실제 렌더 수 기준으로만 걸어(병적 fan-out 방어) 도달 전까지는
        # 전부 그린다. 콜드 파싱 비용(이 도면 기준 +40초)은 gzip 캐시로 파일당 1회.
        BG_ENTITY_BUDGET = 500_000

        # 레이어 카테고리 사전 계산(캐시) — 배경(건축/제외) 레이어 판별용.
        _cat_settings = Remote30Settings()
        _cat_cache: dict[str, str] = {}

        def _layer_category(name: str) -> str:
            cached = _cat_cache.get(name)
            if cached is not None:
                return cached
            # 콘텐츠(HEAD/PIPE/TEXT) 신호가 ARCH 를 이긴다: "SHEET-TEXT" 처럼 arch 키워드가
            # 섞인 콘텐츠 레이어를 건축으로 흡수(정리)해 버리는 오류 방지. EXCLUDE 만 최우선.
            if layer_match(name, _cat_settings.exclude_layer_keywords):
                cat = "EXCLUDE"
            elif layer_match(name, _cat_settings.head_layer_keywords):
                cat = "HEAD"
            elif layer_match(name, _cat_settings.pipe_layer_keywords):
                cat = "PIPE"
            elif layer_match(name, _cat_settings.text_layer_keywords):
                cat = "TEXT"
            elif layer_match(name, _cat_settings.arch_layer_keywords):
                cat = "ARCH"
            else:
                cat = "OTHER"
            _cat_cache[name] = cat
            return cat

        # ezdxf 의 virtual_entities() 가 xscale=-1 mirror INSERT 의 자식 좌표를
        # 잘못 계산하는 버그가 있어, AutoCAD 표준 매트릭스를 직접 빌드해 적용한다.
        # world = M · local,  M = T(insert) · R(rot_z) · S(sx,sy,sz) · T(-base)
        from ezdxf.math import Matrix44, Vec3  # noqa: PLC0415

        def _insert_matrix(insert_entity):
            ix = float(insert_entity.dxf.insert.x)
            iy = float(insert_entity.dxf.insert.y)
            try:
                iz = float(insert_entity.dxf.insert.z)
            except Exception:
                iz = 0.0
            sx = float(getattr(insert_entity.dxf, "xscale", 1.0) or 1.0)
            sy = float(getattr(insert_entity.dxf, "yscale", 1.0) or 1.0)
            sz = float(getattr(insert_entity.dxf, "zscale", 1.0) or 1.0)
            rot_rad = math.radians(float(getattr(insert_entity.dxf, "rotation", 0.0) or 0.0))
            block_name = str(insert_entity.dxf.name)
            block = insert_entity.doc.blocks.get(block_name) if insert_entity.doc else None
            if block is not None:
                try:
                    bx = float(block.base_point.x); by = float(block.base_point.y)
                    bz = float(block.base_point.z) if hasattr(block.base_point, "z") else 0.0
                except Exception:
                    bx = by = bz = 0.0
            else:
                bx = by = bz = 0.0
            # ezdxf chain(A, B, C) — A 먼저 적용 후 B, C 순. 즉 result = C @ B @ A.
            return Matrix44.chain(
                Matrix44.translate(-bx, -by, -bz),
                Matrix44.scale(sx, sy, sz),
                Matrix44.z_rotate(rot_rad),
                Matrix44.translate(ix, iy, iz),
            )

        def _t(matrix, x, y):
            """matrix 가 None 이면 (x, y) 그대로, 아니면 변환 좌표 반환."""
            if matrix is None:
                return float(x), float(y)
            v = matrix.transform(Vec3(float(x), float(y), 0.0))
            return float(v.x), float(v.y)

        # ARC/CIRCLE 반지름 스케일 추정은 remote30_prototype 의 단일 헬퍼를 공유한다.
        from remote30_prototype import _uniform_scale

        def _render_entity(e, *, matrix=None, layer_override: str | None = None, depth: int = 0) -> None:
            """Convert one ezdxf entity to canvas dict(s) and append to entities[].

            INSERT 는 다이아몬드 마커 + virtual_entities 폭발(자식 LINE/CIRCLE/HATCH/...) 까지 함께 렌더.
            중첩 INSERT 도 깊이에 상관없이 재귀 폭발 (MAX_INSERT_DEPTH 가드).
            layer_override 가 있으면 자식이 "0"(BYLAYER) 일 때 부모 INSERT 의 레이어로 대체.
            CAD 화면에 안 보이는 것은 그대로 안 보내도록 다음을 스킵:
              - effective layer 가 hidden_layers (off/frozen/color<0) 에 속한 경우
              - entity 자체의 invisible flag 가 1인 경우
            """
            etype = e.dxftype()
            own_layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
            # BYLAYER 의미: 블록 내부 "0" 레이어는 부모 INSERT 의 레이어를 따른다
            if layer_override is not None and own_layer in ("0", ""):
                layer = layer_override
            else:
                layer = own_layer or (layer_override or "")
            # CAD parity — 숨김 레이어 또는 invisible flag 면 캔버스에 보내지 않음
            if layer in hidden_layers:
                return
            if int(getattr(e.dxf, "invisible", 0) or 0) == 1:
                return
            try:
                if etype == "LINE":
                    x1, y1 = _t(matrix, e.dxf.start.x, e.dxf.start.y)
                    x2, y2 = _t(matrix, e.dxf.end.x, e.dxf.end.y)
                    entities.append({"t": "L", "l": layer, "p": [x1, y1, x2, y2]})
                    _upd(x1, y1); _upd(x2, y2)
                elif etype == "ARC":
                    cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                    r = float(e.dxf.radius) * _uniform_scale(matrix)
                    sa = float(e.dxf.start_angle)
                    ea = float(e.dxf.end_angle)
                    entities.append({"t": "A", "l": layer, "c": [cx, cy], "r": r, "a": [sa, ea]})
                    _upd(cx - r, cy - r); _upd(cx + r, cy + r)
                elif etype == "CIRCLE":
                    cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                    r = float(e.dxf.radius) * _uniform_scale(matrix)
                    entities.append({"t": "C", "l": layer, "c": [cx, cy], "r": r})
                    _upd(cx - r, cy - r); _upd(cx + r, cy + r)
                elif etype == "LWPOLYLINE":
                    pts = [list(_t(matrix, p[0], p[1])) for p in e.get_points()]
                    if pts:
                        for x, y in pts:
                            _upd(x, y)
                        entities.append({"t": "PL", "l": layer, "p": pts})
                elif etype == "POLYLINE":
                    pts = [list(_t(matrix, v.dxf.location.x, v.dxf.location.y)) for v in e.vertices]
                    if pts:
                        for x, y in pts:
                            _upd(x, y)
                        entities.append({"t": "PL", "l": layer, "p": pts})
                elif etype == "INSERT":
                    # 다이아몬드 마커: 최상위 (depth==0) 일 때만, INSERT 위치 (matrix 적용)
                    ix_w, iy_w = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                    if depth == 0:
                        entities.append({"t": "I", "l": layer, "p": [ix_w, iy_w], "n": str(e.dxf.name)})
                    _upd(ix_w, iy_w)
                    # ARCH/EXCLUDE 레이어 블록도 폭발해 건축 배경(건물 외곽선 등)을
                    # 실제 CAD 도면과 동일하게 렌더한다. (이전엔 속도 위해 생략했으나
                    # 화면 누락 문제로 복원) 마커(다이아몬드)도 유지.
                    # AutoCAD 표준 INSERT 매트릭스 빌드 + 부모 매트릭스와 결합
                    if depth >= MAX_INSERT_DEPTH:
                        dropped_types["INSERT(too deep)"] = dropped_types.get("INSERT(too deep)", 0) + 1
                    else:
                        try:
                            my_matrix = _insert_matrix(e)
                        except Exception:
                            my_matrix = None
                        if matrix is not None and my_matrix is not None:
                            # combined: child local → world = matrix @ my_matrix @ local
                            combined = Matrix44.chain(my_matrix, matrix)
                        elif my_matrix is not None:
                            combined = my_matrix
                        else:
                            combined = matrix
                        # 블록 정의의 entity 들을 직접 순회 (virtual_entities() 의 mirror 버그 우회)
                        block = e.doc.blocks.get(e.dxf.name) if e.doc else None
                        if block is not None:
                            for child in block:
                                _render_entity(child, matrix=combined, layer_override=layer, depth=depth + 1)
                elif etype == "TEXT":
                    x, y = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                    raw = str(e.dxf.text)[:60]
                    entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
                    _upd(x, y)
                elif etype in ("MTEXT", "ATTRIB", "ATTDEF"):
                    x, y = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                    raw = str(getattr(e, "text", "") or getattr(e.dxf, "text", ""))[:60]
                    if raw:
                        entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
                    _upd(x, y)
                elif etype == "SPLINE":
                    try:
                        pts = [list(_t(matrix, pt[0], pt[1])) for pt in e.flattening(1.0)]
                    except Exception:
                        pts = []
                    if pts:
                        for x, y in pts:
                            _upd(x, y)
                        entities.append({"t": "PL", "l": layer, "p": pts})
                elif etype == "ELLIPSE":
                    try:
                        pts = [list(_t(matrix, pt[0], pt[1])) for pt in e.flattening(0.5)]
                    except Exception:
                        pts = []
                    if pts:
                        for x, y in pts:
                            _upd(x, y)
                        entities.append({"t": "PL", "l": layer, "p": pts})
                elif etype == "HATCH":
                    paths_out = []
                    for path in e.paths:
                        pts = []
                        # 1) PolylinePath — vertices 직접 사용
                        for vertex in getattr(path, "vertices", []) or []:
                            try:
                                x, y = _t(matrix, vertex[0], vertex[1])
                                pts.append([x, y])
                            except Exception:
                                continue
                        # 2) EdgePath — LineEdge / ArcEdge / EllipseEdge / SplineEdge 들의 정점 추출
                        if not pts:
                            for edge in getattr(path, "edges", []) or []:
                                edge_type = type(edge).__name__
                                try:
                                    if edge_type == "LineEdge":
                                        x1, y1 = _t(matrix, edge.start[0], edge.start[1])
                                        x2, y2 = _t(matrix, edge.end[0], edge.end[1])
                                        pts.append([x1, y1]); pts.append([x2, y2])
                                    elif edge_type == "ArcEdge":
                                        cx, cy = float(edge.center[0]), float(edge.center[1])
                                        r = float(edge.radius)
                                        sa = float(edge.start_angle); ea = float(edge.end_angle)
                                        if ea < sa: ea += 360.0
                                        for k in range(9):
                                            ang = math.radians(sa + (ea - sa) * k / 8)
                                            x, y = _t(matrix, cx + r * math.cos(ang), cy + r * math.sin(ang))
                                            pts.append([x, y])
                                    elif edge_type in ("EllipseEdge", "SplineEdge"):
                                        for attr in ("start", "start_point", "control_points"):
                                            v = getattr(edge, attr, None)
                                            if v is None: continue
                                            try:
                                                x, y = _t(matrix, v[0], v[1])
                                                pts.append([x, y])
                                                break
                                            except Exception:
                                                try:
                                                    x, y = _t(matrix, v[0][0], v[0][1])
                                                    pts.append([x, y])
                                                    break
                                                except Exception:
                                                    continue
                                except Exception:
                                    continue
                        if len(pts) > 1:
                            pts = [pts[0]] + [p for prev, p in zip(pts, pts[1:]) if p != prev]
                        if pts:
                            paths_out.append(pts)
                            for x, y in pts:
                                _upd(x, y)
                    if paths_out:
                        biggest = max(paths_out, key=len)
                        entities.append({"t": "H", "l": layer, "p": biggest})
                    else:
                        dropped_types["HATCH(no-geom)"] = dropped_types.get("HATCH(no-geom)", 0) + 1
                elif etype in ("SOLID", "3DFACE", "TRACE"):
                    verts = []
                    for attr in ("vtx0", "vtx1", "vtx2", "vtx3"):
                        try:
                            v = getattr(e.dxf, attr)
                            x, y = _t(matrix, v.x, v.y)
                            verts.append([x, y])
                        except AttributeError:
                            break
                    if len(verts) >= 2 and verts[-1] == verts[-2]:
                        verts.pop()
                    if len(verts) >= 3:
                        for x, y in verts:
                            _upd(x, y)
                        entities.append({"t": "S", "l": layer, "p": verts})
                elif etype == "DIMENSION":
                    # 치수선 본체는 자식 entity 들로 explode 되어 렌더 (matrix 그대로 전달)
                    try:
                        for v in e.virtual_entities():
                            _render_entity(v, matrix=matrix, layer_override=layer)
                    except Exception:
                        pass
                else:
                    dropped_types[etype] = dropped_types.get(etype, 0) + 1
            except Exception:
                dropped_types[etype] = dropped_types.get(etype, 0) + 1

        # ── 점진적 렌더링 (NDJSON 스트리밍) ─────────────────────────────────
        # 배관/헤드 등 전경(foreground) top-level entity 를 먼저 렌더·전송해 사용자가
        # 즉시 작업을 시작하게 하고, 건축 배경(ARCH/EXCLUDE)은 이어서 스트리밍으로 채운다.
        # 화면 정보는 하나도 누락하지 않으며 첫 페인트까지의 체감 시간만 줄인다.
        foreground_top, background_top = [], []
        for e in msp:
            try:
                _lyr = e.dxf.layer if hasattr(e.dxf, "layer") else ""
            except Exception:
                _lyr = ""
            if _layer_category(str(_lyr)) in ("ARCH", "EXCLUDE"):
                background_top.append(e)
            else:
                foreground_top.append(e)

        FLUSH_N = 20000
        layer_counts: dict[str, int] = {}
        layer_type_counts: dict[str, dict] = {}
        total_count = [0]
        bg_status = {"truncated": False, "entities": 0}
        # 진행률 분모 — 방출되는 entity 수는 INSERT 폭발 때문에 미리 알 수 없지만
        # top-level entity 수는 파싱 직후 확정된다. 클라이언트 진행바가 추측이 아닌
        # 실제 소화량을 표시하려면 이 (done, total) 쌍이 필요하다.
        total_top = len(foreground_top) + len(background_top)
        done_top = [0]

        def _bbox_obj():
            if bbox[0] == float("inf"):
                return {"x_min": 0.0, "y_min": 0.0, "x_max": 1.0, "y_max": 1.0}
            return {"x_min": bbox[0], "y_min": bbox[1], "x_max": bbox[2], "y_max": bbox[3]}

        def _emit(phase: str, with_bbox: bool):
            """현재 entities 버퍼를 FLUSH_N 단위로 yield 하며 레이어 통계 누적 후 비운다."""
            n = len(entities)
            i = 0
            while i < n:
                chunk = entities[i:i + FLUSH_N]
                for ent in chunk:
                    l = ent["l"]
                    layer_counts[l] = layer_counts.get(l, 0) + 1
                    tc = layer_type_counts.setdefault(l, {})
                    tc[ent["t"]] = tc.get(ent["t"], 0) + 1
                total_count[0] += len(chunk)
                msg = {"type": "progress", "phase": phase, "entities": chunk,
                       "done": done_top[0], "total": total_top}
                if with_bbox and i == 0:
                    msg["bbox"] = _bbox_obj()
                yield json.dumps(msg, ensure_ascii=False) + "\n"
                i += FLUSH_N
            del entities[:]

        def _progress_lines():
            # 1) 전경 — 점진적 flush. 첫 chunk 은 빠른 첫 페인트를 위해 작게(FIRST_FLUSH_N),
            #    이후는 FLUSH_N 단위. 이전엔 전경 전부 렌더 후 flush 라 cold 파싱 시
            #    첫 byte 까지 100초+ 무 페인트(체감 무한 로딩) 였다. bbox 는 누적되므로
            #    매 flush 의 첫 chunk 가 "지금까지" 범위를 실어 보낸다(클라가 점진 fit).
            FIRST_FLUSH_N = 2000
            fg_flushed = False
            # 파싱(ezdxf.readfile)까지는 진행 상황을 알릴 방법이 없다. 그 구간이
            # 끝난 이 시점을 먼저 알려 클라이언트가 "몇 % 인지 모름" 표시를 끝내고
            # 렌더 진행률로 넘어가게 한다 — 첫 chunk 는 2000 entity 뒤에나 온다.
            yield json.dumps({"type": "phase", "phase": "render",
                              "total": total_top}, ensure_ascii=False) + "\n"
            for e in foreground_top:
                _render_entity(e)
                done_top[0] += 1
                thresh = FIRST_FLUSH_N if not fg_flushed else FLUSH_N
                if len(entities) >= thresh:
                    yield from _emit("foreground", with_bbox=True)
                    fg_flushed = True
            if entities:
                yield from _emit("foreground", with_bbox=True)
            # 2) 배경(건축/제외) — top-level 마다 렌더하며 점진 flush. 상한은 실제 렌더
            #    수로만 판단하고, 도달하면 그 지점에서 끊되 그린 만큼은 내보낸다.
            bg_start = total_count[0]
            for e in background_top:
                _render_entity(e)
                done_top[0] += 1
                if len(entities) >= FLUSH_N:
                    yield from _emit("background", with_bbox=False)
                if total_count[0] - bg_start >= BG_ENTITY_BUDGET:
                    bg_status["truncated"] = True
                    # 상한에 걸려 남은 top-level 을 건너뛴다 — 할 일이 끝난 건
                    # 맞으므로 분자를 채워 진행바가 중간에 멈춘 채 끝나지 않게 한다.
                    done_top[0] = total_top
                    break
            if entities:
                yield from _emit("background", with_bbox=False)
            bg_status["entities"] = total_count[0] - bg_start
            if bg_status["truncated"]:
                dropped_types[f"건축배경(상한 {BG_ENTITY_BUDGET:,} 도달, 이후 절단)"] = 1

        def _build_layer_list():
            layer_list = []
            for name in sorted(layer_counts.keys()):
                info = doc_layer_info.get(name, {})
                is_off = bool(info.get("is_off", False))
                is_frozen = bool(info.get("is_frozen", False))
                color = int(info.get("color", 7))
                layer_list.append({
                    "name": name,
                    "count": layer_counts[name],
                    "types": layer_type_counts.get(name, {}),
                    "auto_category": _layer_category(name),
                    "is_off": is_off,
                    "is_frozen": is_frozen,
                    "color": color,
                    "visible": (not is_off) and (not is_frozen) and (color >= 0),
                })
            return layer_list

        def _stream():
            # 캐시 미스 → 렌더하며 NDJSON 스트리밍하고, 동시에 gzip 캐시에 tee.
            tmp_ent = None
            gz_out = None
            committed = False
            if cache_ent_path is not None:
                tmp_ent = cache_ent_path.with_suffix(".tmp")
                try:
                    gz_out = gzip.open(tmp_ent, "wt", encoding="utf-8")
                except Exception:
                    gz_out = None
            try:
                for line in _progress_lines():
                    if gz_out is not None:
                        try:
                            gz_out.write(line)
                        except Exception:
                            pass
                    yield line
                # 최종 result — 레이어 통계 / 전체 bbox / 카운트
                layer_list = _build_layer_list()
                bbox_obj = _bbox_obj()
                counts_obj = {"total_entities": total_count[0], "layers": len(layer_counts)}
                yield json.dumps({
                    "type": "result",
                    "ok": True,
                    "dxf_filename": dxf_name,
                    "dxf_token": dxf_name,  # extract 재호출 시 DXF 재업로드 생략 토큰
                    "bbox": bbox_obj,
                    "layers": layer_list,
                    "counts": counts_obj,
                    "dropped_types": dropped_types,
                    "bg_truncated": bg_status["truncated"],
                    "bg_entities": bg_status["entities"],
                    "bg_budget": BG_ENTITY_BUDGET,
                }, ensure_ascii=False) + "\n"
                # 스트림 정상 완료 시에만 캐시 commit (부분/중단 시 미저장)
                if gz_out is not None:
                    gz_out.close()
                    gz_out = None
                    try:
                        meta = {
                            "bbox": bbox_obj,
                            "layers": layer_list,
                            "counts": counts_obj,
                            "dropped_types": dropped_types,
                            "bg_truncated": bg_status["truncated"],
                            "bg_entities": bg_status["entities"],
                            "bg_budget": BG_ENTITY_BUDGET,
                        }
                        tmp_meta = cache_meta_path.with_suffix(".tmp")
                        tmp_meta.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
                        os.replace(tmp_ent, cache_ent_path)
                        os.replace(tmp_meta, cache_meta_path)
                        committed = True
                    except Exception:
                        committed = False
            finally:
                if gz_out is not None:
                    try:
                        gz_out.close()
                    except Exception:
                        pass
                if not committed and tmp_ent is not None and tmp_ent.exists():
                    try:
                        tmp_ent.unlink()
                    except Exception:
                        pass

        return Response(_stream(), mimetype="application/x-ndjson")
