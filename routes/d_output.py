# -*- coding: utf-8 -*-
"""모듈 D 출력 워크벤치 — 결과 파일을 올리고 표시 항목을 골라 PDF 를 받는다.

`대조 서버.py` 에서 `register(app, ...)` 로 등록. 계산은 하지 않는다 — D1~D6 이
읽고 그린 것을 화면에 붙일 뿐이다.

**한 번 읽고 여러 번 그린다.** 지시서 2.2 는 "표시 항목 변경이 재계산 없이 즉시
반영된다" 를 수용 기준으로 둔다. 그래서 올린 파일은 잡에 얹어 두고 파싱한 모델을
붙잡아, 항목을 바꿔 다시 그릴 때 SDF/XML 을 다시 읽지 않는다.

**세트 하나씩이다.** 폴더째 던지는 일괄 처리는 D6 의 CLI(`python -m core.d_batch`)
가 맡는다 — 지시서 D6 이 "웹 UI 는 마지막, 그 전까지는 CLI 로 충분하다" 고 못박았고,
브라우저로 340 세트를 올리는 것보다 폴더를 걸어 두는 편이 실제로 빠르다.
"""
from __future__ import annotations

import shutil
import time
import uuid
from datetime import datetime
from pathlib import Path

from flask import jsonify, render_template, request

from core.d_batch import DEFAULT_PRESETS, FileSet, process_set
from core.d_display_model import load_display_model
from core.d_iso_renderer import LINK_ITEMS, NODE_ITEMS, PRESETS, render_iso
from core.d_report_typeset import render_report
from core.d_result_binder import bind_results

_JOBS: dict[str, dict] = {}
_LINK_NAMES = frozenset(i.name for i in LINK_ITEMS)
_NODE_NAMES = frozenset(i.name for i in NODE_ITEMS)
_NONE_ITEM = "None"


def _job(job_id: str) -> dict:
    found = _JOBS.get(job_id)
    if found is None:
        raise ValueError("올린 파일이 만료되었습니다. 다시 올려 주세요.")
    return found


def _join_summary(bound) -> dict:
    """결합에서 한쪽에만 있던 라벨 — 개수만 세지 않고 전량 싣는다."""
    report = bound.report
    return {"matched": report.matched, "summary": report.summary(),
            "sdf_only_nodes": list(report.sdf_only_nodes),
            "xml_only_nodes": list(report.xml_only_nodes),
            "sdf_only_pipes": list(report.sdf_only_pipes),
            "xml_only_pipes": list(report.xml_only_pipes),
            "sdf_only_nozzles": list(report.sdf_only_nozzles),
            "xml_only_nozzles": list(report.xml_only_nozzles),
            "divergent_pipes": list(report.divergent_pipes),
            "duplicate_labels": list(report.duplicate_labels)}


def register(app, *, D_OUTPUT_DIR, _err500, _register_job, _save_upload, _serve_run_file,
             _sweep_old_run_dirs):

    @app.get("/module-d-output")
    def module_d_output_page():
        # 캐시 금지 헤더는 게이트 대상 경로 전체에 after_request 가 이미 박는다.
        return render_template("d_output.html")

    @app.post("/api/module-d/load")
    def module_d_load():
        """SDF(+결과 XML) 를 올려 잡을 연다. 여기서 한 번만 읽는다."""
        try:
            # 잡 축출은 메모리만 비운다 — 남은 잡 폴더는 여기서 쓸어낸다.
            _sweep_old_run_dirs(D_OUTPUT_DIR)
            sdf_upload = _save_upload("sdf_file", {".sdf"}, True)
            xml_upload = _save_upload("xml_file", {".xml"}, False)

            job_id = uuid.uuid4().hex[:16]
            job_dir = D_OUTPUT_DIR / job_id
            job_dir.mkdir(parents=True, exist_ok=True)
            sdf = Path(shutil.move(str(sdf_upload), job_dir / sdf_upload.name))
            xml = None
            if xml_upload is not None:
                # 결과 XML 은 이름이 달라도 된다 — 짝은 사용자가 지은 것이다.
                xml = Path(shutil.move(str(xml_upload), job_dir / xml_upload.name))

            model = load_display_model(sdf)
            bound, join, notes = None, None, list(model.warnings)
            if xml is None:
                notes.append("결과 XML 이 없다 — 계산서를 만들 수 없고 도면은 "
                             "SDF 만으로 되는 항목(관경·길이·라벨)만 값이 붙는다")
            else:
                bound = bind_results(model, xml)
                join = _join_summary(bound)
                notes.extend(bound.warnings)

            _register_job(_JOBS, job_id, {"dir": job_dir, "sdf": sdf, "xml": xml,
                                          "model": model, "bound": bound})
            return jsonify({
                "ok": True, "job_id": job_id, "stem": sdf.stem,
                "title": list(bound.title) if bound else [],
                "counts": {"nodes": len(model.real_nodes), "pipes": len(model.pipes),
                           "nozzles": len(model.nozzles), "devices": len(model.devices),
                           "texts": len(model.texts)},
                "has_results": bound is not None,
                "link_items": [i.name for i in LINK_ITEMS],
                "node_items": [i.name for i in NODE_ITEMS],
                "presets": {name: {"link": p.link, "node": p.node}
                            for name, p in PRESETS.items()},
                "saved": {"link": model.link_scheme, "node": model.node_scheme},
                "join": join, "notes": notes})
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)

    @app.post("/api/module-d/draw")
    def module_d_draw():
        """표시 항목을 골라 ISO 한 장. 파일은 다시 읽지 않는다."""
        try:
            body = request.get_json(silent=True) or {}
            job = _job(str(body.get("job_id", "")))
            link_item = str(body.get("link_item") or _NONE_ITEM)
            node_item = str(body.get("node_item") or _NONE_ITEM)
            if link_item not in _LINK_NAMES:
                return jsonify({"ok": False,
                                "message": f"관로 표시 항목이 아닙니다: {link_item}"}), 400
            if node_item not in _NODE_NAMES:
                return jsonify({"ok": False,
                                "message": f"노드 표시 항목이 아닙니다: {node_item}"}), 400

            started = time.perf_counter()
            stem = f"iso_{uuid.uuid4().hex[:8]}"
            drawn = render_iso(
                job["bound"] or job["model"], job["dir"] / f"{stem}.pdf",
                preset=None, link_item=link_item, node_item=node_item,
                show_labels=bool(body.get("show_labels", True)),
                show_arrows=bool(body.get("show_arrows", True)),
                section=str(body.get("section", "")),
                also_png=job["dir"] / f"{stem}.png")
            elapsed = time.perf_counter() - started

            return jsonify({
                "ok": True, "pdf": f"{stem}.pdf", "png": f"{stem}.png",
                "seconds": round(elapsed, 2),
                "orientation": drawn.orientation,
                "drawn": {"pipes": drawn.pipes_drawn, "nozzles": drawn.nozzles_drawn,
                          "devices": drawn.devices_drawn, "labels": drawn.labels_drawn},
                "blank_links": list(drawn.blank_link_values),
                "blank_nodes": list(drawn.blank_node_values),
                "unplaced_pipes": list(drawn.pipes_unplaced),
                "undrawn_pipes": list(drawn.undrawn_result_pipes),
                "undrawn_nodes": list(drawn.undrawn_result_nodes),
                "dropped_labels": list(drawn.labels_dropped),
                "warnings": list(drawn.warnings)})
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)

    @app.post("/api/module-d/report")
    def module_d_report():
        """수리계산서 한 부. 결과 XML 이 있어야 한다."""
        try:
            body = request.get_json(silent=True) or {}
            job = _job(str(body.get("job_id", "")))
            if job["xml"] is None:
                return jsonify({"ok": False,
                                "message": "결과 XML 이 없어 계산서를 만들 수 없습니다."}), 400

            name = f"report_{uuid.uuid4().hex[:8]}.pdf"
            written = render_report(job["sdf"], job["xml"], job["dir"] / name,
                                    generated=datetime.now())
            return jsonify({"ok": True, "pdf": name, "pages": written.pages,
                            "rows": written.rows, "font_pt": written.font_pt,
                            "velocity_flags": list(written.velocity_flags),
                            "nozzle_flags": list(written.nozzle_flags),
                            "warnings": list(written.warnings)})
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)

    @app.post("/api/module-d/book")
    def module_d_book():
        """계산서 + ISO 프리셋들을 한 권으로 — D6 과 같은 경로를 탄다."""
        try:
            body = request.get_json(silent=True) or {}
            job = _job(str(body.get("job_id", "")))
            asked = [str(p) for p in (body.get("presets") or DEFAULT_PRESETS)]

            dest = job["dir"] / "book"
            shutil.rmtree(dest, ignore_errors=True)
            result = process_set(FileSet(stem=job["sdf"].stem, sdf=job["sdf"],
                                         xml=job["xml"]),
                                 dest, presets=asked, generated=datetime.now())
            if result.error:
                return jsonify({"ok": False, "message": result.error}), 400

            payload = {
                "pdf": result.combined.relative_to(job["dir"]).as_posix()
                if result.combined else None,
                "line": result.line(), "notes": list(result.notes),
                "parts": [{"name": p.name, "ok": p.ok, "pages": p.pages,
                           "error": p.error, "seconds": round(p.seconds, 2),
                           "warnings": list(p.warnings)} for p in result.parts]}
            if payload["pdf"] is None:
                # 조각이 몇 장 나왔든 합본이 없으면 "한 권" 은 실패다.
                return jsonify({**payload, "ok": False,
                                "message": "합본을 만들지 못했습니다."}), 400
            return jsonify({**payload, "ok": True})
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)

    @app.get("/api/module-d/result/<job_id>/<path:filename>")
    def module_d_result(job_id, filename):
        return _serve_run_file(D_OUTPUT_DIR, job_id, filename)
