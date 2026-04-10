from __future__ import annotations

import base64
import json
from datetime import datetime
from io import BytesIO
from pathlib import Path
import math
import xml.etree.ElementTree as ET
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from flask import Flask, jsonify, make_response, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename

from pipenet_validator import PipenetGuideValidator


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "data" / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
UPDATE_HISTORY_PATH = BASE_DIR / "data" / "update_history.json"

app = Flask(__name__)

# Keep chart text strictly ASCII-safe to prevent tofu/square glyphs on some systems.
plt.rcParams["font.family"] = ["DejaVu Sans", "Arial", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False


EXPORT_SCHEMA = {
    "pipes": {
        "sheet": "배관",
        "columns": [
            ("label", "Pipe"),
            ("input_node", "입력 노드"),
            ("output_node", "출력 노드"),
            ("nominal_bore_mm", "구경(mm)"),
            ("flow_lpm", "유량(L/min)"),
            ("velocity_mps", "유속(m/s)"),
            ("inlet_pressure", "입구압"),
            ("outlet_pressure", "출구압"),
            ("friction_loss", "마찰손실"),
            ("special_equipment", "특수설비"),
        ],
    },
    "nozzles": {
        "sheet": "헤드",
        "columns": [
            ("label", "헤드"),
            ("input_node", "입력 노드"),
            ("inlet_pressure_kgf_cm2", "압력(kg/cm²G)"),
            ("required_flow_lpm", "요구 유량"),
            ("actual_flow_lpm", "실제 유량"),
            ("deviation_percent", "편차(%)"),
        ],
    },
    "equipment": {
        "sheet": "특수설비",
        "columns": [
            ("label", "설비"),
            ("pipe_label", "배관"),
            ("description", "구분"),
            ("equivalent_length_m", "등가길이(m)"),
        ],
    },
    "valves": {
        "sheet": "감압밸브",
        "columns": [
            ("label", "밸브"),
            ("inlet_pressure_kgf_cm2", "입구압"),
            ("outlet_pressure_kgf_cm2", "출구압"),
            ("pressure_drop_kgf_cm2", "압력강하"),
            ("flow_lpm", "유량(L/min)"),
        ],
    },
}


def _save_upload(field_name: str, allowed_suffixes: set[str], required: bool) -> Path | None:
    uploaded = request.files.get(field_name)
    if uploaded is None or not uploaded.filename:
        if required:
            raise ValueError(f"`{field_name}` 파일이 필요합니다.")
        return None

    original_name = Path(uploaded.filename).name
    original_suffix = Path(original_name).suffix.lower()
    filename = secure_filename(original_name)
    if not filename:
        filename = f"{field_name}_{int(datetime.now().timestamp())}{original_suffix}"
    elif Path(filename).suffix == "" and original_suffix:
        filename = f"{filename}{original_suffix}"
    suffix = original_suffix or Path(filename).suffix.lower()
    if suffix not in allowed_suffixes:
        allowed = ", ".join(sorted(allowed_suffixes))
        raise ValueError(f"`{field_name}` 파일 형식이 올바르지 않습니다. 허용 형식: {allowed}")

    saved_path = UPLOAD_DIR / filename
    uploaded.save(saved_path)
    return saved_path


def _to_float(v, default: float = 0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def _fig_to_data_url(fig) -> str:
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return "data:image/png;base64," + base64.b64encode(buf.read()).decode("ascii")


def _load_update_history() -> dict:
    if not UPDATE_HISTORY_PATH.exists():
        return {
            "title": "업데이트 기록",
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "items": [],
        }
    with UPDATE_HISTORY_PATH.open("r", encoding="utf-8") as fp:
        payload = json.load(fp)
    payload.setdefault("title", "업데이트 기록")
    payload.setdefault("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M"))
    payload.setdefault("items", [])
    payload["items"] = sorted(
        payload["items"],
        key=lambda item: str(item.get("timestamp") or item.get("date") or ""),
        reverse=True,
    )
    return payload


def _build_visualizations(validation: dict, report_path: Path, sdf_path: Path | None) -> list[dict]:
    tables = validation.get("tables") or {}
    visuals: list[dict] = []

    # 1) Pipe velocity vs limit
    pipe_rows = tables.get("pipes") or []
    pipe_labels: list[str] = []
    velocities: list[float] = []
    limits: list[float] = []
    for r in pipe_rows:
        label = str(r.get("label", ""))
        vel = _to_float(r.get("velocity_mps"), 0.0)
        pipe_labels.append(label)
        velocities.append(vel)
        limits.append(_to_float(r.get("velocity_limit_mps"), 0.0))
    if pipe_labels:
        fig, ax = plt.subplots(figsize=(10, 3.8))
        x = list(range(len(pipe_labels)))
        ax.plot(x, velocities, marker="o", linewidth=1.6, color="#1d4ed8", label="Velocity")
        ax.plot(x, limits, linestyle="--", linewidth=1.2, color="#dc2626", label="Limit")
        ax.set_title("Pipe Velocity vs Limit")
        ax.set_xlabel("Pipe Label")
        ax.set_ylabel("m/s")
        if len(pipe_labels) <= 30:
            ax.set_xticks(x, pipe_labels, rotation=0)
        else:
            step = max(1, len(pipe_labels) // 20)
            ticks = x[::step]
            ax.set_xticks(ticks, [pipe_labels[i] for i in ticks], rotation=0)
        ax.grid(alpha=0.25)
        ax.legend()
        visuals.append(
            {
                "title": "Pipe Velocity Check",
                "description": "Compares each pipe velocity with topology-based branch/other limits from the validator.",
                "image_data_url": _fig_to_data_url(fig),
            }
        )

    # 2) Nozzle pressure-flow scatter
    noz_rows = tables.get("nozzles") or []
    pressures = [_to_float(r.get("inlet_pressure_kgf_cm2"), 0.0) for r in noz_rows]
    flows = [_to_float(r.get("actual_flow_lpm"), 0.0) for r in noz_rows]
    if pressures and flows:
        fig, ax = plt.subplots(figsize=(6.2, 4.2))
        colors = ["#dc2626" if _to_float(r.get("actual_flow_lpm"), 0.0) < 80.0 else "#16a34a" for r in noz_rows]
        ax.scatter(pressures, flows, c=colors, alpha=0.85)
        ax.axhline(80.0, color="#dc2626", linestyle="--", linewidth=1.2, label="80 L/min")
        ax.axvline(1.0, color="#f59e0b", linestyle="--", linewidth=1.2, label="1.0 kg/cm^2G")
        ax.set_title("Nozzle Pressure-Flow Distribution")
        ax.set_xlabel("Inlet Pressure (kg/cm^2G)")
        ax.set_ylabel("Actual Flow (L/min)")
        ax.grid(alpha=0.25)
        ax.legend(loc="best")
        visuals.append(
            {
                "title": "Nozzle Pressure-Flow",
                "description": "Green points pass the flow threshold, red points are below the flow threshold.",
                "image_data_url": _fig_to_data_url(fig),
            }
        )

    return visuals


def _point_on_polyline(path: list[tuple[float, float]], ratio: float) -> tuple[float, float] | None:
    if len(path) < 2:
        return path[0] if path else None
    ratio = max(0.0, min(1.0, ratio))
    seg_lengths: list[float] = []
    total = 0.0
    for i in range(len(path) - 1):
        x1, y1 = path[i]
        x2, y2 = path[i + 1]
        d = math.hypot(x2 - x1, y2 - y1)
        seg_lengths.append(d)
        total += d
    if total <= 0:
        return path[0]
    target = total * ratio
    acc = 0.0
    for i, d in enumerate(seg_lengths):
        if acc + d >= target:
            t = (target - acc) / d if d > 0 else 0.0
            x1, y1 = path[i]
            x2, y2 = path[i + 1]
            return (x1 + (x2 - x1) * t, y1 + (y2 - y1) * t)
        acc += d
    return path[-1]


def _build_sdf_graph(sdf_path: Path | None, tables: dict | None) -> dict:
    if sdf_path is None or not sdf_path.exists():
        return {}

    tables = tables or {}
    pipe_table = {int(r.get("label")): r for r in tables.get("pipes", []) if str(r.get("label", "")).isdigit()}
    nozzle_table = {int(r.get("label")): r for r in tables.get("nozzles", []) if str(r.get("label", "")).isdigit()}
    equipment_table = {int(r.get("label")): r for r in tables.get("equipment", []) if str(r.get("label", "")).isdigit()}
    valve_table = {int(r.get("label")): r for r in tables.get("valves", []) if str(r.get("label", "")).isdigit()}

    root = ET.parse(sdf_path).getroot()

    node_pos: dict[str, tuple[float, float]] = {}
    for node in root.findall(".//Node"):
        label = node.attrib.get("label")
        pos = node.find("Position")
        if not label or pos is None:
            continue
        try:
            x = float(pos.attrib.get("x", "0"))
            y = float(pos.attrib.get("y", "0"))
        except ValueError:
            continue
        node_pos[label] = (x, y)

    pipes: list[dict] = []
    pipe_paths: dict[int, list[tuple[float, float]]] = {}
    for pipe in root.findall(".//Pipe"):
        label_raw = pipe.attrib.get("label", "")
        if not label_raw.isdigit():
            continue
        label = int(label_raw)
        input_node = pipe.attrib.get("input", "")
        output_node = pipe.attrib.get("output", "")

        path: list[tuple[float, float]] = []
        if input_node in node_pos:
            path.append(node_pos[input_node])
        waypoints = pipe.find("Waypoints")
        if waypoints is not None:
            for wp in waypoints.findall("Position"):
                try:
                    path.append((float(wp.attrib.get("x", "0")), float(wp.attrib.get("y", "0"))))
                except ValueError:
                    continue
        if output_node in node_pos:
            path.append(node_pos[output_node])
        if len(path) < 2:
            continue

        pipe_paths[label] = path
        trow = pipe_table.get(label, {})
        status = "fail" if trow.get("highlight") else "pass"
        pipes.append(
            {
                "label": label,
                "input_node": input_node,
                "output_node": output_node,
                "path": [[x, y] for x, y in path],
                "status": status,
            }
        )

    nozzles: list[dict] = []
    for nozzle in root.findall(".//Nozzle"):
        label_raw = nozzle.attrib.get("label", "")
        input_node = nozzle.attrib.get("input", "")
        if not label_raw.isdigit() or input_node not in node_pos:
            continue
        label = int(label_raw)
        x, y = node_pos[input_node]
        trow = nozzle_table.get(label, {})
        status = "fail" if trow.get("highlight") else "pass"
        nozzles.append({"label": label, "input_node": input_node, "x": x, "y": y, "status": status})

    equipment: list[dict] = []
    equipment_pos_by_label: dict[int, tuple[float, float]] = {}
    for eq in root.findall(".//Equipment"):
        label_raw = eq.attrib.get("label", "")
        if not label_raw.isdigit():
            continue
        label = int(label_raw)
        rel = float(eq.attrib.get("rel-position", "0.5"))
        desc = eq.attrib.get("description", "")
        table_row = equipment_table.get(label, {})
        pipe_label = table_row.get("pipe_label")
        if isinstance(pipe_label, str) and pipe_label.isdigit():
            pipe_label = int(pipe_label)
        if not isinstance(pipe_label, int):
            continue
        path = pipe_paths.get(pipe_label)
        if not path:
            continue
        p = _point_on_polyline(path, rel)
        if p is None:
            continue
        x, y = p
        equipment_pos_by_label[label] = (x, y)
        status = "warn" if table_row.get("warn") else ("fail" if table_row.get("highlight") else "pass")
        equipment.append(
            {
                "label": label,
                "description": desc,
                "pipe_label": pipe_label,
                "x": x,
                "y": y,
                "status": status,
            }
        )

    valves: list[dict] = []
    for label, row in valve_table.items():
        pos = equipment_pos_by_label.get(label)
        if pos is None:
            continue
        x, y = pos
        status = "fail" if row.get("highlight") else "pass"
        valves.append({"label": label, "x": x, "y": y, "status": status})

    return {
        "nodes": [{"id": nid, "x": x, "y": y} for nid, (x, y) in node_pos.items()],
        "pipes": pipes,
        "nozzles": nozzles,
        "equipment": equipment,
        "valves": valves,
    }


def _sdf_counts_only(sdf_path: Path | None) -> dict:
    if sdf_path is None or not sdf_path.exists():
        return {}
    root = ET.parse(sdf_path).getroot()
    return {
        "pipes": len(root.findall(".//Pipe")),
        "nozzles": len(root.findall(".//Nozzle")),
        "equipment": len(root.findall(".//Equipment")),
    }


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/cad-compare-module")
def cad_compare_module():
    response = make_response(render_template("cad_compare_module.html"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/api/update-history")
def update_history():
    return jsonify({"ok": True, "history": _load_update_history()})


@app.post("/api/cad-module/dxf-parse")
def cad_module_dxf_parse():
    try:
        cad_path = _save_upload("cad_file", {".dxf", ".dwg"}, required=True)
        if cad_path.suffix.lower() == ".dwg":
            raise ValueError("현재 6번 모듈은 DXF만 지원합니다. DWG는 DXF로 변환 후 업로드해 주세요.")

        from cad_engine import DXFWorkspace

        workspace = DXFWorkspace(UPLOAD_DIR / "cad_workspace")
        workspace.load_file(cad_path)
        payload = workspace.to_payload(include_network_entities=False, include_network_summary=False)
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "message": f"DXF 파싱 중 오류가 발생했습니다: {exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "message": "DXF 파싱이 완료되었습니다.",
            "cad_payload": {
                "filename": payload.get("filename"),
                "bounds": payload.get("bounds"),
                "layers": payload.get("layers"),
                "entities": payload.get("entities") or [],
                "unsupported": payload.get("unsupported") or {},
            },
        }
    )


@app.post("/api/validate")
def validate_files():
    try:
        report_path = _save_upload("report_file", {".docx", ".pdf"}, required=True)
        sdf_path = _save_upload("sdf_file", {".sdf"}, required=False)
        validation = PipenetGuideValidator(report_path=report_path, sdf_path=sdf_path).validate()
        sdf_graph = _build_sdf_graph(sdf_path, validation.get("tables"))
        visualizations = _build_visualizations(validation, report_path, sdf_path)
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "message": f"검증 중 오류가 발생했습니다: {exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "message": "검증이 완료되었습니다.",
            "filename": validation["report_name"],
            "sdf_filename": validation["sdf_name"],
            "summary": validation["summary"],
            "results": validation["results"],
            "insights": validation["insights"],
            "rules": validation.get("rules", {}),
            "stats": validation["stats"],
            "visualizations": visualizations,
            "tables": validation["tables"],
            "sdf_graph": sdf_graph,
            "report": validation["report"],
        }
    )


@app.post("/api/cad-compare")
def cad_compare():
    try:
        cad_path = _save_upload("cad_file", {".dxf", ".dwg"}, required=True)
        sdf_path = _save_upload("sdf_file", {".sdf"}, required=False)
        if cad_path.suffix.lower() == ".dwg":
            raise ValueError("현재 CAD 대조 모듈은 DXF만 지원합니다. DWG는 DXF로 변환 후 업로드해 주세요.")

        from cad_engine import DXFWorkspace

        workspace = DXFWorkspace(UPLOAD_DIR / "cad_workspace")
        workspace.load_file(cad_path)
        payload = workspace.to_payload(include_network_entities=True, include_network_summary=True)

        network_layers = set(payload.get("networkLayers") or [])
        network_entity_ids = set(payload.get("networkEntityIds") or [])
        entities = payload.get("entities") or []
        if network_entity_ids:
            entities = [e for e in entities if e.get("id") in network_entity_ids]
        if network_layers:
            entities = [e for e in entities if e.get("layer") in network_layers]

        head_boxes: list[dict] = []
        detector_mode = "template"
        use_yolo = str(request.form.get("use_yolo", "")).strip() == "1"
        try:
            if use_yolo:
                from head_detector import TriangleHeadDetector

                model_path = BASE_DIR / "yolo26n.pt"
                if not model_path.exists():
                    model_path = BASE_DIR / "yolo11n.pt"
                detector = TriangleHeadDetector(BASE_DIR / "data" / "head_templates", model_path)
                head_boxes = detector.detect(entities, payload.get("bounds") or {}, network_layers)
                detector_mode = "yolo+template" if detector.yolo_detector.available else "template"
            else:
                from head_detector import TriangleHeadTemplateDetector

                detector = TriangleHeadTemplateDetector(BASE_DIR / "data" / "head_templates")
                head_boxes = detector.detect(entities, payload.get("bounds") or {}, network_layers)
                detector_mode = "template"
        except Exception:
            detector_mode = "unavailable"
            head_boxes = []

        cad_counts = {
            "entities": len(entities),
            "network_layers": len(network_layers),
            "detected_heads": len(head_boxes),
            "lines": sum(1 for e in entities if e.get("type") in {"LINE", "LWPOLYLINE", "ARC"}),
            "circles": sum(1 for e in entities if e.get("type") == "CIRCLE"),
            "texts": sum(1 for e in entities if e.get("type") == "TEXT"),
        }
        sdf_counts = _sdf_counts_only(sdf_path)
        messages: list[str] = []
        if sdf_counts:
            sdf_heads = int(sdf_counts.get("nozzles", 0))
            diff = cad_counts["detected_heads"] - sdf_heads
            if diff == 0:
                messages.append(f"헤드 수 일치: CAD 탐지 {cad_counts['detected_heads']} / SDF {sdf_heads}")
            else:
                messages.append(
                    f"헤드 수 차이: CAD 탐지 {cad_counts['detected_heads']} / SDF {sdf_heads} (차이 {diff:+d})"
                )
            messages.append(
                f"SDF 수량: 배관 {sdf_counts.get('pipes', 0)} / 헤드 {sdf_counts.get('nozzles', 0)} / 특수설비 {sdf_counts.get('equipment', 0)}"
            )
        else:
            messages.append("SDF 미업로드 상태입니다. CAD 단독 추출/탐지 결과만 표시합니다.")
        messages.append(f"탐지 엔진: {detector_mode}")

    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "message": f"CAD 대조 중 오류가 발생했습니다: {exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "message": "CAD 대조가 완료되었습니다.",
            "cad_filename": cad_path.name,
            "sdf_filename": sdf_path.name if sdf_path else None,
            "cad_payload": {
                "filename": payload.get("filename"),
                "bounds": payload.get("bounds"),
                "layers": payload.get("layers"),
                "networkLayers": list(network_layers),
                "entities": entities,
            },
            "detected_heads": head_boxes,
            "cad_counts": cad_counts,
            "sdf_counts": sdf_counts,
            "messages": messages,
        }
    )


def _apply_sheet_style(ws):
    thin = Side(style="thin", color="C9CDD3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="F3F4F6")
    title_fill = PatternFill("solid", fgColor="111827")
    fail_fill = PatternFill("solid", fgColor="FEE2E2")
    warn_fill = PatternFill("solid", fgColor="FEF3C7")
    eng_fill = PatternFill("solid", fgColor="DBEAFE")
    econ_fill = PatternFill("solid", fgColor="DCFCE7")

    max_col = ws.max_column
    max_row = ws.max_row

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    title_cell = ws.cell(row=1, column=1)
    title_cell.font = Font(color="FFFFFF", bold=True, size=12)
    title_cell.fill = title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in range(1, max_col + 1):
        c = ws.cell(row=3, column=col)
        c.fill = header_fill
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    for row in range(4, max_row + 1):
        flag_fail = ws.cell(row=row, column=max_col - 3).value == "Y"
        flag_warn = ws.cell(row=row, column=max_col - 2).value == "Y"
        flag_eng = ws.cell(row=row, column=max_col - 1).value == "Y"
        flag_econ = ws.cell(row=row, column=max_col).value == "Y"

        row_fill = None
        if flag_fail:
            row_fill = fail_fill
        elif flag_warn:
            row_fill = warn_fill
        elif flag_eng:
            row_fill = eng_fill
        elif flag_econ:
            row_fill = econ_fill

        for col in range(1, max_col + 1):
            c = ws.cell(row=row, column=col)
            c.border = border
            c.alignment = Alignment(horizontal="center", vertical="center")
            if row_fill:
                c.fill = row_fill

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{ws.cell(row=3, column=max_col).column_letter}{max_row}"

    for col in range(1, max_col + 1):
        width = 12
        for row in range(1, max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is None:
                continue
            width = max(width, min(len(str(val)) + 2, 40))
        ws.column_dimensions[get_column_letter(col)].width = width


@app.post("/api/export-xlsx")
def export_xlsx():
    try:
        payload = request.get_json(force=True)
        tables = payload.get("tables") or {}
        report_name = payload.get("report_name") or "pipenet_result"
    except Exception:
        return jsonify({"ok": False, "message": "엑셀 내보내기 요청 형식이 올바르지 않습니다."}), 400

    wb = Workbook()
    wb.remove(wb.active)

    for key, meta in EXPORT_SCHEMA.items():
        ws = wb.create_sheet(title=meta["sheet"])
        rows = tables.get(key, [])

        ws.cell(row=1, column=1, value=f"{meta['sheet']} 결과 데이터")
        ws.cell(row=2, column=1, value="빨강=기준 위반, 노랑=확인 필요, 파랑=공학 후보, 초록=경제성 후보")

        headers = [label for _, label in meta["columns"]] + ["기준위반", "확인필요", "공학후보", "경제후보"]
        for col, header in enumerate(headers, start=1):
            ws.cell(row=3, column=col, value=header)

        for idx, row in enumerate(rows, start=4):
            values = [row.get(k, "") for k, _ in meta["columns"]]
            values += [
                "Y" if row.get("highlight") else "N",
                "Y" if row.get("warn") else "N",
                "Y" if row.get("engineering_flag") else "N",
                "Y" if row.get("economy_flag") else "N",
            ]
            for col, v in enumerate(values, start=1):
                ws.cell(row=idx, column=col, value=v)

        _apply_sheet_style(ws)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = Path(report_name).stem
    out_name = f"{safe_name}_결과테이블_{timestamp}.xlsx"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)
