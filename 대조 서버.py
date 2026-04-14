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
FEEDBACK_POSTS_PATH = BASE_DIR / "data" / "feedback_posts.json"
FEEDBACK_UPLOAD_DIR = BASE_DIR / "data" / "feedback_uploads"
FEEDBACK_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

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


def _fig_to_data_url(fig, *, tight: bool = True) -> str:
    buf = BytesIO()
    save_kwargs = {"format": "png", "dpi": 140}
    if tight:
        save_kwargs["bbox_inches"] = "tight"
    fig.savefig(buf, **save_kwargs)
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


def _load_feedback_posts() -> list[dict]:
    if not FEEDBACK_POSTS_PATH.exists():
        return []
    try:
        with FEEDBACK_POSTS_PATH.open("r", encoding="utf-8") as fp:
            payload = json.load(fp)
    except (json.JSONDecodeError, OSError):
        return []
    if isinstance(payload, dict):
        posts = payload.get("posts", [])
    else:
        posts = payload
    if not isinstance(posts, list):
        return []
    return sorted(posts, key=lambda item: str(item.get("created_at", "")), reverse=True)


def _save_feedback_posts(posts: list[dict]) -> None:
    FEEDBACK_POSTS_PATH.parent.mkdir(parents=True, exist_ok=True)
    ordered = sorted(posts, key=lambda item: str(item.get("created_at", "")), reverse=True)
    with FEEDBACK_POSTS_PATH.open("w", encoding="utf-8") as fp:
        json.dump({"posts": ordered}, fp, ensure_ascii=False, indent=2)


def _clean_feedback_text(value: object, limit: int) -> str:
    text = str(value or "").strip()
    text = " ".join(text.split()) if limit <= 80 else text
    return text[:limit]


def _save_feedback_attachment(post_id: str) -> dict | None:
    uploaded = request.files.get("attachment")
    if uploaded is None or not uploaded.filename:
        return None
    original_name = Path(uploaded.filename).name
    safe_name = secure_filename(original_name)
    if not safe_name:
        safe_name = f"attachment_{post_id}"
    saved_name = f"{post_id}_{safe_name}"
    saved_path = FEEDBACK_UPLOAD_DIR / saved_name
    uploaded.save(saved_path)
    return {
        "original_name": original_name,
        "stored_name": saved_name,
        "size": saved_path.stat().st_size if saved_path.exists() else 0,
        "download_url": f"/api/feedback-attachments/{saved_name}",
    }


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


def _build_engineering_visualizations(validation: dict) -> list[dict]:
    tables = validation.get("tables") or {}
    pipe_rows = tables.get("pipes") or []
    ratio_rows: list[dict] = []

    for row in pipe_rows:
        label = row.get("label")
        friction_loss = _to_float(row.get("friction_loss"), 0.0)
        length_m = _to_float(row.get("base_length_m") or row.get("pipe_length_m"), 0.0)
        if label is None or length_m <= 0:
            continue
        ratio = friction_loss / length_m
        ratio_rows.append(
            {
                "label": int(label),
                "ratio": ratio,
                "friction_loss": friction_loss,
                "length_m": length_m,
                "velocity_mps": _to_float(row.get("velocity_mps"), 0.0),
                "velocity_limit_mps": _to_float(row.get("velocity_limit_mps"), 0.0),
                "nominal_bore_mm": _to_float(row.get("nominal_bore_mm"), 0.0),
                "fitting_eq_length_m": _to_float(row.get("fitting_eq_length_m"), 0.0),
                "special_eq_length_m": _to_float(row.get("special_eq_length_m"), 0.0),
                "total_length_m": _to_float(row.get("total_length_m"), length_m),
                "engineering_flag": bool(row.get("engineering_flag")),
            }
        )

    if not ratio_rows:
        return []

    ratio_rows.sort(key=lambda item: item["label"])
    labels = [str(item["label"]) for item in ratio_rows]
    values = [item["ratio"] for item in ratio_rows]
    colors = ["#2563eb" if item["engineering_flag"] else "#9ca3af" for item in ratio_rows]
    threshold = 1.0
    lengths = sorted(item["length_m"] for item in ratio_rows)
    median_length = lengths[len(lengths) // 2] if lengths else 0.0
    max_value = max(values) if values else threshold
    y_max = max(max_value * 1.12, threshold * 1.8)
    spike_points: list[dict] = []

    for idx, item in enumerate(ratio_rows):
        if idx == 0:
            continue
        previous = ratio_rows[idx - 1]
        prev_ratio = previous["ratio"]
        ratio = item["ratio"]
        delta = ratio - prev_ratio
        change_rate = delta / max(prev_ratio, 1e-9)
        spike_delta_limit = max(threshold, abs(prev_ratio) * 0.75)
        if ratio <= threshold or delta <= spike_delta_limit:
            continue

        eq_length = item["fitting_eq_length_m"] + item["special_eq_length_m"]
        eq_share = eq_length / max(item["total_length_m"], 1e-9)
        velocity_limit = item["velocity_limit_mps"]
        velocity_ratio = item["velocity_mps"] / velocity_limit if velocity_limit > 0 else 0.0
        causes: list[str] = []
        actions: list[str] = []

        if median_length > 0 and item["length_m"] >= median_length * 1.5:
            causes.append("긴 배관이라 총 마찰손실이 커질 수 있는 구간입니다.")
            actions.append("배관 경로를 단순화하거나 우회 길이를 줄여 실제 배관길이를 단축하는 방안을 검토하세요.")
        if median_length > 0 and item["length_m"] <= median_length * 0.6 and ratio > threshold:
            causes.append("짧은 배관인데 m당 마찰손실이 높아 손실이 국부적으로 집중된 구간입니다.")
            actions.append("해당 짧은 구간의 급격한 방향 전환, 국부 피팅, 특수설비 연결부를 우선 점검하세요.")
        if velocity_ratio >= 0.85:
            causes.append("유속이 적용 기준에 근접하여 마찰손실 증가에 크게 기여할 수 있습니다.")
            actions.append("구경 상향 또는 유량 분산으로 유속을 낮추는 대안을 검토하세요.")
        if eq_share >= 0.35:
            causes.append("피팅/특수설비 등가길이 비중이 커서 배관 자체 길이보다 부속 손실 영향이 큽니다.")
            actions.append("엘보/티/밸브/후렉시블 배관 수량을 줄이거나 손실이 작은 부속으로 변경하는 방안을 검토하세요.")
        if item["nominal_bore_mm"] <= 50 and item["velocity_mps"] >= 5.0:
            causes.append("소구경 배관에서 비교적 높은 유속이 발생하여 손실 집중 가능성이 있습니다.")
            actions.append("50A 이하 구간은 가지배관 기준 6m/s에 근접하는지 확인하고, 필요 시 한 단계 큰 구경을 검토하세요.")
        if not causes:
            causes.append("직전 배관 대비 m당 마찰손실 증가율이 커서 국부 조건 변화가 의심됩니다.")
            actions.append("해당 배관 전후의 구경 변화, 유속 변화, 피팅 수량, 특수설비 연결 여부를 함께 확인하세요.")

        spike_points.append(
            {
                "label": item["label"],
                "previous_label": previous["label"],
                "ratio": ratio,
                "previous_ratio": prev_ratio,
                "delta": delta,
                "change_rate_percent": change_rate * 100.0,
                "data_index": idx,
                "left_percent": 50.0,
                "top_percent": 50.0,
                "cards": {
                    "criteria": [
                        "직전 배관 대비 m당 마찰손실 변화율이 큰 구간을 급증 후보로 표시합니다.",
                        f"급증 조건: 현재 비율 > {threshold:.3f} kg/cm^2/m AND 증가량 > max({threshold:.3f}, 직전 비율 x 75%)",
                    ],
                    "formula": [
                        "m당 마찰손실 = FLOW IN PIPES Frict. Loss / PIPE CONFIGURATION Length",
                        "증가량 = 현재 m당 마찰손실 - 직전 배관 m당 마찰손실",
                        "변화율 = 증가량 / max(직전 배관 m당 마찰손실, 1e-9)",
                    ],
                    "values": [
                        f"Pipe {item['label']}: {item['friction_loss']:.4f} / {item['length_m']:.3f} = {ratio:.4f} kg/cm^2/m",
                        f"Previous Pipe {previous['label']}: {previous['friction_loss']:.4f} / {previous['length_m']:.3f} = {prev_ratio:.4f} kg/cm^2/m",
                        f"증가량 = {delta:.4f} kg/cm^2/m, 변화율 = {change_rate * 100.0:.1f}%",
                        f"구경 = {item['nominal_bore_mm']:.0f}A, 유속 = {item['velocity_mps']:.3f} m/s, 피팅+특수설비 등가길이 비중 = {eq_share * 100.0:.1f}%",
                    ],
                    "conclusion": [
                        *causes,
                        *actions,
                    ],
                },
            }
        )

    fig, ax = plt.subplots(figsize=(11, 4.2))
    x = list(range(len(labels)))
    ax.bar(x, values, color=colors, width=0.78)
    ax.axhline(threshold, color="#dc2626", linestyle="--", linewidth=1.2, label="Threshold 1.00")
    if spike_points:
        spike_x = [labels.index(str(point["label"])) for point in spike_points if str(point["label"]) in labels]
        spike_y = [point["ratio"] for point in spike_points if str(point["label"]) in labels]
        ax.scatter(
            spike_x,
            spike_y,
            marker="v",
            s=26,
            color="#dc2626",
            edgecolor="#7f1d1d",
            linewidth=0.6,
            zorder=5,
            label="Sharp Increase",
        )
    ax.set_title("Friction Loss Ratio by Pipe")
    ax.set_xlabel("Pipe Label")
    ax.set_ylabel("Friction Loss / Length (kg/cm^2/m)")
    ax.set_ylim(0, y_max)
    if len(labels) <= 35:
        ax.set_xticks(x, labels, rotation=0)
    else:
        step = max(1, len(labels) // 24)
        ticks = x[::step]
        ax.set_xticks(ticks, [labels[i] for i in ticks], rotation=0)
    ax.grid(axis="y", alpha=0.25)
    ax.legend(loc="upper right")
    fig.canvas.draw()
    fig_w, fig_h = fig.canvas.get_width_height()
    if fig_w > 0 and fig_h > 0:
        for point in spike_points:
            px, py = ax.transData.transform((point["data_index"], point["ratio"]))
            point["left_percent"] = max(0.0, min(100.0, (px / fig_w) * 100.0))
            point["top_percent"] = max(0.0, min(100.0, ((fig_h - py) / fig_h) * 100.0))

    return [
        {
            "title": "Friction Loss Ratio by Pipe",
            "description": "Shows friction_loss / base_length_m for each pipe. Red markers indicate sharp increases from the previous pipe.",
            "image_data_url": _fig_to_data_url(fig, tight=False),
            "spike_points": spike_points,
        }
    ]


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


def _analyze_sdf_sprinkler_network(sdf_path: Path) -> dict:
    root = ET.parse(sdf_path).getroot()

    titles = [t.text.strip() for t in root.findall(".//Title") if t.text and t.text.strip()]
    nodes: dict[str, dict] = {}
    for node in root.findall(".//Node"):
        label = node.attrib.get("label", "")
        pos = node.find("Position")
        if not label or pos is None:
            continue
        nodes[label] = {
            "id": label,
            "x": _to_float(pos.attrib.get("x")),
            "y": _to_float(pos.attrib.get("y")),
            "z": _to_float(node.attrib.get("elevation")),
        }

    pipes: list[dict] = []
    equipment: list[dict] = []
    material = "UNKNOWN"
    for pipe_set in root.findall(".//Pipe-set"):
        pipe_type = pipe_set.find("Pipe-type")
        name = pipe_type.find("Name") if pipe_type is not None else None
        if name is not None and name.text:
            material = name.text.strip()
        for pipe in pipe_set.findall("Pipe"):
            label = pipe.attrib.get("label", "")
            input_node = pipe.attrib.get("input", "")
            output_node = pipe.attrib.get("output", "")
            bore_mm = _to_float(pipe.attrib.get("bore")) * 1000.0
            length_m = _to_float(pipe.attrib.get("length"))
            rise_m = _to_float(pipe.attrib.get("rise"))
            c_factor = _to_float(pipe.attrib.get("roughness-or-c"))
            fittings: list[dict] = []
            for fitting in pipe.findall(".//Fitting"):
                fittings.append(
                    {
                        "type": fitting.attrib.get("type", ""),
                        "count": int(_to_float(fitting.attrib.get("count"), 0)),
                    }
                )
            path_nodes = [input_node]
            waypoint_positions: list[dict] = []
            waypoints = pipe.find("Waypoints")
            if waypoints is not None:
                for wp in waypoints.findall("Position"):
                    waypoint_positions.append(
                        {
                            "x": _to_float(wp.attrib.get("x")),
                            "y": _to_float(wp.attrib.get("y")),
                        }
                    )
            pipe_row = {
                "label": label,
                "input_node": input_node,
                "output_node": output_node,
                "bore_mm": bore_mm,
                "length_m": length_m,
                "rise_m": rise_m,
                "c_factor": c_factor,
                "material": material,
                "fittings": fittings,
                "fitting_summary": ", ".join(f"{f['type']}({f['count']})" for f in fittings) or "-",
                "waypoints": waypoint_positions,
            }
            pipes.append(pipe_row)
            for eq in pipe.findall(".//Equipment"):
                equipment.append(
                    {
                        "label": eq.attrib.get("label", ""),
                        "pipe_label": label,
                        "description": eq.attrib.get("description", ""),
                        "equivalent_length_m": _to_float(eq.attrib.get("equivalent-length")),
                        "rel_position": _to_float(eq.attrib.get("rel-position"), 0.5),
                    }
                )

    nozzles: list[dict] = []
    for nozzle in root.findall(".//Nozzle"):
        label = nozzle.attrib.get("label", "")
        input_node = nozzle.attrib.get("input", "")
        node = nodes.get(input_node, {})
        nozzles.append(
            {
                "label": label,
                "input_node": input_node,
                "x": node.get("x"),
                "y": node.get("y"),
                "z": node.get("z"),
            }
        )

    pipe_by_label = {p["label"]: p for p in pipes}
    outgoing: dict[str, list[dict]] = {}
    adjacency: dict[str, list[tuple[str, float, str]]] = {}
    for pipe in pipes:
        outgoing.setdefault(pipe["input_node"], []).append(pipe)
        adjacency.setdefault(pipe["input_node"], []).append((pipe["output_node"], pipe["length_m"], pipe["label"]))
        adjacency.setdefault(pipe["output_node"], []).append((pipe["input_node"], pipe["length_m"], pipe["label"]))

    av_equipment = next((e for e in equipment if (e.get("description") or "").upper().replace(" ", "") in {"A/V", "AV"}), None)
    av_node = ""
    av_pipe_label = ""
    if av_equipment:
        av_pipe_label = str(av_equipment.get("pipe_label") or "")
        av_pipe = pipe_by_label.get(av_pipe_label)
        if av_pipe:
            av_node = av_pipe.get("output_node") or av_pipe.get("input_node") or ""
    if not av_node and pipes:
        av_node = pipes[0]["input_node"]

    # Dijkstra distance from alarm valve anchor to each nozzle node.
    dist = {av_node: 0.0} if av_node else {}
    prev_pipe: dict[str, str] = {}
    visited: set[str] = set()
    while dist:
        current = min((n for n in dist if n not in visited), key=lambda n: dist[n], default=None)
        if current is None:
            break
        visited.add(current)
        for nxt, length, pipe_label in adjacency.get(current, []):
            nd = dist[current] + max(length, 0.0)
            if nxt not in dist or nd < dist[nxt]:
                dist[nxt] = nd
                prev_pipe[nxt] = pipe_label

    farthest_heads = sorted(
        [
            {
                **n,
                "distance_from_av_m": dist.get(str(n.get("input_node")), 0.0),
            }
            for n in nozzles
        ],
        key=lambda r: r.get("distance_from_av_m", 0.0),
        reverse=True,
    )[:30]

    length_checks: list[dict] = []
    for pipe in pipes:
        n1 = nodes.get(pipe["input_node"])
        n2 = nodes.get(pipe["output_node"])
        if not n1 or not n2:
            continue
        pts = [(n1["x"], n1["y"])]
        pts.extend((wp["x"], wp["y"]) for wp in pipe.get("waypoints") or [])
        pts.append((n2["x"], n2["y"]))
        xy_m = 0.0
        for i in range(len(pts) - 1):
            xy_m += math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1]) / 1000.0
        geom_m = math.hypot(xy_m, pipe.get("rise_m") or 0.0)
        diff_m = abs(geom_m - pipe["length_m"])
        tol_m = max(0.5, pipe["length_m"] * 0.05)
        if diff_m > tol_m:
            length_checks.append(
                {
                    "pipe_label": pipe["label"],
                    "sdf_length_m": round(pipe["length_m"], 3),
                    "xy_length_m": round(geom_m, 3),
                    "diff_m": round(diff_m, 3),
                    "reason": "SDF length와 XY 좌표거리 차이가 허용오차(5% 또는 0.5m)를 초과합니다.",
                }
            )

    bore_reductions: list[dict] = []
    for pipe in pipes:
        for child in outgoing.get(pipe["output_node"], []):
            if child["bore_mm"] and pipe["bore_mm"] and child["bore_mm"] < pipe["bore_mm"]:
                bore_reductions.append(
                    {
                        "from_pipe": pipe["label"],
                        "to_pipe": child["label"],
                        "node": pipe["output_node"],
                        "from_bore_mm": round(pipe["bore_mm"], 1),
                        "to_bore_mm": round(child["bore_mm"], 1),
                    }
                )

    node_degree: dict[str, int] = {}
    for pipe in pipes:
        node_degree[pipe["input_node"]] = node_degree.get(pipe["input_node"], 0) + 1
        node_degree[pipe["output_node"]] = node_degree.get(pipe["output_node"], 0) + 1
    branch_nodes = [
        {"node": node, "degree": degree, **nodes.get(node, {})}
        for node, degree in sorted(node_degree.items(), key=lambda x: (-x[1], x[0]))
        if degree >= 3
    ]

    fitting_summary: dict[str, int] = {}
    fitting_hotspots: list[dict] = []
    for pipe in pipes:
        total = 0
        for fitting in pipe["fittings"]:
            fitting_summary[fitting["type"]] = fitting_summary.get(fitting["type"], 0) + fitting["count"]
            total += fitting["count"]
        if total >= 2:
            fitting_hotspots.append(
                {
                    "pipe_label": pipe["label"],
                    "fitting_count": total,
                    "fittings": pipe["fitting_summary"],
                    "reason": "엘보/티 등 부속 집중 구간입니다. CAD 도면의 굴곡/분기 위치와 대조가 필요합니다.",
                }
            )

    vertical_pipes = [
        {
            "pipe_label": p["label"],
            "input_node": p["input_node"],
            "output_node": p["output_node"],
            "length_m": round(p["length_m"], 3),
            "rise_m": round(p["rise_m"], 3),
            "bore_mm": round(p["bore_mm"], 1),
        }
        for p in pipes
        if abs(p.get("rise_m") or 0.0) >= 3.0
    ]

    graph_pipes = []
    for p in pipes:
        n1 = nodes.get(p["input_node"])
        n2 = nodes.get(p["output_node"])
        if not n1 or not n2:
            continue
        path = [[n1["x"], n1["y"]]]
        path.extend([[wp["x"], wp["y"]] for wp in p.get("waypoints") or []])
        path.append([n2["x"], n2["y"]])
        status = "red" if any(x["pipe_label"] == p["label"] for x in length_checks) else "normal"
        if any(x["to_pipe"] == p["label"] or x["from_pipe"] == p["label"] for x in bore_reductions):
            status = "orange" if status == "normal" else status
        graph_pipes.append(
            {
                "label": p["label"],
                "input_node": p["input_node"],
                "output_node": p["output_node"],
                "bore_mm": round(p["bore_mm"], 1),
                "length_m": round(p["length_m"], 3),
                "material": p["material"],
                "status": status,
                "path": path,
            }
        )

    return {
        "title": " / ".join(titles) or sdf_path.name,
        "filename": sdf_path.name,
        "summary": {
            "node_count": len(nodes),
            "pipe_count": len(pipes),
            "nozzle_count": len(nozzles),
            "equipment_count": len(equipment),
            "av_node": av_node,
            "av_pipe_label": av_pipe_label,
            "length_issue_count": len(length_checks),
            "bore_reduction_count": len(bore_reductions),
            "branch_node_count": len(branch_nodes),
            "vertical_pipe_count": len(vertical_pipes),
        },
        "nodes": list(nodes.values()),
        "pipes": graph_pipes,
        "nozzles": nozzles,
        "equipment": equipment,
        "farthest_heads": farthest_heads,
        "length_checks": length_checks[:80],
        "bore_reductions": bore_reductions[:80],
        "branch_nodes": branch_nodes[:80],
        "fitting_summary": [{"type": k, "count": v} for k, v in sorted(fitting_summary.items())],
        "fitting_hotspots": fitting_hotspots[:80],
        "vertical_pipes": vertical_pipes[:80],
        "checklist": [
            "CAD 도면의 알람밸브 위치가 SDF A/V 추정 노드와 일치하는지 확인",
            "SDF 최원단 헤드 30개가 CAD 평면도상 검토 영역의 헤드 30개와 1:1 매칭되는지 확인",
            "배관 길이 불일치 후보는 CAD 실측 길이와 SDF length 값을 대조",
            "구경 축소 지점은 CAD 라벨의 관경 표기와 SDF bore 값을 대조",
            "엘보/티 집중 구간은 도면상 굴곡/분기 개수와 SDF Fittings count를 대조",
            "수직 배관은 건축 단면/층고와 SDF rise 및 length를 대조",
        ],
    }


def _extract_cad_head_candidates(cad_path: Path) -> dict:
    # Lightweight DXF scan for the 7th module. It extracts enough geometry for a
    # quick drawing preview without invoking the heavier CAD graph engine.
    try:
        raw = cad_path.read_text(encoding="utf-8", errors="ignore").splitlines()
    except UnicodeDecodeError:
        raw = cad_path.read_text(encoding="cp949", errors="ignore").splitlines()

    pairs: list[tuple[str, str]] = []
    for i in range(0, len(raw) - 1, 2):
        pairs.append((raw[i].strip(), raw[i + 1].strip()))

    entities: list[dict] = []
    in_entities = False
    current: dict | None = None

    def flush() -> None:
        nonlocal current
        if not current:
            return
        etype = current.get("type")
        if etype in {"CIRCLE", "INSERT"} and current.get("x") is not None and current.get("y") is not None:
            entities.append(current)
        elif etype == "LINE" and current.get("x") is not None and current.get("y") is not None and current.get("x2") is not None and current.get("y2") is not None:
            entities.append(current)
        elif etype in {"LWPOLYLINE", "POLYLINE"} and len(current.get("points") or []) >= 2:
            entities.append(current)
        current = None

    for code, value in pairs:
        if code == "0" and value == "SECTION":
            continue
        if code == "2" and value == "ENTITIES":
            in_entities = True
            continue
        if code == "0" and value == "ENDSEC":
            flush()
            in_entities = False
            continue
        if not in_entities:
            continue
        if code == "0":
            flush()
            if value in {"CIRCLE", "INSERT"}:
                current = {"type": value, "layer": "0", "x": None, "y": None, "radius": 0.0}
            elif value == "LINE":
                current = {"type": value, "layer": "0", "x": None, "y": None, "x2": None, "y2": None}
            elif value in {"LWPOLYLINE", "POLYLINE"}:
                current = {"type": value, "layer": "0", "points": [], "_pending_x": None}
            continue
        if current is None:
            continue
        if code == "8":
            current["layer"] = value
        elif code == "10":
            if current.get("type") in {"LWPOLYLINE", "POLYLINE"}:
                current["_pending_x"] = _to_float(value)
            else:
                current["x"] = _to_float(value)
        elif code == "20":
            if current.get("type") in {"LWPOLYLINE", "POLYLINE"}:
                px = current.pop("_pending_x", None)
                if px is not None:
                    current.setdefault("points", []).append([px, _to_float(value)])
            else:
                current["y"] = _to_float(value)
        elif code == "11":
            current["x2"] = _to_float(value)
        elif code == "21":
            current["y2"] = _to_float(value)
        elif code == "40":
            current["radius"] = _to_float(value)
        elif code == "2" and current.get("type") == "INSERT":
            current["block"] = value
    flush()

    circles = [e for e in entities if e.get("type") == "CIRCLE"]
    inserts = [e for e in entities if e.get("type") == "INSERT"]
    line_entities = [e for e in entities if e.get("type") in {"LINE", "LWPOLYLINE", "POLYLINE"}]
    source = circles if len(circles) >= 3 else inserts
    candidates: list[dict] = []
    for idx, ent in enumerate(source, start=1):
        candidates.append(
            {
                "label": str(idx),
                "entity_id": str(idx),
                "type": ent.get("type", ""),
                "layer": ent.get("layer", ""),
                "x": _to_float(ent.get("x")),
                "y": _to_float(ent.get("y")),
                "radius": _to_float(ent.get("radius")),
            }
        )

    if len(candidates) > 300:
        # Keep the comparison responsive. The user should filter the DXF to sprinkler/head layers for precision.
        candidates = candidates[:300]

    drawing_entities: list[dict] = []
    for idx, ent in enumerate(line_entities[:2500], start=1):
        if ent.get("type") == "LINE":
            drawing_entities.append(
                {
                    "id": f"L{idx}",
                    "type": "LINE",
                    "layer": ent.get("layer", ""),
                    "points": [[_to_float(ent.get("x")), _to_float(ent.get("y"))], [_to_float(ent.get("x2")), _to_float(ent.get("y2"))]],
                }
            )
        else:
            drawing_entities.append(
                {
                    "id": f"P{idx}",
                    "type": ent.get("type", "POLYLINE"),
                    "layer": ent.get("layer", ""),
                    "points": ent.get("points") or [],
                }
            )

    preview_points: list[dict] = []
    for ent in drawing_entities:
        for x, y in ent.get("points") or []:
            preview_points.append({"x": x, "y": y})
    preview_points.extend(candidates)

    return {
        "filename": cad_path.name,
        "bounds": _bbox(preview_points) or _bbox(candidates) or {},
        "layers": sorted({e.get("layer", "") for e in entities if e.get("layer")}),
        "network_layers": [],
        "raw_circle_count": len(circles),
        "raw_insert_count": len(inserts),
        "raw_line_count": len(line_entities),
        "candidate_count": len(candidates),
        "candidates": candidates[:500],
        "drawing_entities": drawing_entities,
    }


def _bbox(points: list[dict]) -> dict | None:
    valid = [p for p in points if p.get("x") is not None and p.get("y") is not None]
    if not valid:
        return None
    xs = [float(p["x"]) for p in valid]
    ys = [float(p["y"]) for p in valid]
    return {"min_x": min(xs), "max_x": max(xs), "min_y": min(ys), "max_y": max(ys)}


def _norm_point(p: dict, box: dict) -> tuple[float, float]:
    w = max(float(box["max_x"]) - float(box["min_x"]), 1e-9)
    h = max(float(box["max_y"]) - float(box["min_y"]), 1e-9)
    return ((float(p["x"]) - float(box["min_x"])) / w, (float(p["y"]) - float(box["min_y"])) / h)


def _compare_cad_heads_to_sdf(cad: dict, sdf: dict) -> dict:
    sdf_heads = sdf.get("farthest_heads") or sdf.get("nozzles") or []
    cad_heads = cad.get("candidates") or []
    sdf_box = _bbox(sdf_heads)
    cad_box = _bbox(cad_heads)
    if not sdf_box or not cad_box or not sdf_heads or not cad_heads:
        return {
            "status": "REVIEW",
            "message": "CAD 헤드 후보 또는 SDF 헤드 좌표가 부족하여 자동 비교를 보류했습니다.",
            "matches": [],
            "unmatched_sdf": sdf_heads,
            "unmatched_cad": cad_heads,
            "mismatch_count": 0,
        }

    unused = set(range(len(cad_heads)))
    matches = []
    for sdf_head in sdf_heads:
        sx, sy = _norm_point(sdf_head, sdf_box)
        best_idx = None
        best_dist = float("inf")
        for idx in unused:
            cx, cy = _norm_point(cad_heads[idx], cad_box)
            d = math.hypot(sx - cx, sy - cy)
            if d < best_dist:
                best_idx = idx
                best_dist = d
        if best_idx is None:
            continue
        unused.remove(best_idx)
        cad_head = cad_heads[best_idx]
        matches.append(
            {
                "sdf_head": sdf_head.get("label"),
                "sdf_node": sdf_head.get("input_node"),
                "sdf_x": sdf_head.get("x"),
                "sdf_y": sdf_head.get("y"),
                "cad_candidate": cad_head.get("label"),
                "cad_layer": cad_head.get("layer"),
                "cad_type": cad_head.get("type"),
                "cad_x": cad_head.get("x"),
                "cad_y": cad_head.get("y"),
                "normalized_error": round(best_dist, 4),
                "status": "FAIL" if best_dist > 0.08 else "PASS",
                "reason": "정규화 좌표 오차가 0.08을 초과합니다." if best_dist > 0.08 else "정규화 좌표 기준 최근접 매칭 허용범위 이내입니다.",
            }
        )

    mismatch_count = sum(1 for m in matches if m["status"] == "FAIL")
    return {
        "status": "FAIL" if mismatch_count else "PASS",
        "message": f"SDF 최원단 헤드 {len(sdf_heads)}개와 CAD 헤드 후보 {len(cad_heads)}개를 정규화 좌표 기준으로 비교했습니다.",
        "matches": matches,
        "unmatched_sdf": sdf_heads[len(matches):],
        "unmatched_cad": [cad_heads[i] for i in sorted(unused)],
        "mismatch_count": mismatch_count,
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


@app.get("/api/feedback-posts")
def feedback_posts():
    return jsonify({"ok": True, "posts": _load_feedback_posts()})


@app.post("/api/feedback-posts")
def create_feedback_post():
    if request.content_type and request.content_type.startswith("multipart/form-data"):
        source = request.form
    else:
        source = request.get_json(silent=True) or {}
    author = _clean_feedback_text(source.get("author") or "익명", 40) or "익명"
    title = _clean_feedback_text(source.get("title"), 80)
    body = str(source.get("body") or "").strip()[:3000]

    if not title:
        return jsonify({"ok": False, "message": "제목을 입력해주세요."}), 400
    if not body:
        return jsonify({"ok": False, "message": "개선의견 내용을 입력해주세요."}), 400

    posts = _load_feedback_posts()
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    post_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
    attachment = _save_feedback_attachment(post_id)
    post = {
        "id": post_id,
        "author": author,
        "title": title,
        "body": body,
        "created_at": created_at,
        "attachment": attachment,
    }
    posts.insert(0, post)
    _save_feedback_posts(posts[:300])
    return jsonify({"ok": True, "message": "개선의견이 등록되었습니다.", "post": post})


@app.get("/api/feedback-attachments/<path:filename>")
def download_feedback_attachment(filename: str):
    safe_name = Path(filename).name
    target = FEEDBACK_UPLOAD_DIR / safe_name
    if not target.exists() or not target.is_file():
        return jsonify({"ok": False, "message": "첨부파일을 찾을 수 없습니다."}), 404
    return send_file(target, as_attachment=True)


@app.post("/api/cad-module/dxf-parse")
def cad_module_dxf_parse():
    try:
        cad_path = _save_upload("cad_file", {".dxf", ".dwg"}, required=True)
        if cad_path.suffix.lower() == ".dwg":
            raise ValueError("현재 6번 모듈은 DXF만 지원합니다. DWG는 DXF로 변환 후 업로드해 주세요.")

        from cad_engine import DXFWorkspace

        workspace = DXFWorkspace(UPLOAD_DIR / "cad_workspace")
        workspace.load_file(cad_path)
        payload = workspace.to_payload(
            include_network_entities=False,
            include_network_summary=False,
            include_graph=False,
        )
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
                "graph": {},
                "unsupported": payload.get("unsupported") or {},
            },
        }
    )


@app.post("/api/sdf-sprinkler-analysis")
def sdf_sprinkler_analysis():
    try:
        sdf_path = _save_upload("sdf_file", {".sdf"}, required=True)
        analysis = _analyze_sdf_sprinkler_network(sdf_path)
        cad_analysis = None
        comparison = None
        cad_path = _save_upload("cad_file", {".dxf"}, required=False)
        if cad_path is not None:
            cad_analysis = _extract_cad_head_candidates(cad_path)
            comparison = _compare_cad_heads_to_sdf(cad_analysis, analysis)
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "message": f"SDF 분석 중 오류가 발생했습니다: {exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "message": "SDF 스프링클러 배관 분석이 완료되었습니다.",
            "analysis": analysis,
            "cad_analysis": cad_analysis,
            "comparison": comparison,
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
        engineering_visualizations = _build_engineering_visualizations(validation)
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
            "insights": {
                **validation["insights"],
                "engineering_visualizations": engineering_visualizations,
            },
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
        payload = workspace.to_payload(
            include_network_entities=True,
            include_network_summary=True,
            include_graph=True,
        )

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
                "graph": payload.get("graph") or {},
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
