from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader


DOCX_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
NUM = r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:E[-+]?\d+)?"


def normalize_desc(raw: str) -> str:
    s = raw.strip().upper().replace(" ", "")
    mapping = {
        "A/V": "AV",
        "AV": "AV",
        "P/V": "PV",
        "PV": "PV",
        "FX": "FX",
        "FLEXIBLE": "FX",
        "후렉": "FX",
        "후렉시블": "FX",
    }
    return mapping.get(s, s)


@dataclass(slots=True)
class DesignInfo:
    equation: str
    pipe_materials: dict[int, str]
    available_sizes: list[dict]


@dataclass(slots=True)
class DesignMaterialRow:
    pipe_type_id: int
    material_name: str


@dataclass(slots=True)
class PipeConfigRow:
    label: int
    input_node: str
    output_node: str
    nominal_bore_mm: float
    length_m: float
    elevation_m: float
    c_factor: float
    fitting_eq_length_m: float


@dataclass(slots=True)
class DesignedPipeRow:
    label: int
    input_node: str
    output_node: str
    pipe_type: int
    flow_lpm: float
    pipe_type_id: int
    actual_bore_mm: float
    nominal_size_mm: float


@dataclass(slots=True)
class RuleResult:
    rule_id: str
    status: str
    severity: str
    subject_type: str
    subject_id: str | int | None
    message: str
    evidence: dict


@dataclass(slots=True)
class HwCheckRow:
    label: int
    length_m: float
    fitting_eq_length_m: float
    special_eq_length_m: float
    total_length_m: float
    c_factor: float
    actual_bore_mm: float
    flow_lpm: float
    reported_friction_loss: float
    calculated_friction_loss: float
    abs_diff: float
    rel_diff: float
    has_special_equipment: bool
    hw_ok: bool


@dataclass(slots=True)
class PipeVelocityCheckRow:
    label: int
    input_node: str
    output_node: str
    nominal_bore_mm: float
    actual_bore_mm: float | None
    velocity_mps: float
    downstream_nozzle_count: int
    subtree_has_cross_split: bool
    pipe_role: str
    velocity_limit_mps: float | None
    velocity_ok: bool | None
    role_reason: str


@dataclass(slots=True)
class PipeValidationRow:
    label: int
    input_node: str
    output_node: str
    nominal_bore_mm: float
    actual_bore_mm: float | None
    material_name: str | None
    c_factor: float | None
    inlet_pressure_kgcm2: float | None
    outlet_pressure_kgcm2: float | None
    max_pressure_kgcm2: float | None
    pipe_rule_results: dict


@dataclass(slots=True)
class SdfPipeRow:
    label: int | None
    input_node: str
    output_node: str
    waypoints: list[tuple[float, float]]
    rise_m: float | None = None
    elevation_m: float | None = None


@dataclass(slots=True)
class NozzleConfigRow:
    label: int
    input_node: str
    nozzle_type: str
    k_factor: float
    req_flow_lpm: float
    min_press_kgcm2: float
    max_press_kgcm2: float


@dataclass(slots=True)
class PipeFittingRow:
    pipe_label: int
    fitting_type_id: int
    count: int
    eq_length_m: float


@dataclass(slots=True)
class FlowAtInletRow:
    inlet_node: str
    pressure_kgcm2: float
    flow_lpm: float


@dataclass(slots=True)
class PumpFlowRow:
    label: int
    pump_setting_head_m: float
    flow_lpm: float


@dataclass(slots=True)
class PipeFlowRow:
    label: int
    input_node: str
    output_node: str
    nominal_bore_mm: float
    inlet_pressure: float
    outlet_pressure: float
    pressure_drop: float
    friction_loss: float
    flow_lpm: float
    velocity_mps: float
    has_special_equipment: bool


@dataclass(slots=True)
class NozzleFlowRow:
    label: int
    input_node: str
    inlet_pressure_kgf_cm2: float
    required_flow_lpm: float
    actual_flow_lpm: float
    deviation_percent: float


@dataclass(slots=True)
class EquipmentRow:
    label: int
    pipe_label: int
    equivalent_length_m: float
    description: str


@dataclass(slots=True)
class ElastomericValveRow:
    label: int
    inlet_pressure_kgf_cm2: float
    outlet_pressure_kgf_cm2: float
    pressure_drop_kgf_cm2: float
    flow_lpm: float


class PipenetGuideValidator:
    BASE_DIR = Path(__file__).resolve().parent
    BRANCH_PIPE_LIMIT_MM = 50.0
    BRANCH_PIPE_V_LIMIT = 6.0
    MAIN_PIPE_V_LIMIT = 10.0
    MIN_HEAD_FLOW = 80.0
    MIN_HEAD_PRESSURE = 1.0
    MAX_HEAD_PRESSURE = 12.0
    FX_EQ_MIN = 13.0
    FX_EQ_MAX = 21.0
    AV_EQ_REF = 12.9
    PV_EQ_REF = 10.1
    VALVE_DROP_TOLERANCE = 0.05
    ECONOMY_V_LOW = 2.0
    FRICTION_SPIKE = 1.0
    HW_DECLARED_TEXT = "USING THE HAZEN-WILLIAMS EQUATION"

    def __init__(
        self,
        report_path: Path,
        sdf_path: Path | None = None,
        cad_path: Path | None = None,
        project_meta_path: Path | None = None,
        design_policy_path: Path | None = None,
    ):
        self.report_path = Path(report_path)
        self.sdf_path = Path(sdf_path) if sdf_path else None
        self.cad_path = Path(cad_path) if cad_path else None
        self.project_meta_path = Path(project_meta_path) if project_meta_path else self.BASE_DIR / "data" / "project_meta.json"
        self.design_policy_path = Path(design_policy_path) if design_policy_path else self.BASE_DIR / "data" / "design_policy.json"

    def validate(self) -> dict:
        report_text = self._read_report_text(self.report_path)
        design_info = self._parse_design_info(report_text)
        design_material_rows = self._parse_design_materials(report_text)
        pipe_config_rows = self._parse_pipe_config_rows(report_text)
        designed_pipe_rows = self._parse_designed_pipe_rows(report_text)
        nozzle_config_rows = self._parse_nozzle_config_rows(report_text)
        pipe_fitting_rows = self._parse_pipe_fitting_rows(report_text)
        inlet_rows = self._parse_flow_at_inlet_rows(report_text)
        pump_rows = self._parse_pump_flow_rows(report_text)
        pipe_rows = self._parse_pipe_flows(report_text)
        nozzle_rows = self._parse_nozzle_flows(report_text)
        equipment_rows = self._parse_equipment(report_text)
        valve_rows = self._parse_elastomeric_valves(report_text)
        pipe_lengths = {row.label: row.length_m for row in pipe_config_rows}
        report_titles = self._parse_report_titles(report_text)
        sdf_network = self._parse_sdf_network(self.sdf_path) if self.sdf_path else {}
        sdf_info = self._parse_sdf(self.sdf_path, sdf_network) if self.sdf_path else {}
        project_meta = self._load_json_file(self.project_meta_path)
        design_policy = self._load_json_file(self.design_policy_path)
        hw_declared = self._check_hw_declared(report_text, design_info)
        hw_checks = self._build_hw_check_rows(pipe_config_rows, designed_pipe_rows, equipment_rows, pipe_rows)
        hw_fail_ids = {row.label for row in hw_checks if not row.hw_ok}
        velocity_checks = self._build_pipe_velocity_check_rows(pipe_config_rows, nozzle_config_rows, pipe_rows, designed_pipe_rows)
        pipe_validation_rows = self._build_pipe_validation_rows(
            pipe_config_rows=pipe_config_rows,
            designed_pipe_rows=designed_pipe_rows,
            pipe_rows=pipe_rows,
            design_material_rows=design_material_rows,
        )
        pipe_rule_results = self._evaluate_pipe_rules(
            pipe_validation_rows=pipe_validation_rows,
            nozzle_config_rows=nozzle_config_rows,
            sdf_network=sdf_network,
            project_meta=project_meta,
            design_policy=design_policy,
        )

        pipe_fail_ids = self._find_velocity_fail_ids(velocity_checks)
        nozzle_fail_ids = self._find_nozzle_fail_ids(nozzle_rows)
        equipment_fail_ids, equipment_warn_ids = self._find_equipment_issue_ids(equipment_rows)
        valve_fail_ids = self._find_valve_fail_ids(valve_rows)

        insights = self._build_optimization_insights(
            pipe_rows=pipe_rows,
            pipe_lengths=pipe_lengths,
            equipment_rows=equipment_rows,
            pipe_config_rows=pipe_config_rows,
            designed_pipe_rows=designed_pipe_rows,
            design_material_rows=design_material_rows,
            hw_checks=hw_checks,
            velocity_checks=velocity_checks,
            nozzle_rows=nozzle_rows,
        )

        results: dict[str, list[str]] = {"PASS": [], "FAIL": [], "WARNING": []}
        self._build_common_hw_messages(hw_declared, hw_checks, results)
        self._build_pipe_rule_messages(pipe_rule_results, results)
        self._build_pipe_velocity_messages(velocity_checks, results)
        self._build_nozzle_messages(nozzle_rows, results)
        self._build_equipment_messages(equipment_rows, equipment_fail_ids, equipment_warn_ids, results)
        self._build_valve_messages(valve_rows, valve_fail_ids, results)
        self._build_cross_check_messages(
            nozzles=nozzle_rows,
            equipment_rows=equipment_rows,
            report_titles=report_titles,
            sdf_info=sdf_info,
            results=results,
        )

        pipe_stats = self._build_pipe_stats(pipe_validation_rows, pipe_rule_results)
        stats = self._build_stats(
            pipe_rows,
            nozzle_rows,
            equipment_rows,
            valve_rows,
            report_titles,
            sdf_info,
            hw_declared,
            hw_checks,
            velocity_checks,
            pipe_stats,
        )
        tables = self._build_tables(
            pipe_rows=pipe_rows,
            pipe_lengths=pipe_lengths,
            pipe_config_rows=pipe_config_rows,
            designed_pipe_rows=designed_pipe_rows,
            hw_checks=hw_checks,
            velocity_checks=velocity_checks,
            pipe_validation_rows=pipe_validation_rows,
            nozzle_rows=nozzle_rows,
            equipment_rows=equipment_rows,
            valve_rows=valve_rows,
            pipe_fail_ids=pipe_fail_ids,
            hw_fail_ids=hw_fail_ids,
            nozzle_fail_ids=nozzle_fail_ids,
            equipment_fail_ids=equipment_fail_ids,
            equipment_warn_ids=equipment_warn_ids,
            valve_fail_ids=valve_fail_ids,
            engineering_pipe_ids=insights["engineering_pipe_ids"],
            economy_pipe_ids=insights["economy_pipe_ids"],
            economy_equipment_ids=insights["economy_equipment_ids"],
            engineering_reason_map=insights.get("engineering_reason_map", {}),
            economy_reason_map=insights.get("economy_reason_map", {}),
        )

        return {
            "ok": True,
            "report_name": self.report_path.name,
            "sdf_name": self.sdf_path.name if self.sdf_path else None,
            "design_info": {
                "equation": design_info.equation,
                "pipe_materials": design_info.pipe_materials,
                "available_sizes": design_info.available_sizes,
            },
            "summary": {
                "pass": len(results["PASS"]),
                "fail": len(results["FAIL"]),
                "warning": len(results["WARNING"]),
            },
            "results": results,
            "insights": {
                "engineering_advice": insights["engineering_advice"],
                "economy_guide": insights["economy_guide"],
            },
            "rules": {
                "pipe": [self._rule_result_to_dict(item) for item in pipe_rule_results],
            },
            "stats": stats,
            "tables": tables,
            "report": self._build_report(results, insights, stats),
            "parsed_context": {
                "design_material_count": len(design_material_rows),
                "pipe_config_count": len(pipe_config_rows),
                "designed_pipe_count": len(designed_pipe_rows),
                "nozzle_config_count": len(nozzle_config_rows),
                "pipe_fitting_count": len(pipe_fitting_rows),
                "flow_at_inlet_count": len(inlet_rows),
                "pump_flow_count": len(pump_rows),
                "flow_in_pipes_count": len(pipe_rows),
                "special_equipment_count": len(equipment_rows),
                "sdf_pipe_count": len(sdf_network.get("pipes", [])),
            },
        }

    def _read_report_text(self, path: Path) -> str:
        suffix = path.suffix.lower()
        if suffix == ".docx":
            with zipfile.ZipFile(path) as zf:
                root = ET.fromstring(zf.read("word/document.xml"))
            paragraphs: list[str] = []
            for paragraph in root.findall(".//w:p", DOCX_NS):
                text = "".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS)).strip()
                if text:
                    paragraphs.append(text)
            return "\n".join(paragraphs)
        if suffix == ".pdf":
            reader = PdfReader(str(path))
            return "\n".join((page.extract_text() or "") for page in reader.pages)
        raise ValueError("지원하지 않는 결과 파일 형식입니다. docx 또는 pdf만 검증할 수 있습니다.")

    def _load_json_file(self, path: Path | None) -> dict | None:
        if path is None or not path.exists():
            return None
        try:
            with path.open("r", encoding="utf-8") as fp:
                return json.load(fp)
        except Exception:
            return None

    def _rule_result_to_dict(self, item: RuleResult) -> dict:
        return {
            "rule_id": item.rule_id,
            "status": item.status,
            "severity": item.severity,
            "subject_type": item.subject_type,
            "subject_id": item.subject_id,
            "message": item.message,
            "evidence": item.evidence,
        }

    def _extract_sections(self, text: str, title: str) -> Iterable[str]:
        pattern = re.compile(
            rf"{re.escape(title)}\n[-]+\n(.*?)(?=\n(?:TITLE :|[A-Z][A-Z /&().-]{{4,}}\n[-]+\n)|\Z)",
            re.S,
        )
        for match in pattern.finditer(text):
            yield match.group(1)

    def _section_text(self, text: str, title: str) -> str:
        return "\n".join(self._extract_sections(text, title))

    def _section_lines(self, text: str, title: str) -> list[str]:
        return [line.strip() for line in self._section_text(text, title).splitlines() if line.strip()]

    def _parse_design_info(self, text: str) -> DesignInfo:
        equation_match = re.search(r"(Using the Hazen-Williams Equation)", text, re.I)
        equation = equation_match.group(1).strip() if equation_match else ""
        pipe_materials = {row.pipe_type_id: row.material_name for row in self._parse_design_materials(text)}

        available_sizes: list[dict] = []
        size_pattern = re.compile(rf"^\s*(\d+)\s+({NUM})\s+({NUM})(?:\s+({NUM}))?.*$")
        for line in self._section_lines(text, "AVAILABLE SIZES"):
            match = size_pattern.match(line)
            if not match:
                continue
            available_sizes.append(
                {
                    "pipe_type_id": int(match.group(1)),
                    "nominal_size_mm": float(match.group(2)),
                    "actual_bore_mm": float(match.group(3)),
                    "wall_thickness_mm": float(match.group(4)) if match.group(4) else None,
                }
            )

        return DesignInfo(equation=equation, pipe_materials=pipe_materials, available_sizes=available_sizes)

    def _parse_design_materials(self, text: str) -> list[DesignMaterialRow]:
        rows: list[DesignMaterialRow] = []
        pattern = re.compile(rf"^\s*(\d+)\s+([A-Za-z0-9_.\-/ ]+?)\s*$")
        inline_pattern = re.compile(r"(\d+)\s+--\s+([A-Za-z0-9_.\-/ ]+?)\s{2,}Not Lined", re.I)
        seen: set[int] = set()
        for title in ("PIPE TYPES", "PIPE MATERIALS"):
            for line in self._section_lines(text, title):
                match = pattern.match(line)
                if not match:
                    continue
                pipe_type_id = int(match.group(1))
                if pipe_type_id in seen:
                    continue
                seen.add(pipe_type_id)
                rows.append(DesignMaterialRow(pipe_type_id=pipe_type_id, material_name=match.group(2).strip()))
        for match in inline_pattern.finditer(text):
            pipe_type_id = int(match.group(1))
            if pipe_type_id in seen:
                continue
            seen.add(pipe_type_id)
            rows.append(DesignMaterialRow(pipe_type_id=pipe_type_id, material_name=match.group(2).strip()))
        return rows

    def _parse_pipe_config_rows(self, text: str) -> list[PipeConfigRow]:
        rows: list[PipeConfigRow] = []
        pattern = re.compile(
            rf"^(\d+)\s+(\d+)\s+(\d+)\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})$"
        )
        for line in self._section_lines(text, "PIPE CONFIGURATION"):
            match = pattern.match(line)
            if not match:
                continue
            rows.append(
                PipeConfigRow(
                    label=int(match.group(1)),
                    input_node=match.group(2),
                    output_node=match.group(3),
                    nominal_bore_mm=float(match.group(4)),
                    length_m=float(match.group(5)),
                    elevation_m=float(match.group(6)),
                    c_factor=float(match.group(7)),
                    fitting_eq_length_m=float(match.group(8)),
                )
            )
        return rows

    def _parse_pipe_configs(self, text: str) -> list[PipeConfigRow]:
        return self._parse_pipe_config_rows(text)

    def _parse_designed_pipe_rows(self, text: str) -> list[DesignedPipeRow]:
        rows: list[DesignedPipeRow] = []
        pattern = re.compile(rf"^(\d+)\s+(\d+)\s+(\d+)\s+({NUM})\s+(\d+)\s+({NUM})\s+({NUM})(?:\s+\S+)?$")
        section = "DESIGNED DIAMETERS & FLOWRATES"
        for line in self._section_lines(text, section):
            match = pattern.match(line)
            if not match:
                continue
            rows.append(
                DesignedPipeRow(
                    label=int(match.group(1)),
                    input_node=match.group(2),
                    output_node=match.group(3),
                    flow_lpm=float(match.group(4)),
                    pipe_type=int(match.group(5)),
                    pipe_type_id=int(match.group(5)),
                    actual_bore_mm=float(match.group(6)),
                    nominal_size_mm=float(match.group(7)),
                )
            )
        return rows

    def _parse_designed_pipes(self, text: str) -> list[DesignedPipeRow]:
        return self._parse_designed_pipe_rows(text)

    def _parse_nozzle_config_rows(self, text: str) -> list[NozzleConfigRow]:
        rows: list[NozzleConfigRow] = []
        pattern = re.compile(
            rf"^(\d+)\s+(\d+)\s+([A-Za-z0-9_.\-/]+)\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})$"
        )
        for line in self._section_lines(text, "NOZZLE CONFIGURATION"):
            match = pattern.match(line)
            if not match:
                continue
            rows.append(
                NozzleConfigRow(
                    label=int(match.group(1)),
                    input_node=match.group(2),
                    nozzle_type=match.group(3),
                    k_factor=float(match.group(4)),
                    req_flow_lpm=float(match.group(5)),
                    min_press_kgcm2=float(match.group(6)),
                    max_press_kgcm2=float(match.group(7)),
                )
            )
        return rows

    def _parse_pipe_fitting_rows(self, text: str) -> list[PipeFittingRow]:
        rows: list[PipeFittingRow] = []
        pattern = re.compile(rf"^(\d+)\s+(\d+)\s+(\d+)\s+({NUM})$")
        for line in self._section_lines(text, "PIPE FITTINGS"):
            match = pattern.match(line)
            if not match:
                continue
            rows.append(
                PipeFittingRow(
                    pipe_label=int(match.group(1)),
                    fitting_type_id=int(match.group(2)),
                    count=int(match.group(3)),
                    eq_length_m=float(match.group(4)),
                )
            )
        return rows

    def _parse_flow_at_inlet_rows(self, text: str) -> list[FlowAtInletRow]:
        rows: list[FlowAtInletRow] = []
        pattern = re.compile(rf"^(\d+)\s+({NUM})\s+({NUM})$")
        for line in self._section_lines(text, "FLOW AT INLET"):
            match = pattern.match(line)
            if not match:
                continue
            rows.append(
                FlowAtInletRow(
                    inlet_node=match.group(1),
                    pressure_kgcm2=float(match.group(2)),
                    flow_lpm=float(match.group(3)),
                )
            )
        return rows

    def _parse_pump_flow_rows(self, text: str) -> list[PumpFlowRow]:
        rows: list[PumpFlowRow] = []
        pattern = re.compile(rf"^(\d+)\s+({NUM})\s+({NUM})$")
        for line in self._section_lines(text, "FLOW THROUGH PUMPS"):
            match = pattern.match(line)
            if not match:
                continue
            rows.append(
                PumpFlowRow(
                    label=int(match.group(1)),
                    pump_setting_head_m=float(match.group(2)),
                    flow_lpm=float(match.group(3)),
                )
            )
        return rows

    def _parse_pipe_flows(self, text: str) -> list[PipeFlowRow]:
        rows: list[PipeFlowRow] = []
        pattern = re.compile(
            rf"^(\d+)\s+(\d+)\s+(\d+)\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s*(E)?$"
        )
        for line in self._section_lines(text, "FLOW IN PIPES"):
            match = pattern.match(line.strip())
            if not match:
                continue
            rows.append(
                PipeFlowRow(
                    label=int(match.group(1)),
                    input_node=match.group(2),
                    output_node=match.group(3),
                    nominal_bore_mm=float(match.group(4)),
                    inlet_pressure=float(match.group(5)),
                    outlet_pressure=float(match.group(6)),
                    pressure_drop=float(match.group(7)),
                    friction_loss=float(match.group(8)),
                    flow_lpm=float(match.group(9)),
                    velocity_mps=float(match.group(10)),
                    has_special_equipment=bool(match.group(11)),
                )
            )
        return rows

    def _parse_pipe_lengths(self, text: str) -> dict[int, float]:
        return {row.label: row.length_m for row in self._parse_pipe_configs(text)}

    def _parse_nozzle_flows(self, text: str) -> list[NozzleFlowRow]:
        rows: list[NozzleFlowRow] = []
        pattern = re.compile(rf"^(\d+)\s+(\d+)\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})")
        for line in self._section_lines(text, "FLOW THROUGH NOZZLES"):
            match = pattern.match(line.strip())
            if not match:
                continue
            rows.append(
                NozzleFlowRow(
                    label=int(match.group(1)),
                    input_node=match.group(2),
                    inlet_pressure_kgf_cm2=float(match.group(3)),
                    required_flow_lpm=float(match.group(4)),
                    actual_flow_lpm=float(match.group(5)),
                    deviation_percent=float(match.group(6)),
                )
            )
        return rows

    def _parse_equipment(self, text: str) -> list[EquipmentRow]:
        rows: list[EquipmentRow] = []
        pattern = re.compile(rf"^(\d+)\s+(\d+)\s+({NUM})\s+(.+?)$")
        for line in self._section_lines(text, "SPECIAL EQUIPMENT"):
            match = pattern.match(line.strip())
            if not match:
                continue
            rows.append(
                EquipmentRow(
                    label=int(match.group(1)),
                    pipe_label=int(match.group(2)),
                    equivalent_length_m=float(match.group(3)),
                    description=normalize_desc(match.group(4)),
                )
            )
        return rows

    def _parse_elastomeric_valves(self, text: str) -> list[ElastomericValveRow]:
        rows: list[ElastomericValveRow] = []
        pattern = re.compile(rf"^(\d+)\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})")
        for line in self._section_lines(text, "FLOW THROUGH ELASTOMERIC VALVES"):
            match = pattern.match(line.strip())
            if not match:
                continue
            rows.append(
                ElastomericValveRow(
                    label=int(match.group(1)),
                    inlet_pressure_kgf_cm2=float(match.group(2)),
                    outlet_pressure_kgf_cm2=float(match.group(3)),
                    pressure_drop_kgf_cm2=float(match.group(4)),
                    flow_lpm=float(match.group(5)),
                )
            )
        return rows

    def _check_hw_declared(self, text: str, design_info: DesignInfo) -> bool:
        equation_text = (design_info.equation or "").strip().upper()
        if self.HW_DECLARED_TEXT in equation_text:
            return True
        return self.HW_DECLARED_TEXT in text.upper()

    def _build_hw_check_rows(
        self,
        pipe_config_rows: list[PipeConfigRow],
        designed_pipe_rows: list[DesignedPipeRow],
        equipment_rows: list[EquipmentRow],
        pipe_rows: list[PipeFlowRow],
    ) -> list[HwCheckRow]:
        config_map = {row.label: row for row in pipe_config_rows}
        designed_map = {row.label: row for row in designed_pipe_rows}
        special_map: dict[int, float] = {}
        for item in equipment_rows:
            special_map[item.pipe_label] = special_map.get(item.pipe_label, 0.0) + item.equivalent_length_m

        checks: list[HwCheckRow] = []
        for flow_row in pipe_rows:
            config = config_map.get(flow_row.label)
            designed = designed_map.get(flow_row.label)
            if not config or not designed:
                continue
            special_eq_length_m = special_map.get(flow_row.label, 0.0)
            total_length_m = config.length_m + config.fitting_eq_length_m + special_eq_length_m
            calculated_friction_loss = self._calc_hw_friction_loss_kgcm2(
                q_lpm=designed.flow_lpm,
                total_length_m=total_length_m,
                c_factor=config.c_factor,
                actual_bore_mm=designed.actual_bore_mm,
            )
            reported_friction_loss = flow_row.friction_loss
            abs_diff = abs(calculated_friction_loss - reported_friction_loss)
            rel_diff = abs_diff / max(reported_friction_loss, 1e-9)
            tolerance = max(0.005, reported_friction_loss * 0.005)
            checks.append(
                HwCheckRow(
                    label=flow_row.label,
                    length_m=config.length_m,
                    fitting_eq_length_m=config.fitting_eq_length_m,
                    special_eq_length_m=special_eq_length_m,
                    total_length_m=total_length_m,
                    c_factor=config.c_factor,
                    actual_bore_mm=designed.actual_bore_mm,
                    flow_lpm=designed.flow_lpm,
                    reported_friction_loss=reported_friction_loss,
                    calculated_friction_loss=calculated_friction_loss,
                    abs_diff=abs_diff,
                    rel_diff=rel_diff,
                    has_special_equipment=flow_row.has_special_equipment,
                    hw_ok=abs_diff <= tolerance,
                )
            )
        return checks

    def _calc_hw_friction_loss_kgcm2(
        self,
        *,
        q_lpm: float,
        total_length_m: float,
        c_factor: float,
        actual_bore_mm: float,
    ) -> float:
        dp_mpa = 6.174e4 * (q_lpm ** 1.85) * total_length_m / ((c_factor ** 1.85) * (actual_bore_mm ** 4.87))
        return dp_mpa / 0.1

    def _build_pipe_validation_rows(
        self,
        pipe_config_rows: list[PipeConfigRow],
        designed_pipe_rows: list[DesignedPipeRow],
        pipe_rows: list[PipeFlowRow],
        design_material_rows: list[DesignMaterialRow],
    ) -> list[PipeValidationRow]:
        config_map = {row.label: row for row in pipe_config_rows}
        designed_map = {row.label: row for row in designed_pipe_rows}
        material_map = {row.pipe_type_id: row.material_name for row in design_material_rows}
        rows: list[PipeValidationRow] = []
        for flow in pipe_rows:
            config = config_map.get(flow.label)
            designed = designed_map.get(flow.label)
            rows.append(
                PipeValidationRow(
                    label=flow.label,
                    input_node=flow.input_node,
                    output_node=flow.output_node,
                    nominal_bore_mm=config.nominal_bore_mm if config else flow.nominal_bore_mm,
                    actual_bore_mm=designed.actual_bore_mm if designed else None,
                    material_name=material_map.get(designed.pipe_type_id) if designed else None,
                    c_factor=config.c_factor if config else None,
                    inlet_pressure_kgcm2=flow.inlet_pressure,
                    outlet_pressure_kgcm2=flow.outlet_pressure,
                    max_pressure_kgcm2=max(flow.inlet_pressure, flow.outlet_pressure),
                    pipe_rule_results={},
                )
            )
        return rows

    def _classify_head_type(self, nozzle_rows: list[NozzleConfigRow]) -> str:
        if not nozzle_rows:
            return "standard"
        return "standard" if max((row.k_factor for row in nozzle_rows), default=80.0) <= 80.0 else "special"

    def _count_sdf_downstream_nozzles(
        self,
        node_id: str,
        outgoing_by_node: dict[str, list[SdfPipeRow]],
        nozzle_count_by_node: dict[str, int],
        memo: dict[str, int],
        stack: set[str],
    ) -> int:
        if node_id in memo:
            return memo[node_id]
        if node_id in stack:
            raise ValueError(f"sdf-cycle:{node_id}")
        stack.add(node_id)
        total = nozzle_count_by_node.get(node_id, 0)
        for pipe in outgoing_by_node.get(node_id, []):
            total += self._count_sdf_downstream_nozzles(pipe.output_node, outgoing_by_node, nozzle_count_by_node, memo, stack)
        stack.remove(node_id)
        memo[node_id] = total
        return total

    def _evaluate_pipe_rules(
        self,
        pipe_validation_rows: list[PipeValidationRow],
        nozzle_config_rows: list[NozzleConfigRow],
        sdf_network: dict,
        project_meta: dict | None,
        design_policy: dict | None,
    ) -> list[RuleResult]:
        results: list[RuleResult] = []
        nozzle_by_input: dict[str, list[NozzleConfigRow]] = {}
        for nozzle in nozzle_config_rows:
            nozzle_by_input.setdefault(nozzle.input_node, []).append(nozzle)

        if not self.cad_path or not self.cad_path.exists():
            results.append(RuleResult("PIPE.001", "REVIEW", "INFO", "meta", None, "배관-도면/스케매틱 일치: CAD 도면 입력이 없어 판정 보류", {"cad_path": str(self.cad_path) if self.cad_path else None}))
        else:
            try:
                from cad_engine import build_explicit_pipe_mismatches

                mismatches = build_explicit_pipe_mismatches(self.cad_path, self.sdf_path)
                results.append(
                    RuleResult(
                        "PIPE.001",
                        "FAIL" if mismatches else "PASS",
                        "ERROR" if mismatches else "INFO",
                        "network",
                        None,
                        "배관-도면/스케매틱 일치: 명시적 mismatch가 발견되지 않았습니다." if not mismatches else f"배관-도면/스케매틱 일치: {len(mismatches)}건 mismatch 발견",
                        {"mismatches": mismatches},
                    )
                )
            except Exception as exc:
                results.append(RuleResult("PIPE.001", "REVIEW", "WARN", "network", None, f"배관-도면/스케매틱 일치: CAD 비교를 완료하지 못해 판정 보류 ({exc})", {}))

        high_pressure_rows = [row for row in pipe_validation_rows if (row.max_pressure_kgcm2 or 0.0) >= 12.0]
        max_pipe_pressure = max((row.max_pressure_kgcm2 or 0.0 for row in pipe_validation_rows), default=0.0)
        if not high_pressure_rows:
            results.append(RuleResult("PIPE.002", "PASS", "INFO", "meta", None, f"배관-고압구간 재질 확인: 1.2MPa(12.0kg/cm2G) 이상 구간이 없습니다. 최대 관압은 {max_pipe_pressure:.3f}kg/cm2G 입니다.", {"max_pipe_pressure_kgcm2": max_pipe_pressure}))
            for row in pipe_validation_rows:
                row.pipe_rule_results["PIPE.002"] = RuleResult("PIPE.002", "PASS", "INFO", "pipe", row.label, f"최대 관압 {row.max_pressure_kgcm2:.3f} kg/cm2G로 1.2MPa 미만", {"max_pressure_kgcm2": row.max_pressure_kgcm2})
        else:
            fail_rows = []
            for row in pipe_validation_rows:
                if (row.max_pressure_kgcm2 or 0.0) >= 12.0:
                    ok = (row.material_name or "").strip().upper() == "KSD 3562"
                    row.pipe_rule_results["PIPE.002"] = RuleResult("PIPE.002", "PASS" if ok else "FAIL", "ERROR" if not ok else "INFO", "pipe", row.label, f"Pipe {row.label}, 최대압력 {row.max_pressure_kgcm2:.3f}kg/cm2G, 현재 {row.material_name or '-'}", {"max_pressure_kgcm2": row.max_pressure_kgcm2, "material_name": row.material_name, "required_material": "KSD 3562"})
                    if not ok:
                        fail_rows.append(row)
                else:
                    row.pipe_rule_results["PIPE.002"] = RuleResult("PIPE.002", "PASS", "INFO", "pipe", row.label, f"최대 관압 {row.max_pressure_kgcm2:.3f} kg/cm2G로 1.2MPa 미만", {"max_pressure_kgcm2": row.max_pressure_kgcm2})
            results.append(
                RuleResult(
                    "PIPE.002",
                    "FAIL" if fail_rows else "PASS",
                    "ERROR" if fail_rows else "INFO",
                    "meta",
                    fail_rows[0].label if fail_rows else None,
                    f"배관-고압구간 재질 불일치: Pipe {fail_rows[0].label}, 최대압력 {fail_rows[0].max_pressure_kgcm2:.3f}kg/cm2G, 현재 {fail_rows[0].material_name or '-'}, 요구 KSD3562" if fail_rows else f"배관-고압구간 재질 확인: 1.2MPa 이상 {len(high_pressure_rows)}개 구간이 모두 KSD3562입니다.",
                    {"high_pressure_pipe_count": len(high_pressure_rows), "failed_pipe_labels": [row.label for row in fail_rows]},
                )
            )

        c_fail_rows = []
        for row in pipe_validation_rows:
            material = (row.material_name or "").strip().upper()
            status = "NA"
            if material.startswith("CPVC"):
                status = "PASS" if row.c_factor == 150 else "FAIL"
            elif material in {"KSD 3507", "KSD 3562"}:
                status = "PASS" if row.c_factor == 120 else "FAIL"
            row.pipe_rule_results["PIPE.003"] = RuleResult("PIPE.003", status, "ERROR" if status == "FAIL" else "INFO", "pipe", row.label, f"Pipe {row.label}, 재질 {row.material_name or '-'}, C={row.c_factor}", {"material_name": row.material_name, "c_factor": row.c_factor})
            if status == "FAIL":
                c_fail_rows.append(row)
        results.append(
            RuleResult(
                "PIPE.003",
                "FAIL" if c_fail_rows else "PASS",
                "ERROR" if c_fail_rows else "INFO",
                "meta",
                c_fail_rows[0].label if c_fail_rows else None,
                f"배관-C값 불일치: Pipe {c_fail_rows[0].label}, 재질 {c_fail_rows[0].material_name or '-'}인데 C={c_fail_rows[0].c_factor}" if c_fail_rows else "배관-C값 확인: 강관 계열은 C=120, CPVC 계열은 C=150으로 일치합니다.",
                {"failed_pipe_labels": [row.label for row in c_fail_rows]},
            )
        )

        unit_internal_labels = set((project_meta or {}).get("unit_internal_pipe_labels") or [])
        if not unit_internal_labels:
            results.append(RuleResult("PIPE.004", "REVIEW", "WARN", "meta", None, "배관-세대내 CPVC 확인: unit_internal 분류 입력이 없어 판정 보류", {}))
            for row in pipe_validation_rows:
                row.pipe_rule_results["PIPE.004"] = RuleResult("PIPE.004", "REVIEW", "WARN", "pipe", row.label, "unit_internal 분류 입력 없음", {})
        else:
            fail_labels = []
            for row in pipe_validation_rows:
                if row.label in unit_internal_labels:
                    ok = (row.material_name or "").upper().startswith("CPVC")
                    row.pipe_rule_results["PIPE.004"] = RuleResult("PIPE.004", "PASS" if ok else "FAIL", "ERROR" if not ok else "INFO", "pipe", row.label, f"Pipe {row.label}, 재질 {row.material_name or '-'}", {"material_name": row.material_name})
                    if not ok:
                        fail_labels.append(row.label)
                else:
                    row.pipe_rule_results["PIPE.004"] = RuleResult("PIPE.004", "NA", "INFO", "pipe", row.label, "세대내 배관 대상 아님", {})
            results.append(RuleResult("PIPE.004", "FAIL" if fail_labels else "PASS", "ERROR" if fail_labels else "INFO", "meta", None, "배관-세대내 CPVC 확인: " + ("일부 배관이 CPVC가 아닙니다." if fail_labels else "대상 배관이 모두 CPVC입니다."), {"failed_pipe_labels": fail_labels}))

        unit_inlet_labels = set((project_meta or {}).get("unit_inlet_pipe_labels") or [])
        if not unit_inlet_labels:
            results.append(RuleResult("PIPE.005", "REVIEW", "WARN", "meta", None, "배관-세대유입 65A 확인: unit_inlet 분류 입력이 없어 판정 보류", {}))
            for row in pipe_validation_rows:
                row.pipe_rule_results["PIPE.005"] = RuleResult("PIPE.005", "REVIEW", "WARN", "pipe", row.label, "unit_inlet 분류 입력 없음", {})
        else:
            fail_labels = []
            for row in pipe_validation_rows:
                if row.label in unit_inlet_labels:
                    ok = row.nominal_bore_mm <= 65.0
                    row.pipe_rule_results["PIPE.005"] = RuleResult("PIPE.005", "PASS" if ok else "FAIL", "ERROR" if not ok else "INFO", "pipe", row.label, f"Pipe {row.label}, 구경 {row.nominal_bore_mm:.0f}A", {"nominal_bore_mm": row.nominal_bore_mm})
                    if not ok:
                        fail_labels.append(row.label)
                else:
                    row.pipe_rule_results["PIPE.005"] = RuleResult("PIPE.005", "NA", "INFO", "pipe", row.label, "세대 유입 배관 대상 아님", {})
            results.append(RuleResult("PIPE.005", "FAIL" if fail_labels else "PASS", "ERROR" if fail_labels else "INFO", "meta", None, "배관-세대유입 65A 확인: " + ("65A 초과 배관이 있습니다." if fail_labels else "대상 배관이 모두 65A 이하입니다."), {"failed_pipe_labels": fail_labels}))

        head_policy = ((design_policy or {}).get("head_count_min_nominal_by_head_type") or {}).get("standard")
        if not head_policy:
            results.append(RuleResult("PIPE.006", "REVIEW", "WARN", "meta", None, "배관-헤드개수별 배관구경 정책: design_policy.json 이 없어 판정 보류", {}))
            for row in pipe_validation_rows:
                row.pipe_rule_results["PIPE.006"] = RuleResult("PIPE.006", "REVIEW", "WARN", "pipe", row.label, "design_policy.json 없음", {})
        else:
            sdf_pipes: list[SdfPipeRow] = sdf_network.get("pipes", [])
            outgoing_by_node: dict[str, list[SdfPipeRow]] = {}
            pipe_by_label: dict[int, SdfPipeRow] = {}
            for pipe in sdf_pipes:
                outgoing_by_node.setdefault(pipe.input_node, []).append(pipe)
                if pipe.label is not None:
                    pipe_by_label[pipe.label] = pipe
            nozzle_count_by_node: dict[str, int] = {}
            for node in sdf_network.get("nozzle_inputs", []):
                nozzle_count_by_node[node] = nozzle_count_by_node.get(node, 0) + 1
            memo: dict[str, int] = {}
            fail_labels = []
            for row in pipe_validation_rows:
                pipe = pipe_by_label.get(row.label)
                if pipe is None:
                    row.pipe_rule_results["PIPE.006"] = RuleResult("PIPE.006", "REVIEW", "WARN", "pipe", row.label, "SDF에서 해당 pipe를 찾지 못했습니다.", {})
                    continue
                count = self._count_sdf_downstream_nozzles(pipe.output_node, outgoing_by_node, nozzle_count_by_node, memo, set())
                head_type = self._classify_head_type(nozzle_by_input.get(pipe.output_node, []))
                policy_rows = ((design_policy or {}).get("head_count_min_nominal_by_head_type") or {}).get(head_type) or head_policy
                required = None
                for item in policy_rows:
                    if count <= int(item.get("max_heads", 0)):
                        required = float(item.get("min_nominal_mm", 0))
                        break
                if required is None and policy_rows:
                    required = float(policy_rows[-1].get("min_nominal_mm", 0))
                if required is None:
                    row.pipe_rule_results["PIPE.006"] = RuleResult("PIPE.006", "REVIEW", "WARN", "pipe", row.label, "적용 가능한 정책 행을 찾지 못했습니다.", {"downstream_nozzle_count": count})
                    continue
                ok = row.nominal_bore_mm >= required
                row.pipe_rule_results["PIPE.006"] = RuleResult("PIPE.006", "PASS" if ok else "FAIL", "ERROR" if not ok else "INFO", "pipe", row.label, f"Pipe {row.label}, downstream 헤드 {count}개, 요구 최소 {required:.0f}A", {"downstream_nozzle_count": count, "required_min_nominal_mm": required, "nominal_bore_mm": row.nominal_bore_mm})
                if not ok:
                    fail_labels.append(row.label)
            results.append(RuleResult("PIPE.006", "FAIL" if fail_labels else "PASS", "ERROR" if fail_labels else "INFO", "meta", None, "배관-헤드개수별 배관구경 정책: " + ("정책 미달 배관이 있습니다." if fail_labels else "정책을 충족합니다."), {"failed_pipe_labels": fail_labels}))
        return results

    def _build_pipe_rule_messages(self, rule_results: list[RuleResult], results: dict[str, list[str]]) -> None:
        for item in rule_results:
            if item.subject_type != "meta":
                continue
            if item.status == "PASS":
                results["PASS"].append(item.message)
            elif item.status == "FAIL":
                results["FAIL"].append(item.message)
            elif item.status == "REVIEW":
                results["WARNING"].append(item.message)

    def _build_pipe_stats(self, pipe_validation_rows: list[PipeValidationRow], pipe_rule_results: list[RuleResult]) -> dict:
        high_pressure_rows = [row for row in pipe_validation_rows if (row.max_pressure_kgcm2 or 0.0) >= 12.0]
        c_factor_material_fail_count = sum(
            1
            for row in pipe_validation_rows
            if row.pipe_rule_results.get("PIPE.003") and row.pipe_rule_results["PIPE.003"].status == "FAIL"
        )
        pipe_review_count = sum(1 for item in pipe_rule_results if item.subject_type == "meta" and item.status == "REVIEW")
        return {
            "max_pipe_pressure_kgcm2": max((row.max_pressure_kgcm2 or 0.0 for row in pipe_validation_rows), default=None),
            "high_pressure_pipe_count": len(high_pressure_rows),
            "ksd3562_required_pipe_count": len(high_pressure_rows),
            "c_factor_material_fail_count": c_factor_material_fail_count,
            "pipe_review_count": pipe_review_count,
        }

    def _build_pipe_topology(
        self,
        pipe_config_rows: list[PipeConfigRow],
        nozzle_config_rows: list[NozzleConfigRow],
    ) -> dict:
        pipe_map = {row.label: row for row in pipe_config_rows}
        outgoing_by_node: dict[str, list[int]] = {}
        for row in pipe_config_rows:
            outgoing_by_node.setdefault(row.input_node, []).append(row.label)
        nozzle_count_by_node: dict[str, int] = {}
        for row in nozzle_config_rows:
            nozzle_count_by_node[row.input_node] = nozzle_count_by_node.get(row.input_node, 0) + 1
        return {
            "pipe_map": pipe_map,
            "outgoing_by_node": outgoing_by_node,
            "nozzle_count_by_node": nozzle_count_by_node,
        }

    def _count_downstream_nozzles(
        self,
        node_id: str,
        topology: dict,
        memo: dict[str, int],
        stack: set[str],
        ambiguous_nodes: set[str],
    ) -> int:
        if node_id in memo:
            return memo[node_id]
        if node_id in stack:
            ambiguous_nodes.add(node_id)
            raise ValueError(f"cycle-detected:{node_id}")
        stack.add(node_id)
        total = topology["nozzle_count_by_node"].get(node_id, 0)
        for pipe_label in topology["outgoing_by_node"].get(node_id, []):
            pipe = topology["pipe_map"].get(pipe_label)
            if pipe is None:
                ambiguous_nodes.add(node_id)
                continue
            total += self._count_downstream_nozzles(pipe.output_node, topology, memo, stack, ambiguous_nodes)
        stack.remove(node_id)
        memo[node_id] = total
        return total

    def _node_is_cross_split(
        self,
        node_id: str,
        topology: dict,
        nozzle_memo: dict[str, int],
        ambiguous_nodes: set[str],
    ) -> bool:
        multi_head_child_count = 0
        for pipe_label in topology["outgoing_by_node"].get(node_id, []):
            pipe = topology["pipe_map"].get(pipe_label)
            if pipe is None:
                ambiguous_nodes.add(node_id)
                continue
            count = self._count_downstream_nozzles(pipe.output_node, topology, nozzle_memo, set(), ambiguous_nodes)
            if count > 1:
                multi_head_child_count += 1
        return multi_head_child_count >= 2

    def _subtree_has_cross_split(
        self,
        pipe_label: int,
        topology: dict,
        nozzle_memo: dict[str, int],
        cross_memo: dict[str, bool],
        visiting_nodes: set[str],
        ambiguous_nodes: set[str],
    ) -> bool:
        pipe = topology["pipe_map"].get(pipe_label)
        if pipe is None:
            ambiguous_nodes.add(str(pipe_label))
            raise ValueError(f"missing-pipe:{pipe_label}")
        return self._subtree_node_has_cross_split(
            pipe.output_node,
            topology,
            nozzle_memo,
            cross_memo,
            visiting_nodes,
            ambiguous_nodes,
        )

    def _subtree_node_has_cross_split(
        self,
        node_id: str,
        topology: dict,
        nozzle_memo: dict[str, int],
        cross_memo: dict[str, bool],
        visiting_nodes: set[str],
        ambiguous_nodes: set[str],
    ) -> bool:
        if node_id in cross_memo:
            return cross_memo[node_id]
        if node_id in visiting_nodes:
            ambiguous_nodes.add(node_id)
            raise ValueError(f"cycle-cross:{node_id}")
        visiting_nodes.add(node_id)
        result = self._node_is_cross_split(node_id, topology, nozzle_memo, ambiguous_nodes)
        if not result:
            for pipe_label in topology["outgoing_by_node"].get(node_id, []):
                pipe = topology["pipe_map"].get(pipe_label)
                if pipe is None:
                    ambiguous_nodes.add(node_id)
                    continue
                if self._subtree_node_has_cross_split(
                    pipe.output_node,
                    topology,
                    nozzle_memo,
                    cross_memo,
                    visiting_nodes,
                    ambiguous_nodes,
                ):
                    result = True
                    break
        visiting_nodes.remove(node_id)
        cross_memo[node_id] = result
        return result

    def _build_pipe_velocity_check_rows(
        self,
        pipe_config_rows: list[PipeConfigRow],
        nozzle_config_rows: list[NozzleConfigRow],
        pipe_rows: list[PipeFlowRow],
        designed_pipe_rows: list[DesignedPipeRow],
    ) -> list[PipeVelocityCheckRow]:
        topology = self._build_pipe_topology(pipe_config_rows, nozzle_config_rows)
        nozzle_memo: dict[str, int] = {}
        cross_memo: dict[str, bool] = {}
        ambiguous_nodes: set[str] = set()
        designed_map = {row.label: row for row in designed_pipe_rows}
        checks: list[PipeVelocityCheckRow] = []
        for pipe in pipe_rows:
            role = "review"
            velocity_limit = None
            velocity_ok = None
            reason = "topology ambiguity로 branch/other 판정이 확정되지 않았습니다."
            downstream_nozzle_count = 0
            subtree_has_cross_split = False
            try:
                downstream_nozzle_count = self._count_downstream_nozzles(
                    pipe.output_node, topology, nozzle_memo, set(), ambiguous_nodes
                )
                subtree_has_cross_split = self._subtree_has_cross_split(
                    pipe.label, topology, nozzle_memo, cross_memo, set(), ambiguous_nodes
                )
                if pipe.nominal_bore_mm <= 50.0 and not subtree_has_cross_split:
                    role = "branch"
                    velocity_limit = 6.0
                    velocity_ok = pipe.velocity_mps <= velocity_limit
                    reason = (
                        f"50A 이하이고 downstream subtree에 multi-head split node가 존재하지 않아 가지배관으로 판정"
                    )
                else:
                    role = "other"
                    velocity_limit = 10.0
                    velocity_ok = pipe.velocity_mps <= velocity_limit
                    if pipe.nominal_bore_mm > 50.0:
                        reason = "구경이 50A를 초과하므로 그 밖의 배관으로 판정"
                    else:
                        reason = (
                            f"50A 이하이지만 downstream subtree에 multi-head split node가 존재하여 교차배관으로 판정"
                        )
            except ValueError:
                role = "review"
                velocity_limit = None
                velocity_ok = None
                reason = "그래프가 비정상적이거나 topology ambiguity가 있어 branch/other 판정이 확정되지 않았습니다."
            designed = designed_map.get(pipe.label)
            checks.append(
                PipeVelocityCheckRow(
                    label=pipe.label,
                    input_node=pipe.input_node,
                    output_node=pipe.output_node,
                    nominal_bore_mm=pipe.nominal_bore_mm,
                    actual_bore_mm=designed.actual_bore_mm if designed else None,
                    velocity_mps=pipe.velocity_mps,
                    downstream_nozzle_count=downstream_nozzle_count,
                    subtree_has_cross_split=subtree_has_cross_split,
                    pipe_role=role,
                    velocity_limit_mps=velocity_limit,
                    velocity_ok=velocity_ok,
                    role_reason=reason,
                )
            )
        return checks

    def _find_velocity_fail_ids(self, rows: list[PipeVelocityCheckRow]) -> set[int]:
        return {row.label for row in rows if row.velocity_ok is False}

    def _build_pipe_velocity_messages(self, rows: list[PipeVelocityCheckRow], results: dict[str, list[str]]) -> None:
        if not rows:
            results["FAIL"].append("배관 유속 데이터를 찾지 못했습니다. FLOW IN PIPES 섹션을 확인해 주세요.")
            return
        branch_rows = [row for row in rows if row.pipe_role == "branch" and row.velocity_ok is not None]
        other_rows = [row for row in rows if row.pipe_role == "other" and row.velocity_ok is not None]
        review_rows = [row for row in rows if row.pipe_role == "review"]
        failed_rows = [row for row in rows if row.velocity_ok is False]

        if review_rows:
            results["WARNING"].append(
                f"배관 유속 분류 경고: {len(review_rows)}개 배관은 topology ambiguity로 branch/other 판정이 확정되지 않았습니다."
            )

        if failed_rows:
            first = failed_rows[0]
            role_label = "가지배관" if first.pipe_role == "branch" else "그 밖의 배관"
            results["FAIL"].append(
                f"배관 유속 검증 부적합: {len(failed_rows)}개 배관이 기준을 초과했습니다. Pipe {first.label}({role_label}, {first.velocity_mps:.3f} m/s > {first.velocity_limit_mps:.3f} m/s)"
            )
            results["FAIL"].append("개선 가이드: 유속 초과 구간은 배관 구경 상향 또는 피팅/특수설비에 따른 마찰손실 저감으로 조정하세요.")
        else:
            results["PASS"].append(
                f"배관 유속 검증 적합: topology 기준 가지배관 {len(branch_rows)}개, 그 밖의 배관 {len(other_rows)}개를 판정했고 모두 기준 이내입니다."
            )
            results["PASS"].append(
                f"가지배관 최대 유속은 {max((row.velocity_mps for row in branch_rows), default=0.0):.3f} m/s, 그 밖의 배관 최대 유속은 {max((row.velocity_mps for row in other_rows), default=0.0):.3f} m/s 입니다."
            )

    def _parse_report_titles(self, text: str) -> dict:
        main_match = re.search(r"Results for\s*:\s*(.+)", text)
        zone_match = re.search(r"Results for\s*:.*?\n([A-Za-z0-9_ -]+)", text, re.S)
        return {
            "report_main_title": main_match.group(1).strip() if main_match else None,
            "report_zone_title": zone_match.group(1).strip() if zone_match else None,
        }

    def _parse_sdf_network(self, path: Path | None) -> dict:
        if path is None or not path.exists():
            return {}
        root = ET.parse(path).getroot()
        titles = [item.text.strip() for item in root.findall(".//Title") if item.text and item.text.strip()]
        nodes: dict[str, dict] = {}
        for node in root.findall(".//Node"):
            label = node.attrib.get("label")
            if not label:
                continue
            pos = node.find("Position")
            x = y = None
            if pos is not None:
                try:
                    x = float(pos.attrib.get("x", "0"))
                    y = float(pos.attrib.get("y", "0"))
                except ValueError:
                    x = y = None
            nodes[label] = {"label": label, "x": x, "y": y}

        pipes: list[SdfPipeRow] = []
        for pipe in root.findall(".//Pipe"):
            label_raw = pipe.attrib.get("label")
            label = int(label_raw) if label_raw and label_raw.isdigit() else None
            input_node = pipe.attrib.get("input", "")
            output_node = pipe.attrib.get("output", "")
            waypoints: list[tuple[float, float]] = []
            for wp in pipe.findall(".//Waypoints/Position"):
                try:
                    waypoints.append((float(wp.attrib.get("x", "0")), float(wp.attrib.get("y", "0"))))
                except ValueError:
                    continue
            rise_m = elevation_m = None
            for key in ("rise", "elevation", "height"):
                raw = pipe.attrib.get(key)
                if raw:
                    try:
                        value = float(raw)
                        if key == "rise":
                            rise_m = value
                        else:
                            elevation_m = value
                    except ValueError:
                        pass
            pipes.append(
                SdfPipeRow(
                    label=label,
                    input_node=input_node,
                    output_node=output_node,
                    waypoints=waypoints,
                    rise_m=rise_m,
                    elevation_m=elevation_m,
                )
            )

        nozzle_inputs: list[str] = []
        for nozzle in root.findall(".//Nozzle"):
            input_node = nozzle.attrib.get("input")
            if input_node:
                nozzle_inputs.append(input_node)
        return {
            "main_title": titles[0] if len(titles) >= 1 else None,
            "zone_title": titles[1] if len(titles) >= 2 else None,
            "nodes": nodes,
            "pipes": pipes,
            "nozzle_inputs": nozzle_inputs,
            "equipment_count": len(root.findall(".//Equipment")),
        }

    def _parse_sdf(self, path: Path | None, network: dict | None = None) -> dict:
        if path is None:
            return {}
        network = network or self._parse_sdf_network(path)
        return {
            "sdf_main_title": network.get("main_title"),
            "sdf_zone_title": network.get("zone_title"),
            "pipe_count": len(network.get("pipes", [])),
            "nozzle_count": len(network.get("nozzle_inputs", [])),
            "equipment_count": network.get("equipment_count", 0),
        }

    def _build_common_hw_messages(self, hw_declared: bool, hw_checks: list[HwCheckRow], results: dict[str, list[str]]) -> None:
        if hw_declared:
            results["PASS"].append("공통-하겐윌리엄 선언 확인: 결과서에 Using the Hazen-Williams Equation 문구가 있습니다.")
        else:
            results["FAIL"].append("공통-하겐윌리엄 선언 확인 실패: 결과서에 Using the Hazen-Williams Equation 문구가 없습니다.")

        if not hw_checks:
            results["FAIL"].append("공통-마찰손실 재계산 실패: HW 재계산에 필요한 배관 데이터가 부족합니다.")
            return

        failed = [row for row in hw_checks if not row.hw_ok]
        if failed:
            labels = ", ".join(f"{row.label}번" for row in failed[:8])
            results["FAIL"].append(
                f"공통-마찰손실 재계산 부적합: {len(failed)}개 배관이 허용오차를 벗어났습니다. 주요 배관: {labels}"
            )
        else:
            results["PASS"].append(
                f"공통-마찰손실 재계산 적합: {len(hw_checks)}/{len(hw_checks)}개 배관이 허용오차 이내입니다."
            )

    def _find_nozzle_fail_ids(self, rows: list[NozzleFlowRow]) -> set[int]:
        fail_ids: set[int] = set()
        for row in rows:
            if row.actual_flow_lpm < self.MIN_HEAD_FLOW:
                fail_ids.add(row.label)
            if row.inlet_pressure_kgf_cm2 < self.MIN_HEAD_PRESSURE or row.inlet_pressure_kgf_cm2 > self.MAX_HEAD_PRESSURE:
                fail_ids.add(row.label)
        return fail_ids

    def _find_equipment_issue_ids(self, rows: list[EquipmentRow]) -> tuple[set[int], set[int]]:
        fail_ids: set[int] = set()
        warn_ids: set[int] = set()
        for row in rows:
            desc = normalize_desc(row.description)
            if desc == "FX" and not (self.FX_EQ_MIN <= row.equivalent_length_m <= self.FX_EQ_MAX):
                fail_ids.add(row.label)
            if desc == "AV" and abs(row.equivalent_length_m - self.AV_EQ_REF) > 0.1:
                fail_ids.add(row.label)
            if desc == "PV" and abs(row.equivalent_length_m - self.PV_EQ_REF) > 0.1:
                fail_ids.add(row.label)
        if not any(normalize_desc(row.description) == "PV" for row in rows):
            warn_ids.add(-1)
        return fail_ids, warn_ids

    def _find_valve_fail_ids(self, rows: list[ElastomericValveRow]) -> set[int]:
        fail_ids: set[int] = set()
        for row in rows:
            expected_drop = row.inlet_pressure_kgf_cm2 - row.outlet_pressure_kgf_cm2
            if abs(expected_drop - row.pressure_drop_kgf_cm2) > self.VALVE_DROP_TOLERANCE:
                fail_ids.add(row.label)
        return fail_ids

    def _build_optimization_insights(
        self,
        pipe_rows: list[PipeFlowRow],
        pipe_lengths: dict[int, float],
        equipment_rows: list[EquipmentRow],
        pipe_config_rows: list[PipeConfigRow],
        designed_pipe_rows: list[DesignedPipeRow],
        design_material_rows: list[DesignMaterialRow],
        hw_checks: list[HwCheckRow],
        velocity_checks: list[PipeVelocityCheckRow],
        nozzle_rows: list[NozzleFlowRow],
    ) -> dict:
        engineering_advice: list[str] = []
        economy_guide: list[str] = []
        engineering_pipe_ids: set[int] = set()
        economy_pipe_ids: set[int] = set()
        economy_equipment_ids: set[int] = set()
        engineering_reason_map: dict[int, list[str]] = {}
        economy_reason_map: dict[int, list[str]] = {}

        def add_engineering_reason(label: int, reason: str) -> None:
            engineering_pipe_ids.add(label)
            engineering_reason_map.setdefault(label, [])
            if reason not in engineering_reason_map[label]:
                engineering_reason_map[label].append(reason)

        def add_economy_reason(label: int, reason: str) -> None:
            economy_pipe_ids.add(label)
            economy_reason_map.setdefault(label, [])
            if reason not in economy_reason_map[label]:
                economy_reason_map[label].append(reason)

        pipe_map = {row.label: row for row in pipe_rows}
        designed_map = {row.label: row for row in designed_pipe_rows}
        material_map = {row.pipe_type_id: row.material_name for row in design_material_rows}
        hw_map = {row.label: row for row in hw_checks}
        velocity_map = {row.label: row for row in velocity_checks}

        engineering_advice.append(
            "최적화 전제: 헤드 최소 방수압 0.1MPa(1.0kg/cm²G), 헤드 유량, 배관 유속, C-Factor, 등가길이, Hazen-Williams 검산을 먼저 만족한 뒤 공학/경제 최적화를 검토합니다."
        )
        engineering_advice.append(
            "공학 관점: 비용 증가 가능성을 감수하더라도 마찰손실 집중, 유속 여유 부족, 압력 안정성 부족, 피팅/특수설비 손실 집중을 줄이는 방향입니다."
        )
        economy_guide.append(
            "경제성 전제: 기준을 위반하면서 관경을 줄이는 것이 아니라, 법적/기술 조건을 유지하는 범위에서 과대 관경, 저유속, 대구경 밸브, CPVC 대구경 구간을 줄이는 방향입니다."
        )

        friction_spikes: list[tuple[int, float, float, float]] = []
        change_spikes: list[tuple[int, float, float]] = []
        tight_velocity_candidates: list[tuple[int, float, float, float]] = []
        fitting_concentration: list[tuple[int, float, float, float]] = []
        low_pressure_candidates: list[tuple[int, float]] = []
        low_velocity_candidates: list[tuple[int, float, float]] = []
        main_downsize_candidates: list[tuple[int, float, float]] = []
        branch_downsize_candidates: list[tuple[int, float, float]] = []
        cpvc_large_candidates: list[tuple[int, float, str]] = []
        pressure_surplus_nozzles: list[tuple[int, float]] = []
        unit_losses: list[tuple[int, float]] = []

        for row in pipe_rows:
            length_m = pipe_lengths.get(row.label)
            if length_m and length_m > 0:
                unit_loss = row.friction_loss / length_m
                unit_losses.append((row.label, unit_loss))
                if unit_loss > self.FRICTION_SPIKE:
                    add_engineering_reason(row.label, "마찰손실")
                    friction_spikes.append((row.label, unit_loss, row.friction_loss, length_m))

            velocity_check = velocity_map.get(row.label)
            if velocity_check and velocity_check.velocity_limit_mps:
                velocity_margin = 1 - row.velocity_mps / max(velocity_check.velocity_limit_mps, 1e-9)
                if 0 <= velocity_margin < 0.10:
                    add_engineering_reason(row.label, "유속여유")
                    tight_velocity_candidates.append((row.label, row.velocity_mps, velocity_check.velocity_limit_mps, velocity_margin))
                if velocity_margin > 0.35 and row.nominal_bore_mm >= 80.0 and velocity_check.pipe_role == "other":
                    add_economy_reason(row.label, "주배관 관경축소")
                    main_downsize_candidates.append((row.label, row.nominal_bore_mm, row.velocity_mps))
                if velocity_margin > 0.35 and row.nominal_bore_mm in {32.0, 40.0, 50.0} and velocity_check.pipe_role == "branch":
                    add_economy_reason(row.label, "가지배관 관경축소")
                    branch_downsize_candidates.append((row.label, row.nominal_bore_mm, row.velocity_mps))

            if row.velocity_mps < self.ECONOMY_V_LOW and row.nominal_bore_mm > 25.0:
                add_economy_reason(row.label, "저유속 과설계")
                low_velocity_candidates.append((row.label, row.nominal_bore_mm, row.velocity_mps))

            if row.outlet_pressure < self.MIN_HEAD_PRESSURE + 0.2:
                add_engineering_reason(row.label, "압력여유")
                low_pressure_candidates.append((row.label, row.outlet_pressure))

            hw = hw_map.get(row.label)
            if hw and hw.total_length_m > 0:
                eq_ratio = (hw.fitting_eq_length_m + hw.special_eq_length_m) / hw.total_length_m
                if eq_ratio > 0.30:
                    add_engineering_reason(row.label, "피팅집중")
                    fitting_concentration.append((row.label, eq_ratio, hw.fitting_eq_length_m, hw.special_eq_length_m))

            designed = designed_map.get(row.label)
            material = material_map.get(designed.pipe_type_id, "") if designed else ""
            if material.upper().startswith("CPVC") and row.nominal_bore_mm > 65.0:
                add_economy_reason(row.label, "CPVC 대구경")
                cpvc_large_candidates.append((row.label, row.nominal_bore_mm, material))

        unit_losses.sort(key=lambda x: x[0])
        for idx in range(1, len(unit_losses)):
            label, unit_loss = unit_losses[idx]
            _, prev_loss = unit_losses[idx - 1]
            if prev_loss <= 0:
                continue
            change_rate = (unit_loss - prev_loss) / prev_loss
            if change_rate > 0.50 and unit_loss > 0.10:
                add_engineering_reason(label, "변화율")
                change_spikes.append((label, change_rate, unit_loss))

        for nozzle in nozzle_rows:
            if nozzle.inlet_pressure_kgf_cm2 >= 1.30:
                pressure_surplus_nozzles.append((nozzle.label, nozzle.inlet_pressure_kgf_cm2))

        if friction_spikes:
            friction_spikes.sort(key=lambda x: x[1], reverse=True)
            top_items = ", ".join(
                f"Pipe {label}(m당 {unit_loss:.3f}, 손실 {loss:.3f}, 길이 {length:.2f}m)"
                for label, unit_loss, loss, length in friction_spikes[:8]
            )
            engineering_advice.append(
                f"m당 마찰손실 기준({self.FRICTION_SPIKE:.2f} kg/cm²/m) 초과 배관이 {len(friction_spikes)}개 있습니다. 주요 구간: {top_items}"
            )
            engineering_advice.append("조치: 절대 손실값이 높은 구간은 구경 상향, 피팅 축소, 특수설비 위치 조정, 배관 경로 단순화를 우선 검토합니다.")

        if change_spikes:
            change_spikes.sort(key=lambda x: x[1], reverse=True)
            top_items = ", ".join(f"Pipe {label}(변화율 {rate*100:.1f}%, m당 {unit_loss:.3f})" for label, rate, unit_loss in change_spikes[:8])
            engineering_advice.append(f"직전 배관 대비 마찰손실 변화율 급증 후보가 {len(change_spikes)}개 있습니다. 주요 구간: {top_items}")
            engineering_advice.append("조치: 짧은 배관이면 피팅/밸브/특수설비 등가길이 집중을 확인하고, 긴 배관이면 경로 분산 또는 루프/그리드 배치를 검토합니다.")

        if tight_velocity_candidates:
            tight_velocity_candidates.sort(key=lambda x: x[3])
            top_items = ", ".join(f"Pipe {label}({velocity:.2f}/{limit:.2f}m/s, 여유 {margin*100:.1f}%)" for label, velocity, limit, margin in tight_velocity_candidates[:8])
            engineering_advice.append(f"유속 기준 여유가 10% 미만인 배관이 {len(tight_velocity_candidates)}개 있습니다. 주요 구간: {top_items}")
            engineering_advice.append("조치: 유속 여유 부족 구간은 관경 상향, 분기 위치 조정, 경로 단순화로 안정성을 확보합니다.")

        if fitting_concentration:
            fitting_concentration.sort(key=lambda x: x[1], reverse=True)
            top_items = ", ".join(f"Pipe {label}(등가길이 비율 {ratio*100:.1f}%, 피팅 {fit:.2f}m, 특수 {special:.2f}m)" for label, ratio, fit, special in fitting_concentration[:8])
            engineering_advice.append(f"피팅/특수설비 등가길이 비율이 30%를 넘는 손실 집중 후보가 {len(fitting_concentration)}개 있습니다. 주요 구간: {top_items}")
            engineering_advice.append("조치: 90도 엘보, Tee, Cross, 신축배관, 밸브류가 몰린 구간은 경로 직선화와 분기 방식 단순화를 검토합니다.")

        if low_pressure_candidates:
            low_pressure_candidates.sort(key=lambda x: x[1])
            top_items = ", ".join(f"Pipe {label}(출구압 {pressure:.3f}kg/cm²G)" for label, pressure in low_pressure_candidates[:8])
            engineering_advice.append(f"출구압력이 1.2kg/cm²G 미만인 압력 여유 부족 후보가 {len(low_pressure_candidates)}개 있습니다. 주요 구간: {top_items}")
            engineering_advice.append("조치: 말단 압력 안정성을 위해 해당 하류 계통의 관경 상향 또는 마찰손실 저감 조치를 우선 검토합니다.")

        engineering_advice.append("절충 원칙: 공학안은 손실과 유속을 안정화하지만 비용 증가가 따르므로, 먼저 피팅/특수설비 집중과 배관 경로를 점검한 뒤 관경 상향을 결정합니다.")

        if low_velocity_candidates:
            low_velocity_candidates.sort(key=lambda x: x[2])
            top_items = ", ".join(f"Pipe {label}({bore:.0f}A/{vel:.2f}m/s)" for label, bore, vel in low_velocity_candidates[:8])
            economy_guide.append(
                f"저유속 과설계 후보가 {len(low_velocity_candidates)}개 있습니다. 조건: 유속 < {self.ECONOMY_V_LOW:.1f}m/s이고 구경 > 25A. 주요 구간: {top_items}"
            )
            economy_guide.append("조치: 법적 압력과 최소유량을 유지하는 범위에서 한 단계 구경 축소 시뮬레이션을 수행합니다.")

        if main_downsize_candidates:
            main_downsize_candidates.sort(key=lambda x: (-x[1], x[2]))
            top_items = ", ".join(f"Pipe {label}({bore:.0f}A/{vel:.2f}m/s)" for label, bore, vel in main_downsize_candidates[:8])
            economy_guide.append(f"주배관/교차배관 구경 축소 검토 후보가 {len(main_downsize_candidates)}개 있습니다. 125A→100A, 100A→80A, 80A→65A 순서로 검토합니다. 주요 구간: {top_items}")
            economy_guide.append("확인조건: 그 밖의 배관 유속 10m/s 이하, 말단 방수압 0.1MPa 이상, 전체 압력 균형, 밸브/특수설비 손실 증가 영향을 재계산해야 합니다.")

        if branch_downsize_candidates:
            branch_downsize_candidates.sort(key=lambda x: (-x[1], x[2]))
            top_items = ", ".join(f"Pipe {label}({bore:.0f}A/{vel:.2f}m/s)" for label, bore, vel in branch_downsize_candidates[:8])
            economy_guide.append(f"가지배관 구경 축소 검토 후보가 {len(branch_downsize_candidates)}개 있습니다. 40A→32A, 32A→25A 방향으로 물량이 많은 말단부를 우선 검토합니다. 주요 구간: {top_items}")
            economy_guide.append("확인조건: topology 기준 가지배관 유속 6m/s 이하, 말단 방수압 0.1MPa 이상, 헤드 유량 기준을 다시 만족해야 합니다.")

        if pressure_surplus_nozzles:
            pressure_surplus_nozzles.sort(key=lambda x: x[1], reverse=True)
            top_items = ", ".join(f"Head {label}({pressure:.3f}kg/cm²G)" for label, pressure in pressure_surplus_nozzles[:8])
            economy_guide.append(f"헤드 압력 여유가 큰 후보가 {len(pressure_surplus_nozzles)}개 있습니다. 경제 목표는 약 1.1~1.2kg/cm²G 수준을 검토하되 최소 1.0kg/cm²G는 반드시 유지해야 합니다. 주요 헤드: {top_items}")
            economy_guide.append("조치: 압력 여유가 큰 계통은 관경 한 단계 축소를 가정하고 Hazen-Williams 재계산, 유속 기준, 헤드 유량을 재확인합니다.")

        if cpvc_large_candidates:
            cpvc_large_candidates.sort(key=lambda x: x[1], reverse=True)
            top_items = ", ".join(f"Pipe {label}({bore:.0f}A/{material})" for label, bore, material in cpvc_large_candidates[:8])
            economy_guide.append(f"CPVC 65A 초과 비용 검토 후보가 {len(cpvc_large_candidates)}개 있습니다. 주요 구간: {top_items}")
            economy_guide.append("조치: CPVC 80A 이상 단일 라인은 50A/65A 복수 라인 분산, 세대 인입부 65A 이하 유지 가능성을 검토합니다.")

        for item in equipment_rows:
            desc = normalize_desc(item.description)
            if desc not in {"AV", "PV", "VALVE"}:
                continue
            linked_pipe = pipe_map.get(item.pipe_label)
            if linked_pipe and linked_pipe.nominal_bore_mm > 100.0:
                economy_equipment_ids.add(item.label)
                economy_guide.append(
                    f"[밸브 계통 {item.label}] 연결 배관 구경이 {linked_pipe.nominal_bore_mm:.0f}A입니다. "
                    "수리 여유가 있으면 100A 이하 대안 검토로 밸브/부속 단가를 낮출 수 있습니다."
                )

        economy_guide.append("루프/그리드 검토: 하류 헤드 수가 많고 한쪽 계통에 유량이 집중된 구간은 양방향 공급으로 손실을 분산해 전체 관경을 줄일 수 있는지 검토합니다.")
        economy_guide.append("C-Factor 활용: C=150 계열 재질은 손실을 줄일 수 있으나 재료 단가가 오를 수 있으므로, 재질비 증가분과 관경/펌프양정/시공비 절감분을 함께 비교합니다.")

        if len(engineering_advice) <= 3:
            engineering_advice.append("현재 데이터 기준으로 즉시 눈에 띄는 마찰손실 급증/유속 여유 부족/압력 여유 부족 후보는 제한적입니다.")
        if len(economy_guide) <= 3:
            economy_guide.append("현재 데이터 기준으로 즉시 눈에 띄는 저유속 과설계 또는 대구경 비용 후보는 제한적입니다.")

        return {
            "engineering_advice": engineering_advice,
            "economy_guide": economy_guide,
            "engineering_pipe_ids": engineering_pipe_ids,
            "economy_pipe_ids": economy_pipe_ids,
            "economy_equipment_ids": economy_equipment_ids,
            "engineering_reason_map": engineering_reason_map,
            "economy_reason_map": economy_reason_map,
        }

    def _build_nozzle_messages(self, rows: list[NozzleFlowRow], results: dict[str, list[str]]) -> None:
        if not rows:
            results["FAIL"].append("헤드 데이터를 찾지 못했습니다. FLOW THROUGH NOZZLES 섹션을 확인해 주세요.")
            return
        min_flow = min(item.actual_flow_lpm for item in rows)
        min_pressure = min(item.inlet_pressure_kgf_cm2 for item in rows)
        max_pressure = max(item.inlet_pressure_kgf_cm2 for item in rows)
        if any(item.actual_flow_lpm < self.MIN_HEAD_FLOW for item in rows):
            results["FAIL"].append(f"헤드 유량 미달 항목이 있습니다. 기준 {self.MIN_HEAD_FLOW:.1f} L/min 이상을 확인하세요.")
        else:
            results["PASS"].append(f"헤드 유량이 적합합니다. 최소 유량은 {min_flow:.2f} L/min 입니다.")
        if any(item.inlet_pressure_kgf_cm2 < self.MIN_HEAD_PRESSURE or item.inlet_pressure_kgf_cm2 > self.MAX_HEAD_PRESSURE for item in rows):
            results["FAIL"].append(
                f"헤드 압력 범위를 벗어난 항목이 있습니다. 허용범위는 {self.MIN_HEAD_PRESSURE:.1f}~{self.MAX_HEAD_PRESSURE:.1f} kg/cm2G 입니다."
            )
        else:
            results["PASS"].append(f"헤드 압력이 적합합니다. 범위는 {min_pressure:.3f}~{max_pressure:.3f} kg/cm2G 입니다.")
        required_count = self._infer_required_head_count(rows)
        if required_count is not None:
            if len(rows) >= required_count:
                results["PASS"].append(f"헤드 개수가 적합합니다. 결과는 {len(rows)}개, 요구 최소 {required_count}개입니다.")
            else:
                results["FAIL"].append(f"헤드 개수 부족입니다. 결과는 {len(rows)}개, 요구 최소 {required_count}개입니다.")

    def _build_equipment_messages(
        self,
        rows: list[EquipmentRow],
        fail_ids: set[int],
        warn_ids: set[int],
        results: dict[str, list[str]],
    ) -> None:
        if not rows:
            results["WARNING"].append("특수설비 데이터를 찾지 못했습니다. SPECIAL EQUIPMENT 섹션을 확인해 주세요.")
            return
        fx_rows = [item for item in rows if normalize_desc(item.description) == "FX"]
        av_rows = [item for item in rows if normalize_desc(item.description) == "AV"]
        pv_rows = [item for item in rows if normalize_desc(item.description) == "PV"]

        if any(item.label in fail_ids and normalize_desc(item.description) == "FX" for item in rows):
            bad_values = sorted({item.equivalent_length_m for item in fx_rows if item.label in fail_ids})
            results["FAIL"].append(
                f"FX 등가길이 부적합: 기준 {self.FX_EQ_MIN:.1f}~{self.FX_EQ_MAX:.1f} m, 현재 {', '.join(f'{v:.2f}' for v in bad_values)} m"
            )
        elif fx_rows:
            results["PASS"].append(f"FX 등가길이가 적합합니다. 총 {len(fx_rows)}개 항목을 확인했습니다.")

        if any(item.label in fail_ids and normalize_desc(item.description) == "AV" for item in rows):
            bad_av = next(item for item in rows if item.label in fail_ids and normalize_desc(item.description) == "AV")
            results["FAIL"].append(f"A/V 등가길이 부적합: 기준 {self.AV_EQ_REF:.1f} m, 현재 {bad_av.equivalent_length_m:.2f} m")
        elif av_rows:
            results["PASS"].append("A/V 등가길이가 적합합니다.")

        if any(item.label in fail_ids and normalize_desc(item.description) == "PV" for item in rows):
            bad_pv_values = sorted({item.equivalent_length_m for item in rows if item.label in fail_ids and normalize_desc(item.description) == "PV"})
            results["FAIL"].append(f"P/V 등가길이 부적합: 기준 {self.PV_EQ_REF:.1f} m, 현재 {', '.join(f'{v:.2f}' for v in bad_pv_values)} m")
        elif pv_rows:
            results["PASS"].append("P/V 등가길이가 적합합니다.")

        if -1 in warn_ids:
            results["WARNING"].append("P/V 항목이 결과서에 없습니다. 프리액션 존이 아니면 무시해도 됩니다.")

    def _build_valve_messages(self, rows: list[ElastomericValveRow], fail_ids: set[int], results: dict[str, list[str]]) -> None:
        if not rows:
            results["WARNING"].append("감압밸브 데이터를 찾지 못했습니다. FLOW THROUGH ELASTOMERIC VALVES 섹션을 확인해 주세요.")
            return
        if fail_ids:
            for row in rows:
                if row.label not in fail_ids:
                    continue
                expected_drop = row.inlet_pressure_kgf_cm2 - row.outlet_pressure_kgf_cm2
                results["FAIL"].append(
                    f"감압밸브 {row.label} 압력강하 불일치: 계산 {expected_drop:.3f}, 결과서 {row.pressure_drop_kgf_cm2:.3f}"
                )
        else:
            results["PASS"].append(f"감압밸브 압력강하 계산이 적합합니다. 총 {len(rows)}개 항목을 확인했습니다.")

    def _build_cross_check_messages(
        self,
        nozzles: list[NozzleFlowRow],
        equipment_rows: list[EquipmentRow],
        report_titles: dict,
        sdf_info: dict,
        results: dict[str, list[str]],
    ) -> None:
        report_zone = report_titles.get("report_zone_title")
        sdf_zone = sdf_info.get("sdf_zone_title")
        if report_zone and sdf_zone:
            if report_zone != sdf_zone:
                results["FAIL"].append(f"결과서 존({report_zone})과 SDF 존({sdf_zone})이 다릅니다.")
            else:
                results["PASS"].append(f"결과서와 SDF 존이 일치합니다. ({report_zone})")

        if "nozzle_count" in sdf_info:
            if sdf_info["nozzle_count"] != len(nozzles):
                results["FAIL"].append(f"SDF 헤드 개수({sdf_info['nozzle_count']})와 결과서 헤드 개수({len(nozzles)})가 다릅니다.")
            else:
                results["PASS"].append(f"SDF와 결과서 헤드 개수가 일치합니다. ({len(nozzles)}개)")

        if "equipment_count" in sdf_info and sdf_info["equipment_count"] != len(equipment_rows):
            results["WARNING"].append(
                f"SDF 특수설비 개수({sdf_info['equipment_count']})와 결과서 특수설비 개수({len(equipment_rows)})가 다릅니다."
            )

    def _build_stats(
        self,
        pipes: list[PipeFlowRow],
        nozzles: list[NozzleFlowRow],
        equipment: list[EquipmentRow],
        valves: list[ElastomericValveRow],
        report_titles: dict,
        sdf_info: dict,
        hw_declared: bool,
        hw_checks: list[HwCheckRow],
        velocity_checks: list[PipeVelocityCheckRow],
        pipe_stats: dict,
    ) -> dict:
        min_flow = min((item.actual_flow_lpm for item in nozzles), default=None)
        hw_failed = [row for row in hw_checks if not row.hw_ok]
        worst_row = max(hw_checks, key=lambda row: row.abs_diff, default=None)
        branch_velocity_rows = [row for row in velocity_checks if row.pipe_role == "branch" and row.velocity_ok is not None]
        other_velocity_rows = [row for row in velocity_checks if row.pipe_role == "other" and row.velocity_ok is not None]
        review_velocity_rows = [row for row in velocity_checks if row.pipe_role == "review"]
        failed_velocity_rows = [row for row in velocity_checks if row.velocity_ok is False]
        worst_velocity = max(failed_velocity_rows or velocity_checks, key=lambda row: row.velocity_mps, default=None)
        return {
            **report_titles,
            "pipe_count_from_report": len(pipes),
            "nozzle_count_from_report": len(nozzles),
            "equipment_count_from_report": len(equipment),
            "valve_count_from_report": len(valves),
            "min_nozzle_flow_lpm": min_flow,
            "min_nozzle_pressure_kgf_cm2": min((item.inlet_pressure_kgf_cm2 for item in nozzles), default=None),
            "max_branch_pipe_velocity_mps": max((item.velocity_mps for item in branch_velocity_rows), default=None),
            "max_other_pipe_velocity_mps": max((item.velocity_mps for item in other_velocity_rows), default=None),
            "velocity_branch_checked_count": len(branch_velocity_rows),
            "velocity_other_checked_count": len(other_velocity_rows),
            "velocity_review_count": len(review_velocity_rows),
            "velocity_failed_count": len(failed_velocity_rows),
            "worst_velocity_pipe_label": worst_velocity.label if worst_velocity else None,
            "hw_declared": hw_declared,
            "hw_checked_pipe_count": len(hw_checks),
            "hw_failed_pipe_count": len(hw_failed),
            "hw_max_abs_diff": max((row.abs_diff for row in hw_checks), default=None),
            "hw_max_rel_diff": max((row.rel_diff for row in hw_checks), default=None),
            "hw_worst_pipe_label": worst_row.label if worst_row else None,
            **pipe_stats,
            **sdf_info,
        }

    def _build_tables(
        self,
        pipe_rows: list[PipeFlowRow],
        pipe_lengths: dict[int, float],
        pipe_config_rows: list[PipeConfigRow],
        designed_pipe_rows: list[DesignedPipeRow],
        hw_checks: list[HwCheckRow],
        velocity_checks: list[PipeVelocityCheckRow],
        pipe_validation_rows: list[PipeValidationRow],
        nozzle_rows: list[NozzleFlowRow],
        equipment_rows: list[EquipmentRow],
        valve_rows: list[ElastomericValveRow],
        pipe_fail_ids: set[int],
        hw_fail_ids: set[int],
        nozzle_fail_ids: set[int],
        equipment_fail_ids: set[int],
        equipment_warn_ids: set[int],
        valve_fail_ids: set[int],
        engineering_pipe_ids: set[int],
        economy_pipe_ids: set[int],
        economy_equipment_ids: set[int],
        engineering_reason_map: dict[int, list[str]] | None = None,
        economy_reason_map: dict[int, list[str]] | None = None,
    ) -> dict:
        pipe_config_map = {row.label: row for row in pipe_config_rows}
        designed_map = {row.label: row for row in designed_pipe_rows}
        hw_map = {row.label: row for row in hw_checks}
        velocity_map = {row.label: row for row in velocity_checks}
        pipe_validation_map = {row.label: row for row in pipe_validation_rows}
        engineering_reason_map = engineering_reason_map or {}
        economy_reason_map = economy_reason_map or {}
        return {
            "nozzles": [
                {
                    "label": item.label,
                    "input_node": item.input_node,
                    "inlet_pressure_kgf_cm2": item.inlet_pressure_kgf_cm2,
                    "required_flow_lpm": item.required_flow_lpm,
                    "actual_flow_lpm": item.actual_flow_lpm,
                    "deviation_percent": item.deviation_percent,
                    "min_flow_limit_lpm": self.MIN_HEAD_FLOW,
                    "min_pressure_limit_kgf_cm2": self.MIN_HEAD_PRESSURE,
                    "max_pressure_limit_kgf_cm2": self.MAX_HEAD_PRESSURE,
                    "highlight": item.label in nozzle_fail_ids,
                }
                for item in nozzle_rows
            ],
            "pipes": [
                {
                    "label": item.label,
                    "input_node": item.input_node,
                    "output_node": item.output_node,
                    "nominal_bore_mm": item.nominal_bore_mm,
                    "material_name": pipe_validation_map[item.label].material_name if item.label in pipe_validation_map else None,
                    "pipe_type_id": designed_map[item.label].pipe_type_id if item.label in designed_map else None,
                    "c_factor": pipe_config_map[item.label].c_factor if item.label in pipe_config_map else None,
                    "base_length_m": pipe_config_map[item.label].length_m if item.label in pipe_config_map else pipe_lengths.get(item.label),
                    "fitting_eq_length_m": pipe_config_map[item.label].fitting_eq_length_m if item.label in pipe_config_map else None,
                    "special_eq_length_m": hw_map[item.label].special_eq_length_m if item.label in hw_map else 0.0,
                    "total_length_m": hw_map[item.label].total_length_m if item.label in hw_map else None,
                    "actual_bore_mm": designed_map[item.label].actual_bore_mm if item.label in designed_map else None,
                    "inlet_pressure": item.inlet_pressure,
                    "outlet_pressure": item.outlet_pressure,
                    "max_pressure_kgcm2": pipe_validation_map[item.label].max_pressure_kgcm2 if item.label in pipe_validation_map else None,
                    "pressure_drop": item.pressure_drop,
                    "friction_loss": item.friction_loss,
                    "hw_expected_friction_loss": hw_map[item.label].calculated_friction_loss if item.label in hw_map else None,
                    "hw_abs_diff": hw_map[item.label].abs_diff if item.label in hw_map else None,
                    "hw_rel_diff": hw_map[item.label].rel_diff if item.label in hw_map else None,
                    "hw_formula_ok": hw_map[item.label].hw_ok if item.label in hw_map else None,
                    "hw_fail": item.label in hw_fail_ids,
                    "flow_lpm": item.flow_lpm,
                    "velocity_mps": item.velocity_mps,
                    "pipe_length_m": pipe_lengths.get(item.label),
                    "velocity_limit_mps": velocity_map[item.label].velocity_limit_mps if item.label in velocity_map else None,
                    "velocity_ok": velocity_map[item.label].velocity_ok if item.label in velocity_map else None,
                    "downstream_nozzle_count": velocity_map[item.label].downstream_nozzle_count if item.label in velocity_map else None,
                    "subtree_has_cross_split": velocity_map[item.label].subtree_has_cross_split if item.label in velocity_map else None,
                    "pipe_role": velocity_map[item.label].pipe_role if item.label in velocity_map else "review",
                    "pipe_type": velocity_map[item.label].pipe_role if item.label in velocity_map else "review",
                    "role_reason": velocity_map[item.label].role_reason if item.label in velocity_map else "topology 판정 정보가 없습니다.",
                    "pipe_rule_results": {
                        key: self._rule_result_to_dict(value)
                        for key, value in (pipe_validation_map[item.label].pipe_rule_results.items() if item.label in pipe_validation_map else [])
                    },
                    "rule_high_pressure_ok": pipe_validation_map[item.label].pipe_rule_results.get("PIPE.002").status if item.label in pipe_validation_map and pipe_validation_map[item.label].pipe_rule_results.get("PIPE.002") else None,
                    "rule_cfactor_ok": pipe_validation_map[item.label].pipe_rule_results.get("PIPE.003").status if item.label in pipe_validation_map and pipe_validation_map[item.label].pipe_rule_results.get("PIPE.003") else None,
                    "rule_unit_internal_cpvc_status": pipe_validation_map[item.label].pipe_rule_results.get("PIPE.004").status if item.label in pipe_validation_map and pipe_validation_map[item.label].pipe_rule_results.get("PIPE.004") else None,
                    "rule_unit_inlet_status": pipe_validation_map[item.label].pipe_rule_results.get("PIPE.005").status if item.label in pipe_validation_map and pipe_validation_map[item.label].pipe_rule_results.get("PIPE.005") else None,
                    "rule_headcount_size_status": pipe_validation_map[item.label].pipe_rule_results.get("PIPE.006").status if item.label in pipe_validation_map and pipe_validation_map[item.label].pipe_rule_results.get("PIPE.006") else None,
                    "friction_spike_limit": self.FRICTION_SPIKE,
                    "economy_velocity_limit": self.ECONOMY_V_LOW,
                    "special_equipment": item.has_special_equipment,
                    "highlight": item.label in pipe_fail_ids or item.label in hw_fail_ids,
                    "warn": False,
                    "engineering_flag": item.label in engineering_pipe_ids,
                    "economy_flag": item.label in economy_pipe_ids,
                    "engineering_reasons": ", ".join(engineering_reason_map.get(item.label, [])),
                    "economy_reasons": ", ".join(economy_reason_map.get(item.label, [])),
                }
                for item in pipe_rows
            ],
            "equipment": [
                {
                    "label": item.label,
                    "pipe_label": item.pipe_label,
                    "equivalent_length_m": item.equivalent_length_m,
                    "description": item.description,
                    "fx_eq_min_m": self.FX_EQ_MIN,
                    "fx_eq_max_m": self.FX_EQ_MAX,
                    "av_eq_ref_m": self.AV_EQ_REF,
                    "pv_eq_ref_m": self.PV_EQ_REF,
                    "eq_tolerance_m": 0.1,
                    "highlight": item.label in equipment_fail_ids,
                    "warn": -1 in equipment_warn_ids and normalize_desc(item.description) == "PV",
                    "economy_flag": item.label in economy_equipment_ids,
                }
                for item in equipment_rows
            ],
            "valves": [
                {
                    "label": item.label,
                    "inlet_pressure_kgf_cm2": item.inlet_pressure_kgf_cm2,
                    "outlet_pressure_kgf_cm2": item.outlet_pressure_kgf_cm2,
                    "pressure_drop_kgf_cm2": item.pressure_drop_kgf_cm2,
                    "flow_lpm": item.flow_lpm,
                    "calculated_pressure_drop_kgf_cm2": item.inlet_pressure_kgf_cm2 - item.outlet_pressure_kgf_cm2,
                    "pressure_drop_tolerance_kgf_cm2": self.VALVE_DROP_TOLERANCE,
                    "highlight": item.label in valve_fail_ids,
                }
                for item in valve_rows
            ],
        }

    def _build_report(self, results: dict[str, list[str]], insights: dict, stats: dict) -> str:
        lines = [
            "PIPENET 수리계산 검증 결과",
            "=" * 60,
            f"결과서 파일: {self.report_path.name}",
        ]
        if self.sdf_path:
            lines.append(f"SDF 파일: {self.sdf_path.name}")
        lines.append("")
        lines.append("[종합 요약]")
        if results["FAIL"]:
            lines.append(f"- 최종 판정: 부적합 ({len(results['FAIL'])}건)")
        elif results["WARNING"]:
            lines.append(f"- 최종 판정: 조건부 적합 ({len(results['WARNING'])}건 확인 필요)")
        else:
            lines.append("- 최종 판정: 적합")
        lines.append(f"- 적합: {len(results['PASS'])}건")
        lines.append(f"- 부적합: {len(results['FAIL'])}건")
        lines.append(f"- 확인 필요: {len(results['WARNING'])}건")
        lines.append("")
        lines.append("[공학적 마찰손실 최적화 원칙]")
        for item in insights["engineering_advice"]:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("[시공사 경제성 확보 방안]")
        for item in insights["economy_guide"]:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("[검증 통계]")
        lines.append(json.dumps(stats, ensure_ascii=False, indent=2))
        for key, title in (("FAIL", "[부적합 항목]"), ("WARNING", "[확인 필요 항목]"), ("PASS", "[적합 항목]")):
            lines.append("")
            lines.append(title)
            for message in results[key] or ["없음"]:
                lines.append(f"- {message}")
        return "\n".join(lines)

    def _infer_required_head_count(self, rows: list[NozzleFlowRow]) -> int | None:
        if not rows:
            return None
        if len(rows) >= 30:
            return 30
        if len(rows) >= 10:
            return 10
        return None
