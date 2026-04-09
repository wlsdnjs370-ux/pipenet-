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
class PipeConfigRow:
    label: int
    input_node: str
    output_node: str
    nominal_bore_mm: float
    length_m: float
    elevation_m: float
    c_factor: float
    fitting_eq_m: float


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
    FRICTION_SPIKE = 0.05
    HW_DECLARED_TEXT = "USING THE HAZEN-WILLIAMS EQUATION"

    def __init__(self, report_path: Path, sdf_path: Path | None = None):
        self.report_path = Path(report_path)
        self.sdf_path = Path(sdf_path) if sdf_path else None

    def validate(self) -> dict:
        report_text = self._read_report_text(self.report_path)
        design_info = self._parse_design_info(report_text)
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
        sdf_info = self._parse_sdf(self.sdf_path) if self.sdf_path else {}
        hw_declared = self._check_hw_declared(report_text, design_info)
        hw_checks = self._build_hw_check_rows(pipe_config_rows, designed_pipe_rows, equipment_rows, pipe_rows)
        hw_fail_ids = {row.label for row in hw_checks if not row.hw_ok}

        pipe_fail_ids = self._find_pipe_fail_ids(pipe_rows)
        nozzle_fail_ids = self._find_nozzle_fail_ids(nozzle_rows)
        equipment_fail_ids, equipment_warn_ids = self._find_equipment_issue_ids(equipment_rows)
        valve_fail_ids = self._find_valve_fail_ids(valve_rows)

        insights = self._build_optimization_insights(
            pipe_rows=pipe_rows,
            pipe_lengths=pipe_lengths,
            equipment_rows=equipment_rows,
        )

        results: dict[str, list[str]] = {"PASS": [], "FAIL": [], "WARNING": []}
        self._build_common_hw_messages(hw_declared, hw_checks, results)
        self._build_nozzle_messages(nozzle_rows, results)
        self._build_pipe_messages(pipe_rows, pipe_fail_ids, results)
        self._build_equipment_messages(equipment_rows, equipment_fail_ids, equipment_warn_ids, results)
        self._build_valve_messages(valve_rows, valve_fail_ids, results)
        self._build_cross_check_messages(
            nozzles=nozzle_rows,
            equipment_rows=equipment_rows,
            report_titles=report_titles,
            sdf_info=sdf_info,
            results=results,
        )

        stats = self._build_stats(pipe_rows, nozzle_rows, equipment_rows, valve_rows, report_titles, sdf_info, hw_declared, hw_checks)
        tables = self._build_tables(
            pipe_rows=pipe_rows,
            pipe_lengths=pipe_lengths,
            pipe_config_rows=pipe_config_rows,
            designed_pipe_rows=designed_pipe_rows,
            hw_checks=hw_checks,
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
            "stats": stats,
            "tables": tables,
            "report": self._build_report(results, insights, stats),
            "parsed_context": {
                "pipe_config_count": len(pipe_config_rows),
                "designed_pipe_count": len(designed_pipe_rows),
                "nozzle_config_count": len(nozzle_config_rows),
                "pipe_fitting_count": len(pipe_fitting_rows),
                "flow_at_inlet_count": len(inlet_rows),
                "pump_flow_count": len(pump_rows),
                "flow_in_pipes_count": len(pipe_rows),
                "special_equipment_count": len(equipment_rows),
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

        pipe_materials: dict[int, str] = {}
        material_pattern = re.compile(rf"^\s*(\d+)\s+([A-Za-z0-9_.\-/]+)\s*$")
        for line in self._section_lines(text, "PIPE TYPES"):
            match = material_pattern.match(line)
            if match:
                pipe_materials[int(match.group(1))] = match.group(2)

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
                    fitting_eq_m=float(match.group(8)),
                )
            )
        return rows

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
        return {row.label: row.length_m for row in self._parse_pipe_config_rows(text)}

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
            total_length_m = config.length_m + config.fitting_eq_m + special_eq_length_m
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
                    fitting_eq_length_m=config.fitting_eq_m,
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

    def _parse_report_titles(self, text: str) -> dict:
        main_match = re.search(r"Results for\s*:\s*(.+)", text)
        zone_match = re.search(r"Results for\s*:.*?\n([A-Za-z0-9_ -]+)", text, re.S)
        return {
            "report_main_title": main_match.group(1).strip() if main_match else None,
            "report_zone_title": zone_match.group(1).strip() if zone_match else None,
        }

    def _parse_sdf(self, path: Path | None) -> dict:
        if path is None:
            return {}
        root = ET.parse(path).getroot()
        titles = root.findall(".//Title")
        return {
            "sdf_main_title": titles[0].text.strip() if len(titles) >= 1 and titles[0].text else None,
            "sdf_zone_title": titles[1].text.strip() if len(titles) >= 2 and titles[1].text else None,
            "pipe_count": len(root.findall(".//Pipe")),
            "nozzle_count": len(root.findall(".//Nozzle")),
            "equipment_count": len(root.findall(".//Equipment")),
        }

    def _find_pipe_fail_ids(self, rows: list[PipeFlowRow]) -> set[int]:
        fail_ids: set[int] = set()
        for row in rows:
            limit = self.BRANCH_PIPE_V_LIMIT if row.nominal_bore_mm <= self.BRANCH_PIPE_LIMIT_MM else self.MAIN_PIPE_V_LIMIT
            if row.velocity_mps > limit:
                fail_ids.add(row.label)
        return fail_ids

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
    ) -> dict:
        engineering_advice: list[str] = []
        economy_guide: list[str] = []
        engineering_pipe_ids: set[int] = set()
        economy_pipe_ids: set[int] = set()
        economy_equipment_ids: set[int] = set()

        pipe_map = {row.label: row for row in pipe_rows}

        friction_spikes: list[tuple[int, float]] = []
        low_velocity_candidates: list[tuple[int, float, float]] = []

        for row in pipe_rows:
            length_m = pipe_lengths.get(row.label)
            if length_m and length_m > 0:
                unit_loss = row.friction_loss / length_m
                if unit_loss > self.FRICTION_SPIKE:
                    engineering_pipe_ids.add(row.label)
                    friction_spikes.append((row.label, unit_loss))

            if row.velocity_mps < self.ECONOMY_V_LOW and row.nominal_bore_mm > 25.0:
                economy_pipe_ids.add(row.label)
                low_velocity_candidates.append((row.label, row.nominal_bore_mm, row.velocity_mps))

        if friction_spikes:
            friction_spikes.sort(key=lambda x: x[1], reverse=True)
            top_items = ", ".join(f"{label}번({loss:.3f})" for label, loss in friction_spikes[:8])
            engineering_advice.append(
                f"m당 마찰손실 급증 배관이 {len(friction_spikes)}개 있습니다. 주요 구간: {top_items}"
            )
            engineering_advice.append("가이드: 유속 저감(구경 상향), 피팅 수량 축소, 배관 경로 단순화를 우선 검토하세요.")

        if low_velocity_candidates:
            low_velocity_candidates.sort(key=lambda x: x[2])
            top_items = ", ".join(f"{label}번({bore:.0f}A/{vel:.2f}m/s)" for label, bore, vel in low_velocity_candidates[:8])
            economy_guide.append(
                f"저유속 과설계 후보 배관이 {len(low_velocity_candidates)}개 있습니다. 주요 구간: {top_items}"
            )
            economy_guide.append("가이드: 법적 압력/최소유량을 유지하는 범위에서 구경 축소 가능성을 검토하세요.")

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

        if not engineering_advice:
            engineering_advice.append("마찰손실 급증 구간은 발견되지 않았습니다.")
        if not economy_guide:
            economy_guide.append("경제성 저하가 우려되는 과설계 구간은 발견되지 않았습니다.")

        return {
            "engineering_advice": engineering_advice,
            "economy_guide": economy_guide,
            "engineering_pipe_ids": engineering_pipe_ids,
            "economy_pipe_ids": economy_pipe_ids,
            "economy_equipment_ids": economy_equipment_ids,
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

    def _build_pipe_messages(self, rows: list[PipeFlowRow], fail_ids: set[int], results: dict[str, list[str]]) -> None:
        if not rows:
            results["FAIL"].append("배관 유속 데이터를 찾지 못했습니다. FLOW IN PIPES 섹션을 확인해 주세요.")
            return
        if fail_ids:
            for row in rows:
                if row.label not in fail_ids:
                    continue
                limit = self.BRANCH_PIPE_V_LIMIT if row.nominal_bore_mm <= self.BRANCH_PIPE_LIMIT_MM else self.MAIN_PIPE_V_LIMIT
                pipe_type = "가지배관" if row.nominal_bore_mm <= self.BRANCH_PIPE_LIMIT_MM else "주배관"
                results["FAIL"].append(
                    f"{pipe_type} 유속 초과: Pipe {row.label}, {row.nominal_bore_mm:.0f}A, {row.velocity_mps:.3f} m/s (기준 {limit:.1f} m/s)"
                )
        else:
            max_branch = max((item.velocity_mps for item in rows if item.nominal_bore_mm <= self.BRANCH_PIPE_LIMIT_MM), default=0.0)
            max_main = max((item.velocity_mps for item in rows if item.nominal_bore_mm > self.BRANCH_PIPE_LIMIT_MM), default=0.0)
            results["PASS"].append(f"가지배관 유속이 적합합니다. 최대 {max_branch:.3f} m/s 입니다.")
            results["PASS"].append(f"주배관 유속이 적합합니다. 최대 {max_main:.3f} m/s 입니다.")

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
    ) -> dict:
        min_flow = min((item.actual_flow_lpm for item in nozzles), default=None)
        hw_failed = [row for row in hw_checks if not row.hw_ok]
        worst_row = max(hw_checks, key=lambda row: row.abs_diff, default=None)
        return {
            **report_titles,
            "pipe_count_from_report": len(pipes),
            "nozzle_count_from_report": len(nozzles),
            "equipment_count_from_report": len(equipment),
            "valve_count_from_report": len(valves),
            "min_nozzle_flow_lpm": min_flow,
            "min_nozzle_pressure_kgf_cm2": min((item.inlet_pressure_kgf_cm2 for item in nozzles), default=None),
            "max_branch_pipe_velocity_mps": max(
                (item.velocity_mps for item in pipes if item.nominal_bore_mm <= self.BRANCH_PIPE_LIMIT_MM),
                default=None,
            ),
            "max_main_pipe_velocity_mps": max(
                (item.velocity_mps for item in pipes if item.nominal_bore_mm > self.BRANCH_PIPE_LIMIT_MM),
                default=None,
            ),
            "hw_declared": hw_declared,
            "hw_checked_pipe_count": len(hw_checks),
            "hw_failed_pipe_count": len(hw_failed),
            "hw_max_abs_diff": max((row.abs_diff for row in hw_checks), default=None),
            "hw_max_rel_diff": max((row.rel_diff for row in hw_checks), default=None),
            "hw_worst_pipe_label": worst_row.label if worst_row else None,
            **sdf_info,
        }

    def _build_tables(
        self,
        pipe_rows: list[PipeFlowRow],
        pipe_lengths: dict[int, float],
        pipe_config_rows: list[PipeConfigRow],
        designed_pipe_rows: list[DesignedPipeRow],
        hw_checks: list[HwCheckRow],
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
    ) -> dict:
        pipe_config_map = {row.label: row for row in pipe_config_rows}
        designed_map = {row.label: row for row in designed_pipe_rows}
        hw_map = {row.label: row for row in hw_checks}
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
                    "c_factor": pipe_config_map[item.label].c_factor if item.label in pipe_config_map else None,
                    "base_length_m": pipe_config_map[item.label].length_m if item.label in pipe_config_map else pipe_lengths.get(item.label),
                    "fitting_eq_length_m": pipe_config_map[item.label].fitting_eq_m if item.label in pipe_config_map else None,
                    "special_eq_length_m": hw_map[item.label].special_eq_length_m if item.label in hw_map else 0.0,
                    "total_length_m": hw_map[item.label].total_length_m if item.label in hw_map else None,
                    "actual_bore_mm": designed_map[item.label].actual_bore_mm if item.label in designed_map else None,
                    "inlet_pressure": item.inlet_pressure,
                    "outlet_pressure": item.outlet_pressure,
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
                    "velocity_limit_mps": self.BRANCH_PIPE_V_LIMIT if item.nominal_bore_mm <= self.BRANCH_PIPE_LIMIT_MM else self.MAIN_PIPE_V_LIMIT,
                    "pipe_type": "branch" if item.nominal_bore_mm <= self.BRANCH_PIPE_LIMIT_MM else "main",
                    "friction_spike_limit": self.FRICTION_SPIKE,
                    "economy_velocity_limit": self.ECONOMY_V_LOW,
                    "special_equipment": item.has_special_equipment,
                    "highlight": item.label in pipe_fail_ids or item.label in hw_fail_ids,
                    "engineering_flag": item.label in engineering_pipe_ids,
                    "economy_flag": item.label in economy_pipe_ids,
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
