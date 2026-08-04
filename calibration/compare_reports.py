# -*- coding: utf-8 -*-
"""수작업본 ↔ 자동본 PIPENET 결과 리포트 대조기 (모듈 A / T0).

FNCADnet 작업지시서 §3 T0. 이후 모든 과제(T1~T9)의 성적표 역할을 한다.
축: 관종 / 표고 / 부속 / 특수설비 / 관경 / 결과.

CLI::

    python -m calibration.compare_reports <수작업.docx> <자동.docx> [--json OUT] [--md OUT]

'모름'을 0으로 대체하지 않는다. 섹션 자체가 없으면 지표는 None 이고
`unresolved` 에 사유가 남는다.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from core.pipenet_validator import PipenetGuideValidator  # noqa: E402

NUM = r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:E[-+]?\d+)?"
# 자동본은 기계실/라이저 관로를 'M1/0', 'R4/0' 같은 문자 라벨로 낸다.
TOKEN = r"[A-Za-z0-9_/.-]+"

_CONFIG_RE = re.compile(
    rf"^({TOKEN})\s+({TOKEN})\s+({TOKEN})\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})$"
)
_DESIGNED_RE = re.compile(
    rf"^({TOKEN})\s+({TOKEN})\s+({TOKEN})\s+({NUM})\s+(\d+)\s+({NUM})\s+({NUM})(?:\s+\S+)?$"
)
_FITTING_LINE_RE = re.compile(rf"^({TOKEN})\s+\d+\s*x\s")
_FITTING_ITEM_RE = re.compile(rf"(\d+)\s*x\s*(\d+)\s+({NUM})")

FITTING_TYPE_NAMES = {
    1: "45° 엘보",
    2: "90° 표준 엘보",
    3: "90° 롱래디우스 엘보",
    4: "티/크로스(90° 전환)",
    5: "게이트 밸브",
    6: "스윙 체크 밸브",
    7: "논리턴 밸브",
    8: "볼 밸브",
    9: "버터플라이 밸브",
}

# 이 값 미만의 표고차는 헤드 니플·FX 드롭 같은 국소 하강으로 본다.
MAJOR_DROP_MIN_M = 0.5


@dataclass(slots=True)
class ReportProfile:
    """PIPENET 결과 리포트 1개에서 뽑은 대조용 지표."""

    source: str
    pipe_count: int | None = None
    materials: list[str] = field(default_factory=list)
    pipes_by_material: dict[str, int] = field(default_factory=dict)
    elevation_nonzero_pipes: int | None = None
    elevation_total_drop_m: float | None = None
    elevation_major_drop_m: float | None = None
    fitting_count: int | None = None
    fittings_by_type: dict[str, int] = field(default_factory=dict)
    fitting_eq_length_m: float | None = None
    special_eq_count: int | None = None
    special_eq_length_m: float | None = None
    special_eq_by_desc: dict[str, int] = field(default_factory=dict)
    bore_histogram: dict[str, int] = field(default_factory=dict)
    nozzle_count: int | None = None
    total_flow_lpm: float | None = None
    deviation_min_pct: float | None = None
    deviation_max_pct: float | None = None
    nozzles_under_min_pressure: int | None = None
    unresolved: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ComparisonReport:
    manual: ReportProfile
    auto: ReportProfile
    flow_match_pct: float | None

    def to_dict(self) -> dict:
        return {
            "flow_match_pct": self.flow_match_pct,
            "manual": asdict(self.manual),
            "auto": asdict(self.auto),
        }

    def to_markdown(self) -> str:
        m, a = self.manual, self.auto
        head = (
            f"# 총유량 일치율 {_pct(self.flow_match_pct)}"
            f"  ({_num(a.total_flow_lpm)} / {_num(m.total_flow_lpm)} L/min)"
        )
        rows = [
            ("관종", "라이브러리 종수", len(m.materials) or None, len(a.materials) or None),
            ("관종", "관종별 관로 수", _kv(m.pipes_by_material), _kv(a.pipes_by_material)),
            ("표고", "관로 수", m.pipe_count, a.pipe_count),
            ("표고", "표고 비영 관로", m.elevation_nonzero_pipes, a.elevation_nonzero_pipes),
            ("표고", f"주요 하강 합 (≥{MAJOR_DROP_MIN_M}m)", m.elevation_major_drop_m, a.elevation_major_drop_m),
            ("표고", "전체 하강 합", m.elevation_total_drop_m, a.elevation_total_drop_m),
            ("부속", "개수", m.fitting_count, a.fitting_count),
            ("부속", "종류별", _kv(m.fittings_by_type), _kv(a.fittings_by_type)),
            ("부속", "등가길이 합 (m)", m.fitting_eq_length_m, a.fitting_eq_length_m),
            ("특수설비", "개수", m.special_eq_count, a.special_eq_count),
            ("특수설비", "종류별", _kv(m.special_eq_by_desc), _kv(a.special_eq_by_desc)),
            ("특수설비", "등가길이 합 (m)", m.special_eq_length_m, a.special_eq_length_m),
            ("관경", "호칭경 분포", _kv(m.bore_histogram), _kv(a.bore_histogram)),
            ("결과", "노즐 수", m.nozzle_count, a.nozzle_count),
            ("결과", "총유량 (L/min)", m.total_flow_lpm, a.total_flow_lpm),
            ("결과", "% Deviation 범위", _dev_range(m), _dev_range(a)),
            ("결과", "최소압 미달 노즐", m.nozzles_under_min_pressure, a.nozzles_under_min_pressure),
        ]
        lines = [head, "", "| 축 | 항목 | 수작업 | 자동화 |", "|---|---|---|---|"]
        lines += [f"| {axis} | {name} | {_cell(mv)} | {_cell(av)} |" for axis, name, mv, av in rows]

        for prof in (m, a):
            if prof.unresolved:
                lines += ["", f"## 미확정 — {prof.source}"]
                lines += [f"- {item}" for item in prof.unresolved]
        return "\n".join(lines)


def _num(v) -> str:
    return "—" if v is None else (f"{v:,.1f}" if isinstance(v, float) else str(v))


def _pct(v) -> str:
    return "—" if v is None else f"{v:.1f}%"


def _cell(v) -> str:
    return v if isinstance(v, str) else _num(v)


def _dev_range(prof: "ReportProfile") -> str | None:
    lo, hi = prof.deviation_min_pct, prof.deviation_max_pct
    return None if lo is None or hi is None else f"{lo:+.2f} ~ {hi:+.2f}"


def _kv(d: dict) -> str | None:
    return ", ".join(f"{k} {v}" for k, v in d.items()) if d else None


def _section(validator, text: str, title: str) -> list[str]:
    return validator._section_lines(text, title)


def _profile(path: Path, source: str) -> ReportProfile:
    v = PipenetGuideValidator.__new__(PipenetGuideValidator)
    return profile_from_text(v._read_report_text(path), source)


def profile_from_text(text: str, source: str) -> ReportProfile:
    v = PipenetGuideValidator.__new__(PipenetGuideValidator)
    prof = ReportProfile(source=source)

    materials = {row.pipe_type_id: row.material_name for row in v._parse_design_materials(text)}
    prof.materials = sorted(set(materials.values()))
    if not materials:
        prof.unresolved.append("DESIGN MATERIALS 섹션 없음 — 관종 판별 불가")

    config_lines = _section(v, text, "PIPE CONFIGURATION")
    if not config_lines:
        prof.unresolved.append("PIPE CONFIGURATION 섹션 없음 — 관로·표고·관경 지표 미산출")
    else:
        rows = [m.groups() for m in map(_CONFIG_RE.match, config_lines) if m]
        elevations = [float(r[5]) for r in rows]
        bores = [float(r[3]) for r in rows]
        prof.pipe_count = len(rows)
        prof.elevation_nonzero_pipes = sum(1 for e in elevations if e != 0.0)
        prof.elevation_total_drop_m = round(sum(e for e in elevations if e < 0), 2)
        prof.elevation_major_drop_m = round(
            sum(e for e in elevations if e <= -MAJOR_DROP_MIN_M), 2
        )
        prof.bore_histogram = {
            f"{b:g}": bores.count(b) for b in sorted(set(bores), reverse=True)
        }

    designed_lines = _section(v, text, "DESIGNED DIAMETERS & FLOWRATES")
    if not designed_lines:
        prof.unresolved.append("DESIGNED DIAMETERS & FLOWRATES 섹션 없음 — 관종별 관로 수 미산출")
    else:
        counts: dict[str, int] = {}
        for line in designed_lines:
            m = _DESIGNED_RE.match(line)
            if not m:
                continue
            name = materials.get(int(m.group(5)), f"미상(type {m.group(5)})")
            counts[name] = counts.get(name, 0) + 1
        prof.pipes_by_material = dict(sorted(counts.items(), key=lambda kv: -kv[1]))

    fitting_lines = _section(v, text, "PIPE FITTINGS")
    if not fitting_lines:
        prof.unresolved.append("PIPE FITTINGS 섹션 없음 — 부속 지표 미산출")
    else:
        by_type: dict[int, int] = {}
        total_eq = 0.0
        total_n = 0
        for line in fitting_lines:
            if not _FITTING_LINE_RE.match(line):
                continue
            # 'Equivalent Length' 는 개당 값이므로 개수를 곱해야 관로 합계와 맞는다.
            for count, type_id, eq_length in _FITTING_ITEM_RE.findall(line):
                n = int(count)
                by_type[int(type_id)] = by_type.get(int(type_id), 0) + n
                total_eq += n * float(eq_length)
                total_n += n
        prof.fitting_count = total_n
        prof.fitting_eq_length_m = round(total_eq, 2)
        prof.fittings_by_type = {
            FITTING_TYPE_NAMES.get(t, f"type {t}"): c for t, c in sorted(by_type.items())
        }

    if not _section(v, text, "SPECIAL EQUIPMENT"):
        prof.unresolved.append("SPECIAL EQUIPMENT 섹션 없음 — 특수설비 지표 미산출")
    else:
        equipment = v._parse_equipment(text)
        by_desc: dict[str, int] = {}
        for row in equipment:
            by_desc[row.description] = by_desc.get(row.description, 0) + 1
        prof.special_eq_count = len(equipment)
        prof.special_eq_length_m = round(sum(r.equivalent_length_m for r in equipment), 2)
        prof.special_eq_by_desc = dict(sorted(by_desc.items()))

    nozzles = v._parse_nozzle_flows(text)
    if not nozzles:
        prof.unresolved.append("FLOW THROUGH NOZZLES 섹션 없음 — 총유량·편차 미산출")
    else:
        prof.nozzle_count = len(nozzles)
        prof.total_flow_lpm = round(sum(r.actual_flow_lpm for r in nozzles), 1)
        prof.deviation_min_pct = round(min(r.deviation_percent for r in nozzles), 2)
        prof.deviation_max_pct = round(max(r.deviation_percent for r in nozzles), 2)
        min_press = {r.label: r.min_press_kgcm2 for r in v._parse_nozzle_config_rows(text)}
        if not min_press:
            prof.unresolved.append("NOZZLE CONFIGURATION 섹션 없음 — 최소압 미달 판정 불가")
        else:
            prof.nozzles_under_min_pressure = sum(
                1
                for r in nozzles
                if r.label in min_press and r.inlet_pressure_kgf_cm2 < min_press[r.label]
            )

    return prof


def compare(manual_docx: str | Path, auto_docx: str | Path) -> ComparisonReport:
    manual = _profile(Path(manual_docx), "수작업")
    auto = _profile(Path(auto_docx), "자동화")
    match = None
    if manual.total_flow_lpm and auto.total_flow_lpm is not None:
        # 이중 반올림(78.747 → 78.75 → 78.8)을 피하려고 표시 자릿수보다 깊게 남긴다.
        match = round(auto.total_flow_lpm / manual.total_flow_lpm * 100.0, 4)
    return ComparisonReport(manual=manual, auto=auto, flow_match_pct=match)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="PIPENET 수작업본 ↔ 자동본 대조")
    ap.add_argument("manual", help="수작업 결과 리포트 (.docx/.pdf)")
    ap.add_argument("auto", help="자동 결과 리포트 (.docx/.pdf)")
    ap.add_argument("--json", dest="json_out", help="JSON 저장 경로")
    ap.add_argument("--md", dest="md_out", help="마크다운 저장 경로")
    args = ap.parse_args(argv)

    report = compare(args.manual, args.auto)
    markdown = report.to_markdown()
    print(markdown)
    if args.md_out:
        Path(args.md_out).write_text(markdown, encoding="utf-8")
    if args.json_out:
        Path(args.json_out).write_text(
            json.dumps(report.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
