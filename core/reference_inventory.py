"""답안 라이브러리 인벤토리 + 도면↔답안 짝맞춤 헬퍼.

스캔 대상: data/reference_library/
출력:
  inventory.json   — 도면 / 답안 SDF 의 전체 목록 + 메타데이터 (한 줄/객체 형식)
  pairs.json       — 도면 ↔ 답안 짝 (휴리스틱)

답안 SDF 파일명 패턴:
  공동주택/{동}/{차수}-{n}. 양주옥정 {동} {zone} {floor}_REMOTE.sdf
  공동주택/{동}/{차수}-{n}. 양주옥정 {동} {zone} {floor}_USER.sdf

도면 키:
  MF-016~ 계통도          → 전체 단지 라이저 (양주옥정 모든 동 자연낙차/펌프가압)
  MF-201~ 동 소화설비 평면도 → 각 동의 단위세대 헤드망
  MF-101~ 지하 소화설비 평면도 → 지하주차장 (LLSP)

전체 자동 매핑은 1:N 관계 (도면 한 개 ↔ 여러 zone 의 답안) 라
일대일이 아니므로 인벤토리만 만들고 짝은 휴리스틱으로 추정 + 사용자 확인용.
"""

from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from pathlib import Path

from sdf_reader import read_sdf, summarize


REF_ROOT = Path("data/reference_library")


# ── 답안 SDF 파일명 패턴 ──
# 예: "01-1. 양주옥정 101동 펌프가압구간 49F_REMOTE.sdf"
# 예: "1-1. 빌딩랙 상부헤드 펌프가압구간_USER.sdf"
_RE_ANSWER = re.compile(
    r"^"
    r"(?P<series>\d+(?:-\d+)?(?:-\d+)?)"  # 01-1 / 01-1-1 / 1-1
    r"\.\s*"
    r"(?P<body>.*?)"                       # building/zone 설명
    r"_(?P<variant>REMOTE|USER|USER_체절감압)\.sdf$"
)

# 동 번호 (101동 / 102동 …)
_RE_BUILDING = re.compile(r"(\d{2,3})\s*동")
# 층 (49F / 31F / RF / B1F)
_RE_FLOOR = re.compile(r"(?:^|\s)(RF|B?\d{1,2}F|옥상|옥탑)(?:\s|$|_)")
# zone 키워드
_RE_ZONE = re.compile(r"(펌프가압|자연낙차\s*감압|자연낙차|체절감압|랙\s*\(?[A-Za-z]*\)?|상부헤드|인랙헤드|소터)")


@dataclass
class DrawingEntry:
    rel_path: str
    project: str          # "다이소" / "양주옥정"
    folder: str           # "도면" / "CAD" / "도면/XR" / ...
    name: str
    is_calced_diagram: bool   # 계통도?
    is_plan_view: bool        # 평면도?


@dataclass
class AnswerEntry:
    rel_path: str
    project: str
    folder: str           # 물류동 / 지원동 / 공동주택/101동 / 공동주택/102동 등
    name: str
    series: str           # 01-1 / 02-1 / ...
    variant: str          # REMOTE / USER / USER_체절감압
    building: str         # 101동 / (다이소는 없음)
    floor: str            # 49F / RF / 31F / 16F / ...
    zone: str             # 펌프가압 / 자연낙차 / 자연낙차감압 / ...
    # 토폴로지 메타 (열어보고 요약)
    nodes: int
    pipes: int
    nozzles: int
    valves: int
    total_length_m: float
    elev_range_m: tuple[float, float]
    main_bore_mm: int     # 길이로 가중한 최대 직경 (라이저 추정)
    main_bore_length_m: float
    bore_distribution: dict


def _classify_drawing(name: str) -> tuple[bool, bool]:
    n = name.lower()
    is_diagram = ("계통도" in name) or ("흐름도" in name)
    is_plan = ("평면도" in name) or ("plan" in n)
    return is_diagram, is_plan


def _project_from_path(p: Path) -> str:
    parts = p.parts
    for part in parts:
        if "다이소" in part:
            return "다이소"
        if "양주옥정" in part:
            return "양주옥정"
    return "?"


def _parse_answer_name(stem: str, name: str) -> dict:
    """파일명에서 series/variant + building/floor/zone 추출."""
    m = _RE_ANSWER.match(name)
    if not m:
        return {"series": "", "variant": "", "body": name}
    info = m.groupdict()
    body = info["body"]
    b = _RE_BUILDING.search(body)
    f = _RE_FLOOR.search(body)
    z = _RE_ZONE.search(body)
    info["building"] = (b.group(1) + "동") if b else ""
    info["floor"]    = f.group(1) if f else ""
    info["zone"]     = z.group(1) if z else ""
    return info


def scan_drawings() -> list[DrawingEntry]:
    out: list[DrawingEntry] = []
    for p in sorted(REF_ROOT.rglob("*.dxf")):
        is_diag, is_plan = _classify_drawing(p.name)
        out.append(DrawingEntry(
            rel_path=str(p.relative_to(REF_ROOT)),
            project=_project_from_path(p),
            folder=str(p.parent.relative_to(REF_ROOT)),
            name=p.name,
            is_calced_diagram=is_diag,
            is_plan_view=is_plan,
        ))
    return out


def scan_answers(quick: bool = False) -> list[AnswerEntry]:
    """답안 SDF 전체 스캔. quick=True 면 토폴로지 메타 생략 (빠른 인벤토리)."""
    out: list[AnswerEntry] = []
    for p in sorted(REF_ROOT.rglob("*.sdf")):
        info = _parse_answer_name(p.stem, p.name)
        nodes = pipes = nozzles = valves = 0
        total_length = 0.0
        elev_min = elev_max = 0.0
        main_bore_mm = 0
        main_bore_length = 0.0
        bore_dist: dict = {}
        if not quick:
            try:
                net = read_sdf(p)
                nodes = len(net.nodes); pipes = len(net.pipes)
                nozzles = len(net.nozzles); valves = len(net.valves)
                total_length = net.total_length_m
                elev_min, elev_max = net.elev_range
                bld = net.bore_length_distribution_mm
                bore_dist = {k: round(v, 2) for k, v in bld.items()}
                if bld:
                    main_bore_mm, main_bore_length = max(bld.items(), key=lambda kv: kv[1])
                    main_bore_length = round(main_bore_length, 2)
            except Exception as e:
                # 파싱 실패 (손상된 SDF 등) — 빈 메타로 기록
                bore_dist = {"_parse_error": str(e)[:80]}
        out.append(AnswerEntry(
            rel_path=str(p.relative_to(REF_ROOT)),
            project=_project_from_path(p),
            folder=str(p.parent.relative_to(REF_ROOT)),
            name=p.name,
            series=info.get("series", ""),
            variant=info.get("variant", ""),
            building=info.get("building", ""),
            floor=info.get("floor", ""),
            zone=info.get("zone", ""),
            nodes=nodes,
            pipes=pipes,
            nozzles=nozzles,
            valves=valves,
            total_length_m=round(total_length, 2),
            elev_range_m=(round(elev_min, 2), round(elev_max, 2)),
            main_bore_mm=main_bore_mm,
            main_bore_length_m=main_bore_length,
            bore_distribution=bore_dist,
        ))
    return out


def write_inventory(out_path: Path = Path("data/reference_library_inventory.json"),
                    quick: bool = False) -> dict:
    drawings = scan_drawings()
    answers = scan_answers(quick=quick)
    inv = {
        "drawings_count": len(drawings),
        "answers_count": len(answers),
        "projects": dict(Counter(d.project for d in drawings)),
        "drawings": [asdict(d) for d in drawings],
        "answers":  [asdict(a) for a in answers],
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(inv, ensure_ascii=False, indent=2),
                        encoding="utf-8")
    return inv


if __name__ == "__main__":
    import sys
    quick = "--quick" in sys.argv
    inv = write_inventory(quick=quick)
    print(f"drawings: {inv['drawings_count']}, answers: {inv['answers_count']}")
    print(f"projects: {inv['projects']}")
    # 답안 메타 통계
    by_bldg = Counter(a["building"] for a in inv["answers"] if a["building"])
    by_floor = Counter(a["floor"] for a in inv["answers"] if a["floor"])
    by_zone = Counter(a["zone"] for a in inv["answers"] if a["zone"])
    by_var = Counter(a["variant"] for a in inv["answers"] if a["variant"])
    print(f"\nBuilding (top 10): {dict(by_bldg.most_common(10))}")
    print(f"Floor (top 10):    {dict(by_floor.most_common(10))}")
    print(f"Zone:              {dict(by_zone.most_common(10))}")
    print(f"Variant:           {dict(by_var.most_common(10))}")
    # 답안 토폴로지 통계
    if not quick:
        with_pipes = [a for a in inv["answers"] if a["pipes"]]
        if with_pipes:
            print(f"\n토폴로지 평균:")
            print(f"  노드 {sum(a['nodes'] for a in with_pipes) / len(with_pipes):.0f}, "
                  f"파이프 {sum(a['pipes'] for a in with_pipes) / len(with_pipes):.0f}, "
                  f"노즐 {sum(a['nozzles'] for a in with_pipes) / len(with_pipes):.0f}, "
                  f"밸브 {sum(a['valves'] for a in with_pipes) / len(with_pipes):.1f}, "
                  f"길이 {sum(a['total_length_m'] for a in with_pipes) / len(with_pipes):.1f}m")
            print(f"  주 라이저 직경 (가장 흔한 main_bore_mm):")
            mb = Counter(a["main_bore_mm"] for a in with_pipes if a["main_bore_mm"])
            for d, c in mb.most_common(8):
                print(f"    {d}mm: {c} 답안")
