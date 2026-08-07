# -*- coding: utf-8 -*-
"""D5 검증 — 만든 PDF 에서 글자를 도로 뽑아 PIPENET 워드 계산서와 표별로 대조한다.

렌더러의 상수를 가져오지 않는다. 계수도 열 위치도 여기서 다시 적는다 — 같은 상수를
돌려 쓰면 계수가 틀려도 양쪽이 똑같이 틀려서 통과한다.
"""
from __future__ import annotations

import re
import sys
import zipfile
from decimal import Decimal
from functools import reduce
from math import gcd, ulp
from pathlib import Path
from typing import Iterable, NamedTuple, Sequence

from pypdf import PdfReader

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

SUB = _ROOT / "routes" / "제출용[최종]"


def word_lines(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    out: list[str] = []
    for para in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
        # <w:t> 만 잡는다. [^>]* 로 열면 <w:tab .../> 까지 걸려 탭 정의가 본문으로 섞인다.
        runs = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", para, re.S)
        joined = "".join(runs)
        joined = joined.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
        out.append(joined)
    return out


def pdf_lines(pdf: Path) -> list[str]:
    reader = PdfReader(str(pdf))
    out: list[str] = []
    for page in reader.pages:
        out.extend(page.extract_text().splitlines())
    return out


def _numbers(cells: list[str]) -> list[float]:
    values = []
    for cell in cells:
        try:
            values.append(float(cell))
        except ValueError:
            pass
    return values


# 양쪽 문서 모두 한 표가 여러 쪽에 걸치면 쪽마다 제목을 다시 찍는다. 첫 쪽에서 멈추면
# 관로 136 개 중 35 개만 대조하고도 통과한다 — 그래서 제목을 만날 때마다 다시 들어가고,
# 다른 절의 제목을 만나야 빠져나온다.
SECTIONS = (
    "CONTROL INFORMATION", "FLUID SYSTEM", "DESIGN INFORMATION",
    "AVAILABLE PIPE SIZES", "PIPE CONFIGURATION", "PIPE FITTINGS", "FITTING TYPES",
    "NOZZLE CONFIGURATION", "SPECIAL EQUIPMENT", "DESIGNED DIAMETERS",
    "FLOW IN PIPES", "FLOW THROUGH NOZZLES", "FLOW AT INLETS",
    "Materials Take", "IMPORTANT NOTICE", "COMMENTS", "WARNINGS", "생성 출처",
)


def read_cells(lines: list[str], heading: str, columns: int) -> dict[str, list[str]]:
    """한 절에서 (라벨 → 그 행의 칸들). 라벨은 첫 칸이다."""
    others = tuple(s for s in SECTIONS if not heading.startswith(s))
    rows: dict[str, list[str]] = {}
    inside = False
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith(heading):
            inside = True
            continue
        if stripped.startswith(others):
            inside = False
            continue
        if not inside or stripped.startswith(("TITLE :", "---", "===")):
            continue
        # 쪽 바닥글은 제목으로 시작해 라벨 행처럼 보인다 — 표 행으로 세면 없는 관로가 하나 는다.
        if re.search(r"Page \d+ of \d+", stripped):
            continue
        cells = stripped.split()
        if len(cells) < columns or not re.match(r"^[-\w/.]+$", cells[0]):
            continue
        if _numbers(cells[1:]):
            rows.setdefault(cells[0], cells)
    return rows


def read_table(lines: list[str], heading: str, columns: int) -> dict[str, list[float]]:
    """한 절에서 (라벨 → 그 행의 수치들). 나머지 칸 중 수치만 추린다."""
    return {label: _numbers(cells[1:])
            for label, cells in read_cells(lines, heading, columns).items()}


def read_nodes(lines: list[str], heading: str, columns: int) -> dict[str, tuple[str, str]]:
    """(라벨 → 입구·출구 노드). 노드는 문자일 수 있어 수치 대조에서 빠진다."""
    return {label: (cells[1], cells[2])
            for label, cells in read_cells(lines, heading, columns).items()}


def compare_nodes(name: str, word: dict[str, tuple[str, str]],
                  mine: dict[str, tuple[str, str]]) -> int:
    bad = 0
    for label, expected in sorted(word.items()):
        got = mine.get(label)
        if got is not None and got != expected:
            print(f"  [{name} 노드] {label}: 워드 {expected} vs 내 {got}")
            bad += 1
    print(f"{name} 노드 라벨: 워드 {len(word)}행 / 내 {len(mine)}행, 불일치 {bad}")
    return bad


def _at(values: list[float], index: int) -> float | None:
    """음수 첨자는 뒤에서부터. 열 하나가 이름 안의 공백으로 갈라져도 끝에서 세면 흔들리지 않는다."""
    if -len(values) <= index < len(values):
        return values[index]
    return None


class Mismatch(NamedTuple):
    table: str
    label: str
    index: int
    word: float
    mine: float


def compare(name: str, word: dict[str, list[float]], mine: dict[str, list[float]],
            picks: list[tuple[int, int]], tol: float = 0.0) -> list[Mismatch]:
    """picks 는 (워드 열, 내 열) 짝. 열 위치는 양쪽에서 따로 세어 적는다.

    기본 허용오차는 0 이다 — 양쪽 모두 같은 원본 double 을 같은 Fortran 서식으로 찍으므로
    찍힌 수는 정확히 같아야 한다. 오차를 열어 두면 서식이 틀려도 통과한다.
    어긋난 자리는 여기서 판정하지 않고 그대로 넘긴다 — 설명은 XML 정밀도로 따로 잰다.
    """
    missing = sorted(set(word) - set(mine))
    extra = sorted(set(mine) - set(word))
    found: list[Mismatch] = []
    for label, expected in sorted(word.items()):
        got = mine.get(label)
        if got is None:
            found.append(Mismatch(name, label, 0, float("nan"), float("nan")))
            continue
        for w_index, m_index in picks:
            a, b = _at(expected, w_index), _at(got, m_index)
            if a is None or b is None:
                print(f"  [{name}] {label}: 열이 모자란다 (워드 {len(expected)} / 내 {len(got)})")
                found.append(Mismatch(name, label, w_index, float("nan"), float("nan")))
                continue
            if abs(a - b) > tol * max(1.0, abs(a)):
                found.append(Mismatch(name, label, w_index, a, b))
    print(f"{name}: 워드 {len(word)}행 / 내 {len(mine)}행, 대조 불일치 {len(found) - len(missing)}"
          f"{f', 내게 없는 라벨 {len(missing)}' if missing else ''}"
          f"{f', 워드에 없는 라벨 {len(extra)}' if extra else ''}")
    return found


# ── 남은 자릿수 차이를 XML 이 실은 정밀도로 설명한다 ────────────────────────
#
# PIPENET 은 자기 내부의 사용자 단위 값에서 계산서를 찍고, XML 로는 SI 로 환산한 값을
# 내보낸다. 그 환산본은 원본이 아니다 — 압력은 성긴 격자에 올라앉고, 나머지도 적힌
# 자릿수까지만 남는다. 그래서 워드가 찍은 마지막 자리를 XML 만 보고 되살릴 수 없는
# 자리가 생긴다.
#
# 판정은 이렇게 한다: 두 표시값 사이의 경계값을 SI 로 되돌리고, XML 이 실제로 내보낸
# 값이 그 경계에서 자기 정밀도 한 칸 안에 있는지 본다. 한 칸을 넘으면 파일이 답을
# 갖고 있는데 내가 틀린 것이므로 설명하지 않는다. 아래에 없는 열에서 어긋나도 마찬가지다.
_PA_PER_KGFCM2 = 98066.5      # kg/cm2 → Pa (정의값)
_M3S_PER_LPM = 1.0 / 60000.0  # lit/min → m3/s (정의값)

SOURCES = {
    # (표, 열): (XML 표, 그 값을 이루는 열들(둘이면 차), 표시값→SI 계수)
    ("FLOW IN PIPES", -6): ("Pipes-results", ("Inlet pressure",), _PA_PER_KGFCM2),
    ("FLOW IN PIPES", -5): ("Pipes-results", ("Outlet pressure",), _PA_PER_KGFCM2),
    ("FLOW IN PIPES", -4): ("Pipes-results", ("Inlet pressure", "Outlet pressure"),
                            _PA_PER_KGFCM2),
    ("FLOW IN PIPES", -3): ("Pipes-results", ("Friction loss",), _PA_PER_KGFCM2),
    ("FLOW IN PIPES", -2): ("Pipes-results", ("Flow",), _M3S_PER_LPM),
    ("FLOW THROUGH NOZZLES", 1): ("Spray-nozzles-results", ("Inlet pressure",),
                                  _PA_PER_KGFCM2),
    ("FLOW THROUGH NOZZLES", 3): ("Spray-nozzles-results", ("Calculated flow",),
                                  _M3S_PER_LPM),
}


def xml_columns(xml: Path) -> dict[tuple[str, str], tuple[str, dict[str, str]]]:
    """XML 원문 → {(표, 열): (단위 이름, {라벨: 적힌 글자})}.

    파싱기를 거치지 않는다 — float 으로 바꾸는 순간 파일이 몇 자리를 실어 왔는지가
    지워지고, 그 자릿수가 여기서 재려는 바로 그것이다.
    """
    raw = xml.read_text(encoding="utf-8", errors="replace")
    out: dict[tuple[str, str], tuple[str, dict[str, str]]] = {}
    for body in re.findall(r"<TABLE[^>]*>.*?</TABLE>", raw, re.S):
        table = re.search(r'<TABLE[^>]*name="([^"]+)"', body).group(1)
        fields = [(re.search(r'name="([^"]+)"', f).group(1),
                   (re.search(r'unit="([^"]*)"', f) or [None, ""])[1])
                  for f in re.findall(r"<FIELD[^>]*>", body)]
        rows = [[c.strip() for c in re.findall(r"<TD>(.*?)</TD>", tr, re.S)]
                for tr in re.findall(r"<TR>(.*?)</TR>", body, re.S)]
        for index, (name, unit) in enumerate(fields):
            out[(table, name)] = (unit, {r[0]: r[index] for r in rows if len(r) > index})
    return out


def measured_grid(cells: Iterable[str]) -> Decimal:
    """값 전부가 올라앉은 가장 성긴 십진 격자. 적힌 자릿수의 최대공약수로 잰다."""
    values = [Decimal(c) for c in cells if _is_number(c)]
    finest = min(Decimal(1).scaleb(v.as_tuple().exponent) for v in values)
    step = reduce(gcd, (abs(int((v / finest).to_integral_value())) for v in values))
    return finest * step


def _is_number(cell: str) -> bool:
    try:
        float(cell)
    except ValueError:
        return False
    return cell.strip().lower() not in ("nan", "inf", "-inf")


def precisions(columns: dict[tuple[str, str], tuple[str, dict[str, str]]]) -> dict[tuple[str, str], float]:
    """열마다 '이 파일이 실제로 실어 나른 최소 단위'.

    같은 단위를 단 열들은 한 덩어리로 본다. 압력은 파일 전체가 0.5 Pa 격자에 올라앉는데,
    한 열만 떼어 재면 그 열에 성긴 값이 없다는 이유로 격자가 더 잘게 보인다.
    """
    buckets: dict[str, list[str]] = {}
    for (_table, _field), (unit, cells) in columns.items():
        if unit:
            buckets.setdefault(unit, []).extend(v for v in cells.values() if _is_number(v))
    grids = {unit: measured_grid(cells) for unit, cells in buckets.items() if cells}
    out: dict[tuple[str, str], float] = {}
    for key, (unit, cells) in columns.items():
        values = [Decimal(v) for v in cells.values() if _is_number(v)]
        if not values:
            continue
        coarsest = max(Decimal(1).scaleb(v.as_tuple().exponent) for v in values)
        out[key] = float(max(coarsest, grids.get(unit, Decimal(0))))
    return out


def explain(found: list[Mismatch],
            columns: dict[tuple[str, str], tuple[str, dict[str, str]]]) -> int:
    """설명되지 않은 불일치 수. 설명된 것도 숨기지 않고 경계까지의 거리를 적는다."""
    quanta = precisions(columns)
    unexplained = 0
    worst: dict[str, tuple[float, float]] = {}
    for m in found:
        source = SOURCES.get((m.table, m.index))
        if source is None or m.word != m.word:  # NaN — 열이 모자라거나 라벨이 없다
            print(f"  설명 못 함: [{m.table}] {m.label} 열{m.index} "
                  f"워드 {m.word!r} vs 내 {m.mine!r}")
            unexplained += 1
            continue
        table, fields, to_si = source
        exported, quantum = 0.0, 0.0
        for sign, field in zip((1.0, -1.0), fields):
            exported += sign * float(columns[(table, field)][1][m.label])
            quantum += quanta[(table, field)]
        # 워드와 내 표시값 사이의 경계 — 여기를 넘느냐에 마지막 자리가 걸려 있다
        boundary = (m.word + m.mine) / 2.0 * to_si
        distance = abs(exported - boundary)
        name = f"{table}/{'-'.join(fields)}"
        seen = worst.get(name, (0.0, quantum))
        worst[name] = (max(seen[0], distance), quantum)
        # 경계에서 정확히 한 칸 떨어진 자리가 있다 (노즐 2). 거리는 크기가 비슷한 두
        # 실수의 차라 마지막 비트가 상쇄로 드러난다 — 그 한두 자리 잡음만큼 열어 둔다.
        if distance > quantum + 4 * ulp(max(abs(exported), abs(boundary))):
            print(f"  설명 못 함: [{m.table}] {m.label} {name} — XML {exported!r} 는 "
                  f"표시 경계 {boundary!r} 에서 {distance:.4g} 떨어져 있다 "
                  f"(정밀도 {quantum:.4g})")
            unexplained += 1
    for name, (distance, quantum) in sorted(worst.items()):
        print(f"  {name}: 표시 경계까지 최대 {distance:.4g} (XML 정밀도 {quantum:.4g})")
    return unexplained


def warned(lines: list[str], patterns: Sequence[str]) -> tuple[set[str], set[str]]:
    """경고문에서 (유속 초과 관로, 최소압 미달 노즐) 라벨. 문구가 아니라 라벨만 본다."""
    pipes, nozzles = set(), set()
    for line in lines:
        for target, pattern in zip((pipes, nozzles), patterns):
            hit = re.search(pattern, line)
            if hit:
                target.add(hit.group(1))
    return pipes, nozzles


def compare_warnings(word: list[str], mine: list[str]) -> int:
    """경고는 옮겨 적는 값이 아니라 우리가 내리는 판정이다 — 원본과 같아야 한다.

    유속 초과는 재질별 규격표의 최대유속을, 최소압 미달은 절대압/게이지압 기준면 차를
    제대로 잡아야 같은 집합이 나온다. 둘 중 하나만 틀려도 여기서 갈린다.
    """
    expected = warned(word, (r"Velocity in pipe\s+(\S+)\s+exceeds",
                             r"Nozzle\s+(\S+)\s+below minimum"))
    got = warned(mine, (r"관로 (\S+) 의 유속이", r"노즐 (\S+) 의 입구압이"))
    bad = 0
    for name, a, b in zip(("유속 초과 관로", "최소압 미달 노즐"), expected, got):
        if a != b:
            print(f"  [{name}] 워드에만 {sorted(a - b)} / 내게만 {sorted(b - a)}")
            bad += 1
    print(f"WARNINGS: 유속 초과 {len(expected[0])}개 / 최소압 미달 {len(expected[1])}개, "
          f"판정 불일치 {bad}")
    return bad


def main() -> int:
    from core.d_report_typeset import render_report

    xml = SUB / "2. Pipenet_auto.xml"
    out = _ROOT / "data" / "_d5_verify_auto.pdf"
    result = render_report(SUB / "2. Pipenet_auto.sdf", xml, out)
    print(f"PDF {result.pages}쪽 / {result.rows}행 / {result.font_pt:.2f}pt")

    word = word_lines(SUB / "3. 수리계산서_Pipenet_자동화.docx")
    mine = pdf_lines(out)

    found: list[Mismatch] = []
    bad = 0
    # 열은 끝에서부터 센다. 노드 라벨이 M1/0 처럼 문자면 앞에서 센 첨자가 통째로 밀린다
    # (auto 136 관로 중 39 개). 노드는 수치가 아니므로 아래 노드 대조에서 따로 본다.
    found += compare(
        "PIPE CONFIGURATION",
        read_table(word, "PIPE CONFIGURATION", 8),
        read_table(mine, "PIPE CONFIGURATION", 8),
        [(-5, -5), (-4, -4), (-3, -3), (-2, -2), (-1, -1)])

    found += compare(
        "FLOW IN PIPES",
        read_table(word, "FLOW IN PIPES", 10),
        read_table(mine, "FLOW IN PIPES", 10),
        [(-7, -7), (-6, -6), (-5, -5), (-4, -4), (-3, -3), (-2, -2), (-1, -1)])

    bad += compare_nodes(
        "PIPE CONFIGURATION",
        read_nodes(word, "PIPE CONFIGURATION", 8),
        read_nodes(mine, "PIPE CONFIGURATION", 8))

    # 노즐 종류를 워드는 번호(1)로, 나는 XML 이 적어 둔 이름(INDOOR HYDRANT)으로 싣는다.
    # 이름 안의 공백이 칸을 가르므로 K-factor 뒤쪽 네 값은 끝에서부터 센다.
    found += compare(
        "NOZZLE CONFIGURATION",
        read_table(word, "NOZZLE CONFIGURATION", 7),
        read_table(mine, "NOZZLE CONFIGURATION", 7),
        [(0, 0), (-4, -4), (-3, -3), (-2, -2), (-1, -1)])

    found += compare(
        "FLOW THROUGH NOZZLES",
        read_table(word, "FLOW THROUGH NOZZLES", 6),
        read_table(mine, "FLOW THROUGH NOZZLES", 6),
        [(0, 0), (1, 1), (2, 2), (3, 3), (4, 4)])

    # 끝의 네 수치는 유량·재질·실지름·호칭경이다. 재질을 워드는 카탈로그 번호(9)로, 나는
    # 이름(KSD 3507)으로 싣는다 — 그 열(-3)은 건너뛴다.
    # Flowrate(-4) 는 워드가 관경 결정 단계의 값을, 나는 수렴된 해의 값을 싣는다 (XML 에 전자가 없다).
    found += compare(
        "DESIGNED DIAMETERS & FLOWRATES",
        read_table(word, "DESIGNED DIAMETERS & FLOWRATES", 8),
        read_table(mine, "DESIGNED DIAMETERS & FLOWRATES", 7),
        [(-2, -2), (-1, -1)])

    bad += compare_nodes(
        "DESIGNED DIAMETERS",
        read_nodes(word, "DESIGNED DIAMETERS & FLOWRATES", 8),
        read_nodes(mine, "DESIGNED DIAMETERS & FLOWRATES", 7))

    designed_flow = compare(
        "DESIGNED DIAMETERS — Flowrate (참고)",
        read_table(word, "DESIGNED DIAMETERS & FLOWRATES", 8),
        read_table(mine, "DESIGNED DIAMETERS & FLOWRATES", 7),
        [(-4, -4)], 1e-4)

    bad += compare_warnings(word, mine)

    print()
    unexplained = explain(found, xml_columns(xml)) + bad
    print(f"\n자릿수 불일치 {len(found)} 건 / 설명 못 한 불일치 {unexplained} 건 "
          f"(관경 결정 단계 유량 참고 대조 별도: {len(designed_flow)})")
    return 1 if unexplained else 0


if __name__ == "__main__":
    raise SystemExit(main())
