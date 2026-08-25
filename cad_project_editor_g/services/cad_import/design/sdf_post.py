# -*- coding: utf-8 -*-
"""[G9~G12] SDF 방출 후처리 — 배관 규격 바인딩과 표시 좌표.

`pipenet_converter` 의 `write_sdf` 는 `<Pipe-type>` 을 전혀 쓰지 않는다.
`Pipe.material` 을 받아만 두고 직렬화에서 버리므로 PIPENET 의 Pipe 표에서
`Type` 이 "None defined" 로 뜬다. **표에서 죽는 게 아니라 XML 에서 죽는다.**

writer 를 고치면 모듈 A·F 의 산출물까지 흔들리므로(§3), 여기서 **방출 후처리**로
바로잡는다 — 모듈 A 가 쓰는 것과 같은 방법이다.

★이 모듈이 건드리는 것은 **표시 전용** 값뿐이다. 수리계산은 `length`·`elevation`·
`rise` 로 하므로, 좌표를 아무리 옮겨도 계산 결과는 바뀌지 않아야 한다(§4 회귀).
"""
from __future__ import annotations

import math
import xml.etree.ElementTree as ET
from pathlib import Path

# ── 6종 스케줄 — 권위 SLF(`2. Pipenet_hand_FX28.slf`)의 Item-name 과 정합.
#
# ★이름은 **철자·공백까지 동일**해야 PIPENET 이 Pipe-type ↔ Schedule(내경)을
#   바인딩한다. "KSD 3507"(공백 있음) · "CPVC2"(숫자 2) 를 흘리면 조용히 안 붙는다.
# 호칭경 집합은 SLF 의 Size-definition.nominal 과 같고, velocity 컨벤션
# (≤50 mm = 6, ≥65 mm = 10)은 레퍼런스 알람밸브 SDF 의 Pipe-type 정의에서 왔다.
SCHEDULE_DEFS: list[tuple[str, str, list[tuple[float, int]]]] = [
    ("KSD 3507", "120", [
        (0.015, 6), (0.02, 6), (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6),
        (0.065, 10), (0.08, 10), (0.09, 10), (0.1, 10), (0.125, 10),
        (0.15, 10), (0.2, 10), (0.25, 10), (0.3, 10),
    ]),
    ("KSD 3562", "120", [
        (0.02, 6), (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6),
        (0.065, 10), (0.08, 10), (0.09, 10), (0.1, 10), (0.125, 10),
        (0.15, 10), (0.2, 10), (0.25, 10), (0.3, 10),
    ]),
    ("KSD 3576", "120", [
        (0.015, 6), (0.02, 6), (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6),
        (0.065, 10), (0.08, 10), (0.09, 10), (0.1, 10), (0.125, 10),
        (0.15, 10), (0.2, 10), (0.25, 10), (0.3, 10),
    ]),
    ("DP", "120", [(0.025, 10)]),
    ("CPVC2", "150", [
        (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6), (0.065, 10), (0.08, 10),
    ]),
    ("FX", "120", [(0.025, 10)]),
]

SCHEDULE_NAMES = [name for name, _c, _s in SCHEDULE_DEFS]
DEFAULT_SCHEDULE = SCHEDULE_NAMES[0]

COS30 = 0.8660254037844387
SIN30 = 0.5


class UnknownSchedule(ValueError):
    """`SCHEDULE_DEFS` 에 없는 관종 이름. 조용히 기본값으로 떨어지지 않는다."""


def check_schedule(name: str) -> str:
    """관종 이름 검사. 오타 하나로 다시 "None defined" 가 되는 것을 막는다."""
    n = str(name or "").strip()
    if n not in SCHEDULE_NAMES:
        raise UnknownSchedule(
            f"모르는 배관 규격입니다: {name!r}\n"
            f"  쓸 수 있는 것: {' · '.join(SCHEDULE_NAMES)}\n"
            f"  (이름은 SLF 의 Item-name 과 철자·공백까지 같아야 합니다)")
    return n


# ────────────────────────────────────────────────── [G11] 좌표 정규화
def normalize_node_coords(tables, *, canvas_units: float = 3000.0) -> float:
    """bbox 중심 → (0,0), 가장 긴 축 → `canvas_units`. 적용한 배율을 돌려준다.

    B1F 실측 도면은 한 변이 수백 m 라 mm 좌표를 그대로 두면 PIPENET 스키매틱
    캔버스에서 망이 한 점에 뭉친다. 모듈 A 가 쓰는 것과 같은 규칙이다.

    `display_z` 가 있으면 **x,y 와 같은 배율**을 곱한다 — 다른 배율을 쓰면 라이저
    입상관이 평면과 비례하지 않는다.

    ★`elevation` 은 건드리지 않는다. 그것은 표시가 아니라 수리계산 입력이다.
    """
    nodes = getattr(tables, "nodes", None) or []
    if not nodes:
        return 1.0
    xs = [float(n.get("x", 0) or 0) for n in nodes]
    ys = [float(n.get("y", 0) or 0) for n in nodes]
    cx = (min(xs) + max(xs)) / 2.0
    cy = (min(ys) + max(ys)) / 2.0
    longest = max(max(xs) - min(xs), max(ys) - min(ys))
    scale = (canvas_units / longest) if longest > 1e-9 else 1.0
    for n in nodes:
        n["x"] = (float(n.get("x", 0) or 0) - cx) * scale
        n["y"] = (float(n.get("y", 0) or 0) - cy) * scale
        dz = n.get("display_z")
        if dz is not None:
            n["display_z"] = float(dz) * scale
    return scale


# ────────────────────────────────────────────── [G12] 아이소매트릭 베이크
def bake_isometric(tables, *, iso_z_scale: float = 1.0,
                   ref_label=None, no_lift_labels=None) -> None:
    """30° 등각투영을 x,y 에 in-place 로 굽는다.

    공식은 `routes/r30_combined.py:_bake_isometric_node_coords` 와 **같아야 한다** —
    다른 공식을 쓰면 같은 망이 SDF·KFP·HAS 에서 다르게 보인다::

        x' = (x - y)·cos30
        y' = (x + y)·sin30 + (elev - e_ref)·lift
        lift = 평면대각선 · 0.5 · iso_z_scale / 표고범위   (표고범위 0 이면 0)

    `ref_label` 은 lift 의 영점 — 보통 알람밸브다. 영점을 bbox 중앙으로 두면
    이음매에서 두 망이 찢어진다(모듈 A 실측: 대명동 약 11.6 m).
    `no_lift_labels` 는 라이저·기계실 계통도처럼 schematic y 가 이미 수직을
    인코딩한 노드 — lift 를 또 더하면 이중부호로 구부러진다.

    ★`elevation` 은 건드리지 않는다(표시 전용 변환이다).
    """
    nodes = getattr(tables, "nodes", None) or []
    if not nodes:
        return
    no_lift = {str(x) for x in (no_lift_labels or ())}
    xs = [float(n.get("x", 0) or 0) for n in nodes]
    ys = [float(n.get("y", 0) or 0) for n in nodes]
    elevs = [float(n.get("elevation", 0) or 0) for n in nodes]
    e_min, e_max = min(elevs), max(elevs)
    e_ref = (e_min + e_max) / 2.0
    if ref_label is not None:
        ref = next((n for n in nodes
                    if str(n.get("label")) == str(ref_label)), None)
        if ref is not None:
            e_ref = float(ref.get("elevation", 0) or 0)
    e_range = e_max - e_min
    diag = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) if xs else 0.0
    lift = (diag * 0.5 * iso_z_scale / e_range) if e_range > 0 else 0.0
    for n in nodes:
        x = float(n.get("x", 0) or 0)
        y = float(n.get("y", 0) or 0)
        e = float(n.get("elevation", 0) or 0)
        dy = 0.0 if str(n.get("label")) in no_lift else (e - e_ref) * lift
        n["x"] = (x - y) * COS30
        n["y"] = (x + y) * SIN30 + dy


# ─────────────────────────────────────────── [G9] Pipe-type 주입
def _make_pipe_type(name: str, c_factor: str, sizes) -> ET.Element:
    pt = ET.Element("Pipe-type", {
        "c-factor": c_factor, "criteria": "velocity", "max-velocity": "10",
    })
    ET.SubElement(pt, "Name").text = name
    ET.SubElement(pt, "Schedule").text = name
    for sz, vel in sizes:
        ET.SubElement(pt, "Pipe-size", {
            "Lagging-thickness": "0", "size": str(sz),
            "use": "1", "velocity": str(vel),
        })
    return pt


def inject_pipe_types(sdf_path, sched_by_pipe: dict) -> None:
    """방출된 SDF 를 다시 읽어 Pipe-set 을 관종별로 재구성하고 Pipe-type 을 넣는다.

    `sched_by_pipe` : {pipe label: schedule name}. 없는 배관은 기본 규격으로 본다.

    최종 `<Links>` 순서::

        [빈 placeholder] + [쓰인 schedule Pipe-set] + [빈 정의 Pipe-set] + Nozzle/Valve

    ★빈 placeholder 를 맨 앞에 유지해야 한다. PIPENET 은 첫 Pipe-set 을
    blank/default 슬롯으로 예약하므로, 없으면 우리 Pipe-type 이 그 슬롯에 흡수돼
    관경이 "Unset" 이 된다(writer 주석 · 레퍼런스 SDF 3종에서 확인).

    쓰이지 않는 나머지 규격도 «Pipe-type 만 있고 Pipe 는 없는» Pipe-set 으로 정의해
    둔다 → PIPENET UI 의 드롭다운에 노출되어 사용자가 관종을 바꿀 수 있다.
    이것이 「라이브러리를 가져오게 한다」의 실제 내용이다.
    """
    path = Path(sdf_path)
    tree = ET.parse(path)
    root = tree.getroot()
    by_name = {n: (n, c, s) for n, c, s in SCHEDULE_DEFS}

    for links in root.iter("Links"):
        # writer 가 만든 것 — [빈 placeholder][파이프 담긴 것]
        populated = None
        for child in list(links):
            if child.tag == "Pipe-set" and child.find("Pipe") is not None:
                populated = child
                break
        if populated is None:
            continue

        pipes = list(populated.findall("Pipe"))
        insert_at = list(links).index(populated)
        links.remove(populated)

        # 규격별로 가른다. 순서는 SCHEDULE_DEFS 를 따라 결정적으로.
        buckets: dict[str, list] = {}
        for p in pipes:
            name = sched_by_pipe.get(p.get("label")) or DEFAULT_SCHEDULE
            buckets.setdefault(check_schedule(name), []).append(p)

        used = [n for n in SCHEDULE_NAMES if n in buckets]
        for name in used:
            ps = ET.Element("Pipe-set")
            ps.append(_make_pipe_type(*by_name[name]))
            for p in buckets[name]:
                ps.append(p)
            links.insert(insert_at, ps)
            insert_at += 1

        # 안 쓰인 규격도 정의만 넣어 드롭다운에 띄운다.
        for name in SCHEDULE_NAMES:
            if name in buckets:
                continue
            ps = ET.Element("Pipe-set")
            ps.append(_make_pipe_type(*by_name[name]))
            links.insert(insert_at, ps)
            insert_at += 1

        # 맨 앞 빈 placeholder 확보 — 없으면 관경이 "Unset" 이 된다.
        first = list(links)[0] if len(links) else None
        if first is None or first.tag != "Pipe-set" or len(first):
            links.insert(0, ET.Element("Pipe-set"))
        break

    tree.write(path, encoding="utf-8", xml_declaration=True)
