# -*- coding: utf-8 -*-
"""[G4] 부속 · 노즐 · 기기.

부속 **판정**은 `core/fitting_rules.py` 의 규칙을 그대로 쓴다(§5 — 다시 구현
하지 말 것). 그 안에 현업 관행 두 가지가 이미 들어 있다:

  · 티는 흐름이 꺾일 때만 분류티다. **직진해 지나가는 갈래는 직류티라 계상하지
    않는다** — 마찰손실 비가 10:1 이라 직류티까지 세면 손실이 통째로 부풀려진다.
  · 엘보는 45°/90° 뿐이다. 어느 쪽으로도 보낼 수 없으면 **지어내지 말고**
    「판정 불가」로 센다.

부속 **등가길이**는 D2 에 따라 `fittings_library_v3.json` 에서 가져온다
(NFPA13 2025 표 28.2.4.4.1, 단위 m). 라이브러리에 값이 없는 조합은 **0 으로
채우지 않고** 미해결로 세어 로그에 남긴다 — 0 은 「손실 없음」이라는 거짓말이다.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

_G_ROOT = Path(__file__).resolve().parents[3]
_REPO = Path(__file__).resolve().parents[4]

# 편향이 이보다 작으면 «똑바로 지나감» 으로 본다 — 부속 없음이 정답이다.
# fitting_rules 의 ELBOW_STRAIGHT_MAX_DEG(22.5°)와 다른 자다: 그쪽은 「직선이라
# 단정할 수 없어 판정 불가」 구간이고, 이쪽은 「전개가 만든 정확한 직선」이다.
_STRAIGHT_EPS_DEG = 0.5

# 부속 종류 → 라이브러리 항목 id. fitting_rules 의 반환값이 왼쪽이다.
FITTING_LIB_ID = {
    "elbow-45": "ELBOW_45",
    "elbow": "ELBOW_90_STD",
    "tee": "TEE_BRANCH",
    "alarm_valve": "VALVE_ALARM",
}


def _rules():
    """`core/fitting_rules.py` — 저장소 루트 아래라 경로를 붙여 준다."""
    core = _REPO / "core"
    if str(core) not in sys.path:
        sys.path.append(str(core))
    import fitting_rules
    return fitting_rules


def load_equivalent_lengths(path: Path | None = None) -> dict:
    """{라이브러리 id: {호칭경(int): 등가길이 m}}. 값이 없는 칸은 담지 않는다."""
    p = Path(path) if path else (_G_ROOT / "fittings_library_v3.json")
    data = json.loads(p.read_text(encoding="utf-8"))
    out: dict = {}
    for item in data.get("items") or ():
        table = {}
        for dn, rec in (item.get("equivalent_lengths") or {}).items():
            v = (rec or {}).get("value")
            if v is None:
                continue          # 비어 있는 칸 — 0 으로 메우지 않는다
            try:
                table[int(dn)] = float(v)
            except (TypeError, ValueError):
                continue
        out[item.get("id")] = table
    return out


def equivalent_length_m(lib: dict, kind: str, dia_mm: int):
    """부속 1개의 등가길이(m). 라이브러리에 없으면 None — 0 이 아니다."""
    table = lib.get(FITTING_LIB_ID.get(kind, kind)) or {}
    return table.get(int(dia_mm))


def _pipe_ends(net, pid):
    p = ((net or {}).get("pipe_data") or {}).get(pid) or {}
    return p.get("start") or p.get("from"), p.get("end") or p.get("to")


def _dir3(a, b):
    """a→b 의 단위 방향(3차원). 길이가 0 이면 None."""
    d = (b[0] - a[0], b[1] - a[1], b[2] - a[2])
    n = math.sqrt(d[0] * d[0] + d[1] * d[1] + d[2] * d[2])
    if n < 1e-9:
        return None
    return (d[0] / n, d[1] / n, d[2] / n)


def _deflect3_deg(u, v):
    """두 방향 사이 편향(도). 0 이면 그대로 직진, 90 이면 직각으로 꺾임."""
    dot = max(-1.0, min(1.0, u[0] * v[0] + u[1] * v[1] + u[2] * v[2]))
    return math.degrees(math.acos(dot))


def _is_vertical(a, b) -> bool:
    """평면에서 겹치고 높이만 다른 구간 — 헤드 접속관·가지 상승이 그렇다."""
    return (math.hypot(b[0] - a[0], b[1] - a[1]) < 1e-6
            and abs(b[2] - a[2]) > 1e-9)


def build_fittings(net, node_xy, bores, *, parents=None, lib=None,
                   node_z=None) -> dict:
    """배관마다 부속 목록과 등가길이 합.

    `net`     : 제한 전개 kfp dict
    `node_xy` : {node_id: (x, y)} — 평면 좌표(m 든 mm 든 각도만 쓰므로 무관)
    `bores`   : {pipe_id: (dia_mm, source)} — G3 결과
    `parents` : {node_id: 상류 node_id} — 급수원 BFS 의 부모. 없으면 티 판정이
                「상류를 모름」이 되어 전부 티로 두고 판정 불가로 센다.
    `node_z`  : {node_id: 표고} — 없으면 전부 0 으로 본다.

    ★평면 좌표만으로는 **세로 구간을 판정할 수 없다**. `bearing_deg` 는 같은 점에
      대해 `atan2(0,0)=0`(동쪽)을 돌려주므로, 헤드 접속관처럼 위아래로만 가는
      구간은 «가로 배관이 어느 방위를 보고 있었나» 에 따라 엘보가 되기도 하고
      직선이 되기도 했다. 지어낸 값이다. 그래서 각을 **3차원으로 잰다** —
      가로에서 세로로 꺾이면 정확히 90° 가 나오고 엘보가 된다.

    반환::

        {"per_pipe": {pipe_id: {"fittings": [종류…], "equivalent_length": m}},
         "counts": {종류: 개수}, "unresolved_kind": n, "unresolved_length": n}
    """
    fr = _rules()
    lib = lib if lib is not None else load_equivalent_lengths()
    pipes = (net or {}).get("pipe_data") or {}
    parents = parents or {}
    node_z = node_z or {}

    def at(nid):
        xy = node_xy.get(nid)
        if xy is None:
            return None
        return (float(xy[0]), float(xy[1]), float(node_z.get(nid, 0.0) or 0.0))

    # 노드마다 붙은 배관 — 엘보·티 판정의 재료
    incident: dict = {}
    for pid in pipes:
        a, b = _pipe_ends(net, pid)
        if a is None or b is None:
            continue
        incident.setdefault(a, []).append((pid, b))
        incident.setdefault(b, []).append((pid, a))

    per_pipe: dict = {pid: {"fittings": [], "equivalent_length": 0.0}
                      for pid in pipes}
    counts: dict = {}
    unresolved_kind = 0
    n_straight = 0

    for nid, links in incident.items():
        here = node_xy.get(nid)
        if here is None:
            continue
        up = parents.get(nid)
        up_xy = node_xy.get(up) if up else None
        here3 = at(nid)
        up3 = at(up) if up else None

        if len(links) == 2 and up3 is not None and here3 is not None:
            # 관통 — 꺾였으면 엘보 1개. 어느 배관에 달아도 손실은 같으므로
            # 하류 쪽(부모가 아닌 쪽)에 단다.
            downs = [(pid, o) for pid, o in links if o != up]
            if len(downs) == 1:
                pid, other = downs[0]
                o3 = at(other)
                if o3 is not None:
                    # 3차원으로 잰다 — 세로 구간도 이 한 줄로 옳게 갈린다.
                    u, v = _dir3(up3, here3), _dir3(here3, o3)
                    if u is None or v is None:
                        continue
                    ang = _deflect3_deg(u, v)
                    # ★똑바로 지나가는 자리는 부속이 없는 것이 정답이다.
                    #   fitting_rules 는 「직선인지 확신 못 함」과 「부속으로 못
                    #   보냄」을 같은 None 으로 돌려주므로, 여기서 편향 0 을
                    #   먼저 갈라낸다. 안 그러면 헤드 관통 노드가 전부 판정
                    #   불가로 잡혀(실측 30건) 진짜 미해결이 묻힌다.
                    if ang <= _STRAIGHT_EPS_DEG:
                        n_straight += 1
                        continue
                    kinds, bad = fr.elbow_fittings([ang])
                    unresolved_kind += bad
                    for k in kinds:
                        per_pipe[pid]["fittings"].append(k)
                        counts[k] = counts.get(k, 0) + 1
            continue

        if len(links) >= 3:
            # 분기 — 직류티는 계상하지 않는다(fitting_rules 가 가린다).
            downs = [(pid, o) for pid, o in links if o != up]
            # ★위아래로 갈라지는 갈래는 평면에서 «같은 점» 이라 방위를 잴 수
            #   없다. 그러나 가로 본관에서 세로로 빠지는 것은 언제나 분류티다 —
            #   따로 세고, 평면 규칙에는 가로 갈래만 넘긴다.
            vert, flat_downs = [], []
            for pid, o in downs:
                o3 = at(o)
                if here3 is not None and o3 is not None and _is_vertical(here3, o3):
                    vert.append(pid)
                elif node_xy.get(o) is not None:
                    flat_downs.append((pid, node_xy.get(o)))
            for pid in vert:
                per_pipe[pid]["fittings"].append(fr.TEE)
                counts[fr.TEE] = counts.get(fr.TEE, 0) + 1
            if len(flat_downs) < 2:
                continue
            labels, bad = fr.tee_fittings(here, up_xy, flat_downs)
            unresolved_kind += bad
            for pid in labels:
                per_pipe[pid]["fittings"].append(fr.TEE)
                counts[fr.TEE] = counts.get(fr.TEE, 0) + 1

    # 등가길이 — 라이브러리에 없으면 0 으로 메우지 않고 센다.
    unresolved_length = 0
    for pid, rec in per_pipe.items():
        dia = (bores.get(pid) or (None, None))[0]
        total = 0.0
        for kind in rec["fittings"]:
            L = None if dia is None else equivalent_length_m(lib, kind, dia)
            if L is None:
                unresolved_length += 1
                continue
            total += L
        rec["equivalent_length"] = round(total, 3)

    if unresolved_length:
        print(f"[G4] 등가길이 미해결 {unresolved_length}건 — 라이브러리에 그 "
              f"호칭경 값이 없다(0 으로 채우지 않았다).")
    if unresolved_kind:
        print(f"[G4] 부속 판정 불가 {unresolved_kind}건 — 지어내지 않고 셌다.")
    return {"per_pipe": per_pipe, "counts": counts,
            "straight": n_straight,
            "unresolved_kind": unresolved_kind,
            "unresolved_length": unresolved_length}


def build_nozzles(net, *, k_factor, required_pressure_bar=0.0) -> list[dict]:
    """헤드 노드마다 1행. K 값·필요압력은 변환 창의 DTO 를 그대로 쓴다(§G4)."""
    meta = (net or {}).get("nodes_meta_runtime") or {}
    rows = []
    for nid, m in meta.items():
        if str((m or {}).get("type_id", "")) != "head":
            continue
        rows.append({"node": nid, "k_factor_si": float(k_factor),
                     "required_pressure_bar": float(required_pressure_bar)})
    return rows


def build_equipment(net, *, valve_nodes=None) -> list[dict]:
    """기기 표 — 알람밸브(A/V) 1행부터.

    손질 단계에서 찍은 알람밸브 위치가 없으면 **행을 만들지 않는다**(§G4).
    """
    rows = []
    for nid in (valve_nodes or ()):
        rows.append({"node": nid, "kind": "alarm_valve",
                     "lib_id": FITTING_LIB_ID["alarm_valve"]})
    return rows
