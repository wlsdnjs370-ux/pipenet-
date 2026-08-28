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
                   node_z=None, overrides=None) -> dict:
    """배관마다 부속 목록과 등가길이 합.

    `net`     : 제한 전개 kfp dict
    `node_xy` : {node_id: (x, y)} — 평면 좌표(m 든 mm 든 각도만 쓰므로 무관)
    `bores`   : {pipe_id: (dia_mm, source)} — G3 결과
    `parents` : {node_id: 상류 node_id} — 급수원 BFS 의 부모. 없으면 티 판정이
                「상류를 모름」이 되어 전부 티로 두고 판정 불가로 센다.
    `node_z`  : {node_id: 표고} — 없으면 전부 0 으로 본다.
    `overrides`: 사람이 손으로 채운 값. 규칙이 못 가린 자리에만 쓴다 —
                 **판정을 덮어쓰지 않는다.** 자동이 답을 낸 자리는 건드리지
                 않으므로, 이 인자를 줘도 «규칙이 옳게 판정한 값» 은 안 바뀐다::

                     {"kind":   [{"node","pipe","kind","note"}],   # 자리 단위
                      "eq_len": [{"kind","dia","m","note"}]}       # (종류,호칭경) 쌍

                 단위가 다른 이유는 문제의 성격이 다르기 때문이다. 부속 판정은
                 자리마다 기하가 달라 묶을 수 없고, 등가길이는 «라이브러리에
                 그 호칭경이 없다» 는 구멍이라 쌍을 한 번 채우면 그 쌍을 쓰는
                 배관이 한꺼번에 풀린다.

                 쓴 것은 `applied_overrides` 로 돌려준다 — 자동이 낸 값과
                 사람이 넣은 값을 같은 얼굴로 두지 않기 위해서다.

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

    # 사람이 채운 값 — 찾기 쉬운 표로 바꿔 둔다. 값이 숫자가 아니거나 칸이
    # 비면 «없는 것» 으로 본다: 잘못 넣은 값을 조용히 계산에 넣지 않는다.
    ov = overrides or {}
    ov_kind: dict = {}
    for r in (ov.get("kind") or ()):
        k = str((r or {}).get("kind") or "").strip()
        if k:
            ov_kind[(str(r.get("node")), str(r.get("pipe")))] = (k, r.get("note"))
    # ★「직선 — 부속 없음」도 사람이 낼 수 있는 정답이다. 22.5° 미만은 45° 엘보
    #   보다 직선에 가깝지만, collinear merge 가 흡수를 거부한 각이라 프로그램이
    #   단정할 수 없어 판정 불가로 셌다(fitting_rules 주석). 사람은 도면을 보고
    #   단정할 수 있으므로 그 답을 받는다 — 부속을 «안 다는» 것으로 해결된다.
    NO_FITTING = "none"
    ov_eq: dict = {}
    for r in (ov.get("eq_len") or ()):
        try:
            m = float((r or {}).get("m"))
        except (TypeError, ValueError):
            continue
        if m < 0:
            continue
        ov_eq[(str(r.get("kind")), int(r.get("dia")))] = (m, r.get("note"))
    applied_overrides: list = []

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
    # ★세는 그 자리에서 «어느 배관인지» 를 함께 남긴다(순수 추가). 예전에는
    #   `+= 1` 만 하고 버려서, 화면이 「3건」만 알고 «어디» 를 몰랐다. 그러면
    #   사람이 손으로 채울 수가 없다. 개수는 여전히 여기서만 정해지므로,
    #   목록과 개수가 어긋날 수 없다.
    unresolved_kind_items: list = []

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
                    if bad:
                        # 사람이 이 자리를 채웠으면 그것을 쓴다. 규칙이 답을
                        # 낸 자리(bad == 0)에는 손대지 않는다.
                        hit = ov_kind.get((str(nid), str(pid)))
                        if hit:
                            if hit[0] != NO_FITTING:
                                kinds = list(kinds) + [hit[0]] * int(bad)
                            bad = 0
                            applied_overrides.append(
                                {"what": "kind", "node": str(nid),
                                 "pipe": str(pid), "kind": hit[0],
                                 "note": hit[1]})
                    unresolved_kind += bad
                    if bad:
                        unresolved_kind_items.append(
                            {"pipe": str(pid), "node": str(nid), "where": "관통",
                             "n": int(bad), "angle_deg": round(float(ang), 1)})
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
            if bad:
                # 분기는 «자리» 가 단위다 — 대표 배관으로 찾는다. 사람이 고른
                # 종류를 그대로 단다(티로 굳히지 않는다 — 직류/분류가 갈린다).
                rep = str(flat_downs[0][0])
                hit = ov_kind.get((str(nid), rep))
                if hit:
                    if hit[0] != NO_FITTING:
                        for _ in range(int(bad)):
                            per_pipe[rep]["fittings"].append(hit[0])
                            counts[hit[0]] = counts.get(hit[0], 0) + 1
                    applied_overrides.append(
                        {"what": "kind", "node": str(nid), "pipe": rep,
                         "kind": hit[0], "note": hit[1]})
                    bad = 0
            unresolved_kind += bad
            if bad:
                # 분기는 «자리» 가 단위다 — 어느 갈래인지까지는 규칙이 못 가른
                # 것이므로, 그 노드에 걸린 갈래들을 함께 남긴다.
                unresolved_kind_items.append(
                    {"pipe": str(flat_downs[0][0]), "node": str(nid),
                     "where": "분기", "n": int(bad),
                     "branches": [str(p) for p, _ in flat_downs]})
            for pid in labels:
                per_pipe[pid]["fittings"].append(fr.TEE)
                counts[fr.TEE] = counts.get(fr.TEE, 0) + 1

    # 등가길이 — 라이브러리에 없으면 0 으로 메우지 않고 센다.
    unresolved_length = 0
    # 세는 자리에서 «어느 배관의 어느 부속·어느 호칭경» 인지 함께 남긴다.
    # 라이브러리 구멍은 (종류, 호칭경) 쌍이 단위라 그 쌍도 따로 모아 둔다 —
    # 사람이 한 번 채우면 같은 쌍을 쓰는 배관이 한꺼번에 풀린다.
    unresolved_length_items: list = []
    unresolved_pairs: dict = {}
    for pid, rec in per_pipe.items():
        dia = (bores.get(pid) or (None, None))[0]
        total = 0.0
        for kind in rec["fittings"]:
            L = None if dia is None else equivalent_length_m(lib, kind, dia)
            if L is None and dia is not None:
                # 라이브러리 구멍을 사람이 채웠으면 그 값을 쓴다. 라이브러리에
                # 값이 «있는» 자리는 건드리지 않는다 — 덮어쓰기가 아니라 채우기다.
                hit = ov_eq.get((str(kind), int(dia)))
                if hit:
                    L = hit[0]
                    applied_overrides.append(
                        {"what": "eq_len", "kind": str(kind), "dia": int(dia),
                         "m": hit[0], "note": hit[1], "pipe": str(pid)})
            if L is None:
                unresolved_length += 1
                unresolved_length_items.append(
                    {"pipe": str(pid), "kind": str(kind),
                     "dia": (int(dia) if dia is not None else None)})
                key = (str(kind), int(dia) if dia is not None else -1)
                unresolved_pairs[key] = unresolved_pairs.get(key, 0) + 1
                continue
            total += L
        rec["equivalent_length"] = round(total, 3)

    if unresolved_length:
        print(f"[G4] 등가길이 미해결 {unresolved_length}건 — 라이브러리에 그 "
              f"호칭경 값이 없다(0 으로 채우지 않았다).")
    if unresolved_kind:
        print(f"[G4] 부속 판정 불가 {unresolved_kind}건 — 지어내지 않고 셌다.")
    if applied_overrides:
        n_k = sum(1 for a in applied_overrides if a["what"] == "kind")
        n_e = sum(1 for a in applied_overrides if a["what"] == "eq_len")
        print(f"[G4] ★직접 입력 적용 — 부속 {n_k}자리 · 등가길이 {n_e}건. "
              f"자동이 못 가린 자리에만 썼다.")
    return {"per_pipe": per_pipe, "counts": counts,
            "straight": n_straight,
            "unresolved_kind": unresolved_kind,
            "unresolved_length": unresolved_length,
            # ★아래 셋은 순수 추가다. 기존 호출자는 안 읽으므로 영향이 없고,
            #   개수는 여전히 위에서만 정해지므로 목록과 어긋날 수 없다.
            "unresolved_kind_items": unresolved_kind_items,
            "unresolved_length_items": unresolved_length_items,
            # 사람이 넣은 값을 쓴 자리 — 자동이 낸 값과 같은 얼굴로 두지 않는다.
            "applied_overrides": applied_overrides,
            "unresolved_pairs": [{"kind": k, "dia": (d if d >= 0 else None),
                                  "n": n}
                                 for (k, d), n in sorted(unresolved_pairs.items())]}


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
