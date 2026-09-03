# -*- coding: utf-8 -*-
"""[G5] 5개 테이블 조립 — 제한 전개 망 → PipeTablesG.

이번 작업에서 **유일하게 새로 짜는 것**이다(§5). 모듈 A 의 `build_input_tables`
는 A 전용 선정 결과 타입에 묶여 있어 그대로 못 쓴다. 키 규약·단위는 A 의
`PipeTables` 를 **그대로** 따른다 — 키 하나만 달라도 SDF 가 조용히 비거나
관경이 "Unset" 이 된다(§3).

★행 순서 = 물이 흐르는 순서. 급수원을 뿌리로 **BFS 한 번**을 돌려 노드 라벨
번호·행 순서·배관 방향(`in`→`out`)을 전부 그 순서에 맞춘다. 표를 위에서 아래로
읽으면 물길이 된다. 트리에 못 들어간 간선(루프 잔여)은 **표 꼬리로 몰아** 둔다 —
그 꼬리가 곧 「길이 잘못 트인」 후보 목록이다.

★단위(§T3) — 같은 dict 안에 mm 와 m 가 섞인다. 헷갈리면 SDF 가 통째로 틀어진다.
    nodes.x / nodes.y      mm   (전개 결과는 m 이므로 되돌려 곱한다)
    nodes.elevation        m
    pipes.length / elev    m
    pipes.dia              호칭경 mm
    압력                    Pa
"""
from __future__ import annotations

import math
from collections import deque
from dataclasses import dataclass, field

from services.cad_import.design.anchor import require_anchor

# 전개 좌표(m) → 표 좌표(mm)
M_TO_MM = 1000.0
# 관종의 권위는 `design/sdf_post.SCHEDULE_DEFS` 다(SLF 의 Item-name 과 정합).
# 여기서 다시 적으면 두 곳이 갈라지므로, 기본값도 그쪽에서 받는다.
DEFAULT_C = 120


@dataclass
class PipeTablesG:
    """모듈 A 의 `PipeTables` 키 규약을 그대로 따른다(§1 · §3)."""

    nodes: list = field(default_factory=list)      # {label,elevation,io_node,x,y}
    pipes: list = field(default_factory=list)      # {label,in,out,type,dia,length,elev,c,status,group}
    nozzles: list = field(default_factory=list)    # {label,in,out,status,lib,flow_m3s,flow_lmin}
    fittings: list = field(default_factory=list)   # {pipe,in,out,type,count}
    equipment: list = field(default_factory=list)  # {pipe,in,out,label,desc,eq_len,rel_pos}
    meta: list = field(default_factory=list)       # [(key, value)]
    # 「부속 판정 불가」·「등가길이 미해결」이 **어느 배관인지**. meta 는 개수만
    # 담아 왔는데, 개수만으로는 사람이 손으로 채울 수가 없다. 세는 자리
    # (`build_fittings`)가 이미 아는 것을 버리지 않고 여기까지 들고 온다.
    #   {"kind_items": [...], "length_items": [...], "pairs": [...]}
    unresolved: dict = field(default_factory=dict)
    # 사람이 관경을 덮은 자리 — **원값·원출처와 함께**(D-F11-3). 관경은 부속과
    # 달리 규칙 값도 덮으므로, 전·후가 같이 안 남으면 나중에 그 수치를 누가
    # 왜 정했는지 알 길이 없다. SDF 형식에는 출처 칸이 없어 여기가 그 자리다.
    #   {pipe_id: {dia, note, orig_dia, orig_src, a, b}}
    bore_overrides: dict = field(default_factory=dict)

    def as_dict(self) -> dict:
        return {"nodes": self.nodes, "pipes": self.pipes,
                "nozzles": self.nozzles, "fittings": self.fittings,
                "equipment": self.equipment,
                "meta": [list(m) for m in self.meta],
                "unresolved": self.unresolved,
                "bore_overrides": self.bore_overrides}


def _ends(pr):
    return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")


def bfs_order(net, root):
    """급수원 뿌리 BFS — (노드 방문 순서, 부모 맵, 트리 간선, 트리 밖 간선).

    부모 맵은 G4 의 부속 판정(상류 방향)에도 쓰인다. 두 곳이 다른 순서를 쓰면
    「가지로는 분류됐는데 티는 안 달린」 조합이 생기므로 한 번만 돈다.
    """
    pipes = (net or {}).get("pipe_data") or {}
    adj: dict = {}
    for pid, pr in pipes.items():
        a, b = _ends(pr)
        if not a or not b:
            continue
        adj.setdefault(a, []).append((b, pid))
        adj.setdefault(b, []).append((a, pid))

    order: list = []
    parent: dict = {}
    tree_pipes: list = []      # (pid, in_node, out_node) — 물 흐르는 방향
    seen = {root} if root else set()
    q = deque([root] if root else [])
    used_pipes = set()
    while q:
        u = q.popleft()
        order.append(u)
        for v, pid in sorted(adj.get(u, ()), key=lambda t: t[1]):
            if v in seen:
                continue
            seen.add(v)
            parent[v] = u
            used_pipes.add(pid)
            tree_pipes.append((pid, u, v))
            q.append(v)
    # 트리에 못 들어간 간선 — 루프 잔여. 버리지 않고 꼬리로 몬다.
    off_tree = [(pid, *_ends(pr)) for pid, pr in pipes.items()
                if pid not in used_pipes]
    return order, parent, tree_pipes, off_tree


def build_design_tables(net, worst, edge_ref, dia_text_pts, *,
                        project_title="Module G 수리계산 입력",
                        bores=None, fittings=None, nozzle_k=80.0,
                        nozzle_flow_lmin=80, valve_nodes=None,
                        excluded_heads=0, board_pts=None,
                        default_schedule=None,
                        schedule_by_pipe=None,
                        tree_loads=None,
                        fitting_overrides=None,
                        bore_overrides=None,
                        origin_mm=None) -> PipeTablesG:
    """제한 전개 망 → 5개 테이블. 지시서 §1 공개 시그니처.

    `bores` / `fittings` 는 G3 · G4 결과를 받는다. 없으면 여기서 만들지 않고
    비워 둔다 — 이 모듈은 «조립» 이지 판정이 아니다.

    `bore_overrides` 는 그 «없으면 만드는» 자리로만 흘러간다 — `bores` 를
    이미 받았으면 그것이 권위다. 두 곳에서 덮으면 어느 것이 이겼는지 알 수 없다.
    """
    from services.cad_import.design.bore import decide_bores, source_counts
    from services.cad_import.design.fitting import build_fittings
    from services.cad_import.design.sdf_post import DEFAULT_SCHEDULE, check_schedule

    # 관종 이름을 **먼저** 검사한다. 오타가 조용히 기본값으로 떨어지면 PIPENET
    # 에서 다시 "None defined" 가 되고, 그때는 원인을 도면 탓으로 오해하게 된다.
    default_schedule = check_schedule(default_schedule or DEFAULT_SCHEDULE)
    schedule_by_pipe = {str(k): check_schedule(v)
                        for k, v in (schedule_by_pipe or {}).items()}

    meta_nodes = (net or {}).get("nodes_meta_runtime") or {}
    pipes_raw = (net or {}).get("pipe_data") or {}

    def xy(nid):
        c = (meta_nodes.get(nid) or {}).get("coords") or (0.0, 0.0, 0.0)
        return float(c[0]), float(c[1])

    def z(nid):
        c = (meta_nodes.get(nid) or {}).get("coords") or (0.0, 0.0, 0.0)
        return float(c[2]) if len(c) > 2 else 0.0

    # ── 뿌리 = 접속점(알람밸브 자리). 없으면 표를 만들 수 없다 — 그러니 던진다.
    #   종전엔 `next(iter(meta_nodes))` 로 눕어, dict 에 먼저 들어온 노드가
    #   Input 경계가 되고 물 흐르는 방향이 통째로 거기서 유도됐다(design/anchor).
    root = require_anchor(meta_nodes, what="설계 표")
    order, parent, tree_pipes, off_tree = bfs_order(net, root)

    # ── 관경·부속 (없으면 지금 만든다)
    if bores is None:
        bores = decide_bores(net, edge_ref, (worst or {}).get("loads") or {},
                             dia_text_pts, pts=board_pts,
                             tree_loads=tree_loads,
                             overrides=bore_overrides)
    node_xy = {n: xy(n) for n in meta_nodes}
    node_z = {n: z(n) for n in meta_nodes}
    if fittings is None:
        # 표고를 함께 넘긴다 — 세로 구간은 평면 좌표만으로 판정할 수 없다(§G19).
        fittings = build_fittings(net, node_xy, bores, parents=parent,
                                  node_z=node_z,
                                  overrides=fitting_overrides)

    # ── ① 노드표 — BFS 순서대로 번호. 뿌리가 Input.
    label_of: dict = {}
    for i, nid in enumerate(order, start=1):
        label_of[nid] = str(i)
    # BFS 에 안 잡힌 노드(고립)도 표에는 있어야 배관표가 고아가 안 된다.
    for nid in meta_nodes:
        if nid not in label_of:
            label_of[nid] = str(len(label_of) + 1)

    tbl = PipeTablesG()
    for nid, lab in sorted(label_of.items(), key=lambda kv: int(kv[1])):
        x_mm, y_mm = xy(nid)
        row = {
            "label": lab,
            "x": int(round(x_mm * M_TO_MM)),      # §T3 — 노드 좌표만 mm
            "y": int(round(y_mm * M_TO_MM)),
            "elevation": round(z(nid), 3),        # m
            "io_node": "Input" if nid == root else "No",
        }
        if nid == root:
            row["pressure_pa"] = 101325.0
        tbl.nodes.append(row)

    # ── ② 배관표 — 트리 순서(물 흐르는 방향) → 꼬리에 루프 잔여
    def pipe_row(pid, a, b, *, off):
        pr = pipes_raw.get(pid) or {}
        dia, src = (bores.get(pid) or (0, "?"))
        row = {
            "label": pid,
            "in": label_of.get(a, "?"), "out": label_of.get(b, "?"),
            "type": schedule_by_pipe.get(pid, default_schedule),
            "dia": int(dia),                       # 호칭경 mm
            "length": round(float(pr.get("length_m") or 0.0), 3),   # m
            "elev": round(z(b) - z(a), 3),                          # m
            "c": DEFAULT_C, "status": "Normal", "group": "Unset",
            # 관경을 무엇이 정했는지 행에 남긴다 — 요약의 집계만으로는 «이 배관»
            # 이 도면 텍스트에서 온 것인지 별표1 폴백인지 확인할 길이 없다(§G16).
            # SDF 방출은 이름 붙인 칸만 읽으므로 이 칸이 파일을 바꾸지 않는다.
            "dia_src": src,
        }
        if off:
            row["off_tree"] = True     # 「길이 잘못 트인」 후보 — 꼬리에 몰린다
        return row

    for pid, a, b in tree_pipes:
        tbl.pipes.append(pipe_row(pid, a, b, off=False))
    for pid, a, b in off_tree:
        tbl.pipes.append(pipe_row(pid, a, b, off=True))

    # ── ③ 노즐표 — 헤드 노드마다 1행
    for i, nid in enumerate(
            [n for n in order if str((meta_nodes.get(n) or {}).get("type_id")) == "head"],
            start=1):
        tbl.nozzles.append({
            "label": str(i), "in": label_of.get(nid, "?"), "out": f"@/{i}",
            "status": "1", "lib": "SP-HEAD",
            "flow_lmin": nozzle_flow_lmin,
            # m³/s 는 L/min 에서 유도한다 — 손으로 자른 상수를 쓰면 어긋난다.
            "flow_m3s": nozzle_flow_lmin / 60000.0,
        })

    # ── ④ 부속표 — 배관표에 있는 라벨만(고아 참조 0)
    pipe_by_label = {r["label"]: r for r in tbl.pipes}
    for pid, rec in (fittings.get("per_pipe") or {}).items():
        prow = pipe_by_label.get(pid)
        if prow is None:
            continue
        for kind in rec.get("fittings") or ():
            tbl.fittings.append({
                "pipe": pid, "in": prow["in"], "out": prow["out"],
                "type": kind, "count": "1",
            })

    # ── ⑤ 기기표 — 알람밸브. 찍은 것이 없으면 행을 만들지 않는다.
    #
    # ★등가길이는 부속표와 **같은 함수**로 정한다(`resolve_eq_len`).
    #   종전에는 여기가 `"eq_len": 0.0` 으로 박혀 있었다 — 라이브러리에 값이
    #   버젓이 있는데도(실측: 알람밸브 100A = 9.5m) SDF 에는 0 이 실렸다.
    #   0 은 「손실이 없다」는 주장이라, 그만큼 계산이 낙관적으로 틀어진다.
    from services.cad_import.design.fitting import (
        parse_eq_len_overrides, resolve_eq_len)
    _ov_eq = parse_eq_len_overrides(fitting_overrides)
    av_unresolved: list = []
    for nid in (valve_nodes or ()):
        lab = label_of.get(nid)
        if lab is None:
            continue
        host = next((r for r in tbl.pipes if lab in (r["in"], r["out"])), None)
        if host is None:
            continue
        eq_m, eq_why = resolve_eq_len("alarm_valve", host.get("dia"),
                                      ov_eq=_ov_eq)
        if eq_m is None:
            # 못 구한 것을 0 으로 채우지 않는다 — 대신 어디가 빈지 남긴다.
            av_unresolved.append({"pipe": host["label"],
                                  "dia": host.get("dia")})
        row = {
            "pipe": host["label"], "in": host["in"], "out": host["out"],
            "label": str(len(tbl.equipment) + 1), "desc": "A/V",
            "eq_len": (0.0 if eq_m is None else round(eq_m, 3)),
            "rel_pos": 0.5,
        }
        if eq_m is not None:
            row["eq_len_src"] = eq_why
        tbl.equipment.append(row)

    # ── meta — 근거를 남긴다(§G5). 화면과 산출물이 같은 말을 하게 한다.
    src = source_counts(bores)
    worst_head_label = label_of.get(
        _worst_head_node(net, worst, meta_nodes,
                         board_pts=board_pts, origin_mm=origin_mm), "?")
    tbl.meta = [
        ("제목", project_title),
        ("기준개수 K", str(len((worst or {}).get("heads") or []))),
        ("기준 헤드 노드", worst_head_label),
        ("최원 유하거리 (m)", str((worst or {}).get("far_m", ""))),
        ("설계면적 폭 (m)", str((worst or {}).get("span_m", ""))),
        ("corridor 총연장 (m)", str((worst or {}).get("total_m", ""))),
        ("주배관 담당 헤드 수", str((worst or {}).get("max_load", ""))),
        ("배관 규격(기본)", default_schedule),
        ("관경 근거 — 도면 텍스트", str(src.get("text", 0))),
        ("관경 근거 — 별표1 보강 (text<min)", str(src.get("nfpc_min", 0))),
        ("관경 근거 — 별표1 폴백 (text 없음)", str(src.get("nfpc_fallback", 0))),
        # 관경은 규칙 값도 덮는다(D-F11-3) — 그래서 덮은 수를 규칙 근거와
        # 나란히 세운다. 한 줄로 「이 도면의 관경을 무엇이 정했나」가 읽힌다.
        ("직접 입력 — 관경", str(src.get("user", 0))),
        # 개수는 여전히 `build_fittings` 한 곳에서만 정해진다 — 아래 목록도
        # 같은 자리에서 나오므로 둘이 어긋날 수 없다.
        ("부속 판정 불가", str(fittings.get("unresolved_kind", 0))),
        # 부속표(배관에 딸린 것)와 기기표(알람밸브)를 **함께** 센다 — 같은
        # 「등가길이」라는 한 칸이므로 두 수를 따로 두면 사람이 하나를 놓친다.
        ("등가길이 미해결",
         str(int(fittings.get("unresolved_length", 0)) + len(av_unresolved))),
        # ★사람이 넣은 값을 쓴 자리는 **산출물에도** 남긴다. 자동이 낸 값과
        #   같은 얼굴로 두면, 나중에 그 수치를 누가 정했는지 알 수 없다.
        ("직접 입력 — 부속 판정",
         str(sum(1 for a in (fittings.get("applied_overrides") or ())
                 if a.get("what") == "kind"))),
        ("직접 입력 — 등가길이",
         str(sum(1 for a in (fittings.get("applied_overrides") or ())
                 if a.get("what") == "eq_len"))),
        ("루프 잔여 배관(표 꼬리)", str(len(off_tree))),
        # ★B4 1안 — 전개가 못 붙인 헤드는 후보에서 뺐다. 조용히 빼면 「더 불리한
        #   헤드가 있는데 못 본 채」 수리계산이 나간다. 산출물에도 남긴다.
        ("전개가 못 붙여 제외한 헤드", str(excluded_heads)),
        ("설계구역 선정", "모듈 G 기준헤드 방식 (SDF 전용 · .kfp 는 솔버가 따로 고른다)"),
    ]
    # 미해결이 «어느 배관인지» — 개수와 같은 자리에서 나온 목록이다.
    tbl.unresolved = {
        "kind_items": list(fittings.get("unresolved_kind_items") or ()),
        # 알람밸브 행도 같은 목록에 넣는다 — 화면이 「어디를 채워야 하나」를
        # 한 곳에서 읽는다. kind 를 붙여 두어야 채울 칸을 특정할 수 있다.
        "length_items": list(fittings.get("unresolved_length_items") or ())
        + [{"pipe": r["pipe"], "kind": "alarm_valve", "dia": r["dia"]}
           for r in av_unresolved],
        "pairs": list(fittings.get("unresolved_pairs") or ()),
        # 사람이 넣은 값을 쓴 자리 — 화면이 「직접 입력」이라고 밝힐 재료다.
        "applied": list(fittings.get("applied_overrides") or ()),
    }
    # 관경을 덮은 자리도 같은 규약으로 들고 온다. `decide_bores` 가 곁에 붙여
    # 보낸 것을 그대로 옮길 뿐이다 — 여기서 다시 세지 않는다(두 벌 금지).
    tbl.bore_overrides = dict(getattr(bores, "overridden", None) or {})
    return tbl


# 부착점 x,y 에서 이만큼 안에 헤드 노드가 «하나만» 있으면 그것이 그 헤드다.
#
# ★숫자의 근거(2026-09-03 · B1F 실측 30개). 세 종류 모두 부착점 x,y 자리에
#   헤드 노드가 하나 생긴다 — 세로 전개는 z 를 옮기지 x,y 를 옮기지 않는다.
#   상하향식만 «하향» 쪽이 combo_2(기본 0.3m)만큼 밀리는데, 같은 헤드의
#   «상향» 쪽은 제자리에 남으므로 여기 걸리는 것은 언제나 있다.
#     최근접 거리   6 ~ 19 mm  (중앙 6)
#     2등/1등 거리비 최소 141.7배 · 500mm 안에 둘 이상 = 0건
#   100mm 는 실측 최대(19mm)의 다섯 배이면서, 이웃 헤드(실측 2,950mm)에는
#   한참 못 미친다. 넘어가면 «못 찾았다» 로 둔다 — 틀린 헤드를 가리키느니.
WORST_HEAD_SNAP_M = 0.10


def _worst_head_node(net, worst, meta_nodes, *, board_pts=None,
                     origin_mm=None):
    """기준 헤드(최원단)에 해당하는 kfp 노드. 못 찾으면 None.

    ★「앵커」라 부르지 않는다 — 이 저장소에서 앵커는 **접속점**(라이저가 붙는
      자리)이다. 여기는 정반대 끝, 급수에서 가장 먼 헤드다(design/anchor).

    ■ 종전에 왜 못 이었나 (BLOCKED §30)

      「board 헤드 번호는 전개 노드와 1:1 이 아니다」— 맞는 말이었지만, 그래서
      **아무것도 안 했다.** 실측으로 원인을 갈라 보니 둘이었다.
        · `node_ref`(kfp 노드 → board 노드)로 되짚기: 뽑힌 헤드 30개 중 12개만
          맞는다. 그 표는 «노드정리 전» 의 id 를 담고 있어(323 → 80) 절반 넘게
          이미 사라진 노드를 가리킨다.
        · 좌표로 맞대기: 30개 전부 6~19mm 안에서 유일하게 걸린다.
      즉 못 이을 이유는 없었고, **되짚는 표를 잘못 고른 것**이었다.

    ■ 지금 잇는 방법

      최불리가 돌려준 `worst_path` 의 마지막 절점이 곧 기준 헤드의 board
      부착 노드다. 그 자리를 kfp 좌표로 옮겨(origin_mm) 가장 가까운 헤드 노드를
      집되, `WORST_HEAD_SNAP_M` 안에 **하나뿐일 때만** 받는다.

      재료(`origin_mm`)가 없으면 종전처럼 None 이다 — 좌표계를 모르는 채로
      «가장 먼 헤드» 같은 추측을 하지 않는다.
    """
    hi = (worst or {}).get("worst_head")
    if hi is None or origin_mm is None or not board_pts:
        return None
    path = (worst or {}).get("worst_path") or ()
    if not path:
        return None
    bn = path[-1]
    if not isinstance(bn, int) or not (0 <= bn < len(board_pts)):
        return None
    bx, by = float(board_pts[bn][0]), float(board_pts[bn][1])
    # board mm → kfp m. 규칙은 `convert/main_walk.xf_mm_to_m` 하나뿐이다
    # (좌표계 계약은 tests/test_module_f_coord_contract.py 가 지킨다).
    tx = (bx - float(origin_mm[0])) / 1000.0 + 1.0
    ty = (by - float(origin_mm[1])) / 1000.0 + 1.0

    near = []
    for nid, meta in (meta_nodes or {}).items():
        if str((meta or {}).get("type_id", "")) != "head":
            continue
        c = (meta or {}).get("coords") or (0.0, 0.0, 0.0)
        d = math.hypot(float(c[0]) - tx, float(c[1]) - ty)
        if d <= WORST_HEAD_SNAP_M:
            near.append((d, nid))
    if len(near) != 1:
        # 0개면 전개가 그 헤드를 못 붙였다는 뜻이고, 2개 이상이면 어느 쪽인지
        # 모른다. 둘 다 «찍어서 맞히는» 것보다 모른다고 두는 편이 옳다.
        return None
    return near[0][1]
