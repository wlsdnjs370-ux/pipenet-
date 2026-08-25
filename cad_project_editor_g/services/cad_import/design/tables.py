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

from collections import deque
from dataclasses import dataclass, field

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

    def as_dict(self) -> dict:
        return {"nodes": self.nodes, "pipes": self.pipes,
                "nozzles": self.nozzles, "fittings": self.fittings,
                "equipment": self.equipment,
                "meta": [list(m) for m in self.meta]}


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
                        tree_loads=None) -> PipeTablesG:
    """제한 전개 망 → 5개 테이블. 지시서 §1 공개 시그니처.

    `bores` / `fittings` 는 G3 · G4 결과를 받는다. 없으면 여기서 만들지 않고
    비워 둔다 — 이 모듈은 «조립» 이지 판정이 아니다.
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

    # ── 뿌리 = 급수원(펌프). 없으면 표를 만들 수 없다.
    root = next((n for n, m in meta_nodes.items()
                 if str((m or {}).get("type_id", "")) == "pump"), None)
    if root is None:
        root = next(iter(meta_nodes), None)
    order, parent, tree_pipes, off_tree = bfs_order(net, root)

    # ── 관경·부속 (없으면 지금 만든다)
    if bores is None:
        bores = decide_bores(net, edge_ref, (worst or {}).get("loads") or {},
                             dia_text_pts, pts=board_pts,
                             tree_loads=tree_loads)
    node_xy = {n: xy(n) for n in meta_nodes}
    node_z = {n: z(n) for n in meta_nodes}
    if fittings is None:
        # 표고를 함께 넘긴다 — 세로 구간은 평면 좌표만으로 판정할 수 없다(§G19).
        fittings = build_fittings(net, node_xy, bores, parents=parent,
                                  node_z=node_z)

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
    for nid in (valve_nodes or ()):
        lab = label_of.get(nid)
        if lab is None:
            continue
        host = next((r for r in tbl.pipes if lab in (r["in"], r["out"])), None)
        if host is None:
            continue
        tbl.equipment.append({
            "pipe": host["label"], "in": host["in"], "out": host["out"],
            "label": str(len(tbl.equipment) + 1), "desc": "A/V",
            "eq_len": 0.0, "rel_pos": 0.5,
        })

    # ── meta — 근거를 남긴다(§G5). 화면과 산출물이 같은 말을 하게 한다.
    src = source_counts(bores)
    anchor_label = label_of.get(_anchor_node(net, worst, meta_nodes), "?")
    tbl.meta = [
        ("제목", project_title),
        ("기준개수 K", str(len((worst or {}).get("heads") or []))),
        ("앵커 노드", anchor_label),
        ("최원 유하거리 (m)", str((worst or {}).get("far_m", ""))),
        ("설계면적 폭 (m)", str((worst or {}).get("span_m", ""))),
        ("corridor 총연장 (m)", str((worst or {}).get("total_m", ""))),
        ("주배관 담당 헤드 수", str((worst or {}).get("max_load", ""))),
        ("배관 규격(기본)", default_schedule),
        ("관경 근거 — 도면 텍스트", str(src.get("text", 0))),
        ("관경 근거 — 별표1 보강 (text<min)", str(src.get("nfpc_min", 0))),
        ("관경 근거 — 별표1 폴백 (text 없음)", str(src.get("nfpc_fallback", 0))),
        ("부속 판정 불가", str(fittings.get("unresolved_kind", 0))),
        ("등가길이 미해결", str(fittings.get("unresolved_length", 0))),
        ("루프 잔여 배관(표 꼬리)", str(len(off_tree))),
        # ★B4 1안 — 전개가 못 붙인 헤드는 후보에서 뺐다. 조용히 빼면 「더 불리한
        #   헤드가 있는데 못 본 채」 수리계산이 나간다. 산출물에도 남긴다.
        ("전개가 못 붙여 제외한 헤드", str(excluded_heads)),
        ("설계구역 선정", "모듈 G 앵커 방식 (SDF 전용 · .kfp 는 솔버가 따로 고른다)"),
    ]
    return tbl


def _anchor_node(net, worst, meta_nodes):
    """앵커 헤드에 해당하는 kfp 노드. 못 찾으면 None."""
    an = (worst or {}).get("anchor")
    if an is None:
        return None
    # 앵커는 board 헤드 번호다. 전개 노드와 1:1 이 아니므로 «헤드 노드 중
    # 가장 먼 것» 으로 되짚지 않고, 못 찾으면 솔직히 못 찾았다고 둔다.
    return None
