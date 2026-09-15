# -*- coding: utf-8 -*-
"""[가지치기·부속판정·두사전 §2] 바꾸기 전에 재는 여섯 줄 — 자만 있는 모듈.

지시서 `ModuleF_가지치기_부속판정_두사전_지시서.md` §2. **아무것도 안 고친다** —
지금 코드가 무엇을 내고 있는지만 센다. 새 규칙은 여기서 «판정만» 돌려 본다
(파일을 쓰지 않는다).

`_probe_plan_preserve.py` 가 이 자들을 불러 쓴다. 옆 모듈로 뺀 이유는
`_pp_checks.py`·`_pp_chain.py` 와 같다 — 프로브 본체는 «순서» 를 읽는 자리이고,
자는 자기 파일에서 혼자 시험될 수 있어야 한다.

■ 왜 이 계측이 필요한가 (그림 16)

  가지치기는 계산 범위를 좁히는 일이지 배관을 뜯어내는 일이 아니다. 그런데
  지금은 가지가 지워져 차수가 3 → 2 로 떨어진 자리가 «부속 없음» 이나 «엘보» 로
  나온다 — 분류티 등가길이 3.0 m 자리에 엘보 1.5 m 가 실린다(50A 기준).
  그 자리가 **몇 개인지** 먼저 센다.
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

_REPO = Path(__file__).resolve().parent.parent
for _p in (str(_REPO), str(_REPO / "core"), str(_REPO / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


# ─────────────────────────────────────────────── 잡동사니
def _ends(pr):
    return (pr.get("start") or pr.get("from"),
            pr.get("end") or pr.get("to"))


def _nodes(kfp):
    return (kfp or {}).get("nodes_meta_runtime") or {}


def _pipes(kfp):
    return (kfp or {}).get("pipe_data") or {}


class Tee:
    """로그를 화면에 그대로 흘리면서 **사본도 모은다**.

    `끝배관1단보호` 는 planar 가 로그로만 내고 반환에는 없다. 반환에 키를 더하면
    지시서 §5 의 「planar.py diff = 키워드 통과뿐」을 깨므로, 로그를 읽는다.
    """

    def __init__(self, real):
        self._real = real
        self.buf = []

    def write(self, s):
        self.buf.append(s)
        return self._real.write(s)

    def flush(self):
        return self._real.flush()

    def text(self):
        return "".join(self.buf)


def stub_count_from_log(text: str) -> int | None:
    """planar 로그에서 «끝배관1단보호 n» 을 읽는다. 여러 번이면 **마지막**.

    ★마지막인 이유: 스냅을 껐다가 회랑이 쪼개져 되돌리면 planar 가 두 번 돈다.
      실제로 쓰인 것은 나중 것이다.
    """
    hits = re.findall(r"끝배관1단보호\s+(\d+)", text or "")
    return int(hits[-1]) if hits else None


# ─────────────────────────────────────────────── [잔류]
def residual_stubs(tbl):
    """표에서 «차수 1 인데 헤드도 급수원도 아닌» 노드와 그 배관 길이 합.

    가지치기가 남긴 잔류 스텁이 표까지 흘러온 자리다(그림 16 ②의 빨간 선).
    """
    deg = defaultdict(int)
    inc = defaultdict(list)
    for r in (getattr(tbl, "pipes", None) or ()):
        a, b = str(r.get("in")), str(r.get("out"))
        deg[a] += 1
        deg[b] += 1
        inc[a].append(r)
        inc[b].append(r)
    heads = {str(z.get("in")) for z in (getattr(tbl, "nozzles", None) or ())}
    io_in = {str(n.get("label")) for n in (getattr(tbl, "nodes", None) or ())
             if str(n.get("io_node")) == "Input"}
    bad = [lab for lab, d in deg.items()
           if d == 1 and lab not in heads and lab not in io_in]
    length = 0.0
    for lab in bad:
        for r in inc[lab]:
            length += float(r.get("length") or 0.0)
    return sorted(bad), round(length, 3)


# ─────────────────────────────────────────────── [차수]
GRID_M = 0.05          # planar 의 격자 — 여기서 바꾸면 자가 갈린다


def fold_clusters(pts, used, origin_mm):
    """격자로 **한 칸에 접히는** board 절점 무리 — planar 의 `fold_key` 그대로.

    ★이것이 없으면 차수를 과소평가한다. `node_ref` 는 한 칸당 **한 vid** 만
      적어 두는데(`node_ref.setdefault(node_id[tgt], vid)` · `used` 가 set 이라
      어느 vid 가 적힐지도 임의다), 그 한 점의 G 차수는 «칸 전체» 의 차수가
      아니다. 티 둘이 한 칸에 들면 크로스 하나가 되는 그 자리다.

    반환: (cell_of_vid, vids_of_cell)
    """
    o = origin_mm or (0.0, 0.0)
    minx, miny = float(o[0]), float(o[1])

    def key(x_mm, y_mm):
        mx = (float(x_mm) - minx) / 1000.0 + 1.0
        my = (float(y_mm) - miny) / 1000.0 + 1.0
        return (round(round(mx / GRID_M) * GRID_M, 3),
                round(round(my / GRID_M) * GRID_M, 3))

    cell_of, of_cell = {}, defaultdict(set)
    for vid in used:
        try:
            c = key(pts[vid][0], pts[vid][1])
        except (IndexError, TypeError):
            continue
        cell_of[int(vid)] = c
        of_cell[c].add(int(vid))
    return cell_of, of_cell


def physical_degree(limited, built, kfp):
    """`phys` — 지우기 **전** 배관망 G 의 차수 (지시서 §1-3).

    `phys[nid] = max(deg_G[node_ref[nid]] + n_vert[nid], deg_kfp[nid])`

    ★`max` 인 이유: `normalize_tee_overlaps` 가 허브에서 관통을 쪼개 만든 접속은
      G 에 없어도 **물리 접속**이다. 그 수(`extra`)를 함께 돌려준다 —
      지시서 §2 의 「deg_kfp > deg_G + n_vert 인 노드」 줄이 그것이다.

    반환: (phys, deg_kfp, n_vert, deg_G, extra_nodes, no_ref_nodes)
    """
    edges_G = [(int(e[0]), int(e[1])) for e in (limited.get("edges") or ())]
    deg_G = defaultdict(int)
    for a, b in edges_G:
        deg_G[a] += 1
        deg_G[b] += 1

    node_ref = {str(k): int(v) for k, v in (built.get("node_ref") or {}).items()}

    # ★칸 차수 — 접힌 형제들까지 묶어 «밖으로 나가는» G 간선을 센다.
    #   칸 안끼리 이은 간선은 밖에서 보면 한 점이므로 세지 않는다.
    pts = limited.get("pts") or []
    # ★묶는 대상은 **G_K 의 절점 전부**다. `node_ref` 의 값만 묶으면 한 칸에
    #   하나씩이라 무리가 언제나 한 점이 되고, 접힘을 영영 못 본다
    #   (한 번 그렇게 「접힌 노드 0개」를 읽었다).
    used_all = set()
    for a, b in edges_G:
        used_all.add(a)
        used_all.add(b)
    cell_of, of_cell = fold_clusters(pts, used_all, built.get("origin_mm"))
    deg_cell = {}
    for c, vids in of_cell.items():
        n = 0
        for a, b in edges_G:
            ina, inb = a in vids, b in vids
            if ina != inb:
                n += 1
        deg_cell[c] = n
    eref = {str(k) for k, v in (built.get("edge_ref") or {}).items()
            if v is not None}

    deg_kfp = defaultdict(int)
    n_vert = defaultdict(int)
    for pid, p in _pipes(kfp).items():
        s, e = _ends(p)
        if s is None or e is None:
            continue
        deg_kfp[str(s)] += 1
        deg_kfp[str(e)] += 1
        if str(pid) not in eref:          # 역참조 없음 = 전개가 만든 세로 토막
            n_vert[str(s)] += 1
            n_vert[str(e)] += 1

    phys, extra, no_ref, folded_gain = {}, [], [], []
    for nid in _nodes(kfp):
        nid = str(nid)
        vid = node_ref.get(nid)
        if vid is None:
            no_ref.append(nid)
            continue                      # 전개가 만든 노드 — 종전 규칙에 맡긴다
        # 칸 차수를 먼저 쓴다 — 접힌 형제가 없으면 둘은 같은 값이다.
        c = cell_of.get(vid)
        dg = deg_cell.get(c, deg_G.get(vid, 0)) if c is not None             else deg_G.get(vid, 0)
        if dg != deg_G.get(vid, 0):
            folded_gain.append((nid, deg_G.get(vid, 0), dg))
        base = dg + n_vert.get(nid, 0)
        phys[nid] = max(base, deg_kfp.get(nid, 0))
        if deg_kfp.get(nid, 0) > base:
            extra.append(nid)
    return (phys, dict(deg_kfp), dict(n_vert), dict(deg_G), extra, no_ref,
            folded_gain)


def flow_parents(kfp):
    """흐름 부모 — 표가 쓰는 그 함수(`bfs_order`)를 그대로 부른다.

    ★프로브가 제 손으로 BFS 를 짜면 표와 다른 나무를 볼 수 있다. 판정의 «자» 는
      하나여야 한다.
    """
    from services.cad_import.design.anchor import require_anchor
    from services.cad_import.design.tables import bfs_order
    meta = _nodes(kfp)
    root = require_anchor(meta, what="계측")
    _order, parent, _tp, _off = bfs_order(kfp, root)
    return parent, root


def lost_junctions(limited, built, kfp):
    """갈래가 지워진 분기점 — phys ≥ 3 인데 kfp 차수가 2 인 자리.

    그리고 그 자리가 회랑에서 «직진» 인지 «꺾임» 인지까지 가른다 —
    직진이면 직류티, 꺾임이면 분류티가 붙어야 한다(그림 16 규칙표).
    """
    import fitting_rules as fr

    phys, deg_kfp, n_vert, deg_G, extra, no_ref, folded = physical_degree(
        limited, built, kfp)
    parent, _root = flow_parents(kfp)
    meta = _nodes(kfp)

    def xy(nid):
        c = (meta.get(str(nid)) or {}).get("coords") or (0.0, 0.0, 0.0)
        return (float(c[0]), float(c[1]))

    inc = defaultdict(list)
    for pid, p in _pipes(kfp).items():
        s, e = _ends(p)
        if s is None or e is None:
            continue
        inc[str(s)].append((str(pid), str(e)))
        inc[str(e)].append((str(pid), str(s)))

    rows, boundary = [], []
    for nid, p in phys.items():
        if p < 3 or deg_kfp.get(nid, 0) != 2:
            continue
        up = parent.get(nid)
        downs = [(pid, o) for pid, o in inc.get(nid, ()) if o != up]
        if up is None or len(downs) != 1:
            rows.append({"nid": nid, "phys": p, "turn": None, "angle": None})
            continue
        here, up_xy, dn_xy = xy(nid), xy(up), xy(downs[0][1])
        ang = fr.deflection_deg(fr.bearing_deg(here, dn_xy),
                                fr.bearing_deg(up_xy, here))
        straight = ang <= fr.TRUNK_TURN_TOL_DEG
        rows.append({"nid": nid, "phys": p, "turn": not straight,
                     "angle": round(ang, 1), "pipe": downs[0][0]})
        if fr.ELBOW_STRAIGHT_MAX_DEG < ang <= fr.ELBOW_45_MAX_DEG:
            boundary.append({"nid": nid, "angle": round(ang, 1)})
    return {"rows": rows, "boundary": boundary, "extra": extra,
            "no_ref": no_ref, "phys": phys, "deg_kfp": deg_kfp,
            "deg_G": deg_G, "parent": parent, "folded": folded}


def vanished_board_junctions(limited, built):
    """G 의 분기점(deg_G ≥ 3) 중 **어느 kfp 노드도 가리키지 않는** 것.

    노드정리(일직선 중간 노드 병합)가 지운 자리다 — 지시서 D6 이 여기에
    `tee-run` 을 1건씩 달라고 한 자리. cover 가 없는 §2 단계에서는 **상한**이다
    (회랑 밖의 분기점까지 세므로) — 그 사실을 보고에 적는다.
    """
    deg_G = defaultdict(int)
    for e in (limited.get("edges") or ()):
        deg_G[int(e[0])] += 1
        deg_G[int(e[1])] += 1
    seen = {int(v) for v in (built.get("node_ref") or {}).values()}
    return sorted(v for v, d in deg_G.items() if d >= 3 and v not in seen)


# ─────────────────────────────────────────────── [부속]
def current_fittings(tbl):
    """지금 부속표가 내고 있는 것 — 종류별 개수와 **판정 불가 건수**.

    ★`tbl.unresolved` 는 «목록» 이다(`kind_items`·`length_items`) — 숫자가 아니다.
      숫자를 얻으려고 `un.get("kind")` 를 부르면 조용히 0 이 나온다(한 번 그렇게
      0 으로 읽었다). 목록의 `n` 을 더한다.
    """
    c = defaultdict(int)
    for r in (getattr(tbl, "fittings", None) or ()):
        c[str(r.get("type"))] += int(r.get("count") or 0)
    un = (getattr(tbl, "unresolved", None) or {})
    n_kind = sum(int((it or {}).get("n") or 1)
                 for it in (un.get("kind_items") or ()))
    n_len = len(un.get("length_items") or ())
    return dict(c), {"kind": n_kind, "length": n_len}


def dry_run_new_rule(limited, built, kfp, bores, node_xy, node_z):
    """새 규칙을 **판정만** 돌려 본다 — 파일도 제품 코드도 안 건드린다.

    지금 `build_fittings` 와 **같은 자**(같은 각·같은 임계·같은 부모)를 쓰되,
    분기 블록의 문턱을 `len(links) >= 3` 이 아니라 `phys >= 3` 으로 두고,
    직진 갈래를 버리지 않고 `tee-run` 으로 센다(그림 16 규칙표).

    반환: {종류: 개수} + 판정 불가 수. 등가길이는 여기서 안 센다 —
    그것은 `resolve_eq_len` 한 곳의 일이라 dry-run 이 흉내내면 두 벌이 된다.
    """
    import fitting_rules as fr
    from services.cad_import.design.fitting import (_deflect3_deg, _dir3,
                                                    _is_vertical,
                                                    _STRAIGHT_EPS_DEG)

    phys, deg_kfp, _nv, _dg, _ex, _nr, _fg = physical_degree(
        limited, built, kfp)
    parent, _root = flow_parents(kfp)

    inc = defaultdict(list)
    for pid, p in _pipes(kfp).items():
        s, e = _ends(p)
        if s is None or e is None:
            continue
        inc[str(s)].append((str(pid), str(e)))
        inc[str(e)].append((str(pid), str(s)))

    def at(nid):
        q = node_xy.get(str(nid))
        if q is None:
            return None
        return (float(q[0]), float(q[1]), float(node_z.get(str(nid), 0.0)))

    counts = defaultdict(int)
    unresolved = 0
    for nid, links in inc.items():
        here = node_xy.get(nid)
        if here is None:
            continue
        up = parent.get(nid)
        up_xy = node_xy.get(up) if up else None
        here3, up3 = at(nid), (at(up) if up else None)
        p = phys.get(nid, len(links))

        if p == 2:
            # 종전 관통 블록과 같은 자 — 3차원 각으로 엘보를 가린다.
            if len(links) != 2 or up3 is None or here3 is None:
                continue
            downs = [(pid, o) for pid, o in links if o != up]
            if len(downs) != 1:
                continue
            o3 = at(downs[0][1])
            if o3 is None:
                continue
            u, v = _dir3(up3, here3), _dir3(here3, o3)
            if u is None or v is None:
                continue
            ang = _deflect3_deg(u, v)
            if ang <= _STRAIGHT_EPS_DEG:
                continue
            kinds, bad = fr.elbow_fittings([ang])
            unresolved += bad
            for kk in kinds:
                counts[kk] += 1
            continue

        if p < 3:
            continue

        # 분기 — 세로 하류는 종전대로 티, 평면 하류는 직진/꺾임으로 가른다.
        downs = [(pid, o) for pid, o in links if o != up]
        vert, flat = [], []
        for pid, o in downs:
            o3 = at(o)
            if here3 is not None and o3 is not None and _is_vertical(here3, o3):
                vert.append(pid)
            elif node_xy.get(o) is not None:
                flat.append((pid, node_xy.get(o)))
        big = p >= 4
        for _pid in vert:
            counts["cross" if big else "tee"] += 1
        if not flat:
            continue
        if up_xy is None:
            unresolved += len(flat)
            continue
        inflow = fr.bearing_deg(up_xy, here)
        turning, straight = [], []
        for pid, pos in flat:
            if fr.deflection_deg(fr.bearing_deg(here, pos),
                                 inflow) <= fr.TRUNK_TURN_TOL_DEG:
                straight.append(pid)
            else:
                turning.append(pid)
        if len(straight) > 1:
            # 종전 규칙 그대로 — spine 이 갈라져 가릴 수 없다.
            unresolved += len(straight)
            for _pid in flat:
                counts["cross" if big else "tee"] += 1
            continue
        for _pid in turning:
            counts["cross" if big else "tee"] += 1
        for _pid in straight:
            counts["cross-run" if big else "tee-run"] += 1
    return dict(counts), unresolved


# ─────────────────────────────────────────────── [하향식]
def pendent_arm_change(flat_kfp, plan, node_head_kinds):
    """스텁을 끄면 «팔 없는 하향식» 이 «팔 있는 하향식» 이 되나 (지시서 §2·§6-1).

    ★**세로 처리 전** 평면 망(`built["kfp"]`)에서 잰다. 세로 처리가 끝난 뒤에는
      헤드가 스텁 끝으로 옮겨 가 차수가 1 이 되므로, 거기서 재면 인라인 헤드가
      **언제나 0** 으로 나온다(한 번 그렇게 0 을 읽었다).

    끝배관 1단 보호는 «인라인 헤드»(평면 차수 ≥ 2)에만 걸린다. 보호를 끄면 그
    헤드는 관말이 되고, 관말 하향식은 팔(①)이 선다 — 그림 16 ③ 의 「스텁은
    지우고 H1 을 관말로」가 그 자리다.
    """
    rows = []
    for h in _head_rows(flat_kfp, plan, node_head_kinds):
        if str(h.get("kind") or "") == "상향식":
            continue
        if int(h.get("deg") or 0) >= 2:
            rows.append(h)
    return rows


def _head_rows(kfp, plan, node_head_kinds):
    nd, pr = _nodes(kfp), _pipes(kfp)
    deg = defaultdict(int)
    for p in pr.values():
        s, e = _ends(p)
        deg[str(s)] += 1
        deg[str(e)] += 1
    out = []
    for nid, m in nd.items():
        if str((m or {}).get("type_id") or "").lower() != "head":
            continue
        c = (m or {}).get("coords") or (0, 0, 0)
        out.append({"nid": str(nid), "xy": plan.mm(c),
                    "z": float(c[2]) if len(c) > 2 else 0.0,
                    "kind": (node_head_kinds or {}).get(str(nid)),
                    "deg": deg.get(str(nid), 0)})
    return out


# ─────────────────────────────────────────────── [kfp]
def worst_kfp_vs_table(payload, board, worst, raised, tbl):
    """**옛 경로**(제한 전개를 한 번 더)와 표의 망이 얼마나 다른가 (§2 [kfp] · D5).

    ★D5 를 넣은 뒤로 최불리 `.kfp` 는 **표에서** 난다(`emit_design_kfp`).
      이 줄은 그 결정의 «근거» 를 계속 보여 주는 자리다 — 옛 경로가 얼마나
      다른 망을 냈는지(배관 수·관경·부속 0건). 제품이 쓰는 길이 아니다.

    반환 dict. 못 돌면 `{"ok": False, "why": …}` — 지어내지 않는다.
    """
    try:
        from services.cad_import.convert.engine import (convert_to_kfp,
                                                        ensure_planar)
        from services.cad_import.dto import default_dto, dto_to_convert_kwargs
        from routes.module_f.remote30 import _restrict_to_worst
    except Exception as exc:                        # noqa: BLE001
        return {"ok": False, "why": f"{type(exc).__name__}: {exc}"}

    import tempfile
    try:
        pl = _restrict_to_worst(dict(payload), board, worst)
        pl = ensure_planar(pl)
        if pl.get("kfp") is None and not pl.get("kfp_path"):
            return {"ok": False, "why": pl.get("_planar_error") or "planar 실패"}
        with tempfile.TemporaryDirectory(prefix="pp_kfp_") as d:
            res = convert_to_kfp(pl, str(Path(d) / "worst.kfp"),
                                 **dto_to_convert_kwargs(default_dto()))
        if not res.get("ok"):
            return {"ok": False, "why": "convert_to_kfp 막힘"}
        kfp = res["kfp"]
    except Exception as exc:                        # noqa: BLE001
        return {"ok": False, "why": f"{type(exc).__name__}: {exc}"}

    a = set(map(str, _pipes(kfp)))
    b = set(map(str, _pipes(raised)))
    dn_kfp = defaultdict(int)
    for p in _pipes(kfp).values():
        dn_kfp[p.get("nominal_mm") or p.get("pipe_dn") or "?"] += 1
    dn_tbl = defaultdict(int)
    for r in (getattr(tbl, "pipes", None) or ()):
        dn_tbl[int(r.get("dia") or 0)] += 1
    n_fit = sum(len((p or {}).get("fittings") or ())
                for p in _pipes(kfp).values())
    n_eq = sum(1 for p in _pipes(kfp).values()
               if float((p or {}).get("equivalent_length") or 0) > 0)
    return {"ok": True, "same": a == b, "only_kfp": len(a - b),
            "only_tbl": len(b - a), "n_kfp": len(a), "n_tbl": len(b),
            "dn_kfp": dict(sorted(dn_kfp.items(), key=lambda t: str(t[0]))),
            "dn_tbl": dict(sorted(dn_tbl.items())),
            "fittings": n_fit, "eq_len_nonzero": n_eq}


# ─────────────────────────────────────────────── 보고
def report(*, key, k, limited, built, raised, tbl, plan, stub_log, flat=None,
           payload=None, board=None, worst=None):
    """§2 여섯 줄을 그대로 찍는다. 돌려주는 것은 수치 dict(시험이 쓴다)."""
    kfp = raised
    nd = _nodes(kfp)
    node_xy = {}
    node_z = {}
    for nid, m in nd.items():
        c = (m or {}).get("coords") or (0.0, 0.0, 0.0)
        node_xy[str(nid)] = (float(c[0]), float(c[1]))
        node_z[str(nid)] = float(c[2]) if len(c) > 2 else 0.0

    res_nodes, res_len = residual_stubs(tbl)
    lost = lost_junctions(limited, built, kfp)
    vanished = vanished_board_junctions(limited, built)
    cur, un = current_fittings(tbl)
    new, new_un = dry_run_new_rule(limited, built, kfp, None, node_xy, node_z)
    pend = pendent_arm_change(flat if flat is not None else built.get("kfp"),
                              plan, built.get("node_head_kinds") or {})

    rows = lost["rows"]
    turn = [r for r in rows if r.get("turn") is True]
    stra = [r for r in rows if r.get("turn") is False]
    unknown = [r for r in rows if r.get("turn") is None]
    big = [r for r in rows if int(r.get("phys") or 0) >= 4]
    angles = defaultdict(int)
    for r in rows:
        a = r.get("angle")
        if a is None:
            angles["모름"] += 1
        elif a <= 0.5:
            angles["0"] += 1
        elif abs(a - 45.0) <= 5.0:
            angles["45"] += 1
        elif abs(a - 90.0) <= 5.0:
            angles["90"] += 1
        else:
            angles["그 외"] += 1

    print(f"\n  [잔류]   끝배관1단보호 {stub_log if stub_log is not None else '?'}개"
          f" (planar 로그) · 표에 남는 차수 1 비헤드·비급수원 노드"
          f" {len(res_nodes)}개 · 그 배관 길이 합 {res_len}m")
    print(f"  [차수]   phys ≥ 3 인데 kfp 차수 2 = **갈래가 지워진 분기점**"
          f" {len(rows)}개")
    print(f"           그중 직진(≤45°) {len(stra)} · 꺾임 {len(turn)}"
          f" · 상류 모름 {len(unknown)} · phys ≥ 4 {len(big)}")
    print(f"           편향각 분포 {dict(angles)}")
    # ★§6-2 — 「전개가 만든 것이 아닌데 역참조가 없는 **평면** 노드」.
    #   역참조 없는 노드 대부분은 세로 토막 끝(①②③④)이라 정상이다. 가로
    #   배관이 붙어 있는데도 역참조가 없으면 그것이 물어야 할 자리다.
    eref = {str(kk) for kk, vv in (built.get("edge_ref") or {}).items()
            if vv is not None}
    flat_at = defaultdict(int)
    for pid, pp_ in _pipes(kfp).items():
        s_, e_ = _ends(pp_)
        if str(pid) in eref and s_ and e_:
            flat_at[str(s_)] += 1
            flat_at[str(e_)] += 1
    no_ref_flat = [n for n in lost["no_ref"] if flat_at.get(n, 0) > 0]
    print(f"           deg_kfp > deg_G + n_vert (planar 가 만든 접속)"
          f" {len(lost['extra'])}개 · 역참조 없는 노드 {len(lost['no_ref'])}개"
          f" (그중 가로 배관이 붙은 것 {len(no_ref_flat)}개 ← §6-2)")
    print(f"           격자로 접혀 차수가 는 노드 {len(lost['folded'])}개"
          f" (node_ref 한 점만 세면 과소평가되는 자리)")
    print(f"           노드정리로 사라진 board 분기점 {len(vanished)}개"
          f" (cover 없이 센 **상한** — 회랑 밖도 포함)")
    print(f"  [부속]   지금  {dict(sorted(cur.items()))}"
          f" · 판정 불가 {un['kind']} · 등가길이 미해결 {un['length']}")
    print(f"           새규칙 {dict(sorted(new.items()))}"
          f" · 판정 불가 {new_un}")
    print(f"  [경계]   분기점 갈래 중 편향 22.5~67.5° {len(lost['boundary'])}개"
          + (f" — {[b['nid'] for b in lost['boundary'][:6]]}"
             if lost["boundary"] else ""))
    print(f"  [하향식] 인라인 하향식 헤드(스텁을 끄면 관말이 된다)"
          f" {len(pend)}개" + (f" — {[h['nid'] for h in pend[:6]]}"
                              if pend else ""))
    wk = None
    if payload is not None and board is not None and worst is not None:
        wk = worst_kfp_vs_table(payload, board, worst, kfp, tbl)
        if not wk.get("ok"):
            print(f"  [kfp]    최불리 .kfp 를 못 냈습니다 — {wk.get('why')}")
        else:
            print(f"  [kfp]    (옛 경로) 두 번째 전개 .kfp 배관 {wk['n_kfp']}"
                  f" vs 표 {wk['n_tbl']}"
                  f" · 같은 집합 {wk['same']}"
                  f" (.kfp 에만 {wk['only_kfp']} · 표에만 {wk['only_tbl']})")
            print(f"           관경 .kfp {wk['dn_kfp']}")
            print(f"           관경 표   {wk['dn_tbl']}")
            print(f"           .kfp 의 부속 {wk['fittings']}건"
                  f" · 등가길이 > 0 인 배관 {wk['eq_len_nonzero']}개")
    return {
        "stub_log": stub_log, "residual_nodes": res_nodes,
        "residual_len_m": res_len, "lost": rows, "straight": len(stra),
        "turn": len(turn), "unknown": len(unknown), "phys4": len(big),
        "extra": lost["extra"], "no_ref": lost["no_ref"],
        "vanished": vanished, "cur": cur, "new": new,
        "new_unresolved": new_un, "boundary": lost["boundary"],
        "pendent_inline": pend, "worst_kfp": wk,
        "no_ref_flat": no_ref_flat,
    }
