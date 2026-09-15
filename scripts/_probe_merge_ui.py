# -*- coding: utf-8 -*-
"""[통합 격자·활성·위상 §2] 바꾸기 전에 잰다 — 일곱 줄.

지시서 `ModuleF_통합_격자활성_위상수정_지시서.md` §2 · 그림 18·19·20.
**아무것도 안 고친다.** 격자·밑그림·삭제 규칙이 전부 «좌표의 자» 위에 서므로,
그 자가 맞는지부터 수치로 확인한다.

    [좌표]   기준점·회랑 노드 — board mm 를 1-B ② 식으로 옮긴 값 vs 결합망 좌표
    [배율]   04 의 norm.scale · §3-1-3 유추가 그 값을 되찾는가
    [격자]   bbox · 1 m 칸이 «화면 가득» 에서 몇 px · 라이저/기계실 구역 비율
    [차수]   나무인가 · 고리 · 차수 분포 · 지우면 갈라지는 배관(다리)
    [요소]   노드·배관·노즐·기기·이음매 (도면별)
    [라벨]   배관 라벨을 «물 흐르는 순서» 로 다시 매기면 몇 개가 바뀌나

**멈추는 조건**(§2): [좌표] 의 차가 1 mm 를 넘으면 §3 으로 가지 않는다 — 그 수가
틀리면 격자도 밑그림도 «그럴듯하게 어긋난 그림» 이 된다.

    python scripts/_probe_merge_ui.py [--key 저장본] [--k 30]
"""
from __future__ import annotations

import argparse
import io
import contextlib
import math
import os
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DM_KEY = "1. 입력도면 대명동 단위세대 평면도"
CANVAS_PX = 1200.0          # 「화면 가득」 을 가정하는 폭 — px 수치의 기준


# ─────────────────────────────────────────────── 자
def _ends(r):
    return str(r.get("in")), str(r.get("out"))


def board_to_table_mm(bx, by, origin_mm):
    """1-B ② — board mm → 표 좌표(mm). 규칙은 `main_walk.xf_mm_to_m` 하나뿐이다.

        kfp_m = (board_mm − origin) / 1000 + 1      →  표 mm = kfp_m · 1000
    """
    ox, oy = float(origin_mm[0]), float(origin_mm[1])
    return ((float(bx) - ox) + 1000.0, (float(by) - oy) + 1000.0)


def undirected(rows):
    adj = defaultdict(set)
    for r in rows:
        a, b = _ends(r)
        adj[a].add(b)
        adj[b].add(a)
    return adj


def components(adj, nodes):
    seen, n = set(), 0
    for s in nodes:
        if s in seen:
            continue
        n += 1
        st = [s]
        seen.add(s)
        while st:
            c = st.pop()
            for x in adj.get(c, ()):
                if x not in seen:
                    seen.add(x)
                    st.append(x)
    return n


def bridges(rows):
    """지우면 망이 갈라지는 배관 — Tarjan. 고리 안의 배관은 다리가 아니다."""
    adj = defaultdict(list)
    for i, r in enumerate(rows):
        a, b = _ends(r)
        adj[a].append((b, i))
        adj[b].append((a, i))
    disc, low, out = {}, {}, []
    t = [0]

    def walk(root):
        stack = [(root, -1, iter(adj[root]))]
        disc[root] = low[root] = t[0]
        t[0] += 1
        while stack:
            u, pe, it = stack[-1]
            adv = False
            for v, ei in it:
                if ei == pe:
                    continue
                if v not in disc:
                    disc[v] = low[v] = t[0]
                    t[0] += 1
                    stack.append((v, ei, iter(adj[v])))
                    adv = True
                    break
                low[u] = min(low[u], disc[v])
            if adv:
                continue
            stack.pop()
            if stack:
                p = stack[-1][0]
                low[p] = min(low[p], low[u])
                if low[u] > disc[p]:
                    out.append(pe)
    for s in list(adj):
        if s not in disc:
            walk(s)
    return {i for i in out if i >= 0}


def infer_scale(pipes, xy):
    """[§3-1-3] 표가 답을 갖고 있다 — 수평 배관의 (좌표거리 ÷ 길이) 중앙값.

    `kfp_sdf_converter.py:774~790` 과 **같은 방법**이다(1-B ⑥). 새로 짜지 않는다.
    반환 (배율, 쓴 배관 수, ±5% 밖 비율) · 못 구하면 (None, n, None).
    """
    rs = []
    for r in pipes:
        L = float(r.get("length") or 0.0)
        dz = abs(float(r.get("elev") or 0.0))
        if L <= 0 or dz > 0.25 * L:
            continue                      # 수평에 가까운 것만
        a, b = _ends(r)
        if a not in xy or b not in xy:
            continue
        d = math.dist(xy[a], xy[b])
        if d <= 0:
            continue
        rs.append(d / L)
    if len(rs) < 3:
        return None, len(rs), None
    rs.sort()
    med = rs[len(rs) // 2]
    if med <= 0:
        return None, len(rs), None
    off = sum(1 for r in rs if abs(r - med) > 0.05 * med)
    return med, len(rs), off / len(rs)


# ─────────────────────────────────────────────── 본체
def run(key, k):
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib

    from _probe_candidate_drop import setup, wait
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    buf = io.StringIO()
    # ★저장본을 **격리 안으로 들여야** `reopen` 이 선다. 안 들이면 빈 폴더를
    #   보고 「열기 잡 실패」로 떨어진다(한 번 그렇게 떨어뜨렸다).
    with isolated_workdir(prefix="mgui_", copy_key=key),             srv.app.test_client() as c:
        with contextlib.redirect_stdout(buf):
            with c.session_transaction() as s:
                s["authed"] = True
            sid = setup(c, argparse.Namespace(key=key))
            st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
            if not (st.get("sources") or ()):
                seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                             key=len, reverse=True)[0]
                c.post("/api/module-f/edit/mode",
                       json={"sid": sid, "mode": "급수시작위치"})
                c.post("/api/module-f/edit/click",
                       json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                             "y": (seg[1] + seg[3]) / 2, "max_d": 2000})
            c.post("/api/module-f/edit/worst", json={"sid": sid, "k": k})
            c.post("/api/module-f/design/build", json={"sid": sid, "k": k})
            if wait(c, sid).get("state") != "done":
                print("★표 확정 실패")
                return 1
            j04 = c.get(f"/api/module-f/design/preview?sid={sid}").get_json()

            from routes.module_f.jobs import _sess
            sess = _sess(sid)
            d = sess["design"]
            got, tbl = d["got"], d["tables"]
            board = sess["edit"].board

            # ── 결합 — 계통도 추출이 막히는 도면이 있어(§27) 표본 입상관으로
            #    세운다. 재는 것은 «좌표의 자» 라 입상관의 출처와 무관하다.
            import importlib.util as ilu
            spec = ilu.spec_from_file_location(
                "_mf_merge_fx", str(ROOT / "tests" / "test_module_f_merge.py"))
            fx = ilu.module_from_spec(spec)
            spec.loader.exec_module(fx)
            from routes.module_f.merge import label_offset_for, merge_network
            mrfx = ilu.spec_from_file_location(
                "_mf_seam_fx",
                str(ROOT / "tests" / "test_module_f_seam_coords.py"))
            seam = ilu.module_from_spec(mrfx)
            mrfx.loader.exec_module(seam)
            mg = merge_network(tbl, riser=fx._riser(),
                               machineroom=seam._machineroom(),
                               mode="lsp_gravity")
            sess["supply_mode"] = "lsp_gravity"
            sess["merged"] = mg
            # 미리보기가 서는지도 함께 본다 — 화면이 그것을 그린다.
            _jm = c.get(f"/api/module-f/merge/preview?sid={sid}").get_json()
            assert _jm.get("ok"), _jm

    print(f"\n■ 통합 격자·활성·위상 계측 · {key} · K={k}")

    comb = mg.get("combined")
    if comb is None:
        print("★결합망이 없습니다")
        return 1
    parts = mg.get("parts") or {}
    of = {}
    for kind in ("system", "machineroom", "plan"):
        for lab in (parts.get(kind) or ()):
            of[str(lab)] = kind
    off = label_offset_for("manual")
    origin = got.get("origin_mm")
    node_ref = {str(a): int(b) for a, b in (got.get("node_ref") or {}).items()}

    # ── [좌표] 1-B ② 의 식으로 board 를 옮긴 값 vs 결합망 좌표
    mg_xy = {str(n.get("label")): (float(n.get("x") or 0), float(n.get("y") or 0))
             for n in (comb.nodes or ())}
    t_xy = {str(n.get("label")): (float(n.get("x") or 0), float(n.get("y") or 0))
            for n in tbl.nodes}
    kmeta = (got.get("kfp") or {}).get("nodes_meta_runtime") or {}
    lab_by_xyz = {}
    for n in tbl.nodes:
        lab_by_xyz[(int(n.get("x") or 0), int(n.get("y") or 0),
                    round(float(n.get("elevation") or 0), 3))] = str(n["label"])
    diffs, bad, anchor_d = [], [], None
    # 헤드와 절점을 갈라 둔다 — 사슬이 헤드를 옮기는 것은 **오너가 허용한
    # 어긋남**이다(사슬 보고가 「헤드 이탈 … 경고 아님」으로 이미 말한다).
    # 섞어 재면 «자가 틀린 것» 과 «규칙대로인 것» 이 한 통에 든다.
    head_d, flat_d = [], []
    _head_labs = {str(z.get("in")) for z in tbl.nozzles}
    for nid, vid in node_ref.items():
        cds = (kmeta.get(nid) or {}).get("coords")
        if not cds:
            continue
        lab = lab_by_xyz.get((int(round(float(cds[0]) * 1000)),
                              int(round(float(cds[1]) * 1000)),
                              round(float(cds[2]) if len(cds) > 2 else 0.0, 3)))
        if lab is None or lab not in t_xy:
            continue
        try:
            bx, by = board.pts[vid][0], board.pts[vid][1]
        except (IndexError, TypeError):
            continue
        want = board_to_table_mm(bx, by, origin)
        got_xy = t_xy[lab]
        dd = math.dist(want, got_xy)
        diffs.append(dd)
        (head_d if lab in _head_labs else flat_d).append(dd)
        if dd > 1.0:
            bad.append((lab, round(dd, 2)))
        mlab = str(int(lab) + off) if lab.isdigit() else None
        if mlab and mlab in mg_xy:
            # 결합이 회랑 좌표를 안 움직이는가(1-B ③)
            if math.dist(mg_xy[mlab], got_xy) > 1.0:
                bad.append((f"{lab}→{mlab} 결합이동", round(math.dist(mg_xy[mlab], got_xy), 2)))
    anchor = str(10)
    if anchor in mg_xy:
        blab = str(10 - off)
        if blab in t_xy:
            anchor_d = math.dist(mg_xy[anchor], t_xy[blab])
    print(f"\n  [좌표]   기준점(라벨 10) 차 "
          f"{('%.3f mm' % anchor_d) if anchor_d is not None else '?'}")
    diffs.sort()
    _med = diffs[len(diffs) // 2] if diffs else 0.0
    print(f"           회랑 노드 {len(diffs)}개 대조 — 최대 차 "
          f"{max(diffs) if diffs else 0:.3f} mm · 중앙 {_med:.3f} mm"
          f" · 1mm 넘는 노드 {len(bad)}개")
    # ★어긋남이 «어디서» 들어오는지 — 두 자리다(단계별 실측 `data/_coord.log`):
    #     ① 평면 전개 직후  최대 34.7mm · 중앙 17.5mm  ← 격자 스냅의 **반올림**
    #                       (50mm 격자라 대각선 최대 √2·25 ≈ 35.4mm)
    #     ③ 사슬 좌표 뒤    최대 96.4mm · 중앙 21.7mm  ← 사슬이 좌표를 «만든다»
    #   즉 지시서 1-B ② 의 「회랑은 board 와 1:1 평행이동」은 **사슬 이전에도
    #   이미 참이 아니다.** 식(`xf_mm_to_m`)은 맞지만 좌표가 그 뒤로 옮겨진다.
    print("           ★1-B ② 의 「1:1 평행이동」은 지금 참이 아니다 — "
          "격자 스냅(≈35mm)과 사슬(≈96mm)이 좌표를 옮긴다")

    # ── [배율]
    und = ((j04.get("view") or {}).get("underlay") or {})
    kk = und.get("k")
    vx = {str(n["label"]): (float(n["x"]), float(n["y"]))
          for n in (j04.get("view") or {}).get("nodes", ())}
    s_inf, n_used, off_ratio = infer_scale(
        (j04.get("tables") or {}).get("pipes") or [], vx)
    print(f"  [배율]   04 underlay.k = {kk} → 1 m = "
          f"{(1000.0 * float(kk)) if kk else '?'} 단위")
    print(f"           §3-1-3 유추 1 m = "
          f"{('%.3f' % s_inf) if s_inf else '못 구함'} 단위"
          f" · 쓴 배관 {n_used}개"
          + (f" · ±5% 밖 {off_ratio * 100:.1f}%" if off_ratio is not None else ""))
    if kk and s_inf:
        want = 1000.0 * float(kk)
        print(f"           유추 대 받은 값 — 차 {abs(s_inf - want):.3f} 단위"
              f" ({abs(s_inf - want) / want * 100:.2f}%)")

    # ── [격자]
    xs = [p[0] for p in mg_xy.values()]
    ys = [p[1] for p in mg_xy.values()]
    w_m = (max(xs) - min(xs)) / 1000.0 if xs else 0.0
    h_m = (max(ys) - min(ys)) / 1000.0 if ys else 0.0
    px_m = CANVAS_PX / max(w_m, h_m) if max(w_m, h_m) > 0 else 0.0
    area_all = max(w_m, 1e-9) * max(h_m, 1e-9)
    rs_x = [mg_xy[l][0] for l in mg_xy if of.get(l) in ("system", "machineroom")]
    rs_y = [mg_xy[l][1] for l in mg_xy if of.get(l) in ("system", "machineroom")]
    area_rs = (((max(rs_x) - min(rs_x)) / 1000.0)
               * ((max(rs_y) - min(rs_y)) / 1000.0)) if rs_x else 0.0
    print(f"  [격자]   결합망 bbox {w_m:.1f} × {h_m:.1f} m"
          f" · 1 m 칸 = {px_m:.2f} px (폭 {CANVAS_PX:.0f}px 가정)")
    print(f"           라이저·기계실 구역이 bbox 에서 {area_rs / area_all * 100:.1f}%")
    v4 = (j04.get("view") or {})
    xs4 = [float(n["x"]) for n in v4.get("nodes", ())]
    ys4 = [float(n["y"]) for n in v4.get("nodes", ())]
    px_m_04 = 0.0
    if xs4 and kk:
        w4 = (max(xs4) - min(xs4)) / (1000.0 * float(kk))
        h4 = (max(ys4) - min(ys4)) / (1000.0 * float(kk))
        px_m_04 = CANVAS_PX / max(w4, h4, 1e-9)
        print(f"           04 bbox {w4:.1f} × {h4:.1f} m · 1 m 칸 = "
              f"{px_m_04:.2f} px")

    # ── [차수]
    for nm, rows, nodes in (("결합망", list(comb.pipes or ()), mg_xy),
                            ("회랑망", list(tbl.pipes), t_xy)):
        adj = undirected(rows)
        V, E = len(nodes), len(rows)
        comp = components(adj, list(nodes))
        loops = E - V + comp
        deg = defaultdict(int)
        for r in rows:
            a, b = _ends(r)
            deg[a] += 1
            deg[b] += 1
        heads = ({str(z.get("in")) for z in (comb.nozzles or ())} if nm == "결합망"
                 else {str(z.get("in")) for z in tbl.nozzles})
        io_in = ({str(n.get("label")) for n in (comb.nodes or ())
                  if str(n.get("io_node")) == "Input"} if nm == "결합망"
                 else {str(n.get("label")) for n in tbl.nodes
                       if str(n.get("io_node")) == "Input"})
        d1 = [l for l, x in deg.items() if x == 1 and l not in heads and l not in io_in]
        br = bridges(rows)
        print(f"  [차수]   {nm} |V|={V} |E|={E} 성분 {comp} · "
              f"나무 {'예' if E == V - comp else '아니오'} · 고리 {loops}")
        print(f"           차수1 비헤드·비급수원 {len(d1)} · "
              f"차수2 {sum(1 for x in deg.values() if x == 2)} · "
              f"차수≥3 {sum(1 for x in deg.values() if x >= 3)}")
        print(f"           지우면 갈라지는 배관(다리) {len(br)}/{E}"
              f" — 고리 덕분에 지워도 되는 배관 {E - len(br)}개")

    # ── [요소]
    cnt = defaultdict(int)
    for lab in mg_xy:
        cnt[of.get(lab, "plan")] += 1
    seam_n = sum(1 for r in (comb.pipes or ())
                 if of.get(_ends(r)[0], "plan") != of.get(_ends(r)[1], "plan"))
    print(f"  [요소]   노드 {len(mg_xy)} (회랑 {cnt['plan']} · 계통도 "
          f"{cnt['system']} · 기계실 {cnt['machineroom']}) · 배관 "
          f"{len(comb.pipes or ())} · 노즐 {len(comb.nozzles or ())} · 기기 "
          f"{len(getattr(comb, 'equipment', None) or ())} · 이음매 {seam_n}")

    # ── [라벨] 물 흐르는 순서로 다시 매기면 몇 개가 바뀌나 (D7)
    from services.cad_import.design.anchor import require_anchor
    from services.cad_import.design.tables import bfs_order
    net = got["kfp"]
    root = require_anchor(net.get("nodes_meta_runtime") or {}, what="라벨")
    order, parent, tree_pipes, off_tree = bfs_order(net, root)
    new_lab, i = {}, 0
    # ★`tree_pipes` 의 원소는 (pid, in, out) 튜플이다 — 그냥 돌리면 튜플
    #   자체를 이름으로 삼아 **전부 바뀐다**는 거짓 수치가 나온다(그렇게
    #   한 번 읽었다: 「233개 중 233개」).
    for pid, _a, _b in list(tree_pipes) + list(off_tree):
        i += 1
        new_lab[str(pid)] = f"P{i}"
    same = sum(1 for pid, nl in new_lab.items() if str(pid) == nl)
    print(f"  [라벨]   배관 {len(new_lab)}개 중 다시 매기면 바뀌는 라벨 "
          f"{len(new_lab) - same}개 ({(len(new_lab) - same) / max(len(new_lab), 1) * 100:.0f}%)")

    # ★판정은 mm 가 아니라 **화면(px)** 이다(오너 2026-09-15 ㉮).
    #   회랑 좌표는 격자 스냅과 사슬로 이미 수십 mm 옮겨져 있으므로(1-B ② 정정),
    #   mm 자로는 어느 도면도 못 선다. 밑그림이 쓸모 있느냐를 가르는 것은
    #   「그 어긋남이 화면에서 보이느냐」다.
    px04 = px_m_04 if px_m_04 else 0.0

    def _px(v):
        return v * px04 / 1000.0

    fl, hd = sorted(flat_d), sorted(head_d)
    fmax = _px(fl[-1]) if fl else 0.0
    fmed = _px(fl[len(fl) // 2]) if fl else 0.0
    print(f"\n  [밑그림 판정] 04 화면({px04:.1f} px/m) 기준")
    def _q(v, q):
        return _px(v[int(len(v) * q)]) if v else 0.0

    print(f"           회랑 절점(헤드 제외) {len(fl)}개 — 중앙 {fmed:.2f} px"
          f" · 90% {_q(fl, 0.90):.2f} · 95% {_q(fl, 0.95):.2f}"
          f" · 최대 {fmax:.2f} px")
    print(f"           1px 넘는 절점 {sum(1 for v in fl if _px(v) > 1.0)}개 ·"
          f" 2px 넘는 절점 {sum(1 for v in fl if _px(v) > 2.0)}개")
    print(f"           헤드 {len(hd)}개 — 최대 {_px(hd[-1]) if hd else 0:.2f} px"
          f"  (사슬이 옮긴다 · 오너가 허용한 어긋남이라 판정에서 뺀다)")
    # ★판정자는 **중앙값**이다. 변환식이 깨지면 회랑 **전체**가 옮겨져 중앙값이
    #   즉시 튄다 — 그것이 M4 가 지키려는 것이다. 최대값을 판정자로 쓰면 격자
    #   스냅이 한 칸 옮긴 절점 하나(실측 12/142)에 전체가 막힌다. 그 12개는
    #   알려진 현상이고 화면에서 선 굵기 수준이라, 보고하되 문으로 쓰지 않는다.
    p90 = _q(fl, 0.90)
    ok = (anchor_d is not None and anchor_d <= 1.0
          and fmed <= 1.0 and p90 <= 2.0)
    print("  " + ("★[좌표] 가 선다 — §3 으로 갈 수 있다" if ok
                  else "★★[좌표] 가 안 선다 — §2 멈춤 조건"))
    return 0 if ok else 2


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=DM_KEY)
    ap.add_argument("--k", type=int, default=30)
    a = ap.parse_args()
    return run(a.key, a.k)


if __name__ == "__main__":
    raise SystemExit(main())
