# -*- coding: utf-8 -*-
"""[통합 격자·활성·위상 §4] 위상 수정을 **실제 도면**에 걸어 본다.

합성 망 시험(`tests/test_element_ops.py`)이 규칙을 지키는지 보는 자라면,
이것은 그 규칙이 **대명동 실측 망**에서도 서는지 보는 자다. 둘 다 필요하다 —
합성 망은 차수 2 절점도, 고리도, 헤드도 내가 만든 대로라서, 진짜 도면이 내는
모양(차수 2 가 174개 · 고리 0 · 다리 233/233)을 못 보여 준다.

재는 것:

    [삭제]   차수 2 절점 하나를 지운다 → 배관이 합쳐지고 총연장이 보존되는가
    [추가]   배관 하나를 t=0.37 로 쪼갠다 → 길이 합 · 관경 복사 · 라벨 밀기
    [기기]   밸브 하나를 단다 → 기기 행 · 등가길이 · 노드 fitting_id 불변
    [멱등]   같은 목록을 두 번 적용하면 완전히 같은 표
    [M9]     회랑 삭제 하나가 04 표와 결합망 **양쪽**에서 사라지는가
"""
from __future__ import annotations

import argparse
import contextlib
import copy
import io
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for p in (ROOT, ROOT / "cad_project_editor_g", ROOT / "core", ROOT / "scripts",
          ROOT / "tests"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

DM_KEY = "1. 입력도면 대명동 단위세대 평면도"


def _tbl_pipes(tbl):
    return {str(r["label"]): r for r in tbl.pipes}


def _total(tbl):
    return round(sum(float(r.get("length") or 0.0) for r in tbl.pipes), 3)


def run(key, k):
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib

    from _probe_candidate_drop import setup, wait
    from _workdir_iso import isolated_workdir
    from routes.module_f import overrides as ov
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    buf = io.StringIO()
    with isolated_workdir(prefix="opse2e_", copy_key=key), \
            srv.app.test_client() as c:
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

            from routes.module_f.jobs import _sess
            sess = _sess(sid)
            d = sess["design"]
            got0, tbl0 = d["got"], d["tables"]
            board = sess["edit"].board
            keys = d["keys"]

        print(f"■ 위상 수정 e2e · {key} · K={k}")
        print(f"  표: 절점 {len(tbl0.nodes)} · 배관 {len(tbl0.pipes)} · "
              f"총연장 {_total(tbl0)} m · 노즐 {len(tbl0.nozzles)}")

        # ── 어디를 고칠지 고른다 — 차수 2 절점 하나와 그 위 배관 하나
        deg = {}
        for r in tbl0.pipes:
            deg[str(r["in"])] = deg.get(str(r["in"]), 0) + 1
            deg[str(r["out"])] = deg.get(str(r["out"]), 0) + 1
        heads = {str(r["in"]) for r in tbl0.nozzles}
        node_keys = keys.get("node") or {}
        pipe_keys = keys.get("pipe") or {}
        # ★급수원을 빼야 한다 — 차수 2 여도 지울 수 없고(망의 정의), 그러면
        #   계측이 「거절되더라」만 다시 보여 준다. 헤드도 뺀다(기준개수가 준다).
        root = next((str(r["label"]) for r in tbl0.nodes
                     if str(r.get("io_node")) == "Input"), None)
        cand_node = next((lab for lab, n in deg.items()
                          if n == 2 and lab in node_keys and lab not in heads
                          and lab != root), None)
        cand_pipe = next((str(r["label"]) for r in tbl0.pipes
                          if str(r["label"]) in pipe_keys
                          and float(r.get("length") or 0) > 0.5), None)
        # 기기는 **가장 굵은** 배관 위에 단다 — 라이브러리는 가는 관에 구멍이
        # 많아서(게이트 밸브는 50A 부터 있다), 아무 데나 달면 재는 것이
        # 「미해결로 세는가」로 바뀐다. 그것은 §4 M8 이고 여기 것이 아니다.
        _with_key = [r for r in tbl0.pipes if str(r["label"]) in pipe_keys]
        eq_pipe = (str(max(_with_key,
                           key=lambda r: int(r.get("dia") or 0))["label"])
                   if _with_key else cand_pipe)
        _dias = sorted({int(r.get("dia") or 0) for r in tbl0.pipes})
        print(f"  호칭경 — {_dias}")
        print(f"  고른 자리 — 차수2 절점 {cand_node} · 배관 {cand_pipe}"
              f" · 기기 자리 {eq_pipe}")
        if not cand_node or not cand_pipe:
            print("  ★고칠 자리를 못 골랐다 — 이 도면으로는 잴 수 없다")
            return 2

        def rebuild(ops):
            """같은 입력으로 표를 다시 만든다 — 목록만 갈아 끼운다."""
            g = copy.deepcopy(got0)
            with contextlib.redirect_stdout(io.StringIO()):
                _n, missed, rep = ov.apply_ops_to_kfp(g, board, ops)
                from services.cad_import.design.tables import build_design_tables
                t = build_design_tables(
                    g["kfp"], g["worst"], g["edge_ref"], [],
                    board_pts=board.pts, default_schedule=d["schedule"],
                    tree_loads=g.get("tree_loads"), phys=g.get("phys"),
                    interior_junctions=g.get("interior_junctions"),
                    origin_mm=g.get("origin_mm"))
                ov.apply_ops_to_tables(t, rep)
            return g, t, missed, rep

        # ── [삭제] 차수 2 절점
        op_del = {"id": "d1", "op": "delete",
                  "target": list(node_keys[cand_node]), "payload": {},
                  "reason": "계측", "at": ""}
        _g, t1, m1, _r = rebuild([op_del])
        print(f"  [삭제]   절점 {len(tbl0.nodes)} → {len(t1.nodes)} · "
              f"배관 {len(tbl0.pipes)} → {len(t1.pipes)} · "
              f"총연장 {_total(tbl0)} → {_total(t1)} m "
              f"(차 {abs(_total(t1) - _total(tbl0)):.3f})"
              + (f" · 못 함 {len(m1)}" if m1 else ""))

        # ── [추가] 배관 쪼개기
        op_add = {"id": "n1", "op": "add_node",
                  "target": list(pipe_keys[cand_pipe]), "payload": {"t": 0.37},
                  "reason": "계측", "at": ""}
        _g2, t2, m2, r2 = rebuild([op_add])
        was = _tbl_pipes(tbl0)[cand_pipe]
        new_pid = (r2.get("added", {}).get("pipes") or {}).get("n1")
        # ★조각 둘은 **pid 로** 찾는다 — 이름은 다시 매겨져서(§3-6) 옛 라벨로
        #   찾으면 엉뚱한 배관이 잡힌다(한 번 그렇게 읽어 「1.27 → 0.539+0.917」
        #   이라는 말이 안 되는 줄을 냈다).
        rows2 = _tbl_pipes(t2)
        # ★`cand_pipe` 는 **표 이름**이고 `pipe_labels` 의 키는 **pid** 다.
        #   둘은 이제 같은 글자꼴(P1·P2…)이라 그냥 넣으면 조용히 다른 배관이
        #   잡힌다 — 실제로 그렇게 「1.27 → 0.539+0.917」을 읽었다. 이름에서
        #   pid 로 한 번 되짚는다.
        pid_of_lab0 = {lb: pid for pid, lb in tbl0.pipe_labels.items()}
        cand_pid = pid_of_lab0.get(cand_pipe, cand_pipe)
        up = rows2.get(t2.pipe_labels.get(cand_pid, ""), {})
        dn = rows2.get(t2.pipe_labels.get(new_pid, ""), {})
        print(f"  [추가]   배관 {len(tbl0.pipes)} → {len(t2.pipes)} · "
              f"총연장 {_total(tbl0)} → {_total(t2)} m "
              f"(차 {abs(_total(t2) - _total(tbl0)):.3f}) · "
              f"원 {was['length']} m ({was['dia']}A) → "
              f"{up.get('length')} + {dn.get('length')} m "
              f"({up.get('dia')}A · {dn.get('dia')}A)"
              + (f" · 못 함 {len(m2)}" if m2 else ""))
        moved = sum(1 for pid, lb in t2.pipe_labels.items()
                    if pid in tbl0.pipe_labels and lb != tbl0.pipe_labels[pid])
        print(f"           이름이 민 배관 {moved}개 / {len(tbl0.pipe_labels)}")

        # ── [기기] 밸브
        op_eq = {"id": "e1", "op": "add_equip",
                 "target": list(pipe_keys[eq_pipe]),
                 "payload": {"lib_id": "VALVE_GATE", "t": 0.5},
                 "reason": "계측", "at": ""}
        g3, t3, m3, _r3 = rebuild([op_eq])
        base_eq = len(rebuild([])[1].equipment)      # 같은 조건의 «전» 을 센다
        added = [r for r in t3.equipment if str(r.get("lib")) == "VALVE_GATE"]
        fid0 = {k2: (v2 or {}).get("fitting_id")
                for k2, v2 in (got0["kfp"]["nodes_meta_runtime"] or {}).items()}
        fid3 = {k2: (v2 or {}).get("fitting_id")
                for k2, v2 in (g3["kfp"]["nodes_meta_runtime"] or {}).items()}
        print(f"  [기기]   기기 행 {base_eq} → {len(t3.equipment)}"
              f" · 등가길이 "
              f"{added[0]['eq_len'] if added else '?'} m "
              f"({_tbl_pipes(t3).get(str(added[0]['pipe']), {}).get('dia')}A)"
              f" · 노드 fitting_id 바뀐 수 "
              f"{sum(1 for k2 in fid0 if fid0[k2] != fid3.get(k2))}"
              + (f" · 못 함 {len(m3)}" if m3 else ""))

        # ── [⑤] 방금 만든 절점의 «값» 을 고칠 수 있는가
        #
        #    ★순서가 ①값 → ②삭제 → ③노드 → ④기기 → ⑤새 요소의 값 인 이유가
        #      이것이다. ⑤ 를 ① 과 한 벌로 돌리면 그 요소가 아직 없어서
        #      「범위 밖」으로 조용히 떨어진다 — 실제로 그랬다.
        #
        #    ★재는 자리는 **표** 다. 표고는 kfp 메타에 칸이 없는 «표 전용»
        #      칸이라(적용 ③ · `_META_PIPE`·`_META_HEAD` 에 없다), kfp 좌표의
        #      z 를 보면 영영 안 바뀐 것처럼 보인다 — 한 번 그렇게 읽었다.
        #
        #    ★행을 (x, y) 로 찾으면 안 된다. 헤드 접속관·가지 상승이 같은
        #      x·y 에 세로로 쌓여 서로 덮는다 — **표고까지** 넣어야 1:1 이다.
        g5 = copy.deepcopy(got0)
        with contextlib.redirect_stdout(io.StringIO()):
            from services.cad_import.design.tables import build_design_tables
            _n, _m, r5 = ov.apply_ops_to_kfp(g5, board, [op_add])
            new_nid = (r5.get("added", {}).get("nodes") or {}).get("n1")

            def _table5():
                return build_design_tables(
                    g5["kfp"], g5["worst"], g5["edge_ref"], [],
                    board_pts=board.pts, default_schedule=d["schedule"],
                    tree_loads=g5.get("tree_loads"), phys=g5.get("phys"),
                    interior_junctions=g5.get("interior_junctions"),
                    origin_mm=g5.get("origin_mm"))

            t5 = _table5()
            c5 = (g5["kfp"]["nodes_meta_runtime"].get(new_nid) or {}).get("coords")
            keyxyz = (int(round(float(c5[0]) * 1000)),
                      int(round(float(c5[1]) * 1000)),
                      round(float(c5[2]) if len(c5) > 2 else 0.0, 3))
            trow = next((r for r in t5.nodes
                         if (int(r["x"]), int(r["y"]),
                             round(float(r["elevation"]), 3)) == keyxyz), None)
            before = None if trow is None else float(trow["elevation"])
            rows5 = ov.put([], ov.key_added("n1"), "node", "elevation",
                           4.321, reason="계측")
            n5, m5, rep5 = ov.apply_to_kfp(g5, board, rows5)
            _n6, m6 = ov.apply_to_tables(t5, g5, board, rows5, rep5)
        after = None if trow is None else float(trow["elevation"])
        print(f"  [⑤]      새 절점({new_nid}) 표고 고치기 — "
              f"못 옮김 {len(m5) + len(m6)}건 · 표 {before} → {after} "
              f"· 원값 기록 {rows5[0].get('old')}")
        five_ok = (not m5 and not m6 and trow is not None
                   and after is not None and abs(after - 4.321) < 1e-9
                   and rows5[0].get("old") == before)

        # ── [멱등]
        _g4, t4, _m4, _r4 = rebuild([op_del, op_add, op_eq])
        _g5, t5, _m5, _r5 = rebuild([op_del, op_add, op_eq])
        same = (json.dumps(t4.as_dict(), sort_keys=True, default=str)
                == json.dumps(t5.as_dict(), sort_keys=True, default=str))
        print(f"  [멱등]   두 번 적용한 표가 같은가 — {'예' if same else '★아니오'}")

        # ── [M9] 회랑 삭제가 결합에도 실리는가
        #    결합은 04 의 표를 그대로 받으므로, 표에서 사라졌으면 결합에도 없다.
        # ★라벨로 보면 안 된다 — 절점을 지우면 **모든 라벨이 다시 매겨져서**
        #   「2」는 여전히 있다(다른 절점이다). 지워졌는지는 kfp 이름으로 본다.
        nid_of = keys.get("nid") or {}
        cand_nid = nid_of.get(cand_node)
        gone = bool(cand_nid) and cand_nid not in (
            _g["kfp"]["nodes_meta_runtime"] or {})
        print(f"  [M9]     회랑 절점({cand_nid})이 망에서 사라졌는가 — "
              f"{'예' if gone else '★아니오'} (결합은 이 표를 그대로 받는다)")
        # 등가길이는 «라이브러리에 그 호칭경이 있으면 그 값, 없으면 None» 이
        # 정답이다 — 없는데 0 이 실리면 그때가 사고다(§4 M8).
        from services.cad_import.design.fitting import load_equivalent_lengths
        _dia = int(_tbl_pipes(t3).get(str(added[0]["pipe"]), {}).get("dia") or 0)
        _want = load_equivalent_lengths().get("VALVE_GATE", {}).get(_dia)
        eq_ok = len(added) == 1 and added[0]["eq_len"] == _want
        print(f"           등가길이 판정 — 라이브러리 {_want} vs 표 "
              f"{added[0]['eq_len'] if added else '?'} → "
              f"{'맞다' if eq_ok else '★다르다'}")
        ok = (five_ok
              and abs(_total(t1) - _total(tbl0)) <= 0.001
              and abs(_total(t2) - _total(tbl0)) <= 0.001
              and same and gone and not m1 and not m2 and not m3 and eq_ok
              and all(fid0[k2] == fid3.get(k2) for k2 in fid0
                      if k2 in fid3))
        print("  " + ("★위상 수정이 실측 망에서도 선다"
                      if ok else "★★위상 수정이 실측 망에서 안 선다"))
        return 0 if ok else 3


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=DM_KEY)
    ap.add_argument("--k", type=int, default=30)
    a = ap.parse_args()
    return run(a.key, a.k)


if __name__ == "__main__":
    raise SystemExit(main())
