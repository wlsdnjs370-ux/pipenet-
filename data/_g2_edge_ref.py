# -*- coding: utf-8 -*-
"""[G2] planar.py 에 역참조(edge_ref / node_ref) 노출.

지시서 §G2: `node_id`·`remap` 은 이미 안에 있으나 밖으로 안 나간다. 이것을
결과에 실어 보낸다. **더하기만 한다** — 기존 계산·반환 키는 손대지 않으므로
전체망 .kfp 는 비트 동일하게 남는다(§3·§4).
"""
import io

p = "cad_project_editor_g/services/cad_import/convert/planar.py"
s = io.open(p, encoding="utf-8").read()

# ── ① snap_edges 를 만들 때 «어느 board 간선에서 왔는지» 를 함께 기록 ────────
old = '''    snap_edges = {tuple(sorted((remap[i], remap[j])))
                  for (i, j) in edges if remap[i] != remap[j]}'''
new = '''    snap_edges = {tuple(sorted((remap[i], remap[j])))
                  for (i, j) in edges if remap[i] != remap[j]}
    # [G2] 역참조 — 눌린 간선이 «어느 board 간선에서 왔는지». 여러 board 간선이
    # 한 자리로 눌리면 첫 것을 대표로 둔다(관경 매칭은 선분 위치만 쓰므로
    # 대표 하나로 충분하다). 이 표가 없으면 G3 의 관경이 엉뚱한 배관에 붙는다.
    _snap_origin: dict = {}
    for (i, j) in edges:
        if remap[i] == remap[j]:
            continue
        _snap_origin.setdefault(tuple(sorted((remap[i], remap[j]))), (i, j))'''
assert old in s, "snap_edges 생성부를 못 찾음"
s = s.replace(old, new, 1)

# ── ② 배관을 만들 때 pipe_id → board 간선 을 채운다 ──────────────────────
old2 = '''    pipe_seq = 0
    seen_pairs = set()
    for ri, rj in sorted(snap_edges):'''
new2 = '''    pipe_seq = 0
    seen_pairs = set()
    edge_ref: dict = {}          # [G2] kfp 배관 id → (board_i, board_j)
    for ri, rj in sorted(snap_edges):'''
assert old2 in s, "배관 루프 시작을 못 찾음"
s = s.replace(old2, new2, 1)

old3 = '''        pipe = models.Pipe(f"P{pipe_seq}", a_id, b_id)'''
new3 = '''        origin = _snap_origin.get((ri, rj))
        if origin is not None:
            edge_ref[f"P{pipe_seq}"] = origin
        pipe = models.Pipe(f"P{pipe_seq}", a_id, b_id)'''
assert old3 in s, "배관 생성부를 못 찾음"
s = s.replace(old3, new3, 1)

# ── ③ 반환에 두 키를 «더한다» ────────────────────────────────────────────
old4 = '''        "ho": ho,
        "sources": list(user_sources or ()),
    }'''
new4 = '''        "ho": ho,
        "sources": list(user_sources or ()),
        # [G2] 역참조 — 관경(G3)이 평면 mm 좌표에서 매칭하려면 kfp 배관에서
        # 원 board 간선으로 되짚을 수 있어야 한다. 기존 키는 그대로 두고 더한다.
        "edge_ref": edge_ref,
        "node_ref": node_ref,
    }'''
assert old4 in s, "반환 dict 를 못 찾음"
s = s.replace(old4, new4, 1)

# ── ④ node_ref — 노드가 어느 board 노드에서 왔는지 ──────────────────────
old5 = '''    for vid in used:
        tgt = remap[vid]
        if tgt in node_id:
            continue
        x, y = final_xy(tgt)
        node_id[tgt] = make_node(x, y, OPT_BASE_Z)'''
new5 = '''    node_ref: dict = {}          # [G2] kfp 노드 id → board 노드 인덱스
    for vid in used:
        tgt = remap[vid]
        if tgt in node_id:
            continue
        x, y = final_xy(tgt)
        node_id[tgt] = make_node(x, y, OPT_BASE_Z)
        node_ref.setdefault(node_id[tgt], vid)'''
assert old5 in s, "노드 생성 루프를 못 찾음"
s = s.replace(old5, new5, 1)

io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("planar.py — edge_ref / node_ref 노출 (더하기만)")
