# -*- coding: utf-8 -*-
"""G 엔진 그래프 색인 «증분 유지 = 전체 재구축» 동등성.

노드정리 937초의 몸통을 걷어내며 NetworkGraph 색인을 증분 유지로 바꿨다
(종전: 변경 한 건마다 dirty → 다음 조회 때 O(N+P) 전체 재구축). 증분이
재구축과 **한 자리라도** 다르면 편집기의 이웃 순회·배관 조회가 조용히 다른
답을 낸다 — 그 동등성을 여기서 무작위 연산으로 두들긴다.

같은 이유로 겹침 검사 색인(frozen_geometry)도 «전수 훑기와 같은 답» 을
무작위 좌표로 확인한다. 색인은 superset 필터일 뿐이어야 한다.
"""
from __future__ import annotations

import os
import random
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
for _p in (_ROOT, _G):
    if _p not in sys.path:
        sys.path.append(_p)


def _fresh_rebuild_view(g):
    """전체 재구축이 «지금» 만들었을 색인 — 별도 사본으로 계산한다."""
    adjacency: dict[str, list[str]] = {n: [] for n in g.nodes}
    node_to_pipes: dict[str, list[str]] = {n: [] for n in g.nodes}
    key_index: dict[tuple, str] = {}
    for pid, p in g.pipes.items():
        if p.start in adjacency:
            adjacency[p.start].append(p.end)
        if p.end in adjacency:
            adjacency[p.end].append(p.start)
        if p.start in node_to_pipes:
            node_to_pipes[p.start].append(pid)
        if p.end in node_to_pipes:
            node_to_pipes[p.end].append(pid)
        key_index[(p.start, p.end)] = pid
        key_index[(p.end, p.start)] = pid
    return adjacency, node_to_pipes, key_index


def _assert_index_equal(g, ctx: str):
    g._ensure_indices()
    adjacency, node_to_pipes, key_index = _fresh_rebuild_view(g)
    assert g._adjacency == adjacency, f"[{ctx}] 인접 목록이 재구축과 다르다"
    assert g._node_to_pipes == node_to_pipes, \
        f"[{ctx}] 노드-배관 목록이 재구축과 다르다"
    assert g._pipe_key_index == key_index, \
        f"[{ctx}] 배관 키 색인이 재구축과 다르다"


def test_무작위_연산_후에도_증분_색인은_재구축과_같다():
    from domain.models import Node, Pipe
    from domain.network_graph import NetworkGraph

    rng = random.Random(20260901)
    g = NetworkGraph()
    node_seq = pipe_seq = 0
    live_nodes: list[str] = []
    live_pipes: list[str] = []

    for step in range(2000):
        op = rng.random()
        if op < 0.35 or len(live_nodes) < 2:
            node_seq += 1
            nid = f"N{node_seq}"
            g.add_node(Node(id=nid, coords=(rng.uniform(0, 50),
                                            rng.uniform(0, 50), 0.0),
                            elevation_m=0.0))
            live_nodes.append(nid)
        elif op < 0.65:
            pipe_seq += 1
            pid = f"P{pipe_seq}"
            # ★자기고리(start==end)도 섞는다. 재구축은 한 노드의 목록에 이
            #   배관을 **두 칸** 넣으므로, 증분 삭제가 두 칸을 다 빼야 한다 —
            #   `rng.sample` 만 쓰면 이 경로가 한 번도 안 돈다(손으로만 따져
            #   본 코드가 시험을 통과한 것처럼 보였다).
            if rng.random() < 0.08:
                a = b = rng.choice(live_nodes)
            else:
                a, b = rng.sample(live_nodes, 2)
            g.add_pipe(Pipe(pid, a, b))
            live_pipes.append(pid)
        elif op < 0.80 and live_pipes:
            pid = live_pipes.pop(rng.randrange(len(live_pipes)))
            g.remove_pipe(pid)
        elif op < 0.90 and live_nodes:
            nid = live_nodes.pop(rng.randrange(len(live_nodes)))
            g.remove_node(nid)      # dangling 배관을 일부러 남긴다
            live_pipes = [p for p in live_pipes if p in g.pipes]
        elif live_pipes:
            # 값 갱신 — 색인과 무관해야 한다(끝점은 아래 별도 시험).
            g.update_pipe(rng.choice(live_pipes), length_m=rng.uniform(0, 9))
        # 조회를 «중간중간» 끼운다 — 증분과 재구축이 갈리는 순간을 잡기 위해.
        if step % 97 == 0:
            _assert_index_equal(g, f"step {step}")
    _assert_index_equal(g, "마지막")


def test_평행배관과_겹쳐쓰기와_끝점변경도_재구축과_같다():
    from domain.models import Node, Pipe
    from domain.network_graph import NetworkGraph

    g = NetworkGraph()
    for nid in ("A", "B", "C"):
        g.add_node(Node(id=nid, coords=(0.0, 0.0, 0.0), elevation_m=0.0))
    # 평행 배관 — 재구축의 last-write-wins 승자와 같아야 한다.
    g.add_pipe(Pipe("P1", "A", "B"))
    g.add_pipe(Pipe("P2", "B", "A"))
    assert g.find_pipe_by_nodes("A", "B") == "P2"
    g.remove_pipe("P2")             # 주인 삭제 → 남은 평행 배관으로 되돌기
    _assert_index_equal(g, "평행 배관 주인 삭제")
    assert g.find_pipe_by_nodes("A", "B") == "P1"
    # 같은 id 겹쳐쓰기 — 드문 길은 dirty 로 빠져 재구축이 맡는다.
    g.add_pipe(Pipe("P1", "B", "C"))
    _assert_index_equal(g, "id 겹쳐쓰기")
    # 미등록 끝점 배관 → 그 노드가 뒤늦게 들어오는 경우.
    g.add_pipe(Pipe("P9", "C", "Z"))
    g.add_node(Node(id="Z", coords=(1.0, 1.0, 0.0), elevation_m=0.0))
    _assert_index_equal(g, "dangling 뒤 노드 등장")
    # update_pipe 로 끝점을 바꾸는 못된 경우 — 이제는 dirty 로 잡힌다.
    g.update_pipe("P1", end="A")
    _assert_index_equal(g, "끝점 변경")


def _editor_with(pipes_xy):
    """작은 편집기 하나 — (x1,y1)-(x2,y2) 목록으로 배관을 깐다(z=0).

    같은 좌표는 같은 노드다 — 안 그러면 «가운데 점» 이 두 노드로 갈라져
    일직선 병합 후보 자체가 안 생긴다.
    """
    from editor_core import PipeEditor
    e = PipeEditor()
    seq = [0]
    by_xy: dict[tuple, str] = {}

    def node_at(x, y):
        key = (float(x), float(y))
        nid = by_xy.get(key)
        if nid is None:
            seq[0] += 1
            nid = f"N{seq[0]}"
            e.add_node(nid, key[0], key[1], 0.0)
            by_xy[key] = nid
        return nid

    for (x1, y1, x2, y2) in pipes_xy:
        e.create_pipe(node_at(x1, y1), node_at(x2, y2))
    return e


def test_frozen_geometry_는_전수훑기와_같은_답을_낸다():
    rng = random.Random(7)
    # 격자 배관밭 — 겹치는 것·안 겹치는 것·대각까지 섞는다.
    pipes = []
    for i in range(40):
        x, y = rng.randrange(0, 20), rng.randrange(0, 20)
        if i % 3 == 0:
            pipes.append((x, y, x + rng.randrange(1, 4), y))          # X축
        elif i % 3 == 1:
            pipes.append((x, y, x, y + rng.randrange(1, 4)))          # Y축
        else:
            d = rng.randrange(1, 3)
            pipes.append((x, y, x + d, y + d))                        # 대각
    e = _editor_with(pipes)
    pm = e._pipe_mgr

    # 같은 질문 100개를 색인 있음/없음으로 각각 — 답이 전부 같아야 한다.
    questions = []
    for _ in range(100):
        x, y = rng.randrange(0, 22), rng.randrange(0, 22)
        kind = rng.randrange(3)
        if kind == 0:
            questions.append((x, y, x + rng.randrange(1, 5), y))
        elif kind == 1:
            questions.append((x, y, x, y + rng.randrange(1, 5)))
        else:
            d = rng.randrange(1, 4)
            questions.append((x, y, x + d, y + d))

    def ask_all():
        # ★임시 노드는 그래프에 «직접» 넣는다. e.add_node 는 기존 배관 위에
        #   앉는 노드로 배관을 쪼개는(편집기 동작) 부작용이 있어, 질문이
        #   그래프를 바꿔 버린다 — 두 번 물으면 두 번째 답이 달라진다.
        from domain.models import Node
        out = []
        seq = [90000]

        def tmp_node(x, y):
            seq[0] += 1
            nid = f"T{seq[0]}"
            e.graph.add_node(Node(id=nid, coords=(float(x), float(y), 0.0),
                                  elevation_m=0.0))
            return nid
        for (x1, y1, x2, y2) in questions:
            a, b = tmp_node(x1, y1), tmp_node(x2, y2)
            out.append(pm.validate_new_segment(a, b))
            e.graph.remove_node(a)
            e.graph.remove_node(b)
        return out

    plain = ask_all()
    with pm.frozen_geometry():
        indexed = ask_all()
    assert plain == indexed, "색인 답이 전수 훑기와 다르다"
    # 색인 구간 안에서 «생긴» 배관도 검사에 걸려야 한다(이벤트 추종).
    # ★NC·ND 는 그래프에 직접 넣는다 — e.add_node 는 배관 위에 앉는 노드로
    #   그 배관을 쪼개므로(NA-NB 가 사라진다) 시나리오가 성립하지 않는다.
    from domain.models import Node
    with pm.frozen_geometry():
        e.add_node("NA", 100.0, 100.0, 0.0)
        e.add_node("NB", 104.0, 100.0, 0.0)
        assert e.create_pipe("NA", "NB")[0]
        e.graph.add_node(Node(id="NC", coords=(101.0, 100.0, 0.0),
                              elevation_m=0.0))
        e.graph.add_node(Node(id="ND", coords=(103.0, 100.0, 0.0),
                              elevation_m=0.0))
        ok, _msg = pm.validate_new_segment("NC", "ND")
        assert not ok, "새로 만든 배관과의 겹침을 색인이 못 봤다"
        e.delete_pipe("NA", "NB")   # 지우면 다시 허용되어야 한다(삭제 추종)
        ok, _msg = pm.validate_new_segment("NC", "ND")
        assert ok, "지운 배관이 색인에 남아 헛걸림"


def test_색인이_못_따라가면_스스로_꺼진다():
    """구멍 난 색인은 «겹치는데 안 겹친다» 고 답한다 — 그럴 바엔 꺼야 한다.

    `Event.emit` 이 구독자 예외를 삼키므로(로그만) 아무도 안 알려 준다.
    그래서 담기에 실패하는 순간 스스로 물러나는지, 물러난 뒤의 답이 전수
    훑기와 같은지 본다.
    """
    e = _editor_with([(0, 0, 4, 0)])
    pm = e._pipe_mgr
    with pm.frozen_geometry():
        assert pm._geom_ready()
        # 키를 못 만드는 상황을 만든다 — 담다가 튀면 색인을 접어야 한다.
        boom = lambda _p: (_ for _ in ()).throw(RuntimeError("키 실패"))  # noqa: E731
        orig, pm._geom_key = pm._geom_key, boom
        try:
            e.add_node("Z1", 0.0, 9.0, 0.0)
            e.add_node("Z2", 4.0, 9.0, 0.0)
            e.create_pipe("Z1", "Z2")
        finally:
            pm._geom_key = orig
        assert pm._geom_index is None, "색인이 구멍 난 채로 살아 있다"
        assert not pm._geom_ready()
        # 꺼진 뒤에도 답은 옳아야 한다(전수 훑기로 되돌아간다).
        from domain.models import Node
        e.graph.add_node(Node(id="Q1", coords=(1.0, 0.0, 0.0), elevation_m=0.0))
        e.graph.add_node(Node(id="Q2", coords=(3.0, 0.0, 0.0), elevation_m=0.0))
        ok, _m = pm.validate_new_segment("Q1", "Q2")
        assert not ok, "색인을 끈 뒤 겹침을 놓쳤다"


def test_그래프가_바뀌면_색인을_믿지_않는다():
    """`_rebind_managers` 는 실제로 `_pipe_mgr.graph` 를 갈아끼운다.

    그때 옛 색인을 계속 쓰면 «남의 그래프» 를 보고 답한다 — 조용히 틀린다.
    """
    from domain.network_graph import NetworkGraph

    e = _editor_with([(0, 0, 4, 0)])
    pm = e._pipe_mgr
    with pm.frozen_geometry():
        assert pm._geom_ready()
        pm.graph = NetworkGraph()          # 바꿔치기
        assert not pm._geom_ready(), "바뀐 그래프인데 옛 색인을 믿는다"
        assert pm._geom_index is None


def test_내부_CRUD_는_색인을_무효화하지_않는다():
    """★이것이 937초의 정체였다 — 배관 하나마다 색인 전체 재구축.

    `_refresh_pipe_indices` 가 다시 무효화를 부르기 시작하면 그 값이 통째로
    돌아온다. 「색인이 깨끗한 채로 남는가」를 시험이 지킨다.
    """
    e = _editor_with([(0, 0, 4, 0)])
    e.graph._ensure_indices()
    assert not e.graph._index_dirty
    e.add_node("K1", 0.0, 5.0, 0.0)
    e.add_node("K2", 4.0, 5.0, 0.0)
    e.create_pipe("K1", "K2")
    assert not e.graph._index_dirty, \
        "배관을 하나 만들었을 뿐인데 색인이 통째로 무효화됐다"
    e.delete_pipe("K1", "K2")
    assert not e.graph._index_dirty, \
        "배관을 하나 지웠을 뿐인데 색인이 통째로 무효화됐다"
    # 반대로 «바깥에서 바뀌었다» 신호는 진짜로 무효화해야 한다.
    e._mark_pipe_key_index_dirty()
    assert e.graph._index_dirty, "외부 신호가 색인을 무효화하지 않는다"


def test_노드정리는_색인_구간_안에서도_같은_결과다():
    """일직선 3점 → 가운데가 지워지고, 결과가 종전 알고리즘과 같은 모양."""
    e = _editor_with([(0, 0, 2, 0), (2, 0, 5, 0), (5, 0, 5, 3)])
    res = e.cleanup_collinear_intermediate_nodes()
    assert len(res["removed"]) == 1 and not res["failed"]
    assert len(e.graph.pipes) == 2
    lengths = sorted(round(p.length_m, 6) for p in e.graph.pipes.values())
    assert lengths == [3.0, 5.0], lengths     # 2+3 병합 · 세로 3 유지
