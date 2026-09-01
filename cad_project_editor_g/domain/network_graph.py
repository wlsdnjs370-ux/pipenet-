"""
NetworkGraph - 통합 네트워크 자료구조

기존의 분산 dict 구조를
하나의 통합된 자료구조로 관리하는 핵심 도메인 클래스.

주요 기능:
- 단일 저장소: dict[str, Node], dict[str, Pipe]
- 인접 인덱스: 빠른 토폴로지 쿼리 (O(1) 이웃 조회)
- 이벤트 시스템: 변경 시 자동 알림
- 스냅샷/복원: Undo/Redo 지원
"""

from typing import Any
from .models import Node, Pipe, NodeTypeId
from .events import Event


class NetworkGraph:
    """
    네트워크 그래프 통합 자료구조
    
    [핵심 설계 원칙]
    1. 노드/배관 데이터는 이 클래스만이 소유한다
    2. 외부에서는 반드시 이 클래스의 메서드를 통해 접근한다
    3. 인덱스는 지연 계산 (dirty flag 기반)
    4. 모든 변경은 이벤트로 알린다
    """
    
    def __init__(self):
        # 1. 주 저장소
        self.nodes: dict[str, Node] = {}
        self.pipes: dict[str, Pipe] = {}
        
        # 2. 인접 인덱스 (캐시)
        self._adjacency: dict[str, list[str]] = {}           # node_id -> [neighbor_ids]
        self._node_to_pipes: dict[str, list[str]] = {}       # node_id -> [pipe_ids]
        self._pipe_key_index: dict[tuple[str, str], str] = {} # (start, end) -> pipe_id
        self._index_dirty: bool = False
        
        # 3. 이벤트
        self.on_node_added = Event()
        self.on_node_removed = Event()
        self.on_node_changed = Event()
        self.on_pipe_added = Event()
        self.on_pipe_removed = Event()
        self.on_pipe_changed = Event()
        self.on_changed = Event()  # 범용 변경 이벤트
        self.on_batch_begin = Event()
        self.on_batch_end = Event()
        self._batch_depth = 0

    @property
    def is_in_batch(self) -> bool:
        """현재 배치 작업(begin_batch) 중인지 여부"""
        return self._batch_depth > 0

    def begin_batch(self, reason: str = "") -> None:
        """복합 작업 시작을 알린다. 뷰는 이 신호를 받으면 갱신을 보류한다."""
        self._batch_depth = getattr(self, "_batch_depth", 0) + 1
        if self._batch_depth == 1:
            self.on_batch_begin.emit({"reason": reason})

    def end_batch(self, reason: str = "") -> None:
        """복합 작업 종료를 알린다. 뷰는 보류했던 갱신을 한 번에 수행한다."""
        self._batch_depth = getattr(self, "_batch_depth", 0) - 1
        if self._batch_depth <= 0:
            self._batch_depth = 0
            self.on_batch_end.emit({"reason": reason})

    def _emit_with_delta(self, event: Event, *payload: Any) -> None:
        """
        Event.emit 다중 인자/단일 인자 구현을 모두 지원한다.
        - 다중 인자 지원 시: 원본 인자를 그대로 전달
        - 단일 인자만 지원 시: 인자 튜플 하나로 전달
        """
        try:
            event.emit(*payload)
        except TypeError:
            event.emit(payload)
    
    # ========== CRUD: 노드 ==========
    
    def add_node(self, node: Node) -> None:
        """
        노드 추가

        Args:
            node: 추가할 노드 객체
        """
        # ★증분 유지 — «더럽히고 다음 조회 때 전체 재구축» 이 아니라 그 자리에서
        #   고친다. 종전에는 변경 한 건마다 _index_dirty=True 였고, CAD 변환의
        #   노드정리(병합 7,204건 · 건마다 조회가 끼므로 건마다 O(N+P) 재구축)가
        #   이 한 줄 때문에 **937초** 를 썼다. 증분 결과는 전체 재구축과 완전히
        #   같아야 한다 — 그 동등성은 각 메서드 주석과
        #   tests/test_g_graph_index_incremental.py(무작위 연산 대조)가 지킨다.
        #
        #   색인이 깨끗할 때, 이 노드를 미리 참조하는 배관(dangling)은 있을 수
        #   없다: 그런 배관이 들어오는 순간 add_pipe 가 색인을 더럽히기 때문.
        #   따라서 빈 슬롯만 만들면 전체 재구축과 같다.
        if not self._index_dirty:
            self._adjacency.setdefault(node.id, [])
            self._node_to_pipes.setdefault(node.id, [])
        self.nodes[node.id] = node
        delta = {
            "op": "added",
            "entity": "node",
            "id": node.id,
            "coords": tuple(node.coords) if getattr(node, "coords", None) is not None else None,
        }
        self._emit_with_delta(self.on_node_added, node.id, delta)
        self.on_changed.emit(delta)
    
    def remove_node(self, node_id: str) -> Node | None:
        """
        노드 제거
        
        Args:
            node_id: 제거할 노드 ID
            
        Returns:
            제거된 노드 객체 (없으면 None)
        """
        if node_id not in self.nodes:
            return None

        node = self.nodes.pop(node_id)
        # 증분 유지: 재구축은 self.nodes 만 돌므로 이 노드의 슬롯만 사라진다.
        # 남은 배관이 이 id 를 여전히 가리켜도(dangling) 다른 노드의 목록에는
        # 그대로 남는 것이 재구축과 같은 결과다(재구축도 끝점 존재를 안 본다).
        if not self._index_dirty:
            self._adjacency.pop(node_id, None)
            self._node_to_pipes.pop(node_id, None)
        delta = {"op": "removed", "entity": "node", "id": node_id}
        self._emit_with_delta(self.on_node_removed, node_id, delta)
        self.on_changed.emit(delta)
        return node
    
    def get_node(self, node_id: str) -> Node | None:
        """노드 조회"""
        return self.nodes.get(node_id)
    
    def has_node(self, node_id: str) -> bool:
        """노드 존재 여부"""
        return node_id in self.nodes
    
    def update_node(self, node_id: str, **kwargs) -> None:
        """
        노드 속성 업데이트
        
        Args:
            node_id: 대상 노드 ID
            **kwargs: 변경할 속성들
        """
        if node_id not in self.nodes:
            return
        
        node = self.nodes[node_id]
        for key, value in kwargs.items():
            if hasattr(node, key):
                setattr(node, key, value)

        changed_fields = list(kwargs.keys()) if isinstance(kwargs, dict) else []
        delta = {
            "op": "updated",
            "entity": "node",
            "id": node_id,
            "changed_fields": changed_fields,
        }
        self._emit_with_delta(self.on_node_changed, node_id, delta)
        self.on_changed.emit(delta)
    
    # ========== CRUD: 배관 ==========
    
    def add_pipe(self, pipe: Pipe) -> None:
        """
        배관 추가
        
        Args:
            pipe: 추가할 배관 객체
        """
        # 증분 유지. 두 경우만 전체 재구축에 맡긴다(드문 길):
        #   · 같은 id 겹쳐쓰기 — 옛 배관의 흔적을 목록에서 걷어내야 하는데
        #     그 비용/복잡도가 이득보다 크다.
        #   · 끝점 노드가 아직 없음 — 나중에 그 노드가 add_node 로 들어오면
        #     재구축이 이 배관을 주워 담아야 하므로, 지금 더럽혀 둔다
        #     (add_node 의 «깨끗하면 dangling 없음» 전제가 여기서 성립한다).
        if not self._index_dirty:
            if (pipe.id in self.pipes
                    or pipe.start not in self.nodes
                    or pipe.end not in self.nodes):
                self._index_dirty = True
            else:
                # 재구축과 같은 자리: 목록 끝 추가(= pipes dict 삽입 순서),
                # 키 색인은 나중 것이 이긴다(재구축의 last-write-wins 그대로).
                self._adjacency[pipe.start].append(pipe.end)
                self._adjacency[pipe.end].append(pipe.start)
                self._node_to_pipes[pipe.start].append(pipe.id)
                self._node_to_pipes[pipe.end].append(pipe.id)
                self._pipe_key_index[(pipe.start, pipe.end)] = pipe.id
                self._pipe_key_index[(pipe.end, pipe.start)] = pipe.id
        self.pipes[pipe.id] = pipe
        pipe_key = tuple(sorted((pipe.start, pipe.end)))
        delta = {
            "op": "added",
            "entity": "pipe",
            "id": pipe.id,
            "start": pipe.start,
            "end": pipe.end,
            "pipe_key": pipe_key,
        }
        self._emit_with_delta(self.on_pipe_added, pipe.id, delta)
        self.on_changed.emit(delta)
    
    def remove_pipe(self, pipe_id: str) -> Pipe | None:
        """
        배관 제거
        
        Args:
            pipe_id: 제거할 배관 ID
            
        Returns:
            제거된 배관 객체 (없으면 None)
        """
        if pipe_id not in self.pipes:
            return None

        pipe = self.pipes[pipe_id]
        pipe_key = tuple(sorted((pipe.start, pipe.end)))
        delta = {
            "op": "removed",
            "entity": "pipe",
            "id": pipe_id,
            "start": pipe.start,
            "end": pipe.end,
            "pipe_key": pipe_key,
        }

        pipe = self.pipes.pop(pipe_id)
        if not self._index_dirty:
            self._unindex_pipe(pipe)
        self._emit_with_delta(self.on_pipe_removed, pipe_id, delta)
        self.on_changed.emit(delta)
        return pipe

    def _unindex_pipe(self, pipe: Pipe) -> None:
        """배관 하나를 색인에서 걷어낸다 — 전체 재구축과 같은 결과가 되도록.

        ★목록의 «자리» 까지 지킨다. `_adjacency[n][k]` 와 `_node_to_pipes[n][k]`
          는 재구축에서 같은 배관이 나란히 앉는 짝이다(둘 다 배관마다 한 칸씩
          append). 첫 등장만 지우면(list.remove) 평행 배관이 있을 때 남는
          목록의 «순서» 가 재구축과 달라져, get_neighbors 순회에 기대는 코드가
          다른 답을 낼 수 있다 — 그래서 pid 의 자리(index)를 찾아 짝으로 뺀다.
        """
        start, end, pid = pipe.start, pipe.end, pipe.id
        for a in (start, end):
            ntp = self._node_to_pipes.get(a)
            adj = self._adjacency.get(a)
            if ntp is None or adj is None:
                continue
            while True:     # 자기고리(start==end)면 두 칸 — 일반형으로 돈다
                try:
                    i = ntp.index(pid)
                except ValueError:
                    break
                ntp.pop(i)
                if i < len(adj):
                    adj.pop(i)
            if start == end:
                break       # 한 노드에서 이미 두 칸을 다 지웠다
        # 키 색인 — 내가 주인일 때만. 평행 배관이 남아 있으면 재구축이 골랐을
        # 것(pipes dict 뒤쪽 = 목록 뒤쪽)으로 되돌린다.
        if (self._pipe_key_index.get((start, end)) == pid
                or self._pipe_key_index.get((end, start)) == pid):
            self._pipe_key_index.pop((start, end), None)
            self._pipe_key_index.pop((end, start), None)
            # 평행 배관은 두 끝점을 모두 만지므로 어느 한쪽의 목록이면 전부
            # 들어 있다. 두 끝점이 «둘 다» 지워진 채라면(온통 dangling) 목록이
            # 없어 못 찾는다 — 그때만 재구축에 맡긴다(드문 길·정확성 우선).
            lst = self._node_to_pipes.get(start)
            if lst is None:
                lst = self._node_to_pipes.get(end)
            if lst is None:
                self._index_dirty = True
                return
            repl = None
            for other_id in lst:
                other = self.pipes.get(other_id)
                if other is not None and {other.start, other.end} == {start, end}:
                    repl = other        # 마지막 것이 남는다 = last-write-wins
            if repl is not None:
                self._pipe_key_index[(repl.start, repl.end)] = repl.id
                self._pipe_key_index[(repl.end, repl.start)] = repl.id
    
    def get_pipe(self, pipe_id: str) -> Pipe | None:
        """배관 조회"""
        return self.pipes.get(pipe_id)
    
    def has_pipe(self, pipe_id: str) -> bool:
        """배관 존재 여부"""
        return pipe_id in self.pipes
    
    def update_pipe(self, pipe_id: str, **kwargs) -> None:
        """
        배관 속성 업데이트
        
        Args:
            pipe_id: 대상 배관 ID
            **kwargs: 변경할 속성들
        """
        if pipe_id not in self.pipes:
            return
        
        pipe = self.pipes[pipe_id]
        for key, value in kwargs.items():
            if hasattr(pipe, key):
                setattr(pipe, key, value)
        # 끝점이 바뀌면 색인이 낡는다. 종전에는 여기서도 안 더럽혀 조용히
        # 낡은 채로 남았다 — 증분 유지를 넣으며 이 구멍도 막는다.
        if "start" in kwargs or "end" in kwargs:
            self._index_dirty = True

        changed_fields = list(kwargs.keys()) if isinstance(kwargs, dict) else []
        pipe_key = tuple(sorted((pipe.start, pipe.end)))
        delta = {
            "op": "updated",
            "entity": "pipe",
            "id": pipe_id,
            "start": pipe.start,
            "end": pipe.end,
            "pipe_key": pipe_key,
            "changed_fields": changed_fields,
        }
        self._emit_with_delta(self.on_pipe_changed, pipe_id, delta)
        self.on_changed.emit(delta)
    
    def find_pipe_by_nodes(self, start_id: str, end_id: str) -> str | None:
        """
        시작/끝 노드로 배관 ID 찾기
        
        Args:
            start_id: 시작 노드 ID
            end_id: 끝 노드 ID
            
        Returns:
            배관 ID (없으면 None)
        """
        self._ensure_indices()
        return self._pipe_key_index.get((start_id, end_id))
    
    # ========== 쿼리 메서드 ==========
    
    def get_neighbors(self, node_id: str) -> list[str]:
        """
        인접 노드 목록 조회
        
        Args:
            node_id: 노드 ID
            
        Returns:
            인접한 노드 ID 목록
        """
        self._ensure_indices()
        return self._adjacency.get(node_id, []).copy()
    
    def get_connected_pipes(self, node_id: str) -> list[str]:
        """
        노드에 연결된 배관 목록 조회
        
        Args:
            node_id: 노드 ID
            
        Returns:
            연결된 배관 ID 목록
        """
        self._ensure_indices()
        return self._node_to_pipes.get(node_id, []).copy()
    
    def get_nodes_by_type(self, type_id: str | NodeTypeId) -> list[Node]:
        """
        특정 타입의 노드 목록 조회
        
        Args:
            type_id: 노드 타입 ID (NodeTypeId enum 또는 문자열)
            
        Returns:
            해당 타입의 노드 목록
        """
        if isinstance(type_id, NodeTypeId):
            type_id = type_id.value
        
        return [node for node in self.nodes.values() if node.type_id == type_id]
    
    def get_all_nodes(self) -> list[Node]:
        """모든 노드 목록 (복사본)"""
        return list(self.nodes.values())
    
    def get_all_pipes(self) -> list[Pipe]:
        """모든 배관 목록 (복사본)"""
        return list(self.pipes.values())
    
    # ========== 인덱스 관리 ==========
    
    def _ensure_indices(self) -> None:
        """인덱스가 dirty하면 재빌드"""
        if self._index_dirty:
            self._rebuild_indices()
            self._index_dirty = False
    
    def _rebuild_indices(self) -> None:
        """
        인접 인덱스 재구축
        
        [성능 최적화 포인트]
        - 변경 시에만 호출됨 (lazy evaluation)
        - O(P) 시간 복잡도 (P = 배관 수)
        """
        self._adjacency.clear()
        self._node_to_pipes.clear()
        self._pipe_key_index.clear()
        
        # 모든 노드를 빈 리스트로 초기화
        for node_id in self.nodes:
            self._adjacency[node_id] = []
            self._node_to_pipes[node_id] = []
        
        # 배관을 순회하며 인덱스 구축
        for pipe_id, pipe in self.pipes.items():
            start, end = pipe.start, pipe.end
            
            # 인접 리스트
            if start in self._adjacency:
                self._adjacency[start].append(end)
            if end in self._adjacency:
                self._adjacency[end].append(start)
            
            # 노드-배관 매핑
            if start in self._node_to_pipes:
                self._node_to_pipes[start].append(pipe_id)
            if end in self._node_to_pipes:
                self._node_to_pipes[end].append(pipe_id)
            
            # 배관 키 인덱스 (양방향)
            self._pipe_key_index[(start, end)] = pipe_id
            self._pipe_key_index[(end, start)] = pipe_id
    
    def invalidate_indices(self) -> None:
        """인덱스를 수동으로 무효화 (다음 쿼리 시 재빌드)"""
        self._index_dirty = True
    
    # ========== 직렬화 ==========
    
    def to_dict(self) -> dict[str, Any]:
        """
        전체 네트워크를 딕셔너리로 직렬화
        
        Returns:
            저장 가능한 딕셔너리
        """
        return {
            'nodes': {node_id: node.to_dict() for node_id, node in self.nodes.items()},
            'pipes': {pipe_id: pipe.to_dict() for pipe_id, pipe in self.pipes.items()}
        }
    
    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> 'NetworkGraph':
        """
        딕셔너리에서 NetworkGraph 복원
        
        Args:
            data: to_dict()로 생성된 딕셔너리
            
        Returns:
            복원된 NetworkGraph 객체
        """
        graph = cls()
        
        # 노드 복원
        nodes_data = data.get('nodes', {})
        for node_id, node_dict in nodes_data.items():
            node = Node.from_dict(node_dict)
            graph.nodes[node.id] = node
        
        # 배관 복원
        pipes_data = data.get('pipes', {})
        for pipe_id, pipe_dict in pipes_data.items():
            pipe = Pipe.from_dict(pipe_id, pipe_dict)
            graph.pipes[pipe.id] = pipe
        
        # 인덱스 재빌드
        graph._index_dirty = True
        
        return graph
    
    # ========== 스냅샷 (Undo/Redo 지원) ==========
    
    def snapshot(self) -> dict[str, Any]:
        """
        현재 상태의 스냅샷 생성
        
        Returns:
            복원 가능한 스냅샷 딕셔너리
        """
        return self.to_dict()
    
    def restore(self, snapshot: dict[str, Any]) -> None:
        """
        스냅샷에서 상태 복원
        
        Args:
            snapshot: snapshot()으로 생성된 딕셔너리
        """
        self.nodes.clear()
        self.pipes.clear()
        
        # 노드 복원
        for node_id, node_dict in snapshot.get('nodes', {}).items():
            node = Node.from_dict(node_dict)
            self.nodes[node.id] = node
        
        # 배관 복원
        for pipe_id, pipe_dict in snapshot.get('pipes', {}).items():
            pipe = Pipe.from_dict(pipe_id, pipe_dict)
            self.pipes[pipe.id] = pipe
        
        # 인덱스 재빌드 및 이벤트 발생
        self._index_dirty = True
        self.on_changed.emit({"op": "bulk", "entity": "graph", "type": "restore"})
    
    # ========== 유틸리티 ==========
    
    def clear(self) -> None:
        """모든 노드와 배관 제거"""
        self.nodes.clear()
        self.pipes.clear()
        self._adjacency.clear()
        self._node_to_pipes.clear()
        self._pipe_key_index.clear()
        self._index_dirty = False
        self.on_changed.emit({"op": "bulk", "entity": "graph", "type": "clear"})
    
    def node_count(self) -> int:
        """노드 개수"""
        return len(self.nodes)
    
    def pipe_count(self) -> int:
        """배관 개수"""
        return len(self.pipes)
    
    def __repr__(self) -> str:
        return f"NetworkGraph(nodes={self.node_count()}, pipes={self.pipe_count()})"
