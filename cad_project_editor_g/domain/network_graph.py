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
        self.nodes[node.id] = node
        self._index_dirty = True
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
        self._index_dirty = True
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
        self.pipes[pipe.id] = pipe
        self._index_dirty = True
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
        self._index_dirty = True
        self._emit_with_delta(self.on_pipe_removed, pipe_id, delta)
        self.on_changed.emit(delta)
        return pipe
    
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
