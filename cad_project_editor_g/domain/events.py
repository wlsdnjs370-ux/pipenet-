"""
간단한 이벤트 시스템 (콜백 리스트 기반)

Qt Signal을 사용하지 않고도 도메인 이벤트를 전파할 수 있는 경량 시스템.
"""

import logging
from typing import Callable, Any

logger = logging.getLogger(__name__)


class Event:
    """
    경량 이벤트 클래스
    
    사용 예:
        >>> event = Event()
        >>> event.subscribe(lambda data: print(f"Received: {data}"))
        >>> event.emit("Hello")
        Received: Hello
    """
    
    def __init__(self, strict_debug: bool = False):
        self.strict_debug = bool(strict_debug)
        self._subscribers: list[Callable[[Any], None]] = []
    
    def subscribe(self, callback: Callable[[Any], None]) -> None:
        """
        이벤트 구독자 추가
        
        Args:
            callback: 이벤트 발생 시 호출될 함수
        """
        if callback not in self._subscribers:
            self._subscribers.append(callback)
    
    def unsubscribe(self, callback: Callable[[Any], None]) -> None:
        """
        이벤트 구독 해제
        
        Args:
            callback: 제거할 콜백 함수
        """
        if callback in self._subscribers:
            self._subscribers.remove(callback)
    
    def emit(self, data: Any = None) -> None:
        """
        이벤트 발생 - 모든 구독자에게 데이터 전달
        
        Args:
            data: 전달할 데이터 (선택)
        """
        for callback in self._subscribers:
            try:
                callback(data)
            except Exception as e:
                if self.strict_debug:
                    raise
                # 한 구독자의 에러가 다른 구독자에게 영향을 주지 않도록
                logger.warning("[Event] 콜백 실행 중 에러: %s", e, exc_info=True)
    
    def clear(self) -> None:
        """모든 구독자 제거"""
        self._subscribers.clear()
