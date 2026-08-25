from typing import Protocol, runtime_checkable, Any, Optional

@runtime_checkable
class Viewer3DProtocol(Protocol):
    """3D 뷰어 생성을 위한 인터페이스"""

    def generate_viewer_html(self, graph: Any, results: Optional[Any] = None) -> str:
        """
        주어진 NetworkGraph 객체를 사용하여 3D 뷰어 HTML 파일을 생성하고
        파일의 절대 경로를 반환한다.

        Args:
            graph: PipeEditor (domain/topology.py) 단일 진실 공급원 객체
            results: HydraulicModel 수리계산 결과 (옵션)

        Returns:
            str: 생성된 HTML 파일의 절대 경로
        """
        ...
