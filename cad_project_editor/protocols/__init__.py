"""protocols — 교체 가능한 인터페이스 정의 계층"""
from protocols.solver_protocol import HydraulicSolverProtocol
from protocols.report_protocol import ReportGeneratorProtocol
from protocols.project_serializer_protocol import ProjectSerializerProtocol
from protocols.network_canvas_protocol import NetworkCanvasProtocol
from protocols.viewer_protocol import Viewer3DProtocol

__all__ = [
    "HydraulicSolverProtocol",
    "ReportGeneratorProtocol",
    "ProjectSerializerProtocol",
    "NetworkCanvasProtocol",
    "Viewer3DProtocol",
]
