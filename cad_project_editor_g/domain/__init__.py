"""
K-Fire Hydraulic Solver - Domain Layer

도메인 계층: PySide6 의존성이 없는 순수 비즈니스 로직 계층
"""

from .models import Node, Pipe, HydraulicModel, NodeTypeId, normalize_node_type_id
from .events import Event
from .network_graph import NetworkGraph
from .node_manager import NodeManager
from .pipe_manager import PipeManager
from .flow_analyzer import FlowAnalyzer
from .geometry import (
    TOLERANCE, ISO_COS, ISO_SIN,
    is_close, is_same_point, dist_3d,
    snap_value, snap_point,
    normalize_vector, dot_product, angle_between_vectors, get_vector,
    to_screen_coord,
)
from .constants import (
    PIPE_DIAMETER_DN,
)
from .topology import (
    analyze_connectivity,
    analyze_pump_topology,
    get_neighbors,
    get_subtree_nodes,
    get_subtree_one_direction,
)
from .validation import (
    validate_topology_directions,
    check_isolation,
)
from .node_visual import classify_node_visual, NodeVisual
from .pump_utils import calculate_pump_curve_points, calculate_nozzle_k_factor

__all__ = [
    # models
    'Node',
    'Pipe',
    'HydraulicModel',
    'NodeTypeId',
    'normalize_node_type_id',
    # events
    'Event',
    # graph
    'NetworkGraph',
    # managers
    'NodeManager',
    'PipeManager',
    'FlowAnalyzer',
    # geometry
    'TOLERANCE',
    'ISO_COS',
    'ISO_SIN',
    'is_close',
    'is_same_point',
    'dist_3d',
    'snap_value',
    'snap_point',
    'normalize_vector',
    'dot_product',
    'angle_between_vectors',
    'get_vector',
    'to_screen_coord',
    # constants
    'PIPE_DIAMETER_DN',
    # topology
    'analyze_connectivity',
    'analyze_pump_topology',
    'get_neighbors',
    'get_subtree_nodes',
    'get_subtree_one_direction',
    # validation
    'validate_topology_directions',
    'check_isolation',
    # node_visual
    'classify_node_visual',
    'NodeVisual',
    # pump_utils
    'calculate_pump_curve_points',
    'calculate_nozzle_k_factor',
]
