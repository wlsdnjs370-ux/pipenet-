"""Module F adapter: saved text mapping + full source graph → bore evidence."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Sequence
from ezdxf.lldxf.const import DXFError

from src.pipenet_converter.dxf.diameter_annotations import enrich_diameter_annotations
from src.pipenet_converter.graph.diameter_inference import (
    DiameterAnnotation, DiameterContext, InferenceConfig, infer_diameter_annotations,
)
from src.pipenet_converter.graph.flow import FlowTree

CONFIG_PATH = Path(__file__).resolve().parents[2] / 'configs' / 'diameter_inference.json'


def build_bore_context(sess: dict, board: Any, flow: FlowTree,
                       text_points: Sequence[Sequence[float]]) -> DiameterContext:
    """Read only native text metadata; never modify the live drawing or its caches."""
    with CONFIG_PATH.open(encoding='utf-8') as stream:
        config = InferenceConfig(**json.load(stream))
    h_policy = (sess.get('design_settings') or {}).get('diameter_policy') == 'drawing_first_v1'
    annotations = [DiameterAnnotation(str(i),float(p[0]),float(p[1]),int(p[2]),source_ambiguous=h_policy)
                   for i,p in enumerate(text_points)]
    source = sess.get('dxf')
    if not isinstance(source,(str,Path)) or not Path(source).is_file():
        from services.cad_import.pipeline import handoff
        spec = Path(handoff.pick_out_dir()) / f"{sess.get('key')}_찍은스펙.json"
        if spec.is_file():
            with spec.open(encoding='utf-8') as stream:
                source = json.load(stream).get('source_dxf')
    metadata_status = 'not_available'
    if annotations and isinstance(source,(str,Path)) and Path(source).is_file():
        source = Path(source)
        stamp = (str(source.resolve()),source.stat().st_mtime_ns,source.stat().st_size,
                 tuple(tuple(p) for p in text_points), 'h-segments-v2' if h_policy else 'f')
        cached = sess.get('_diameter_annotations')
        try:
            if cached and cached['stamp'] == stamp:
                annotations = cached['annotations']
            else:
                print("[관경 위상] 원본 DXF 문자 방향 보완 중 · 도면별 최초 1회")
                native = sess.get('_h_native_texts') or {}
                annotations = enrich_diameter_annotations(source,text_points,
                    native_texts=native.get('texts') if native.get('stamp')==stamp[:3] and
                        (not h_policy or native.get('extended')) else None, extended=h_policy)
                sess['_diameter_annotations'] = {'stamp':stamp,'annotations':annotations}
            metadata_status = 'native_dxf'
        except (OSError,ValueError,RuntimeError,DXFError) as exc:
            if h_policy:
                raise ValueError(f'원본 관경 문자 근거를 읽지 못했습니다. 자동 대입을 중단했습니다: {exc}') from exc
            print(f"[관경 위상] 문자 방향 보완 실패: {type(exc).__name__}: {exc}")
            metadata_status = 'read_failed'
    physical_edges = board.edges
    if getattr(board, 'network_mode', 'tree') != 'tree':
        from services.cad_import.design.flow import network_for_board
        from src.pipenet_converter.graph.diameter_inference import AnnotationGraph
        network = network_for_board(board, index=list(board.sources).index(flow.roots[0]))
        flow = AnnotationGraph(network.edges, tuple(network.reference.roots),
                               network.reference.representatives, network.revision)
        physical_edges = (set(board.edges) | set(network.edges)
                          if (sess.get('design_settings') or {}).get('diameter_policy') == 'drawing_first_v1'
                          else network.edges)
    if h_policy:
        from src.pipenet_converter.graph.diameter_definition import define_diameters, DefinitionConfig
        with CONFIG_PATH.with_name('h_diameter_definition.json').open(encoding='utf-8') as stream:
            h_config = DefinitionConfig(**json.load(stream))
        heads = {}
        for index, nodes in enumerate(board.hnodes):
            for node in nodes:
                heads[node] = board.disk_kinds[index]
        context = define_diameters(board.pts,physical_edges,flow,annotations,
            heads=heads,barriers=getattr(board,'valves',()),inference=config,config=h_config)
        sess['_h_diameter_context'] = context
        sess['_h_diameter_annotations'] = annotations
    else:
        context = infer_diameter_annotations(board.pts,physical_edges,flow,annotations,
            barriers=getattr(board,'valves',()),config=config)
    context.summary.update(metadata_status=metadata_status,
        rotations=sum(t.rotation_deg is not None for t in annotations),
        native_sources=sum(bool(t.raw_text) for t in annotations),
        leader_sources=sum(t.leader_xy_mm is not None for t in annotations),config=str(CONFIG_PATH.name))
    print(f"[관경 위상] 전체 헤드 {context.summary['full_heads']} · 문자 {len(annotations)} · "
          f"방향 확인 {context.summary['rotations']} · 관로 대응 {context.summary['owned']}")
    return context
