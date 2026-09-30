"""Native diameter labels remain display-only observations throughout tracing."""
from dataclasses import asdict
import uuid

import ezdxf
import pytest

from src.pipenet_converter.dxf.diameter_annotations import enrich_diameter_annotations
from src.pipenet_converter.graph.diameter_definition import define_diameters
from src.pipenet_converter.graph.diameter_inference import DiameterAnnotation
from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.progress import ProgressEvent, reporting
from routes.module_h_progress import PreviewStore


def test_raw_native_prefix_multiline_and_full_rotation_preserved(tmp_path):
    doc = ezdxf.new()
    doc.layers.new('FIRE-DIA')
    ms = doc.modelspace()
    ms.add_text('SP 150', dxfattribs=dict(insert=(100, 200), height=50, rotation=270, layer='FIRE-DIA'))
    ms.add_mtext('H\\P100', dxfattribs=dict(insert=(1000, 200), char_height=50, rotation=270, layer='FIRE-DIA'))
    block = doc.blocks.new('diameter')
    block.add_text('25', dxfattribs=dict(insert=(0, 0), height=50))
    ms.add_blockref('diameter', (2000, 200), dxfattribs=dict(rotation=270, layer='FIRE-DIA'))
    path = tmp_path / 'native-source.dxf'
    doc.saveas(path)
    before = path.read_bytes()
    result = enrich_diameter_annotations(path, [(100, 200, 150), (1000, 200, 100), (2000, 200, 25)])
    assert [a.raw_text for a in result] == ['SP 150', 'H\n100', '25']
    assert [a.rotation_deg for a in result] == pytest.approx([270, 270, 270])
    assert all(a.layer == 'FIRE-DIA' for a in result)
    assert [(a.x, a.y) for a in result] == [(100, 200), (1000, 200), (2000, 200)]
    assert path.read_bytes() == before


def test_original_labels_stream_before_inference_including_unmatched_text(monkeypatch):
    import src.pipenet_converter.graph.diameter_definition as definition
    points, edges = [(0, 0), (0, 5000)], [(0, 1)]
    flow = build_flow_tree(points, edges, [[1]], [0])
    annotations = [DiameterAnnotation('t1', 100, 1000, 150, 270, 50, 'FIRE-DIA', 'SP 150'),
                   DiameterAnnotation('unmatched', 90000, 90000, 25, 0, 50, 'FIRE-DIA', '25')]
    expected = define_diameters(points, edges, flow, annotations)
    events = []
    infer = definition.infer_diameter_annotations

    def inspect(*args, **kwargs):
        assert events[0].kind == 'diameter' and events[0].data['reset']
        assert events[0].data['annotations'] == [asdict(a) for a in annotations]
        return infer(*args, **kwargs)

    monkeypatch.setattr(definition, 'infer_diameter_annotations', inspect)
    with reporting(events.append):
        actual = define_diameters(points, edges, flow, annotations)
    assert actual == expected  # Rendering observations do not change calculation evidence.
    assert actual.decisions[0, 1]['text_mm'] == 150


def test_source_labels_survive_progress_gaps_and_clear_at_next_reset():
    store = PreviewStore()
    operation = uuid.uuid4().hex
    labels = [asdict(DiameterAnnotation('t1', 100, 200, 150, 270, 50, 'FIRE-DIA', 'SP 150'))]
    store.publish(operation, ProgressEvent('diameter', 'read', {'reset': True, 'annotations': labels}))
    labels[0]['raw_text'] = 'changed outside store'
    for i in range(250):
        store.publish(operation, ProgressEvent('diameter', 'trace', {'edge': [i, i + 1]}))
    result = store.read(operation, 0)
    assert result['gap'] and result['annotation_snapshot'][0]['raw_text'] == 'SP 150'
    assert len(result['diameter_snapshot']) == 250
    assert result['snapshot_cursor'] == 251
    assert not store.read(operation, 251)['gap']
    store.publish(operation, ProgressEvent('diameter', 'read', {'reset': True, 'annotations': []}))
    for i in range(250):
        store.publish(operation, ProgressEvent('graph', 'trace', {'count': i}))
    cleared = store.read(operation, 0)
    assert cleared['gap'] and cleared['annotation_snapshot'] == [] and cleared['diameter_snapshot'] == []
    # Unrelated decoding/graph operations do not erase completed diameter text.
    other = uuid.uuid4().hex
    for i in range(250):
        store.publish(other, ProgressEvent('graph', 'trace', {'count': i}))
    assert store.read(other, 0)['annotation_snapshot'] is None
