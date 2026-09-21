"""Merged exports retain standalone plan scale and local nozzle spacing."""
from copy import deepcopy
from pathlib import Path
import math
import xml.etree.ElementTree as ET

import pytest

from src.pipenet_converter.render.framing import reference_display_scale
from routes.module_f.common import _boot
from routes.module_f.merge import merge_network, bake_combined_iso
from routes.module_f.emit import emit_merged
from test_module_f_system_layout import heads, system


def points(path):
    return {n.get('label'): tuple(float(n.find('Position').get(k)) for k in ('x', 'y'))
            for n in ET.parse(path).findall('.//Nodes/Node')}


def calculation_signature(path):
    """Ignore only display positions; all calculation elements must match."""
    root = ET.parse(path).getroot()
    for parent in root.iter():
        for position in parent.findall('Position'):
            parent.remove(position)
    return [(el.tag, tuple(sorted(el.attrib.items())), (el.text or '').strip())
            for tag in ('Nodes', 'Links') for el in root.find('.//' + tag).iter()]


def test_reference_scale_is_uniform_and_explicit():
    assert reference_display_scale([(10000, -3000), (20000, 7000)]) == .3
    assert reference_display_scale([(10, 10)]) == 1
    for bad in ([], [(float('nan'), 0)]):
        with pytest.raises(ValueError):
            reference_display_scale(bad)
    with pytest.raises(ValueError):
        reference_display_scale([(0, 0)], canvas_units=0)


@pytest.mark.parametrize('height', [4, 80])
def test_display_scale_and_nozzle_gap_do_not_depend_on_riser_height(tmp_path, height):
    _boot()
    from services.cad_import.design.emit import emit_design_sdf
    plan = heads()
    for p in plan.pipes:
        p['type'] = 'KSD 3507'
    riser = system(True)
    for n in riser['nodes']:
        n['elevation'] = height if n['label'] in ('1', 'n2') else 0
    riser['pipes'][1].update(length=height, elev=-height)
    before = deepcopy(plan.__dict__)
    got = merge_network(plan, riser=riser, mode='lsp_gravity')
    combined_before = deepcopy(got['combined'].__dict__)
    ref = got['parts']['plan']
    old = emit_merged(got['combined'], tmp_path/'old', iso_nodes=bake_combined_iso(got)[0])
    new = emit_merged(got['combined'], tmp_path/'new', iso_nodes=bake_combined_iso(got)[0],
                      display_reference_labels=ref)
    standalone = tmp_path/'standalone.sdf'
    emit_design_sdf(plan, standalone, iso=True, iso_ref_label='1')
    from routes.module_f.export_compat import prepare_sdf_export
    prepare_sdf_export(standalone)
    a, b = points(standalone), points(new['sdf_iso'])
    for first, second in [('1', '2'), ('2', '3'), ('1', '3')]:
        pa, pb = str(int(first)+9), str(int(second)+9)
        assert (b[pb][0]-b[pa][0], b[pb][1]-b[pa][1]) == pytest.approx(
            (a[second][0]-a[first][0], a[second][1]-a[first][1]), abs=.02)
    nozzle = ET.parse(new['sdf_iso']).find('.//Nozzle')
    assert math.dist(b[nozzle.get('input')], b[nozzle.get('output')]) == pytest.approx(
        math.dist(a['3'], a['@/1']), abs=.02)
    for variant in ('sdf', 'sdf_iso'):
        assert calculation_signature(old[variant]) == calculation_signature(new[variant])
    assert new.get('kfp') and new.get('kfp_iso'), new['warnings']
    assert Path(new['kfp']).read_bytes() == Path(new['kfp_iso']).read_bytes()
    assert plan.__dict__ == before
    assert got['combined'].__dict__ == combined_before


def test_bad_reference_is_rejected_before_writing(tmp_path):
    got = merge_network(heads(), riser=system(True), mode='lsp_gravity')
    with pytest.raises(ValueError, match='기준 절점'):
        emit_merged(got['combined'], tmp_path, display_reference_labels=['missing'])
    assert not list(tmp_path.glob('*.sdf'))


def test_production_api_passes_plan_reference():
    source = (Path(__file__).resolve().parents[1] / 'routes/module_f/api_merge.py').read_text(encoding='utf-8')
    assert 'display_reference_labels=' in source
