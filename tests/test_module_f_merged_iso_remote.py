"""[오너 2026-09-22] 통합 .sdf 두 가지.

① 펌프 가압인데 펌프 제원이 없으면 PIPENET «가장 먼 헤드» 방식으로 저장한다.
   수원(맨 아래)에 0 g 만 걸려 PIPENET 이 헤드→펌프로 거꾸로 흐른다고 계산하던
   문제. 급수원 지정 한 줄과 계산 방식만 바뀌고 나머지는 한 글자도 같다.
② 아이소 .sdf 는 좌표만 4배로 넓혀 노드·기호가 작아 보이게 한다. 노즐 꼬리
   길이와 모양(방향·비율)은 그대로, KFP·HAS·평면 .sdf 는 바이트까지 같다.
"""
from pathlib import Path
import math
import sys
import xml.etree.ElementTree as ET

import pytest

ROOT = Path(__file__).resolve().parents[1]
for path in (ROOT, ROOT / 'core'):
    sys.path.insert(0, str(path))

from routes.module_f.merge import bake_combined_iso, merge_network
from routes.module_f.emit import ISO_SPREAD, emit_merged
from test_module_f_system_layout import heads, system

SAME_BYTES = ('kfp', 'kfp_iso', 'has', 'has_iso', 'slf', 'slf_iso')


def _root(path):
    return ET.parse(path).getroot()


def _positions(path):
    return {n.get('label'): (float(n.find('Position').get('x')),
                             float(n.find('Position').get('y')))
            for n in _root(path).iter('Node')}


def _signature(path, *, drop_input_spec=False):
    """좌표를 뺀 계산 요소 전부(Nodes · Links)."""
    root = _root(path)
    for parent in root.iter():
        for position in parent.findall('Position'):
            parent.remove(position)
    if drop_input_spec:
        for node in root.iter('Node'):
            if node.get('io-node') == 'Input':
                for spec in node.findall('Calculation-spec'):
                    node.remove(spec)
    return [(el.tag, tuple(sorted(el.attrib.items())), (el.text or '').strip())
            for tag in ('Nodes', 'Links') for el in root.find('.//' + tag).iter()]


def _emit(out, got, **kw):
    return emit_merged(got['combined'], out, iso_nodes=bake_combined_iso(got)[0],
                       display_reference_labels=got['parts']['plan'], **kw)


def _pumped_riser():
    """실제 추출처럼 급수원에 대기압(0 g) 경계가 걸린 계통도."""
    raw = system(True)
    raw['nodes'][0]['pressure_pa'] = 101325.0
    return raw


def test_pump_mode_without_pump_is_saved_as_most_remote_nozzle(tmp_path):
    got = merge_network(heads(), riser=_pumped_riser(), mode='hsp_pump')
    assert not got['combined'].pumps
    old = _emit(tmp_path / 'old', got)
    new = _emit(tmp_path / 'new', got, remote_nozzle=True)
    for kind in ('sdf', 'sdf_iso'):
        before, after = _root(old[kind]), _root(new[kind])
        was = before.find('.//Design-options')
        assert was.get('specification-type') == 'user-defined'
        opts = after.find('.//Design-options')
        assert opts.get('specification-type') == 'remote-nozzle'
        assert opts.find('Nozzle-specification') is None
        assert {k: v for k, v in opts.attrib.items() if k != 'specification-type'} == \
            {k: v for k, v in was.attrib.items() if k != 'specification-type'}
        old_in = [n for n in before.iter('Node') if n.get('io-node') == 'Input']
        new_in = [n for n in after.iter('Node') if n.get('io-node') == 'Input']
        assert len(old_in) == len(new_in) == 1
        assert old_in[0].find('Calculation-spec').get('pressure') == '101325'
        assert new_in[0].find('Calculation-spec') is None
        # 급수원 지정 한 줄만 빠진다 — 배관·노즐·표고·0유량 출구·좌표는 그대로.
        assert _signature(new[kind]) == _signature(old[kind], drop_input_spec=True)
        assert _positions(new[kind]) == _positions(old[kind])
        # 나머지 설정(단위·계산 옵션·그림 설정)도 그대로다.
        for tag in ('Units', 'Calc-options', 'Graphics', 'Libraries'):
            a, b = before.find('.//' + tag), after.find('.//' + tag)
            assert ET.tostring(a) == ET.tostring(b), tag
    for kind in SAME_BYTES:
        assert Path(old[kind]).read_bytes() == Path(new[kind]).read_bytes(), kind
    assert sum('가장 먼 헤드' in w for w in new['warnings']) == 2


def test_most_remote_nozzle_refuses_pumped_or_ambiguous_networks(tmp_path):
    from routes.module_f.export_compat import use_remote_nozzle_export
    got = merge_network(heads(), riser=_pumped_riser(), mode='hsp_pump')
    files = _emit(tmp_path, got)
    sdf = Path(files['sdf'])
    text = sdf.read_text(encoding='utf-8')
    pumped = tmp_path / 'pumped.sdf'
    pumped.write_text(text.replace(
        '</Links>', '<Pump-fan efficiency="100" input="1" label="1" output="n2" status="1" /></Links>'),
        encoding='utf-8')
    two = tmp_path / 'two_inputs.sdf'
    two.write_text(text.replace('io-node="No"', 'io-node="Input"', 1), encoding='utf-8')
    for path, why in ((pumped, '펌프'), (two, '2개')):
        before = path.read_bytes()
        messages = use_remote_nozzle_export(path)
        assert len(messages) == 1 and '못 바꿨습니다' in messages[0] and why in messages[0]
        assert path.read_bytes() == before


def test_iso_spread_enlarges_only_iso_positions(tmp_path):
    assert ISO_SPREAD == 4
    got = merge_network(heads(), riser=system(True), mode='lsp_gravity')
    old = _emit(tmp_path / 'old', got)
    new = _emit(tmp_path / 'new', got, iso_spread=ISO_SPREAD)
    a, b = _positions(old['sdf_iso']), _positions(new['sdf_iso'])
    assert a.keys() == b.keys()
    stubs = {nz.get('output'): nz.get('input') for nz in _root(new['sdf_iso']).iter('Nozzle')}
    for label in a:
        if label in stubs:
            continue
        # .sdf 좌표는 유효숫자 6자리로 적힌다 — 그 반올림만큼만 허용한다.
        assert b[label] == pytest.approx((a[label][0] * ISO_SPREAD, a[label][1] * ISO_SPREAD), rel=1e-5, abs=.02)
    # 노즐 꼬리는 종전 길이 그대로 — 배관만 넓혀 상대적 기호 크기를 줄인다.
    for outlet, head in stubs.items():
        assert math.dist(b[head], b[outlet]) == pytest.approx(math.dist(a[head], a[outlet]), abs=.1)
    assert _signature(new['sdf_iso']) == _signature(old['sdf_iso'])
    old_root, new_root = _root(old['sdf_iso']), _root(new['sdf_iso'])
    for tag in ('Attributes', 'Graphics', 'Libraries'):
        assert ET.tostring(old_root.find('.//' + tag)) == ET.tostring(new_root.find('.//' + tag)), tag
    for kind in ('sdf', *SAME_BYTES):
        assert Path(old[kind]).read_bytes() == Path(new[kind]).read_bytes(), kind


def test_iso_spread_needs_the_plan_reference(tmp_path):
    got = merge_network(heads(), riser=system(True), mode='lsp_gravity')
    for bad in (dict(iso_spread=3.0), dict(iso_spread=0.5, display_reference_labels=['10']),
                dict(iso_spread=float('nan'), display_reference_labels=['10'])):
        with pytest.raises(ValueError, match='아이소 좌표 배수'):
            emit_merged(got['combined'], tmp_path, iso_nodes=bake_combined_iso(got)[0], **bad)
    assert not list(tmp_path.glob('*.sdf'))


def test_production_api_spreads_iso_and_uses_remote_nozzle_only_without_pump():
    source = (ROOT / 'routes/module_f/api_merge.py').read_text(encoding='utf-8')
    body = source[source.index('def module_f_merge_emit('):source.index('def module_f_merge_download(')]
    assert 'iso_spread=ISO_SPREAD' in body
    assert 'remote_nozzle=(got.get("mode") in PUMP_MODES' in body
    assert 'not (got["combined"].pumps or ())' in body
