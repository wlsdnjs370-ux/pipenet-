"""Source display is complete without adding symbols/text to the path graph."""
from copy import deepcopy
from pathlib import Path

import ezdxf
import pytest

from routes.module_f.common import _boot
from routes.module_f.subdrawing import entities_to_world, layer_colors
from routes.module_f.world import _world_payload
from routes.module_h_subdrawing import display_payload
from src.pipenet_converter.render.dxf_text import read_display_texts

ROOT = Path(__file__).resolve().parents[1]


def test_multiline_block_attribute_and_original_metadata(tmp_path):
    doc=ezdxf.new()
    block=doc.blocks.new('TAG')
    block.add_text('SP 65',dxfattribs=dict(insert=(10,20),height=5,rotation=90))
    ref=doc.modelspace().add_blockref('TAG',(100,200),dxfattribs=dict(xscale=2,yscale=2))
    ref.add_attrib('NAME','급수 펌프',(120,220),dxfattribs=dict(height=10))
    ref.add_attrib('HIDDEN','숨긴 속성',(120,240),dxfattribs=dict(height=10,flags=1))
    doc.modelspace().add_mtext('기계실\\P'+('가'*130),dxfattribs=dict(insert=(0,0),char_height=12))
    path=tmp_path/'labels.dxf';doc.saveas(path);before=path.read_bytes()
    rows=read_display_texts(path)
    assert len(rows)==3
    sp=next(r for r in rows if r['text']=='SP 65')
    assert (sp['x'],sp['y'],sp['height'],sp['rotation'])==pytest.approx((120,240,10,90))
    assert any(r['text']=='급수 펌프' for r in rows)
    assert any(r['text']=='기계실\n'+'가'*130 for r in rows)
    assert path.read_bytes()==before


@pytest.mark.parametrize('kind,count', [('계통도',785),('기계실',232)])
def test_retained_daemyeong_native_display_not_dropped(kind,count):
    path=ROOT/'routes'/'제출용[최종]'/f'1. 입력도면 대명동 단위세대 {kind}.dxf'
    if not path.is_file():pytest.skip('Local drawing not distributed')
    _boot()
    from remote30_prototype import parse_dxf_for_view
    parsed=parse_dxf_for_view(path);entities=parsed['entities'];before=deepcopy(entities)
    colors=layer_colors(parsed)
    old=_world_payload(entities_to_world(entities,colors))
    result=display_payload(entities,colors,path)
    assert 'texts' not in old  # The original F display contract is not changed.
    assert len(result['texts'])==count
    assert result['shown']==result['counts']
    assert not any(result['dropped'].values())
    assert result['counts']['segs']>old['counts']['segs']  # H/S outlines restored.
    assert entities==before  # No new connection candidates or changed path geometry.
    assert all(t['bundle_id'] in {b['id'] for b in result['bundles']} for t in result['texts'])

