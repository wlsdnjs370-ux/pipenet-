"""Opt-in, read-only checks against local saved drawings (no cache rebuilds).

Run with MODULE_F_REAL_BORE_CHECK=1. No private drawing is committed as a fixture.
"""
import os
from copy import deepcopy

import pytest

pytestmark = pytest.mark.skipif(os.environ.get('MODULE_F_REAL_BORE_CHECK') != '1',
                               reason='Local saved drawing check is opt-in')


@pytest.mark.parametrize('key',[
    'B1F 현장조사 소화설비 평면도',
    '1. 입력도면 대명동 단위세대 평면도',
])
def test_saved_drawing_direction_and_geometry(key):
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import _disp_cache_load, _board_from_data, load_edits
    from services.cad_import.design.flow import flow_for_board
    from services.cad_import.design.bore import decide_bores
    from routes.module_f.api_design import _dia_texts
    from routes.module_f.bore_context import build_bore_context
    data = _disp_cache_load(key)
    if data is None:
        pytest.skip('No current local cache; never rebuild a user cache in this test')
    board = _board_from_data(key,data)
    load_edits(board)
    if len(board.sources)!=1:
        pytest.skip('Saved drawing needs an explicit supply source')
    before = deepcopy((board.pts,board.edges,getattr(board,'edge_len_mm',{})))
    flow = flow_for_board(board)
    sess = {'key':key}
    texts = _dia_texts(sess)
    assert texts, 'Saved drawing annotation source must be available for this check'
    context = build_bore_context(sess,board,flow,texts)
    assert context.summary['metadata_status']=='native_dxf'
    assert context.summary['rotations']>0
    assert (board.pts,board.edges,getattr(board,'edge_len_mm',{}))==before
    print('DRAWING CHECK',key,context.summary)
    if key.startswith('B1F '):
        target=[]
        for edge,rec in context.decisions.items():
            rejected=rec.get('excluded_nearest') or {}
            xy=rejected.get('text_xy_mm')
            a,b=(board.pts[i] for i in edge)
            if (xy and abs(xy[0]-753168.48)<1 and abs(xy[1]-210366.11)<1
                    and abs(a[1]-b[1])<1 and rec['full_head_count']==13):
                target.append((edge,rec))
        assert target, 'Screenshot main with 13 downstream heads must be located'
        for edge,rec in target:
            assert rec.get('text_mm')!=25
            out=decide_bores({'pipe_data':{'screenshot-main':{}}},
                {'screenshot-main':edge},flow.loads,texts,pts=board.pts,context=context)
            assert out['screenshot-main'][0]==65
            assert out.evidence['screenshot-main'].get('text_mm')!=25
        print('SCREENSHOT MAIN',[(e,r.get('text_mm'),r['method']) for e,r in target])
