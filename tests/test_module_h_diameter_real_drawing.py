"""Opt-in read-only representative analysis against a retained local drawing."""
import os
import json
from pathlib import Path
from copy import deepcopy
import pytest


@pytest.mark.skipif(os.environ.get('MODULE_H_REAL_BORE_CHECK')!='1',reason='Private local drawing opt-in')
def test_saved_h_drawing_evidence_is_inspectable_without_geometry_changes():
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import _disp_cache_load,_board_from_data,load_edits
    from services.cad_import.design.flow import flow_for_board
    from routes.module_f.api_design import _dia_texts
    from routes.module_f.bore_context import build_bore_context
    key='B1F 현장조사 소화설비 평면도_컨셉2-수정본 (1)'
    data=_disp_cache_load(key)
    snapshot=False
    if data is None:
        from services.cad_import.pipeline.disp_cache import _disp_cache_path
        path=Path(_disp_cache_path(key))
        if not path.is_file(): pytest.skip('No retained snapshot; never rebuild user files')
        data=json.loads(path.read_text(encoding='utf-8'))['data'];snapshot=True
    board=_board_from_data(key,data);load_edits(board)
    if not board.sources: pytest.skip('No saved supply selection')
    before=deepcopy((board.pts,board.edges,board.disks))
    sess={'key':key,'design_settings':{'diameter_policy':'drawing_first_v1'}}
    context=build_bore_context(sess,board,flow_for_board(board),_dia_texts(sess))
    assert context.summary['metadata_status']=='native_dxf'
    assert context.summary['full_drawing_edges']>=len(board.edges)
    assert before==(board.pts,board.edges,board.disks)
    assert all(r.get('source_edge') and r.get('policy')=='drawing_first_v1' for r in context.decisions.values())
    print('H_READ_ONLY_DRAWING',{'historical_snapshot':snapshot,**context.summary})
