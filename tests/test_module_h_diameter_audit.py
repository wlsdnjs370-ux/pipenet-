"""Read-only report comparison must reject changed or ambiguous source geometry."""
from copy import deepcopy

import pytest

from scripts.audit_h_diameter_tracing import compare


def report():
    return dict(summary={},decisions=[dict(edge=[0,1],xy=[[0,0],[100,0]],evidence={'text_mm':40})])


def test_compare_uses_physical_xy_and_reports_value_changes():
    before=report();after=deepcopy(before)
    after['decisions'][0].update(edge=[8,6],xy=[[100,0],[0,0]],evidence={'text_mm':65})
    result=compare(before,after)
    assert result['geometry_unchanged'] and result['changed_known_values']==1
    assert result['changes'][0]['after_mm']==65
    after['decisions'][0]['evidence']['block_export']=True
    assert compare(before,after)['transitions']=={'known -> unknown':1}


def test_compare_rejects_geometry_changes_or_coincident_edges():
    before=report();after=deepcopy(before)
    after['decisions'][0]['xy'][1][0]=101
    with pytest.raises(ValueError,match='differ'): compare(before,after)
    after=report();after['decisions']*=2
    with pytest.raises(ValueError,match='Coincident'): compare(before,after)
