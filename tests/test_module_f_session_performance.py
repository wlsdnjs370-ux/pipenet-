"""Timing instrumentation stays bounded and does not alter copy isolation."""
from routes.module_f.cancellation import _session_copy
from routes.module_f.performance import report
from routes.module_f.cancellation import Operation
from routes.module_f.session_policy import SELECTION_FIELDS,snapshot_fields


def test_timing_records_only_counts_and_keeps_copy_isolation():
    sess={'id':'performance-test','edit':{'heads':[1,2]},'world':{'geometry':[1,2]}}
    out=_session_copy(sess)
    assert out['world'] is sess['world']
    out['edit']['heads'].append(3)
    assert sess['edit']['heads']==[1,2]
    event=report(sess['id'])[-1]
    assert event['phase']=='session_snapshot' and event['elapsed_ms']>=0
    assert 'heads' not in event and 'world' not in event
    assert report('other-session')==[]


def test_scoped_rollback_does_not_copy_or_restore_unrelated_drawing():
    class NoCopy:
        def __deepcopy__(self,memo):
            raise AssertionError('Unrelated heavy state was copied')
    sess={'id':'scoped-test','edit':{'head':1},'slots':{'system':{'large':NoCopy()}},'design':NoCopy()}
    token=Operation('scoped-test-operation')
    token.reserve()
    token.snapshot(sess,SELECTION_FIELDS)
    sess['edit']['head']=2
    sess['worst']={'new':True}
    token.cancelled=True
    token.release()
    assert sess['edit']=={'head':1} and 'worst' not in sess
    assert isinstance(sess['design'],NoCopy)


def test_unknown_route_keeps_full_backup():
    assert snapshot_fields('/api/module-f/new-route',{}) is None
