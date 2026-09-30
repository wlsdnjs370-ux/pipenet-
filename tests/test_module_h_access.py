"""Real slot material determines H readiness; F's optional merge stays intact."""
from copy import deepcopy

from flask import Flask

from routes.module_f import jobs
from routes.module_f.slots import _slot_add_system, _slot_switch
from routes.module_h_access import access_state, register
from test_module_f_network_editor import design


def prepared_session():
    session=jobs._new_session(key='h-access-fixture')
    session['design']=design()
    for kind,key in [('system','riser'),('machineroom','machineroom')]:
        session['slots'][kind].update(key=kind,world={'bounds':[0,0,1,1]},
            **{key:{'nodes':[{'label':'1'},{'label':'2'}], 'pipes':[{'in':'1','out':'2','length':1}]}})
    return session


def test_material_readiness_is_not_a_visit_upload_or_sdf_flag():
    sess=jobs._new_session(key='h-access-empty')
    try:
        sess.update(world={'bounds':[0,0,1,1]},design_sdf_path='old.sdf',stage='conv')
        assert not access_state(sess)['can_merge']
        sess['design']=design()
        state=access_state(sess)
        assert state['groups']=={'plan':True,'system':False,'machineroom':False}
        sess['design']['h_attributes_dirty']=True
        assert not access_state(sess)['groups']['plan']
    finally: jobs._SESSIONS.pop(sess['id'],None)


def test_all_slots_extras_and_active_slot_restore_preserve_readiness():
    sess=prepared_session()
    try:
        app=Flask(__name__);register(app)
        before=deepcopy(sess['slots'])
        assert app.test_client().get('/api/module-h/access-state',query_string={'sid':sess['id']}).json['can_merge']
        assert sess['slots']==before
        _slot_switch(sess,'system')
        assert access_state(sess)['can_merge']
        kind=_slot_add_system(sess)
        assert not access_state(sess)['groups']['system']
        sess['slots'][kind]['riser']={'nodes':[1,2],'pipes':[{}]}
        assert access_state(sess)['can_merge']
        sess['slots']['machineroom']['machineroom']['pipes']=[]
        assert not access_state(sess)['can_merge']
    finally: jobs._SESSIONS.pop(sess['id'],None)


def test_expired_access_is_an_error_not_a_ready_result():
    app=Flask(__name__);register(app)
    response=app.test_client().get('/api/module-h/access-state?sid=missing')
    assert response.status_code==410


def test_h_replacement_clears_old_paths_but_f_upload_is_unchanged(tmp_path,monkeypatch):
    from routes.module_f import api_slot
    sess=prepared_session()
    app=Flask(__name__)
    api_slot.register(app,_save_upload=lambda *_a,**_kw:tmp_path/'replacement.dxf')
    display_modes=[]
    def sub_job(*_args, **kwargs):
        display_modes.append(kwargs.get('h_display', False))
        return lambda:None
    monkeypatch.setattr(api_slot,'_sub_open_job',sub_job)
    monkeypatch.setattr(api_slot,'_run_job',lambda *_a:None)
    try:
        for is_h in (False,True):
            _slot_switch(sess,'system')
            sess['riser']={'nodes':[1,2],'pipes':[{}]}
            form={'sid':sess['id'],'kind':'system'}
            if is_h:form['h_access']='1'
            response=app.test_client().post('/api/module-f/slot/open',data=form)
            assert response.status_code==200,response.json
            assert bool(sess.get('riser')) is not is_h
        assert display_modes==[False,False,True]
    finally:jobs._SESSIONS.pop(sess['id'],None)
