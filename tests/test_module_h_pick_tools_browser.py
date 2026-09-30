"""Switching to objects releases the pen; no edits to the user's live session."""
import json

from test_module_h_browser import h_ui, seed, settle


def test_head_selection_exits_pen_and_sends_canvas_click(h_ui):
    page,sess,requests,_=h_ui
    state={'armed':True,'mode':'헤드','mat_done':True,'heads':[], 'materials':[],
           'head_label':'상향하향','highlight':{'pipe_segs':[], 'pipe_bundles':[],
           'head_circles':[], 'tri_segs':[]},'clicks':[]}
    def response(route):
        body=route.request.post_data_json or {}
        route.fulfill(content_type='application/json',body=json.dumps({
            'ok':True,'applied':True,'message':'선택','state':state,
            'report':{'모드':'헤드','동작':'추가','픽':'테스트'},
        },ensure_ascii=False))
    page.route('**/api/module-f/pick/mode',response)
    clicks=[]
    def click_response(route):
        clicks.append(route.request.post_data_json)
        response(route)
    page.route('**/api/module-f/pick/click',click_response)
    seed(page,sess,'pick')
    page.evaluate('(state)=>{__mf.pick=state}',state)
    page.evaluate("document.getElementById('pk-crop-pen').click()")
    assert page.locator('#pk-crop-pen').get_attribute('aria-pressed')=='true'
    page.evaluate("document.querySelector('.slot').click()")
    page.wait_for_function("document.getElementById('pk-crop-pen').getAttribute('aria-pressed')==='false'")
    page.locator('#cv').click(position={'x':250,'y':250})
    page.wait_for_timeout(200)
    assert len(clicks)==1
