"""H underlay controls reuse native per-drawing rendering without model writes."""
from copy import deepcopy

from test_module_h_browser import h_ui
from test_module_h_integration_browser import prepare
from test_module_f_merge_underlays import drawings


def with_drawings(h_ui):
    """Use three real, aligned source layers in an isolated session."""
    page, sess, requests, tmp = prepare(h_ui)
    sess.pop('merge_editor', None)
    drawings(sess)
    page.evaluate('async()=>{await __hTest.merged();ModuleHUI.queue();}')
    page.keyboard.press('Escape')
    page.wait_for_selector('#hm-under-controls:not([hidden])')
    return page, sess, requests, tmp


def observe_frames(page):
    """Capture the actual native source-frame strokes, not inferred visibility."""
    page.evaluate('''()=>{
      const c=document.querySelector('#cv').getContext('2d'),stroke=c.stroke;
      window.__underFrames=[];
      c.stroke=function(...args){
        if(this.strokeStyle==='#ffffff'&&this.globalAlpha===.75&&this.getLineDash().length)
          __underFrames.push(1);
        return stroke.apply(this,args);
      };
    }''')


def frame_count(page):
    return page.evaluate('''async()=>{
      __underFrames=[];__hTest.draw();
      await new Promise(requestAnimationFrame);await new Promise(requestAnimationFrame);
      return __underFrames.length;
    }''')


def test_each_source_and_master_toggle_native_rendering_without_model_changes(h_ui):
    page, sess, requests, tmp = with_drawings(h_ui)
    before = deepcopy(vars(sess['merged']['combined']))
    start = len(requests)
    assert page.locator('#hm-under-label').inner_text() == '원본 도면'
    assert page.locator('#hm-under-kinds label').all_text_contents() == ['평면도', '계통도', '기계실']
    for kind in ('plan', 'system', 'machineroom'):
        assert page.locator(f'#hm-under-kinds #mg-under-{kind}').is_disabled()
    observe_frames(page)
    assert frame_count(page) == 0
    page.check('#hm-under')
    page.wait_for_function('__underFrames.length>=3')
    assert frame_count(page) == 3
    for remaining, kind in zip((2, 1, 0), ('plan', 'system', 'machineroom')):
        page.uncheck('#mg-under-' + kind)
        assert frame_count(page) == remaining
    page.check('#mg-under-plan')
    assert frame_count(page) == 1
    page.uncheck('#hm-under')
    assert frame_count(page) == 0
    page.check('#hm-under')
    assert frame_count(page) == 1  # The master does not reset the per-source choices.
    assert not page.locator('#mg-under-system').is_checked()
    page.check('#mg-under-machineroom')
    assert frame_count(page) == 2
    page.screenshot(path=str(tmp / 'separate-original-drawings.png'))
    # The view drawer still uses the same master, in either direction.
    page.click('#hm-view')
    page.uncheck('#mg-under')
    page.wait_for_function('!document.querySelector("#hm-under").checked')
    assert frame_count(page) == 0
    page.check('#mg-under')
    page.wait_for_function('document.querySelector("#hm-under").checked')
    assert frame_count(page) == 2
    assert before == vars(sess['merged']['combined'])
    assert not [r for r in requests[start:] if r[0] == 'POST']
    assert page.locator('#mg-under-plan').count() == 1


def test_multiple_systems_and_unavailable_sources_keep_native_state(h_ui):
    page, _, _, _ = with_drawings(h_ui)
    page.evaluate('''()=>{
      const rows=__mf.mergeView.underlays,second={...rows[1],kind:'system2',label:'계통도 2',token:'second-test',available:false,reason:'공통노드 없음'};
      rows[1].label='계통도 1';rows.splice(2,0,second);
      document.querySelector('#mg-under').dispatchEvent(new Event('change',{bubbles:true}));
      ModuleHIntegration.sync();
    }''')
    assert page.locator('#hm-under-kinds label').all_text_contents() == ['평면도', '계통도 1', '계통도 2', '기계실']
    page.check('#hm-under')
    assert page.locator('#mg-under-system').is_enabled()
    assert page.locator('#mg-under-system2').is_disabled()
    assert page.locator('#mg-under-system2').get_attribute('title') == '공통노드 없음'
    page.uncheck('#mg-under-system')
    assert page.locator('#mg-under-plan').is_checked()
    assert page.locator('#mg-under-machineroom').is_checked()
    page.uncheck('#hm-under')
    page.evaluate('''()=>{
      __mf.mergeView.underlays=__mf.mergeView.underlays.filter(r=>r.kind!=='system2');
      document.querySelector('#mg-under').dispatchEvent(new Event('change',{bubbles:true}));
    }''')
    assert page.locator('#mg-under-system2').count() == 0
    assert not page.locator('#mg-under-system').is_checked()


def test_underlay_bar_wraps_and_is_merge_only(h_ui):
    page, _, _, _ = with_drawings(h_ui)
    for width in (1500, 1024, 390):
        page.set_viewport_size({'width': width, 'height': 960})
        for selector in ('#hm-under-controls', '#hm-under-label', '#hm-under-kinds label'):
            for control in page.locator(selector).all():
                box = control.bounding_box()
                assert box['x'] >= 0 and box['x'] + box['width'] <= width
    page.evaluate("__hTest.setStage('edit');ModuleHIntegration.sync()")
    assert page.locator('#hm-under-controls').is_hidden()
