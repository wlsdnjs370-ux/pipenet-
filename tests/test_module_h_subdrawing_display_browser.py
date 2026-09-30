"""Actual canvas text painting and visibility remain independent of extraction."""
import pytest

from test_module_h_browser import h_ui, settle
from test_module_h_subdrawing_display import ROOT


@pytest.mark.parametrize('kind,slot,count',[('기계실','machineroom',232),('계통도','system',785)])
def test_daemyeong_source_text_is_drawn_and_layer_toggle_still_works(h_ui, kind, slot, count):
    path=ROOT/'routes'/'제출용[최종]'/f'1. 입력도면 대명동 단위세대 {kind}.dxf'
    if not path.is_file():pytest.skip('Local drawing not distributed')
    from remote30_prototype import parse_dxf_for_view
    from routes.module_f.subdrawing import layer_colors
    from routes.module_h_subdrawing import display_payload
    parsed=parse_dxf_for_view(path)
    world=display_payload(parsed['entities'],layer_colors(parsed),path)
    # Isolate canvas behavior: actual drawing data, no user sessions opened.
    page,_,requests,tmp=h_ui
    page.evaluate('''({world,slot})=>{
      Object.assign(__mf,{world,slot,stage:'sub',pick:null});
      __mf.hidden=__hTest.sysDefaultHidden();
      const c=document.querySelector('#cv').getContext('2d'),original=c.fillText.bind(c);
      window.__subLabels=[];c.fillText=(...a)=>{__subLabels.push(a[0]);return original(...a);};
      const b=world.bounds,h=document.querySelector('#cv').clientHeight,w=document.querySelector('#cv').clientWidth;
      const scale=Math.min((w-80)/(b.maxx-b.minx),(h-80)/(b.maxy-b.miny));
      __mf.view={scale,ox:b.minx-40/scale,oy:b.miny-40/scale};__hTest.draw();
    }''',dict(world=world,slot=slot))
    # Tiny native labels may be below one pixel at overview scale; retain them
    # in data instead of inflating all of them to a 3 px text cloud.
    assert len(world['texts']) == count
    page.wait_for_function('__subLabels.length>0')
    assert page.evaluate('__mf.hidden.size')==sum(b['source_hidden'] for b in world['bundles'])
    assert any(t['text'].splitlines()[0] in page.evaluate('__subLabels') for t in world['texts'])
    page.screenshot(path=str(tmp/f'daemyeong-{slot}-labels.png'))
    page.evaluate('''()=>{__mf.hidden=new Set(__mf.world.bundles.map(b=>b.id));__subLabels=[];__hTest.draw();}''')
    page.wait_for_timeout(100)
    assert page.evaluate('__subLabels')==[]
    assert not [p for method,p in requests if method=='POST']


def test_native_hidden_layers_fit_and_subpixel_text(h_ui):
    import ezdxf
    from routes.module_f.common import _boot
    from routes.module_h_subdrawing import display_payload
    from src.pipenet_converter.render.dxf_source import source_display
    _boot()
    doc=ezdxf.new()
    doc.layers.new('OFF').off()
    doc.modelspace().add_line((0,0),(1000,1000))
    doc.modelspace().add_line((1e8,0),(1e8,1000),dxfattribs={'layer':'OFF'})
    doc.modelspace().add_text('tiny native',dxfattribs={'insert':(100,100),'height':.1})
    doc.modelspace().add_text('visible native',dxfattribs={'insert':(300,300),'height':20})
    world=display_payload([],{},'unused',source=source_display(doc))
    page,_,requests,tmp=h_ui
    page.evaluate('''world=>{
      Object.assign(__mf,{world,slot:'system',stage:'sub',pick:null});
      __mf.hidden=__hTest.sysDefaultHidden();__mf.world._grpInit=true;
      const c=document.querySelector('#cv').getContext('2d'),orig=c.fillText.bind(c);
      window.__nativePaint=[];c.fillText=(t,...a)=>{__nativePaint.push([t,c.font]);return orig(t,...a);};
      __hTest.draw();document.querySelector('#btn-fit').click();
    }''',world)
    page.wait_for_function('__nativePaint.some(x=>x[0]==="visible native")')
    assert page.evaluate('__mf.hidden.size')==1
    assert page.evaluate('__mf.view.scale')>.1
    assert not page.evaluate('__nativePaint.some(x=>x[0]==="tiny native")')
    assert page.evaluate('__mf.world.texts.length')==2
    # A deliberate layer toggle expands fit; it does not remove cached objects.
    page.evaluate('''()=>{__mf.hidden.clear();document.querySelector('#btn-fit').click();}''')
    assert page.evaluate('__mf.view.scale')<.001
    assert not [p for method,p in requests if method=='POST']


def test_negative_sweep_renders_clockwise_and_subpixel_labels_reappear(h_ui):
    page,_,_,_=h_ui
    world=dict(bounds=dict(minx=0,miny=0,maxx=100,maxy=100),h_source_display=True,
        bundles=[dict(id='a',css='#fff',segs=[],circles=[],arcs=[50,50,10,180,-90])],
        texts=[dict(bundle_id='a',x=10,y=10,height=.1,text='tiny',rotation=0,css='#fff')])
    page.evaluate('''world=>{
      Object.assign(__mf,{world,stage:'sub',slot:'system',pick:null});__mf.hidden=new Set();
      const c=document.querySelector('#cv').getContext('2d'),arc=c.arc.bind(c),text=c.fillText.bind(c);
      window.__arcs=[];window.__texts=[];
      c.arc=(...a)=>{__arcs.push(a);return arc(...a);};
      c.fillText=(t,...a)=>{__texts.push([t,c.font]);return text(t,...a);};__hTest.draw();
    }''',world)
    settle(page)
    page.evaluate('''()=>{__mf.view={scale:10,ox:0,oy:0};__arcs=[];__texts=[];__hTest.draw();}''')
    page.wait_for_function('__texts.some(x=>x[0]==="tiny")')
    assert page.evaluate('__texts.find(x=>x[0]==="tiny")[1].startsWith("1px")')
    assert page.evaluate('__arcs.some(a=>a[5]===false)')


def test_b1f_native_overview_review(h_ui):
    import json
    path=ROOT/'outputs/h_source_display_review/world.json'
    if not path.is_file():pytest.skip('Local B1F source review not generated')
    world=json.loads(path.read_text(encoding='utf-8'))
    page,_,requests,_=h_ui
    page.evaluate('''world=>{
      Object.assign(__mf,{world,slot:'system',stage:'sub',pick:null});
      __mf.hidden=__hTest.sysDefaultHidden();world._grpInit=true;
      __hTest.setStage('sub');__hTest.draw();
    }''',world)
    settle(page)
    page.screenshot(path=str(path.parent/'b1f-native-overview.png'))
    original_scale=page.evaluate('__mf.view.scale')
    page.click('#btn-fit')
    settle(page)
    page.screenshot(path=str(path.parent/'b1f-full-visible.png'))
    if world.get('source_view_bounds'):
        assert page.evaluate('__mf.view.scale')<original_scale/2
        page.click('#h-source-view')
        settle(page)
        assert page.evaluate('__mf.view.scale')==pytest.approx(original_scale)
    assert page.evaluate('__mf.hidden.size')==sum(b['source_hidden'] for b in world['bundles'])
    assert page.evaluate('__mf.world.counts.texts')==7080
    assert not [p for method,p in requests if method=='POST']
