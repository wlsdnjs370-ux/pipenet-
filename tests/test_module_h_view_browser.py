"""H view toggles preserve calculation data and show the actual DXF frame."""
from copy import deepcopy

import pytest

from test_module_h_browser import h_ui
from test_module_h_bottom_properties_browser import prepare


def original_world(page):
    """Install raw CAD geometry only in this isolated browser fixture."""
    page.evaluate('''()=>{
      __mf.world={bounds:{minx:-1000,miny:-1000,maxx:16000,maxy:6000},bundles:[
        {id:'drawing',css:'#ff0000',segs:[0,1000,15000,1000],circles:[7500,3000,400],arcs:[]}
      ]};ModuleHView.sync();__hTest.draw();
    }''')


def test_checkbox_modes_selection_and_original_are_display_only(h_ui, monkeypatch):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    original_world(page)
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    stored = deepcopy(sess['design']['tables'].pipes)
    start = len(h_ui[2])
    assert page.locator('#ha-original').is_checked()
    assert not page.locator('#ha-iso').is_checked()
    assert page.locator('#ha-original').locator('..').inner_text() == '원본 도면'
    assert page.locator('.ha-view-options').bounding_box()['y'] > page.locator('.ha-summary').bounding_box()['y']
    page.check('#ha-iso')
    page.wait_for_function("!document.querySelector('#dg-plan').checked")
    page.evaluate("ModuleHProperties.select('pipe',__mf.design.tables.pipes[0].label)")
    page.wait_for_timeout(500)
    assert page.locator('#ha-iso').is_checked()
    assert not page.locator('#dg-plan').is_checked()  # Selection must not force plan.
    assert page.locator('#h-properties').is_hidden()
    page.screenshot(path=str(tmp/'iso-original.png'))
    page.uncheck('#ha-original')
    assert page.locator('#ha-iso').is_checked()
    assert page.evaluate('__mf.design.sel')
    page.uncheck('#ha-iso')
    assert page.locator('#dg-plan').is_checked()
    assert not page.locator('#ha-original').is_checked()
    page.check('#ha-original')
    page.screenshot(path=str(tmp/'plan-original.png'))
    page.evaluate("__hTest.busy(true,'fixture');ModuleHView.sync()")
    assert page.locator('#ha-iso').is_disabled()
    page.uncheck('#ha-original')  # Pure drawing visibility remains usable while busy.
    page.evaluate('__hTest.busy(false);ModuleHView.sync()')
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before
    assert sess['design']['tables'].pipes == stored
    assert not [r for r in h_ui[2][start:] if r[0] == 'POST']


@pytest.mark.parametrize('iso', [False, True])
def test_original_uses_server_transform_white_dashed_frame_and_hidden_layers(h_ui, monkeypatch, iso):
    page, _, _ = prepare(h_ui, monkeypatch)
    original_world(page)
    if iso:
        page.check('#ha-iso')
    result = page.evaluate('''()=>{
      const s=__mf,u=s.design.view.underlay,iso=document.querySelector('#ha-iso').checked;
      const convert=([x,y])=>{
        if(!iso)return [x,y];
        const nx=u.k*x+u.tx,ny=u.k*y+u.ty;
        return [(nx-ny)*u.cos30,(nx+ny)*u.sin30+(u.e-u.e_ref)*u.lift];
      };
      const corners=[[-1000,-1000],[16000,-1000],[16000,6000],[-1000,6000]].map(convert);
      const xs=corners.map(p=>p[0]),ys=corners.map(p=>p[1]);
      const scale=Math.min(1000/(Math.max(...xs)-Math.min(...xs)),700/(Math.max(...ys)-Math.min(...ys)));
      const sx=x=>30+(x-Math.min(...xs))*scale,sy=y=>30+(Math.max(...ys)-y)*scale;
      const c=document.createElement('canvas');c.width=1100;c.height=800;const ctx=c.getContext('2d');
      const frames=[],stroke=ctx.stroke.bind(ctx);
      ctx.stroke=p=>{if(ctx.strokeStyle==='#ffffff')frames.push({dash:ctx.getLineDash().map(n=>n*scale),width:ctx.lineWidth*scale,alpha:ctx.globalAlpha});stroke(p);};
      const testState={...s,view:{...s.view,scale}};
      const draw=()=>{ctx.clearRect(0,0,c.width,c.height);ModuleHView.drawOriginal(ctx,testState,sx,sy);};
      const color=point=>{
        const p=convert(point),px=Math.round(sx(p[0])),py=Math.round(sy(p[1]));
        const bytes=ctx.getImageData(px-2,py-2,5,5).data;let maxAlpha=0;
        for(let i=0;i<bytes.length;i+=4)if(bytes[i]>200&&bytes[i+1]<10)maxAlpha=Math.max(maxAlpha,bytes[i+3]);
        return maxAlpha;
      };
      draw();const shown=color([7000,1000]);
      s.hidden.add('drawing');draw();const hidden=color([7000,1000]);s.hidden.delete('drawing');
      document.querySelector('#ha-original').checked=false;draw();const off=color([7000,1000]);
      const empty=!ctx.getImageData(0,0,c.width,c.height).data.some(v=>v);
      return {shown,hidden,off,empty,frames};
    }''')
    assert result['shown'] > 20  # Actual CAD path, aligned by the server transform.
    assert result['hidden'] == result['off'] == 0
    assert result['empty']
    assert len(result['frames']) == 2
    for frame in result['frames']:
        assert frame['dash'] == pytest.approx([6, 4])
        assert frame['width'] == pytest.approx(1)
        assert frame['alpha'] == pytest.approx(.75)


def test_missing_transform_does_not_guess_and_small_bar_is_reachable(h_ui, monkeypatch):
    page, _, _ = prepare(h_ui, monkeypatch)
    original_world(page)
    page.check('#ha-iso')
    page.evaluate('delete __mf.design.view.underlay;ModuleHView.sync();__hTest.draw()')
    assert page.locator('#ha-original').is_disabled()
    assert '좌표 변환' in page.locator('#ha-original').get_attribute('title')
    page.uncheck('#ha-iso')
    assert page.locator('#ha-original').is_enabled()
    page.set_viewport_size({'width':390, 'height':760})
    bar = page.locator('#h-attributes-bar').bounding_box()
    assert bar['x'] >= 0 and bar['x']+bar['width'] <= 390
    assert page.locator('#ha-iso').is_visible() and page.locator('#ha-original').is_visible()


def test_flat_native_preview_switches_to_iso_and_rolls_back_on_failure(h_ui, monkeypatch):
    page, _, _ = prepare(h_ui, monkeypatch)
    original_world(page)
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    start = len(h_ui[2])
    page.evaluate("async()=>{document.querySelector('#dg-iso').checked=false;await document.querySelector('#dg-iso').onchange();ModuleHView.sync()}")
    assert page.evaluate('__mf.design.view.underlay.iso') is False
    page.route('**/api/module-f/design/preview?*', lambda route: route.fulfill(status=500, json={'ok': False, 'error': 'fixture preview unavailable'}))
    page.check('#ha-iso')
    page.wait_for_function("!document.querySelector('#ha-iso').checked && !document.querySelector('#ha-iso').disabled")
    assert page.locator('#dg-plan').is_checked()
    assert not page.locator('#dg-iso').is_checked()
    page.unroute('**/api/module-f/design/preview?*')
    page.check('#ha-iso')
    page.wait_for_function("__mf.design.view.underlay.iso && !document.querySelector('#ha-iso').disabled")
    assert not page.locator('#dg-plan').is_checked()
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before
    assert not [r for r in h_ui[2][start:] if r[0] == 'POST']
