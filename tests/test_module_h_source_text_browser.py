"""Actual canvas text painting follows native CAD labels and view transforms."""
import pytest

from src.pipenet_converter.progress import ProgressEvent
from test_module_h_browser import h_ui, settle
from test_module_h_bottom_properties_browser import prepare
from test_module_h_activity_browser import begin


def capture_text(page) -> None:
    """Capture the source-only canvas at each redraw, including text placement."""
    page.evaluate('''()=>{
      const c=document.querySelector('#h-source-text-canvas').getContext('2d');
      window.__sourcePaint=[];
      const reset=c.setTransform.bind(c),fill=c.fillText.bind(c);
      c.setTransform=(...args)=>{__sourcePaint=[];return reset(...args);};
      c.fillText=(...args)=>{const m=c.getTransform();
        __sourcePaint.push({text:args[0],xy:[m.e,m.f],angle:Math.atan2(m.b,m.a),
          local:args.slice(1),font:c.font,color:c.fillStyle});return fill(...args);};
      __mf.view.ox-=1;window.dispatchEvent(new Event('resize'));
    }''')


def test_completed_native_labels_original_positions_pan_zoom_and_toggle(h_ui, monkeypatch):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    start = len(h_ui[2])
    capture_text(page)
    page.wait_for_function('__sourcePaint.length===3')
    rows = page.evaluate('__sourcePaint')
    assert [r['text'] for r in rows] == ['65A', '40A', '40A']
    assert rows[0]['xy'] == pytest.approx(page.evaluate('[__mf.toScreenX(2500),__mf.toScreenY(100)]'))
    assert rows[1]['xy'] == pytest.approx(page.evaluate('[__mf.toScreenX(5100),__mf.toScreenY(2000)]'))
    assert rows[0]['angle'] == pytest.approx(0)
    assert all(r['angle'] == pytest.approx(0) for r in rows)
    assert all(r['color'] == '#f3f4f6' for r in rows)
    assert page.locator('#h-source-text-canvas').evaluate("n=>getComputedStyle(n).pointerEvents") == 'none'
    page.uncheck('#ha-show')  # Raw source text does not depend on A/B evidence links.
    assert page.locator('#h-source-text-canvas').is_visible()
    page.uncheck('#ha-original')
    page.wait_for_function("document.querySelector('#h-source-text-canvas').hidden")
    page.check('#ha-original')
    page.wait_for_function("!document.querySelector('#h-source-text-canvas').hidden")
    page.evaluate('__mf.view.scale=.08;__mf.view.ox=-900;__hTest.draw()')
    expected = page.evaluate('[__mf.toScreenX(5100),__mf.toScreenY(2000)]')
    page.wait_for_function('p=>Math.abs(__sourcePaint[1]?.xy[0]-p[0])<.01', arg=expected)
    assert page.evaluate('__sourcePaint[1].xy') == pytest.approx(expected)
    page.screenshot(path=str(tmp / 'original-diameter-text-plan.png'))
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before
    assert not [r for r in h_ui[2][start:] if r[0] == 'POST']


def test_native_text_keeps_iso_position_but_displays_upright(h_ui, monkeypatch):
    page, _, tmp = prepare(h_ui, monkeypatch)
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    page.check('#ha-iso')
    settle(page)
    capture_text(page)
    expected = page.evaluate('''()=>{
      const s=__mf,v=s.design.view,u=v.underlay,h=document.querySelector('#cv').clientHeight;
      const ns=v.nodes,xs=ns.map(n=>n.x),ys=ns.map(n=>n.y);
      const scale=Math.min(900/(Math.max(...xs)-Math.min(...xs)),400/(Math.max(...ys)-Math.min(...ys)));
      s.view={scale,ox:Math.min(...xs)-100/scale,oy:Math.max(...ys)-(h-240)/scale};
      __hTest.draw();
      const screen=([x,y])=>{
        const nx=u.k*x+u.tx,ny=u.k*y+u.ty;
        return [s.toScreenX((nx-ny)*u.cos30),s.toScreenY((nx+ny)*u.sin30+(u.e-u.e_ref)*u.lift)];
      };
      const xy=screen([5100,2000]),end=screen([5100,2001]);
      return {xy,angle:Math.atan2(end[1]-xy[1],end[0]-xy[0])};
    }''')
    page.wait_for_function('p=>__sourcePaint.length===3 && Math.abs(__sourcePaint[1].xy[0]-p[0])<.01', arg=expected['xy'])
    assert page.evaluate('__sourcePaint[1].xy') == pytest.approx(expected['xy'])
    assert page.evaluate('__sourcePaint[1].angle') == pytest.approx(0)
    assert page.evaluate('__sourcePaint.map(r=>r.text)') == ['65A', '40A', '40A']
    page.screenshot(path=str(tmp / 'original-diameter-text-iso.png'))
    page.evaluate('delete __mf.design.view.underlay')
    page.wait_for_function("document.querySelector('#h-source-text-canvas').hidden")
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before


def test_live_raw_text_survives_missed_progress_and_failure_then_clears_for_new_drawing(h_ui, monkeypatch):
    page, _, tmp = prepare(h_ui, monkeypatch)
    start = len(h_ui[2])
    capture_text(page)
    labels = [dict(identity='sp', x=2500, y=100, nominal_mm=150, raw_text='SP 150', rotation_deg=0, height_mm=120, layer='FIRE-DIA'),
              dict(identity='h', x=5100, y=2000, nominal_mm=100, raw_text='H\n100', rotation_deg=270, height_mm=120, layer='FIRE-DIA'),
              dict(identity='empty', x=5100, y=1000, nominal_mm=25, raw_text='')]
    labels.append(dict(labels[0], identity='duplicate'))
    # Use the real busy operation; its activity observer owns the poller. A
    # synthetic start(id) races that observer and is not a production workflow.
    page.wait_for_function("!document.body.classList.contains('h-processing')")
    begin(page, '입력값 반영')
    operation = page.evaluate('__mf.activityOperation')
    store = page._h_preview_store
    store.publish(operation, ProgressEvent('diameter', 'read', dict(reset=True, annotations=labels)))
    for i in range(250):
        store.publish(operation, ProgressEvent('diameter', 'trace', dict(edge=[i, i+1])))
    page.wait_for_function("__sourcePaint.map(r=>r.text).join('|')==='SP 150|H|100'")
    rows = page.evaluate('__sourcePaint')
    assert rows[1]['angle'] == pytest.approx(0)  # 270 degrees is display-only upright.
    assert rows[1]['local'] == [0, 0] and rows[2]['local'][1] > 0
    assert page.locator('#h-source-text-canvas').is_visible()
    page.evaluate("ModuleHProgress.stop();__hTest.busy(false);__hTest.say('fixture input failure',true)")
    page.wait_for_timeout(250)
    assert page.evaluate('__sourcePaint.map(r=>r.text)') == ['SP 150', 'H', '100']
    page.screenshot(path=str(tmp / 'original-diameter-text-after-failure.png'))
    page.evaluate("__mf.sid=null;__hTest.setStage('open')")
    page.wait_for_function("document.querySelector('#h-source-text-canvas').hidden")
    assert not [r for r in h_ui[2][start:] if r[0] == 'POST']
