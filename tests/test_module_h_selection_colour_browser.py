"""H selection is solid red; unselected evidence and hydraulic data stay intact."""
from test_module_h_browser import h_ui, menu, settle
from test_module_h_bottom_properties_browser import prepare


def pixel(page,x,y,dx=0):
    return page.evaluate('''([x,y,dx])=>{
      const c=document.querySelector('#h-attributes-canvas'),r=c.getBoundingClientRect(),d=c.width/r.width;
      return [...c.getContext('2d').getImageData(Math.round((__mf.toScreenX(x)+dx)*d),Math.round(__mf.toScreenY(y)*d),1,1).data];
    }''',[x,y,dx])


def wait_red(page,x,y):
    page.wait_for_function('''([x,y])=>{
      const c=document.querySelector('#h-attributes-canvas'),r=c.getBoundingClientRect(),d=c.width/r.width;
      const p=c.getContext('2d').getImageData(Math.round(__mf.toScreenX(x)*d),Math.round(__mf.toScreenY(y)*d),1,1).data;
      return p[0]===255&&p[1]===59&&p[2]===69&&p[3]===255;
    }''',arg=[x,y])


def test_canvas_and_table_selection_are_red_without_recolouring_network(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch)
    page.wait_for_timeout(100)
    before=page.evaluate('JSON.stringify(__mf.design.tables)')
    original_a=pixel(page,5000,2000)
    original_b=pixel(page,10000,2000)
    start=len(h_ui[2])
    pos=page.evaluate('''()=>{const b=document.querySelector('#cv').getBoundingClientRect();return [b.left+__mf.toScreenX(5000),b.top+__mf.toScreenY(2000)]}''')
    page.mouse.click(*pos)
    wait_red(page,5000,2000)
    assert pixel(page,5000,2000,2)==[255,59,69,255]  # A clearly thick stroke.
    assert pixel(page,10000,2000)==original_b
    assert page.locator('#h-properties').is_hidden()
    menu(page,'#h-table')
    target=page.evaluate("ModuleHAttributes.getState().records.find(r=>r.kind==='pipe'&&r.xy.some(([a,b])=>a[0]===10000&&b[0]===10000)).label")
    page.locator(f'#hp-rows tr[data-label="{target}"] td').first.click()
    wait_red(page,10000,2000)
    assert pixel(page,5000,2000)==original_a
    # Evidence toggle must not hide the current active pipe.
    page.uncheck('#ha-show')
    settle(page)  # The selected pixel was already red; wait for the toggle repaint.
    wait_red(page,10000,2000)
    assert pixel(page,5000,2000)==[0,0,0,0]
    assert page.evaluate('JSON.stringify(__mf.design.tables)')==before
    assert not [r for r in h_ui[2][start:] if r[0]=='POST']
    page.check('#ha-show')
    page.screenshot(path=str(tmp/'active-pipe-red.png'))


def test_unknown_diameter_selection_is_solid_not_dashed(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch)
    selected=page.evaluate('''()=>{
      const r=ModuleHAttributes.getState().records.find(r=>r.kind==='pipe'&&r.xy.some(([a,b])=>a[0]===5000&&b[0]===5000));
      r.nominal_mm=null;r.source='unresolved';
      window.dispatchEvent(new CustomEvent('module-h-evidence-select',{detail:{ids:[r.id]}}));return r.label;
    }''')
    wait_red(page,5000,2000)
    for y in range(1900,2150,25):
        assert pixel(page,5000,y)==[255,59,69,255],(selected,y)
