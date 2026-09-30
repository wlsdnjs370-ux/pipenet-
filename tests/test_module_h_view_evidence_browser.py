"""H table toggle and evidence share the native plan/isometric display geometry."""
from copy import deepcopy

import pytest

from test_module_h_browser import h_ui, menu, settle
from test_module_h_bottom_properties_browser import prepare
from test_module_h_evidence_browser import install_evidence, assert_source_border


def capture_canvas(page, selector: str) -> None:
    """Observe real overlay strokes, resetting the log at each repaint."""
    page.evaluate('''selector=>{
      const c=document.querySelector(selector).getContext('2d');let path=[];
      window.__overlay={strokes:[],texts:[],rects:[]};
      for(const name of ['setTransform','beginPath','moveTo','lineTo','stroke','fillText','roundRect']){
        const old=c[name].bind(c);c[name]=(...args)=>{
          if(name==='setTransform')window.__overlay={strokes:[],texts:[],rects:[]};
          if(name==='beginPath')path=[];
          if(name==='moveTo'||name==='lineTo')path.push(args);
          if(name==='stroke')__overlay.strokes.push({path:path.slice(),color:c.strokeStyle,width:c.lineWidth,alpha:c.globalAlpha,dash:c.getLineDash()});
          if(name==='fillText')__overlay.texts.push(args);
          if(name==='roundRect')__overlay.rects.push(args);
          return old(...args);
        };
      }
    }''', selector)


def test_third_checkbox_opens_all_rows_and_tracks_close_fold_and_menu(h_ui, monkeypatch):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    start = len(h_ui[2])
    before = deepcopy(sess['design']['tables'].pipes)
    assert page.locator('.ha-view-options label').all_text_contents() == ['등각투상', '원본 도면', '전체 테이블']
    assert not page.locator('#ha-table').is_checked()
    page.check('#ha-table')
    page.wait_for_function("document.querySelectorAll('#hp-rows tr').length===__mf.design.tables.pipes.length")
    assert page.locator('#h-properties').is_visible()
    assert page.locator('#h-properties').bounding_box()['y'] > 400
    page.locator('#hp-filter').fill('P1')
    page.click('#hp-fold')
    assert page.locator('#h-properties').bounding_box()['height'] <= 50
    assert page.locator('#ha-table').is_checked()
    page.uncheck('#ha-table')
    assert page.locator('#h-properties').is_hidden()
    page.check('#ha-table')
    assert not page.locator('#h-properties').evaluate("n=>n.classList.contains('hp-folded')")
    assert page.locator('#hp-filter').input_value() == ''
    page.wait_for_function("document.querySelectorAll('#hp-rows tr').length===__mf.design.tables.pipes.length")
    page.screenshot(path=str(tmp / 'third-checkbox-table.png'))
    page.click('#hp-close')
    page.wait_for_function("!document.querySelector('#ha-table').checked")
    page.evaluate("ModuleHProperties.select('pipe',__mf.design.tables.pipes[0].label)")
    page.wait_for_timeout(350)
    assert page.locator('#h-properties').is_hidden()
    assert not page.locator('#ha-table').is_checked()
    menu(page, '#h-table')
    page.wait_for_function("document.querySelector('#ha-table').checked")
    page.keyboard.press('Escape')
    page.wait_for_function("!document.querySelector('#ha-table').checked")
    page.evaluate("__hTest.busy(true,'fixture');ModuleHView.sync()")
    page.check('#ha-table')  # Viewing is safe during a job; edits remain guarded.
    assert page.locator('#h-properties').is_visible()
    page.uncheck('#ha-table')
    page.evaluate('__hTest.busy(false)')
    page.set_viewport_size({'width': 390, 'height': 760})
    box = page.locator('.ha-view-options').bounding_box()
    assert box['x'] >= 0 and box['x'] + box['width'] <= 390
    assert page.locator('#ha-table').is_visible()
    assert sess['design']['tables'].pipes == before
    assert not [r for r in h_ui[2][start:] if r[0] == 'POST']


def test_iso_evidence_matches_preview_including_vertical_pipe_and_annotation(h_ui, monkeypatch):
    page, _, tmp = prepare(h_ui, monkeypatch)
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    page.check('#ha-iso')
    settle(page)
    capture_canvas(page, '#h-attributes-canvas')
    expected = page.evaluate('''()=>{
      const s=__mf,v=s.design.view,u=v.underlay;
      const r=ModuleHAttributes.getState().records.find(r=>r.kind==='pipe'&&!r.generated);
      const p=v.pipes.find(p=>p.label===r.label),a=v.nodes.find(n=>n.label===p.a),b=v.nodes.find(n=>n.label===p.b);
      const annotation=ModuleHAttributes.getState().annotations[0];
      document.querySelector('#ha-annotation').value=annotation.identity;
      document.querySelector('#ha-annotation').onchange();
      const xs=v.nodes.map(n=>n.x),ys=v.nodes.map(n=>n.y),h=document.querySelector('#cv').getBoundingClientRect().height;
      const scale=Math.min(900/(Math.max(...xs)-Math.min(...xs)),450/(Math.max(...ys)-Math.min(...ys)));
      s.view={scale,ox:Math.min(...xs)-70/scale,oy:Math.max(...ys)-(h-160)/scale};
      __hTest.draw();
      const screen=n=>[s.toScreenX(n.x),s.toScreenY(n.y)];
      const nx=u.k*annotation.x+u.tx,ny=u.k*annotation.y+u.ty;
      return {label:r.label,pipe:[screen(a),screen(b)],text:[s.toScreenX((nx-ny)*u.cos30),s.toScreenY((nx+ny)*u.sin30+(u.e-u.e_ref)*u.lift)]};
    }''')
    page.evaluate("label=>ModuleHProperties.select('pipe',label)", expected['label'])
    page.wait_for_function("__overlay.rects.length>0 && __overlay.strokes.some(s=>s.path.length===2)")
    assert page.locator('#h-attributes-canvas').is_visible()
    assert any(s['path'] == expected['pipe'] for s in page.evaluate('__overlay.strokes'))
    box=page.evaluate('__overlay.rects[0]')
    assert box[0] <= expected['text'][0]-6 and box[1] < expected['text'][1]
    assert box[-1] == 6  # Rounded padded source border, no duplicate text overlay.
    page.evaluate("label=>ModuleHProperties.select('pipe',label)", expected['label'])
    page.wait_for_function("__overlay.strokes.some(s=>s.color==='#ff3b45' && s.width===7)")
    red = next(s for s in page.evaluate('__overlay.strokes') if s['color'] == '#ff3b45')
    assert red['path'] == expected['pipe'] and red['dash'] == []
    # Generated vertical segments can have coincident CAD XY, but not in iso.
    vertical = page.evaluate('''()=>{
      const v=__mf.design.view,at=new Map(v.nodes.map(n=>[String(n.label),n]));
      for(const r of ModuleHAttributes.getState().records.filter(r=>r.generated)){
        const p=v.pipes.find(p=>p.label===r.label);if(!p)continue;
        const a=at.get(String(p.a)),b=at.get(String(p.b));
        if(a.x!==b.x||a.y!==b.y)return {record:r,expected:[[[a.x,a.y],[b.x,b.y]]],actual:ModuleHView.geometry(r).segments};
      }
    }''')
    assert vertical and vertical['actual'] == vertical['expected']
    page.uncheck('#ha-show')
    page.wait_for_function("__overlay.strokes.some(s=>s.color==='#ff3b45') && !__overlay.strokes.some(s=>s.path.length===2 && s.color!=='#ff3b45')")
    assert page.locator('#h-evidence-canvas').is_visible()  # Selected text link remains with all-links off.
    assert page.locator('#ha-iso').is_checked()
    assert page.locator('#h-properties').is_hidden()
    page.screenshot(path=str(tmp / 'iso-red-selection.png'))
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before


def test_iso_text_link_uses_exact_native_annotation_and_preview_target(h_ui):
    page, sess, requests, tmp = h_ui
    install_evidence(page, sess)
    page.uncheck('#ha-show')
    page.evaluate('''()=>{
      __mf.design={...__mf.design,view:{nodes:[{label:'a',x:620,y:320},{label:'b',x:720,y:400}],
        pipes:[{label:'P25',a:'a',b:'b',load:1}],worst_path:['a','b'],
        underlay:{iso:true,k:.8,cos30:Math.cos(Math.PI/6),sin30:.5,tx:150,ty:-20,e:3,e_ref:0,lift:20}},
        hilite:new Set(),gaps:new Map(),sel:null};
      document.querySelector('#dg-plan').checked=false;document.querySelector('#dg-iso').checked=true;
      ModuleHView.sync();__hTest.draw();
    }''')
    settle(page)
    page.evaluate('__mf.view={scale:1,ox:0,oy:0};__hTest.draw()')
    page.locator('#he-targets button[data-id=p25]').click()
    page.wait_for_function('__paint.links.length===1 && __paint.labels.length===0')
    assert page.locator('#ha-iso').is_checked()
    assert page.locator('#h-evidence-canvas').is_visible()
    target = page.evaluate('[__mf.toScreenX(670),__mf.toScreenY(360)]')
    link = page.evaluate('__paint.links[0]')
    assert_source_border(page,link[1])
    assert link[0] == pytest.approx(target)
    page.evaluate('__mf.view.scale=1.2;__mf.view.ox=-30;__hTest.draw()')
    page.wait_for_function('old=>JSON.stringify(__paint.links[0])!==JSON.stringify(old)', arg=link)
    assert page.evaluate('__paint.links[0][0]') == pytest.approx(page.evaluate('[__mf.toScreenX(670),__mf.toScreenY(360)]'))
    # A fresh preview object must invalidate cached positions, not reuse old nodes.
    page.evaluate("__mf.design.view={...__mf.design.view,nodes:[{label:'a',x:520,y:320},{label:'b',x:620,y:400}]};__hTest.draw()")
    new_target = page.evaluate('[__mf.toScreenX(570),__mf.toScreenY(360)]')
    page.wait_for_function('p=>JSON.stringify(__paint.links[0]?.[0])===JSON.stringify(p)', arg=new_target)
    page.screenshot(path=str(tmp / 'iso-single-evidence-link.png'))
    page.evaluate('delete __mf.design.view.underlay;__hTest.draw()')
    page.wait_for_function('__paint.links.length===0 && __paint.labels.length===0')  # No guessed CAD alignment.
    assert not [r for r in requests if r[0] == 'POST']


def test_h_iso_uses_white_corridor_red_path_without_legacy_pink(h_ui, monkeypatch):
    page, _, tmp = prepare(h_ui, monkeypatch)
    page.check('#ha-iso')
    page.uncheck('#ha-show')
    page.uncheck('#ha-original')
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    result = page.evaluate('''async()=>{
      const c=document.querySelector('#cv').getContext('2d'),calls=[],texts=[],old={};let path=[];
      // Deliberately keep legacy provenance enabled with all pipes in review.
      __mf.boreColor=true;__mf.design.view.pipes.forEach(p=>p.bore_info={category:'review',review:true});
      let paints=0;const paint=ModuleFBore.paint;ModuleFBore.paint=(...a)=>{paints++;return paint(...a);};
      for(const name of ['beginPath','moveTo','lineTo','stroke','fillText']){
        old[name]=c[name];c[name]=function(...a){
          if(name==='beginPath')path=[];
          if(name==='moveTo'||name==='lineTo')path.push(a);
          if(name==='stroke')calls.push({path:path.slice(),color:c.strokeStyle,width:c.lineWidth,dash:c.getLineDash(),shadow:c.shadowBlur});
          if(name==='fillText')texts.push(a[0]);return old[name].apply(this,a);
        };
      }
      try{__hTest.draw();await new Promise(requestAnimationFrame);await new Promise(requestAnimationFrame);}
      finally{Object.assign(c,old);ModuleFBore.paint=paint;}
      return {calls,texts,paints,legacy:__mf.boreColor};
    }''')
    assert result['legacy'] is True  # H presentation did not mutate F/merge preferences.
    assert result['paints'] == 0 and '!' not in result['texts']
    assert any(c['color'] == '#ffffff' and c['width'] >= 3 and c['dash'] == [] for c in result['calls'])
    assert any(c['color'] == '#ff3b3b' and c['dash'] == [9, 5] for c in result['calls'])
    assert not any(c['color'] == '#e7a1ff' or c['shadow'] > 0 for c in result['calls'])
    assert page.locator('#bore-overlay').is_hidden()
    assert page.locator('#dg-bore-color').is_hidden()
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before
    page.screenshot(path=str(tmp / 'iso-white-red-network.png'))


def test_iso_geometry_keeps_heads_and_unknown_reference_alignment(h_ui, monkeypatch):
    page, _, _ = prepare(h_ui, monkeypatch)
    page.check('#ha-iso')
    result = page.evaluate('''()=>{
      const head=ModuleHAttributes.getState().records.find(r=>r.kind==='head');
      const node=__mf.design.view.nodes.find(n=>String(n.label)===String(head.node_label));
      const actual=ModuleHView.geometry(head).point;
      delete __mf.design.view.underlay;
      return {actual,expected:[node.x,node.y],missing:ModuleHView.geometry({kind:'reference',xy:[[[1,2],[3,4]]]}).segments,
        annotation:ModuleHView.projectPoint([1,2])};
    }''')
    assert result['actual'] == result['expected']
    assert result['missing'] == [] and result['annotation'] is None
