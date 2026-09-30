"""Canvas regression using the MF20 review payload, never a live user session."""
import json
from pathlib import Path

import pytest

from test_module_h_browser import h_ui, settle


def test_mf20_labels_and_c_arcs_render_in_the_same_view(h_ui):
    path=Path('outputs/mf20_display_review/world.json')
    if not path.is_file():
        pytest.skip('Run scripts.audit_mf20_display against the local MF20 drawing first')
    world=json.loads(path.read_text(encoding='utf-8'))
    page,_,requests,tmp=h_ui
    page.evaluate('''world=>{
      Object.assign(__mf,{world,slot:'plan',stage:'pick',pick:null,suggest:null});
      __mf.hidden=new Set();__mf.view={scale:.075,ox:188000,oy:3017800};
      const c=document.querySelector('#cv').getContext('2d');
      const text=c.fillText.bind(c),arc=c.arc.bind(c);
      window.__paintLabels=[];window.__paintArcs=[];
      c.fillText=(...a)=>{__paintLabels.push(a[0]);return text(...a);};
      c.arc=(...a)=>{__paintArcs.push(a);return arc(...a);};
      __hTest.draw();
    }''',world)
    settle(page)  # Let the normal first-stage fit finish before the detail view.
    page.evaluate('''()=>{
      __mf.view={scale:.075,ox:188000,oy:3017800};
      __paintLabels=[];__paintArcs=[];__hTest.draw();
    }''')
    settle(page)
    page.wait_for_function('__paintLabels.filter(t=>t==="50").length>=2')
    # C-shaped 270-degree arcs (large sweep) are painted, not replaced by chords.
    assert page.evaluate('__paintArcs.filter(a=>Math.abs(Math.abs(a[4]-a[3])-Math.PI*1.5)<.001).length')>=20
    page.screenshot(path=str(path.parent/'mf20-corrected.png'))
    page.evaluate('''()=>{
      __mf.hidden=new Set(__mf.world.bundles.map(b=>b.id));
      __paintLabels=[];__paintArcs=[];__hTest.draw();
    }''')
    page.wait_for_timeout(100)
    assert page.evaluate('__paintLabels')==[]
    assert page.evaluate('__paintArcs')==[]
    assert not [r for r in requests if r[0]=='POST']
