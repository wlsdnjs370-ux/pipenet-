"""Evidence links now end at actual native text, never representative edges."""
from copy import deepcopy

import pytest

from test_module_h_browser import h_ui, seed, settle


def test_no_resolved_evidence_can_render_repeatedly_without_stale_selection(h_ui):
    page,_,_,_=h_ui
    page.evaluate('''()=>{
      for(let i=0;i<3;i++){
        window.dispatchEvent(new CustomEvent('module-h-evidence-state',{detail:{records:[],pending:2}}));
        window.dispatchEvent(new CustomEvent('module-h-attributes-selection',{detail:{ids:[]}}));
      }
    }''')
    assert page.locator('#he-rows tr,#he-targets button').count()==0
    assert page.locator('#he-status').text_content()=='참조 0 · 미참조 0 · 전체 0'


def test_canonical_pipes_partition_once_and_outside_references_are_not_targets(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,'design')
    def source(identity, dn=40):
        return dict(text_mm=dn,annotation_id=identity,definition_source='drawing_direct')
    rows=[dict(id='outside',label='outside',kind='reference',nominal_mm=40,evidence=source('elsewhere')),
          dict(id='run',label='P1',kind='pipe',nominal_mm=40,evidence=dict(continuous_sources=[
              dict(bore_provenance=source('a')),dict(bore_provenance=source('b'))])),
          dict(id='unknown',label='P2',kind='pipe',nominal_mm=None,evidence={}),
          dict(id='manual',label='P3',kind='pipe',nominal_mm=50,evidence=dict(
              **source('stale'),manual={'value_mm':50},reference_annotation=dict(identity='current',nominal_mm=50)))]
    page.evaluate("records=>window.dispatchEvent(new CustomEvent('module-h-evidence-state',{detail:{records,pending:1}}))",rows)
    assert page.locator('#he-status').text_content()=='참조 2 · 미참조 1 · 전체 3'
    assert page.locator('#he-rows tr').count()==2
    assert page.locator('#he-rows td:nth-child(3)').all_text_contents()==['1','1']
    assert '50' in page.locator('#he-rows').text_content()
    assert not [p for method,p in requests if method=='POST']


def install_evidence(page,sess):
    seed(page,sess,'design')
    rows=[dict(id='donor',label='A',kind='reference',edges=[[1,2],[8,9]],
               xy=[[[100,600],[100,700]],[[180,200],[180,300]]],nominal_mm=100,
               evidence=dict(text_mm=100,definition_source='drawing_direct',annotation_id='text-a'))]
    for i in range(40):
        rows.append(dict(id=f'p{i}',label=f'P{i}',kind='pipe',editable=True,edges=[[100+i*2,101+i*2]],
                         xy=[[[480,200+i*6],[680,200+i*6]]],nominal_mm=100,
                         evidence=dict(text_mm=100,text_xy_mm=[210,180],text_raw='100A',definition_source='drawing_repeat',
                                       source_edge=[8,9],representative_nodes=[1,2,8,9])))
    page.evaluate('''rows=>{
      __mf.view={scale:1,ox:0,oy:0};__hTest.draw();
      const c=document.querySelector('#h-evidence-canvas').getContext('2d');
      window.__paint={labels:[],rects:[],links:[]};let path=[];
      const reset=c.setTransform.bind(c),begin=c.beginPath.bind(c),move=c.moveTo.bind(c),
        line=c.lineTo.bind(c),stroke=c.stroke.bind(c),text=c.fillText.bind(c),rect=c.fillRect.bind(c);
      c.setTransform=(...a)=>{window.__paint={labels:[],rects:[],links:[]};return reset(...a);};
      c.beginPath=()=>{path=[];return begin();};
      c.moveTo=(...a)=>{path.push(a);return move(...a);};
      c.lineTo=(...a)=>{path.push(a);return line(...a);};
      c.stroke=()=>{if(c.getLineDash().length&&path.length===2)__paint.links.push(path.slice());return stroke();};
      c.fillText=(...a)=>{__paint.labels.push(a);return text(...a);};
      c.fillRect=(...a)=>{__paint.rects.push(a);return rect(...a);};
      window.addEventListener('module-h-evidence-select',e=>window.__picked=e.detail.ids);
      window.dispatchEvent(new CustomEvent('module-h-evidence-state',{detail:{records:rows,pending:0}}));
    }''',rows)
    page.wait_for_function("document.querySelector('#he-targets button') && __paint.links.length===40")
    return rows


def assert_source_border(page, endpoint, x=210, y=180, raw_text='100A'):
    """The connector ends at the padded text boundary, not through its glyphs."""
    rect=page.evaluate('''a=>ModuleHView.annotationLayout(a,document.querySelector('#h-evidence-canvas').getContext('2d')).rect''',dict(x=x,y=y,raw_text=raw_text))
    left,top,width,height=rect
    ex,ey=endpoint
    assert left-1e-6<=ex<=left+width+1e-6 and top-1e-6<=ey<=top+height+1e-6
    assert min(abs(ex-left),abs(ex-left-width),abs(ey-top),abs(ey-top-height))<1e-6


def test_all_links_then_selected_pipe_to_actual_text_without_ab_labels(h_ui):
    page,sess,requests,tmp=h_ui
    install_evidence(page,sess)
    before=page.evaluate('JSON.stringify([__mf.edit,__mf.design])')
    assert page.locator('#he-targets button').count()==40
    assert page.evaluate('__paint.links.length')==40
    assert page.evaluate('__paint.labels')==[]
    page.uncheck('#ha-show')
    page.locator('#he-targets button[data-id=p25]').click()
    page.wait_for_function('__paint.links.length===1 && __paint.labels.length===0')
    assert page.evaluate('__picked')==['p25']
    assert_source_border(page,page.evaluate('__paint.links[0][1]'))
    assert page.locator('#he-targets button[aria-pressed=true]').count()==1
    assert page.locator('#he-targets .he-target-list').evaluate('e=>e.scrollTop')>0
    assert before==page.evaluate('JSON.stringify([__mf.edit,__mf.design])')
    assert not [p for method,p in requests if method=='POST']
    page.screenshot(path=str(tmp/'single-reference.png'))
    # A group row now inspects one pipe; it cannot silently bulk-select all 40.
    page.locator('#he-rows tr.he-selected').click()
    assert len(page.evaluate('__picked'))==1


def test_labels_do_not_overlap_or_reappear_outside_view(h_ui):
    page,sess,_,_=h_ui
    rows=install_evidence(page,sess)
    page.uncheck('#ha-show')
    rows[1]['xy']=[[[170,250],[190,250]]]
    page.evaluate("records=>window.dispatchEvent(new CustomEvent('module-h-evidence-state',{detail:{records,pending:0}}))",rows)
    page.locator('#he-targets button[data-id=p0]').click()
    page.wait_for_function('__paint.links.length===1')
    assert page.evaluate('__paint.labels')==[]
    page.evaluate('__mf.view.ox=10000;__hTest.draw()')
    page.wait_for_function('__paint.links.length===0')


def test_missing_text_does_not_link_to_a_representative_origin(h_ui):
    page,sess,_,_=h_ui
    rows=deepcopy(install_evidence(page,sess))
    page.uncheck('#ha-show')
    rows[2]['evidence']['source_edge']=[999,1000]
    rows[2]['evidence'].pop('text_xy_mm');rows[2]['evidence'].pop('text_raw')
    page.evaluate("records=>window.dispatchEvent(new CustomEvent('module-h-evidence-state',{detail:{records,pending:0}}))",rows)
    page.locator('#he-targets button[data-id=p1]').click()
    page.wait_for_function("__picked[0]==='p1' && __paint.links.length===0")
    assert page.evaluate('__paint.links')==[]


def test_multi_selection_and_live_batches_remain_uncluttered(h_ui):
    page,sess,_,_=h_ui
    rows=install_evidence(page,sess)
    page.uncheck('#ha-show')
    page.locator('#he-targets button[data-id=p0]').click()
    page.wait_for_function('__paint.links.length===1')
    page.evaluate("window.dispatchEvent(new CustomEvent('module-h-attributes-selection',{detail:{ids:['p0','p1']}}))")
    page.wait_for_function('__paint.links.length===2')
    snapshot=[dict(edge=r['edges'][0],xy=r['xy'][0],evidence=r['evidence']) for r in rows[1:]]
    page.evaluate("snapshot=>window.dispatchEvent(new CustomEvent('module-h-diameter-progress',{detail:{phase:'반복 구조 대응',data:{reset:true,snapshot}}}))",snapshot)
    page.wait_for_function("document.querySelector('#he-status').textContent.includes('40구간')")
    page.wait_for_function('__paint.links.length===0')
    assert page.locator('#he-targets button').count()==40
    page.evaluate("window.dispatchEvent(new CustomEvent('module-h-attributes-selection',{detail:{ids:['unresolved-pipe']}}))")
    page.wait_for_function('__paint.labels.length===0 && __paint.links.length===0')
