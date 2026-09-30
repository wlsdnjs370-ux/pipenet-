"""The compact button and inspector select/edit the actual full canonical run."""
from test_module_f_bore_provenance_browser import bore_ui
from test_module_f_continuous_runs import design_chain


def test_compact_button_selection_evidence_edit_and_undo(bore_ui):
    page,sess=bore_ui
    sess['design']=design_chain()
    page.evaluate("async()=>{await __boreUi.preview();__boreUi.select('pipe','P1');}")
    page.wait_for_function("document.querySelector('#ne-count-note').textContent.includes('노드 6')")
    page.on('dialog',lambda dialog:dialog.accept())
    page.click('#ne-compact')
    page.wait_for_function("__mf.design.view.pipes.length===2 && document.querySelector('#busy').classList.contains('hidden')")
    page.evaluate("__boreUi.select('pipe','P0')")
    card=page.locator('.bore-card')
    assert '원본 4개 구간 → 배관 1개' in card.inner_text()
    assert '검토용 미확정 상태 유지' in card.inner_text()
    assert card.get_attribute('data-bore-category')=='review'
    page.click('.bore-segments summary')
    assert 'P1 · 2 → 3' in card.inner_text()
    assert 'P3 · 4 → 5' in card.inner_text()
    shown=page.evaluate("__mf.design.view.pipes.find(p=>p.label==='P0')")
    assert shown['a']=='1' and shown['b']=='5'
    page.click('.bore-edit')
    page.wait_for_function("document.querySelector('#ne-op').value==='pipe'")
    page.select_option('#ne-material','CPVC2')
    page.select_option('#ne-dn','65')
    page.fill('#ne-note','통합 구간 전체 수정')
    page.click('#ne-apply')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && __mf.design.view.pipes.find(p=>p.label==='P0').bore_info.manual?.note==='통합 구간 전체 수정'")
    row=sess['design']['tables'].pipes[0]
    assert row['in']=='1' and row['out']=='5' and row['length']==4 and row['type']=='CPVC2'
    page.click('#ne-undo')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && !__mf.design.view.pipes.find(p=>p.label==='P0').bore_info.manual")
    page.click('#ne-undo')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && __mf.design.view.pipes.length===5")
    page.click('#ne-redo')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && __mf.design.view.pipes.length===2")
