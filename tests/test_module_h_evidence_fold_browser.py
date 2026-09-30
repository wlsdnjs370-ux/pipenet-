"""Folding a resized evidence table frees the canvas, retaining its geometry."""
import pytest

from test_module_h_browser import h_ui
from test_module_h_evidence_browser import install_evidence


def test_resized_evidence_folds_to_header_then_restores_even_after_drag(h_ui):
    page, sess, requests, tmp = h_ui
    install_evidence(page, sess)
    panel = page.locator('#h-evidence')
    panel.evaluate("n=>{n.style.cssText='left:400px!important;top:150px!important;width:450px!important;height:350px!important;right:auto!important';n.dataset.dragged='true'}")
    handle = panel.locator('[data-edge=se]').bounding_box()
    x, y = handle['x']+3, handle['y']+3
    page.mouse.move(x, y); page.mouse.down()
    page.mouse.move(x+120, y+100, steps=4); page.mouse.up()
    expanded = panel.bounding_box()
    before = page.evaluate('JSON.stringify([__mf.edit,__mf.design])')
    for _ in range(2):
        page.click('#he-fold')
        assert page.locator('#he-fold').get_attribute('aria-expanded') == 'false'
        header = panel.locator('header').bounding_box()
        assert panel.bounding_box()['height'] == pytest.approx(header['height']+2, abs=1)
        assert panel.bounding_box()['height'] < 65
        assert page.locator('#he-status').is_hidden()
        assert panel.locator('[data-edge=s]').is_hidden()
        page.mouse.move(header['x']+100, header['y']+20); page.mouse.down()
        page.mouse.move(header['x']+125, header['y']+45, steps=3); page.mouse.up()
        assert panel.bounding_box()['height'] < 65
        page.click('#he-fold')
        assert panel.bounding_box()['height'] == pytest.approx(expanded['height'], abs=1)
        assert panel.bounding_box()['width'] == pytest.approx(expanded['width'], abs=1)
        assert page.locator('#he-status').is_visible()
    page.click('#he-fold')
    panel.locator('header strong').dblclick()
    assert panel.bounding_box()['height'] < 65
    page.screenshot(path=str(tmp/'evidence-folded.png'))
    assert page.evaluate('JSON.stringify([__mf.edit,__mf.design])') == before
    assert not [p for method, p in requests if method == 'POST']
