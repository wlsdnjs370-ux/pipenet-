"""Actual H controls -> solver -> proposal table -> SDF, without live data."""
from copy import deepcopy
import json
from pathlib import Path
import xml.etree.ElementTree as ET

import pytest
from test_module_h_browser import h_ui
from test_module_h_integration_browser import prepare as integration
from test_module_f_sizing import tables, merged
from routes.module_f.hydraulic_sizing import fingerprint


def prepare(h_ui):
    page, sess, requests, tmp = integration(h_ui)
    sess['merged'] = merged(tables())
    sess.pop('merge_editor', None)
    page.evaluate('async()=>{await __hTest.merged();await moduleFNetworkEditor.refresh();}')
    page.wait_for_selector('#hh-duration', state='attached')
    page.click('#hm-sizing')
    page.click('#sz-load')
    page.wait_for_function("document.querySelector('#hh-duration').value==='20'")
    page.check('#sz-confirm')
    return page, sess, requests, tmp


def test_joint_review_populates_readonly_inner_table_and_exports(h_ui):
    page, sess, requests, tmp = prepare(h_ui)
    before = fingerprint(sess['merged']['combined'])
    page.fill('#sz-branch', '1')
    page.click('#hh-joint')
    page.wait_for_function('ModuleHHydraulics.getReport()?.feasible===true', timeout=20000)
    report=sess['sizing_report']
    assert report['options']['diameter_scope']=='enlarge'
    assert report['water_supply']['required_volume_m3']==pytest.approx(report['pump']['flow_lpm']*.02)
    assert report['pump']['usable'] and all(p['inner_mm']>0 for p in report['pipes'])
    assert 'm³' in page.locator('#hh-result').inner_text()
    assert page.locator('#h-output-basis').input_value()=='proposal'
    page.locator('#hh-result > button').click()
    page.wait_for_function("document.querySelector('#hp-basis').value==='proposal' && document.querySelector('#hp-head').textContent.includes('실제')")
    assert page.locator('#hp-rows input').count()==0
    assert page.locator('#hp-rows tr').count()==5
    assert page.locator('#hp-undo').is_disabled()
    for row,result in zip(page.locator('#hp-rows tr').all(),report['pipes']):
        assert float(row.locator('td').nth(3).inner_text().replace(',',''))==pytest.approx(result['inner_mm'])
    page.screenshot(path=str(tmp/'hydraulic-proposal-table.png'))
    page.click('#hm-output')
    page.click('#mg-emit')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && !document.querySelector('#sz-download').disabled",timeout=20000)
    assert set(sess['sizing_files'])=={'sdf','slf','sdf_iso','slf_iso','report'}
    assert len(list(ET.parse(sess['sizing_files']['sdf']).getroot().iter('Pipe')))==5
    exported=json.loads(Path(sess['sizing_files']['report']).read_text(encoding='utf-8'))
    assert exported['water_supply']==report['water_supply']
    assert before==fingerprint(sess['merged']['combined'])
    page.click('#hm-sizing')
    page.fill('#hh-duration','40')
    assert page.locator('#sz-emit').is_disabled()
    assert '재계산' in page.locator('#hh-result').inner_text()


def test_failed_heads_and_pipes_red_location_and_dirty_clear(h_ui):
    page,sess,requests,tmp=prepare(h_ui)
    page.select_option('#sz-mode','fixed_supply')
    page.fill('#sz-min','9');page.fill('#sz-branch','1')
    page.click('#sz-run')
    page.wait_for_function('ModuleHHydraulics.getReport()?.feasible===false',timeout=20000)
    page.wait_for_selector('#hh-fail-canvas:not([hidden])')
    assert page.locator('#sz-emit').is_disabled()
    assert not sess['sizing_report']['pump']['usable']
    assert sess['sizing_report']['water_supply']['required_volume_m3'] is None
    assert page.locator('.hh-issues button[data-kind=pipe]').count()>0
    assert page.locator('.hh-issues button[data-kind=node]').count()>0
    page.locator('.hh-issues button[data-kind=node]').first.click()
    assert page.evaluate('moduleFNetworkEditor.getSelection().label')=='h'
    page.wait_for_timeout(200)
    assert page.locator('#hh-fail-canvas').evaluate('''c=>{const d=c.getContext('2d').getImageData(0,0,c.width,c.height).data;for(let i=0;i<d.length;i+=4)if(d[i]>200&&d[i+1]<100&&d[i+3]>200)return true;return false;}''')
    page.screenshot(path=str(tmp/'hydraulic-failures-red.png'))
    page.uncheck('#hh-show-fail');page.wait_for_selector('#hh-fail-canvas',state='hidden')
    page.check('#hh-show-fail');page.wait_for_selector('#hh-fail-canvas',state='visible')
    sess['merged']['combined'].pipes[0]['length']+=1
    page.evaluate('__hTest.merged()')
    page.wait_for_selector('#hh-fail-canvas',state='hidden')
    assert page.locator('#sz-emit').is_disabled()
    assert '재계산' in page.locator('#hh-result').inner_text()


def test_cosmetics_and_underlay_same_native_control(h_ui):
    page,sess,requests,tmp=prepare(h_ui)
    assert '관경과 펌프를 함께 검토' not in page.locator('#mg-sizing').inner_text()
    before=fingerprint(sess['merged']['combined'])
    page.click('#hm-connect')
    assert page.locator('.hm-source').first.evaluate('e=>getComputedStyle(e).borderLeftColor===getComputedStyle(e).borderRightColor')
    page.keyboard.press('Escape')
    page.check('#hm-under')
    assert page.locator('#mg-under').is_checked()
    page.uncheck('#hm-under')
    assert not page.locator('#mg-under').is_checked()
    assert before==fingerprint(sess['merged']['combined'])


def test_validation_error_not_hidden_by_compact_result(h_ui):
    page,sess,requests,tmp=prepare(h_ui)
    page.fill('#hh-duration','0');page.click('#hh-joint')
    assert '지속시간' in page.locator('#hh-result').inner_text()
    assert page.locator('#sz-emit').is_disabled()
    assert not sess.get('sizing_report')
