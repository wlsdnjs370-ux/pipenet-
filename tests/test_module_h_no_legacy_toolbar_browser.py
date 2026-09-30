"""The legacy editor toolbar is removed from H in every loading/stage state."""

from test_module_h_browser import h_ui, seed


def test_legacy_toolbar_stays_hidden_before_ready_and_through_busy_transitions(h_ui):
    page, sess, requests, _ = h_ui
    start = len(requests)
    for stage in ('open', 'pick', 'edit', 'design', 'conv', 'merge'):
        seed(page, sess, stage)
        result = page.evaluate('''()=>{
            const bar=document.querySelector('#ne-toolbar');
            bar.classList.remove('hidden');
            const normal=getComputedStyle(bar).display;
            __hTest.busy(true,'표 생성 중');
            const busy=getComputedStyle(bar).display;
            __hTest.busy(false);
            return [normal,busy];
        }''')
        assert result == ['none', 'none'], stage
    # Shared F controls remain in the DOM, and the override is H-only even
    # before the delayed property panel/class initialisation has run.
    assert page.locator('#ne-toolbar button').count() == 6
    assert page.evaluate('''()=>{
        const body=document.body,bar=document.querySelector('#ne-toolbar');
        const original=body.className;
        try {
            body.className='module-h';
            const early=getComputedStyle(bar).display;
            body.className='';
            return [early,getComputedStyle(bar).display];
        } finally { body.className=original; }
    }''') == ['none', 'flex']
    assert not [r for r in requests[start:] if r[0]=='POST']
