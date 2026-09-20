"""Browser regressions for Module F camera retention and rectangle zoom.

Runs the real template/JavaScript with isolated HTTP fixtures; no live server
or user drawing is changed. Test hooks are injected into this browser only.
"""
from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import urlparse

import pytest

pw_api = pytest.importorskip("playwright.sync_api")
ROOT = Path(__file__).resolve().parents[1]
BOUNDS = dict(minx=10000, miny=20000, maxx=110000, maxy=100000)
WORLD = dict(bounds=BOUNDS, bundles=[dict(id="pipe", css="#ffffff", layer="pipe", name="pipe",
    color=7, segs=[10000, 20000, 110000, 100000, 10000, 100000, 110000, 20000],
    circles=[60000, 60000, 10000], arcs=[60000, 60000, 20000, 20, 140])])
EDIT = dict(bounds=BOUNDS, body_groups=[], wet_pipes=[], heads=[], multi_heads=[],
            sources=[], valves=[], palette={}, kinds={}, counts=dict(pts=0, edges=0, heads=0, bodies=0))


def install_page(page, source: str | None = None) -> list[str]:
    """Serve repository UI locally, adding test-only access to rendering calls."""
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    source = source or (ROOT / "static/module_f.js").read_text(encoding="utf-8")
    hook = """window.__viewTest = {setStage, draw, fitDesignView, resetSlotViewport,
        paint: typeof paint === 'function' ? paint : draw};\n"""
    # Baseline builds do not yet have the reset helper.
    if "function resetSlotViewport" not in source:
        hook = hook.replace("resetSlotViewport,", "")
    source = source.replace('  setStage("open");\n  loadSaved();', hook + '  setStage("open");\n  loadSaved();')

    def route(req):
        path = urlparse(req.request.url).path
        if path == "/module-f":
            req.fulfill(body=(ROOT / "templates/module_f.html").read_text(encoding="utf-8"), content_type="text/html")
        elif path == "/static/module_f.js":
            req.fulfill(body=source, content_type="application/javascript")
        elif path.startswith("/static/"):
            file = (ROOT / path.lstrip("/")).resolve()
            if not file.is_relative_to(ROOT / "static") or not file.is_file():
                req.fulfill(status=404)
            else:
                req.fulfill(path=str(file))
        else:
            req.fulfill(json={"ok": True, "items": [], "rows": [], "view": None, "fields": {}})

    page.route("http://module-f.test/**", route)
    page.goto("http://module-f.test/module-f")
    page.wait_for_function("!!window.__viewTest")
    return errors


def settle(page) -> None:
    page.evaluate("() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))")


@pytest.fixture()
def ui():
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        errors = install_page(page)
        page.evaluate("""({world, edit}) => {
            Object.assign(__mf, {sid:'viewport-test', slot:'plan', key:'drawing',
                method:'manual', world, edit, design: {view:null, marks:{}}});
            __viewTest.setStage('open');
        }""", {"world": WORLD, "edit": EDIT})
        settle(page)
        yield page
        assert not errors, errors
        browser.close()


def camera(page):
    return page.evaluate("({...__mf.view})")


def rectangle(page, start=(0.35, 0.35), end=(0.65, 0.65), button="#open-zoom-window"):
    page.click(button)
    box = page.locator("#cv").bounding_box()
    page.mouse.move(box["x"] + box["width"] * start[0], box["y"] + box["height"] * start[1])
    page.mouse.down()
    page.mouse.move(box["x"] + box["width"] * end[0], box["y"] + box["height"] * end[1], steps=4)
    page.mouse.up()
    settle(page)


def test_rectangle_zoom_changes_only_camera_and_is_one_shot(ui):
    before = camera(ui)
    data = ui.evaluate("JSON.stringify([__mf.world, __mf.edit, __mf.zones])")
    rectangle(ui)
    after = camera(ui)
    # Native mouse positions are rounded to device pixels by Chromium.
    assert after["scale"] == pytest.approx(before["scale"] * 0.96 / 0.30, rel=0.004)
    assert ui.evaluate("JSON.stringify([__mf.world, __mf.edit, __mf.zones])") == data
    assert ui.get_attribute("#open-zoom-window", "aria-pressed") == "false"
    assert ui.locator("#zoom-window").is_hidden()


def test_camera_retained_across_cad_stages_and_refresh(ui):
    rectangle(ui)
    expected = camera(ui)
    for stage in ("pick", "edit", "design", "conv", "edit", "open", "pick"):
        ui.evaluate("stage => __viewTest.setStage(stage)", stage)
        settle(ui)
        assert camera(ui) == expected, stage
    ui.evaluate("__viewTest.fitDesignView()")
    settle(ui)
    assert camera(ui) == expected


def test_auto_result_keeps_zoom_until_explicit_fit(ui):
    rectangle(ui)
    before=camera(ui)
    ui.evaluate('''()=>{
      __mf.autoDone=true;__mf.autoView={nodes:[{label:'10',x:58000,y:59000},{label:'11',x:62000,y:61000}],
        pipes:[{label:'p',a:'10',b:'11'}]};
      __viewTest.setStage('auto');__viewTest.draw();
    }''')
    settle(ui)
    assert camera(ui)==before
    ui.click('#btn-fit');settle(ui)
    assert camera(ui)['scale']>before['scale']*2


def test_rectangle_does_not_place_a_head_or_design_region(ui):
    ui.evaluate("__viewTest.setStage('edit'); document.querySelector('#ed-zone-arm').checked = true")
    settle(ui)
    rectangle(ui, button="#btn-zoom-window")
    assert ui.evaluate("__mf.zones") == []


def test_escape_and_tiny_rectangle_do_not_zoom(ui):
    before = camera(ui)
    ui.click("#open-zoom-window")
    ui.keyboard.press("Escape")
    assert ui.get_attribute("#open-zoom-window", "aria-pressed") == "false"
    rectangle(ui, end=(0.351, 0.351))
    assert camera(ui) == before


def test_reverse_rectangle_and_release_outside_canvas(ui):
    before = camera(ui)
    rectangle(ui, start=(0.65, 0.65), end=(0.35, 0.35))
    assert camera(ui)["scale"] > before["scale"] * 3
    ui.click("#btn-fit")
    settle(ui)
    rectangle(ui, start=(0.60, 0.60), end=(-0.05, 0.10))
    assert ui.locator("#zoom-window").is_hidden()
    assert camera(ui)["scale"] > before["scale"]


def test_drawing_slots_keep_independent_cameras(ui):
    rectangle(ui)
    plan = camera(ui)
    ui.evaluate("__mf.slot='system'; __mf.key='system'; __viewTest.setStage('open')")
    settle(ui)
    assert camera(ui) != plan
    rectangle(ui, start=(0.1, 0.1), end=(0.8, 0.8))
    system = camera(ui)
    ui.evaluate("__mf.slot='plan'; __mf.key='drawing'; __viewTest.setStage('open')")
    settle(ui)
    assert camera(ui) == plan
    ui.evaluate("__mf.slot='system'; __mf.key='system'; __viewTest.setStage('open')")
    settle(ui)
    assert camera(ui) == system


def test_projection_camera_returns_to_cad_zoom(ui):
    rectangle(ui)
    cad = camera(ui)
    # No nodes rendered: use explicit bounds through a tiny network of nodes.
    ui.evaluate("""() => {
        __mf.design.view = {nodes:[{x:0,y:0,label:'1'},{x:3000,y:3000,label:'2'}], pipes:[],nozzles:[]};
        document.querySelector('#dg-plan').checked=false;
        __viewTest.setStage('design');
    }""")
    settle(ui)
    iso = camera(ui)
    assert iso["scale"] > cad["scale"]
    ui.evaluate("__viewTest.setStage('edit')")
    settle(ui)
    assert camera(ui) == cad
    ui.evaluate("__viewTest.setStage('design')")
    settle(ui)
    assert camera(ui) == iso


def test_wheel_keeps_pointer_anchor_and_resize_keeps_center(ui):
    box = ui.locator("#cv").bounding_box()
    x, y = int(box["width"] * 0.3), int(box["height"] * 0.4)
    world = ui.evaluate("p => [__mf.toWorldX(p[0]), __mf.toWorldY(p[1])]", [x, y])
    ui.mouse.move(box["x"] + x, box["y"] + y)
    ui.mouse.wheel(0, -100)
    settle(ui)
    after = ui.evaluate("p => [__mf.toWorldX(p[0]), __mf.toWorldY(p[1])]", [x, y])
    assert after == pytest.approx(world)
    center = ui.evaluate("() => {const r=document.querySelector('#cv').getBoundingClientRect(); return [__mf.toWorldX(r.width/2), __mf.toWorldY(r.height/2)]}")
    scale = camera(ui)["scale"]
    ui.set_viewport_size({"width": 1600, "height": 1000})
    settle(ui)
    resized = ui.evaluate("() => {const r=document.querySelector('#cv').getBoundingClientRect(); return [__mf.toWorldX(r.width/2), __mf.toWorldY(r.height/2)]}")
    assert camera(ui)["scale"] == scale
    assert resized == pytest.approx(center)


def test_cached_lines_crossing_viewport_are_visible_and_geometry_refreshes(ui):
    ui.evaluate("""() => {
        __mf.world.bundles = [{css:'#ffffff', id:'cross', circles:[], arcs:[],
            segs:[-100000,60000,200000,60000]}];
        __viewTest.draw();
    }""")
    rectangle(ui)
    # Both endpoints lie outside the zoomed view: culling must retain the line.
    lit = ui.evaluate("""() => {
        const cv=document.querySelector('#cv'), c=cv.getContext('2d');
        const x=Math.round(__mf.toScreenX(60000)), y=Math.round(__mf.toScreenY(60000));
        return Math.max(...c.getImageData(x-2,y-2,5,5).data.filter((_, i)=>i%4!==3));
    }""")
    assert lit > 200
    ui.evaluate("__mf.world.bundles[0].segs=[]; __viewTest.draw()")
    settle(ui)
    assert ui.evaluate("""() => {
        const c=document.querySelector('#cv').getContext('2d');
        const x=Math.round(__mf.toScreenX(60000)), y=Math.round(__mf.toScreenY(60000));
        return Math.max(...c.getImageData(x-2,y-2,5,5).data.filter((_, i)=>i%4!==3));
    }""") == 0


def test_new_upload_with_same_name_initializes_a_new_view(ui):
    original = camera(ui)
    rectangle(ui)
    ui.evaluate("__viewTest.resetSlotViewport(); __viewTest.draw()")
    settle(ui)
    assert camera(ui) == original


def test_escape_during_drag_does_not_click_the_drawing(ui):
    ui.evaluate("__viewTest.setStage('pick')")
    settle(ui)
    requests = []
    ui.on('request', lambda req: requests.append(req.url) if req.method == 'POST' else None)
    before = camera(ui)
    ui.click('#btn-zoom-window')
    box = ui.locator('#cv').bounding_box()
    ui.mouse.move(box['x']+100,box['y']+100)
    ui.mouse.down()
    ui.mouse.move(box['x']+200,box['y']+200)
    ui.keyboard.press('Escape')
    ui.mouse.up()
    settle(ui)
    assert not requests
    assert camera(ui) == before
