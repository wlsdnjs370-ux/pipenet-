"""Exercise the real freehand UI in an isolated browser, never the live session."""
import json
from pathlib import Path

from test_module_f_viewport_browser import ui, settle


def prepare(page):
    page.evaluate("""() => {
      __mf.world.bounds={minx:0,miny:0,maxx:10000,maxy:10000};
      __mf.edit.bounds=__mf.world.bounds;
      __mf.edit.heads=[[2000,7000,80,'#fff'],[7000,2000,80,'#fff'],[7000,7000,80,'#fff']];
      __mf.edit.sources=[[1500,1500]];
      __mf.edit.mode='알람밸브위치'; __mf.edit.head_kinds=[];
      __viewTest.resetSlotViewport(); __viewTest.setStage('edit');
    }""")
    def route(req):
        data=req.request.post_data_json
        state=page.evaluate("JSON.parse(JSON.stringify(__mf.edit))")
        state["selection_zones"]=data["zones"]
        state["worst"]=None
        req.fulfill(json={"ok":True,"state":state})
    page.route("**/api/module-f/edit/zones",route)
    page.locator("#ed-zone-arm").check()
    page.locator("#ed-zone-shape").select_option("lasso")
    settle(page)


def screen(page, points):
    return page.evaluate("""pts => {const r=document.getElementById('cv').getBoundingClientRect(),v=__mf.view;
      return pts.map(([x,y])=>[r.x+(x-v.ox)*v.scale,r.y+r.height-(y-v.oy)*v.scale]);}""",points)


def stroke(page,points,release=True):
    ps=screen(page,points)
    page.mouse.move(*ps[0]);page.mouse.down()
    for p in ps[1:]: page.mouse.move(*p,steps=8)
    if release: page.mouse.up()
    settle(page)


def test_freehand_concavity_count_camera_and_save_payload(ui):
    prepare(ui)
    calls=[]
    ui.on("request",lambda r:calls.append(r.url))
    pts=[[1000,1000],[8500,1000],[8500,3500],[3500,3500],[3500,8500],[1000,8500],[1000,1000]]
    stroke(ui,pts)
    ui.wait_for_function("__mf.edit.selection_zones?.length === 1")
    zones=ui.evaluate("__mf.zones")
    assert zones[0]["type"]=="polygon"
    assert ui.evaluate("ModuleFRegions.contains(__mf.zones[0],[2000,7000])")
    assert not ui.evaluate("ModuleFRegions.contains(__mf.zones[0],[7000,7000])")
    assert "헤드 2개" in ui.locator("#ed-worst-why").inner_text()
    assert not any("/edit/click" in p for p in calls),"Drawing must not move the alarm valve"
    ui.mouse.wheel(0,-150);settle(ui)
    assert ui.evaluate("__mf.zones")==zones
    saved=[]
    ui.route("**/api/module-f/edit/save",lambda req:(saved.append(req.request.post_data_json),req.fulfill(json={"ok":True,"message":"saved"})))
    ui.locator("#ed-save").click()
    ui.wait_for_timeout(100)
    assert saved[0]["zones"]==zones
    output=Path(__file__).resolve().parents[1]/"data/module_f_lasso_preview.png"
    ui.screenshot(path=str(output))


def test_escape_and_ctrl_z_do_not_edit_network(ui):
    prepare(ui)
    stroke(ui,[[1000,1000],[8000,1000],[5000,8000]],release=False)
    ui.keyboard.press("Escape");ui.mouse.up();settle(ui)
    assert ui.evaluate("__mf.zones")==[]
    stroke(ui,[[1000,1000],[8000,1000],[5000,8000],[1000,1000]])
    ui.wait_for_function("__mf.zones.length === 1")
    ui.locator("#cv").click(position={"x":1,"y":1})
    ui.keyboard.press("Control+z")
    ui.wait_for_function("__mf.zones.length === 0")


def test_rectangle_still_supported_alongside_lasso(ui):
    prepare(ui)
    ui.locator("#ed-zone-shape").select_option("rect")
    stroke(ui,[[1000,1000],[3000,5000]])
    assert len(ui.evaluate("__mf.zones[0]"))==4
    ui.locator("#ed-zone-shape").select_option("lasso")
    stroke(ui,[[5000,5000],[9000,5000],[7000,9000],[5000,5000]])
    zones=ui.evaluate("__mf.zones")
    assert len(zones)==2 and zones[1]["type"]=="polygon"
