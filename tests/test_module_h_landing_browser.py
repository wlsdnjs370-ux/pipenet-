"""Architecture-inspired H start screen; presentation only, no engine writes."""

import pytest

from test_module_h_browser import h_ui


@pytest.mark.parametrize("size", [(1500, 960), (390, 844), (844, 390)])
def test_architecture_fits_and_upload_remains_keyboard_accessible(h_ui, size):
    page, _, requests, tmp_path = h_ui
    page.set_viewport_size({"width": size[0], "height": size[1]})
    welcome = page.locator("#h-welcome")
    mark = page.locator(".h-conversion-mark")
    button = page.locator("#h-welcome-open")
    assert welcome.is_visible() and mark.is_visible() and button.is_visible()
    assert "폰 노이만" in mark.get_attribute("aria-label")
    assert "f(x)" not in welcome.inner_text()
    assert mark.locator('.h-architecture-label').text_content() == 'Pipe Network'
    assert mark.locator('[data-unit=processor] text').count() == 0
    assert mark.locator('.h-network-graph circle').count() == 9
    assert mark.locator('.h-graph-edges, .h-graph-route').count() == 2
    assert welcome.locator("p, h2").count() == 0
    for element in (mark, button):
        box = element.bounding_box()
        stage = welcome.bounding_box()
        assert box["x"] >= stage["x"] and box["y"] >= stage["y"]
        assert box["x"] + box["width"] <= stage["x"] + stage["width"] + 1
        assert box["y"] + box["height"] <= stage["y"] + stage["height"] + 1
    assert button.bounding_box()["height"] >= 44
    page.screenshot(path=str(tmp_path / f"h-architecture-{size[0]}.png"))
    before = [request for request in requests if request[0] == "POST"]
    button.focus()
    page.keyboard.press("Enter")
    page.wait_for_function('!document.querySelector("#h-access-plan").disabled')
    assert page.locator('#h-access').is_visible()
    assert page.locator('#dxf').is_hidden()
    page.click('#h-access-plan')
    page.wait_for_function('!document.querySelector("#h-drawer").hidden')
    assert page.locator("#dxf").is_visible()
    assert [request for request in requests if request[0] == "POST"] == before
