# -*- coding: utf-8 -*-
"""PR-5b 실 편집 런타임 검증 — 일회용.

편집은 캔버스 위 손짓으로만 시작된다. 그래서 함수가 있는지가 아니라 **진짜 마우스로
찍어 골라지는지**, 그리고 고친 폴리곤이 서버를 거쳐 되돌아오는지를 본다. 마지막에
`building.json` 의 `gate.edits` 를 직접 읽어 기록까지 확인한다 — 화면만 바뀌고 기록이
없으면 C1 개선의 학습 데이터가 통째로 빈다.
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 5097
DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else (
    ROOT / "data" / "reference_library" / "2. 고가수조_양주옥정 중상1블럭"
    / "CAD" / "XR" / "XR-단위세대 평면도 (공동주택).dxf")

os.environ["DESIGN_WORKBENCH_ENABLED"] = "1"
os.environ["LOGIN_PASSWORD"] = "dwverify"
sys.path.insert(0, str(ROOT))

spec = importlib.util.spec_from_file_location("daejo_server", ROOT / "대조 서버.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

threading.Thread(
    target=lambda: mod.app.run(host="127.0.0.1", port=PORT, threaded=True),
    daemon=True,
).start()

from playwright.sync_api import sync_playwright  # noqa: E402

from core.design import room_edit as RE  # noqa: E402

STATE: dict = {"rooms": {}, "unit": 1.0, "pick": {}}


def inside(poly, x, y) -> bool:
    n = len(poly)
    hit = False
    for i in range(n):
        xi, yi = poly[i]
        xj, yj = poly[(i - 1) % n]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
            hit = not hit
    return hit


def poly_of(room):
    unit = STATE["unit"]
    return [(x / unit, y / unit) for x, y in room["polygon"]]


def room_at(x, y):
    """화면의 roomAt 과 같은 규칙 — 담고 있는 실 중 면적이 가장 작은 것."""
    rooms = STATE["rooms"]
    best = None
    for rid, room in rooms.items():
        poly = poly_of(room)
        if len(poly) < 3 or not inside(poly, x, y):
            continue
        if best is None or (room.get("area_m2") or 0) < (rooms[best].get("area_m2") or 0):
            best = rid
    return best


def interior(rid):
    """그 실이 **찍히는** 한 점(도면 단위).

    화면은 겹친 실 중 작은 쪽을 고른다. 무게중심이 실 안이라는 것만으로는 모자라고,
    그 점을 눌렀을 때 화면이 같은 실을 고르는지까지 맞아야 클릭 검증이 성립한다.
    """
    if rid in STATE["pick"]:
        return STATE["pick"][rid]
    poly = poly_of(STATE["rooms"][rid])
    cx = sum(p[0] for p in poly) / len(poly)
    cy = sum(p[1] for p in poly) / len(poly)
    found = None
    for mx, my in [(cx, cy)] + [((p[0] + cx) / 2, (p[1] + cy) / 2) for p in poly]:
        if inside(poly, mx, my) and room_at(mx, my) == rid:
            found = (mx, my)
            break
    STATE["pick"][rid] = found
    return found


def edges_of(room) -> set:
    """서버와 같은 0.1mm 눈금의 무향 변 — 맞닿았는지 판정용."""
    poly = room["polygon"]
    out = set()
    for i in range(len(poly)):
        a = (round(poly[i][0] * 10), round(poly[i][1] * 10))
        b = (round(poly[(i + 1) % len(poly)][0] * 10),
             round(poly[(i + 1) % len(poly)][1] * 10))
        out.add((a, b) if a < b else (b, a))
    return out


errors: list[str] = []
with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={"width": 1600, "height": 950})
    page.on("console", lambda m: errors.append(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))

    base = f"http://127.0.0.1:{PORT}"
    page.goto(f"{base}/login", wait_until="domcontentloaded")
    page.fill("input[type=password]", "dwverify")
    page.press("input[type=password]", "Enter")
    page.wait_for_load_state("networkidle")

    page.goto(f"{base}/design-workbench", wait_until="networkidle")
    print("인식 전 도구:", "선택", page.is_enabled("#dw-tool-select"),
          "/ 자르기", page.is_enabled("#dw-tool-split"),
          "/", page.inner_text("#dw-edit-note"))

    page.set_input_files("#dw-dxf", str(DXF))
    page.wait_for_selector("#dw-layer-list .layer-row", timeout=120_000)
    page.select_option("#dw-wall-layers", "ARCHI")
    page.click("#dw-recognize-btn")
    page.wait_for_function(
        "() => !document.getElementById('dw-gate-open').disabled"
        " || /오류/.test(document.getElementById('dw-recognize-status').textContent)",
        timeout=180_000)
    print("인식:", page.inner_text("#dw-recognize-status"))
    session_id = page.inner_text("#dw-session").replace("session", "").strip()
    root = mod.DESIGN_SESSION_DIR / session_id
    print("인식 후 도구:", "선택", page.is_enabled("#dw-tool-select"),
          "/ 자르기", page.is_enabled("#dw-tool-split"),
          "/ 합치기", page.is_enabled("#dw-merge-btn"),
          "/ 지우기", page.is_enabled("#dw-delete-btn"))

    # ── 화면 변환 역산 ────────────────────────────────────────────────
    # 캔버스의 뷰 변환은 스크립트 스코프 안이라 읽을 수 없다. 커서 표시(월드 좌표)
    # 를 두 지점에서 읽어 화면이 실제로 쓰는 변환을 되풀어 쓴다.
    cv = page.locator("#dw-canvas").bounding_box()

    def world_at(sx, sy):
        page.mouse.move(cv["x"] + sx, cv["y"] + sy)
        page.wait_for_timeout(20)
        x, y = page.inner_text("#dw-overlay-cursor").strip("( )").split(",")
        return float(x), float(y)

    T: dict = {}

    def calibrate():
        ax, ay = world_at(200, 200)
        bx, by = world_at(1000, 700)
        T.update(ax=ax, ay=ay, kx=(bx - ax) / 800.0, ky=(by - ay) / 500.0)

    def to_screen(wx, wy):
        return (cv["x"] + 200 + (wx - T["ax"]) / T["kx"],
                cv["y"] + 200 + (wy - T["ay"]) / T["ky"])

    def gap_px(a, b):
        sa, sb = to_screen(*a), to_screen(*b)
        return ((sa[0] - sb[0]) ** 2 + (sa[1] - sb[1]) ** 2) ** 0.5

    def zoom_pair(a, b, want=40.0, limit=24):
        """실이 몇 px 밖에 안 되면 클릭 좌표의 반올림이 옆 실을 집는다. 사람이
        하듯 두 점이 충분히 떨어져 보일 때까지 확대하고 변환을 다시 잰다."""
        mid = ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)
        for _ in range(limit):
            if gap_px(a, b) >= want:
                break
            page.mouse.move(*to_screen(*mid))
            page.mouse.wheel(0, -200)
            page.wait_for_timeout(40)
            calibrate()

    calibrate()

    def reload_rooms():
        building = json.loads((root / "building.json").read_text(encoding="utf-8"))
        STATE["rooms"] = {r["id"]: r for r in building["rooms"]}
        STATE["pick"] = {}
        return building

    building = reload_rooms()
    # draft 에는 단위가 없다. 실 폴리곤(mm) 폭과 화면에 보이는 도면 폭을 견줘
    # mm/cm/m 중 하나를 고른다 — 후보가 10배씩 떨어져 있어 거친 비율로도 갈린다.
    xs = [p[0] for r in building["rooms"] for p in (r.get("polygon") or [])]
    ys = [p[1] for r in building["rooms"] for p in (r.get("polygon") or [])]
    mm_span = max(max(xs) - min(xs), max(ys) - min(ys))
    lo = world_at(30, 30)
    hi = world_at(cv["width"] - 30, cv["height"] - 30)
    world_span = max(abs(hi[0] - lo[0]), abs(hi[1] - lo[1]))
    STATE["unit"] = min((1.0, 10.0, 1000.0),
                        key=lambda u: abs(mm_span / u / world_span - 1.0))
    print("실", len(STATE["rooms"]), "개 / 폭(mm)", round(mm_span),
          "/ 화면 폭(도면단위)", round(world_span), "→ unit_to_mm", STATE["unit"])

    def click_room(rid, shift=False):
        pt = interior(rid)
        if shift:
            page.keyboard.down("Shift")
        page.mouse.click(*to_screen(*pt))
        if shift:
            page.keyboard.up("Shift")
        page.wait_for_timeout(80)

    def wait_edit():
        page.wait_for_function(
            "() => !/반영 중/.test(document.getElementById('dw-edit-note').textContent)",
            timeout=60_000)
        return page.inner_text("#dw-edit-note")

    def on_screen(rid):
        pt = interior(rid)
        if not pt:
            return False
        sx, sy = to_screen(*pt)
        return (cv["x"] + 20 <= sx <= cv["x"] + cv["width"] - 20
                and cv["y"] + 20 <= sy <= cv["y"] + cv["height"] - 20)

    def find_pair(mergeable: bool):
        """합쳐지는(또는 확실히 안 합쳐지는) 쌍 중 둘 다 가장 큰 것.

        맞닿았다고 다 합쳐지지는 않는다 — 이웃 실의 꼭짓점이 내 변 위에 얹힌 T 접합은
        서버가 거절한다(실도면 실측: 변 공유 9쌍 중 5쌍만 성공). 화면 배선을 보려는
        검증이므로 쌍 고르기는 서버와 같은 기하로 미리 가려 낸다. 작은 실은 화면에서
        몇 px 이라 클릭이 흔들리므로 큰 쪽을 고른다.
        """
        ids = [rid for rid in STATE["rooms"] if on_screen(rid)]
        area = lambda r: STATE["rooms"][r].get("area_m2") or 0
        best = None
        for i, a in enumerate(ids):
            for b in ids[i + 1:]:
                if best is not None and min(area(a), area(b)) <= min(area(best[0]),
                                                                    area(best[1])):
                    continue
                try:
                    RE.merge_polygons([STATE["rooms"][a]["polygon"],
                                       STATE["rooms"][b]["polygon"]])
                    joined = True
                except RE.RoomEditError:
                    joined = False
                if joined is mergeable:
                    best = (a, b)
        return best

    # ── 선택 + 합치기 ─────────────────────────────────────────────────
    page.click("#dw-tool-select")
    pair = find_pair(True)
    a_pt, b_pt = interior(pair[0]), interior(pair[1])
    zoom_pair(a_pt, b_pt)
    print("맞닿은 실 쌍:", pair,
          [STATE["rooms"][r].get("area_m2") for r in pair],
          "— 확대 후 두 점 거리(px):", round(gap_px(a_pt, b_pt), 1))
    click_room(pair[0])
    print("클릭 1회 →", page.inner_text("#dw-edit-note"),
          "| 지우기", page.is_enabled("#dw-delete-btn"),
          "| 합치기", page.is_enabled("#dw-merge-btn"))
    click_room(pair[1], shift=True)
    print("shift 클릭 →", page.inner_text("#dw-edit-note"),
          "| 합치기", page.is_enabled("#dw-merge-btn"))
    page.screenshot(path=str(ROOT / "data" / "_dw_edit_selected.png"))

    before = len(STATE["rooms"])
    page.click("#dw-merge-btn")
    print("합치기 →", wait_edit())
    reload_rooms()
    print("  실 수:", before, "→", len(STATE["rooms"]))

    # ── 자르기 ────────────────────────────────────────────────────────
    # 실 대부분이 오목해(실측 57 중 51) 아무 데나 그으면 무한직선이 경계를 4번 이상
    # 지나 거절된다. 서버와 같은 기하로 통과하는 가로선을 먼저 찾아 그 선을 긋는다.
    def splittable():
        for rid, room in sorted(STATE["rooms"].items(),
                                key=lambda kv: -(kv[1].get("area_m2") or 0)):
            pt = interior(rid)
            if not pt or not on_screen(rid):
                continue
            y_mm = pt[1] * STATE["unit"]
            try:
                RE.split_polygon(room["polygon"], (-1e9, y_mm), (1e9, y_mm))
            except RE.RoomEditError:
                continue
            return rid, pt
        return None, None

    page.click("#dw-tool-split")
    # 합치기 때 바짝 당겨 둔 화면을 되돌린다 — 후보가 화면 밖이면 클릭할 수 없다.
    page.click("#dw-fit-btn")
    page.wait_for_timeout(120)
    calibrate()
    STATE["pick"] = {}
    target, pt = splittable()
    print("자를 실:", target, STATE["rooms"][target].get("area_m2"), "㎡")
    tpoly = poly_of(STATE["rooms"][target])
    zoom_pair((min(p[0] for p in tpoly), min(p[1] for p in tpoly)),
              (max(p[0] for p in tpoly), max(p[1] for p in tpoly)), want=300.0)

    # 화면이 실제로 읽는 좌표로 다시 확인한다 — px 반올림 때문에 의도한 y 와 몇 mm
    # 어긋나면 교차 횟수가 달라진다.
    sx, sy = to_screen(*pt)
    line_y = None
    for dy in (0, -6, 6, -12, 12, -24, 24):
        wx, wy = world_at(sx - cv["x"], sy + dy - cv["y"])
        try:
            RE.split_polygon(STATE["rooms"][target]["polygon"],
                             (-1e9, wy * STATE["unit"]), (1e9, wy * STATE["unit"]))
        except RE.RoomEditError:
            continue
        line_y = sy + dy
        break
    page.mouse.move(sx - 60, line_y)
    page.mouse.down()
    page.mouse.move(sx, line_y, steps=5)
    page.mouse.move(sx + 60, line_y, steps=5)
    page.mouse.up()
    print("자르기 →", wait_edit())
    before = len(STATE["rooms"])
    reload_rooms()
    print("  실 수:", before, "→", len(STATE["rooms"]))
    page.screenshot(path=str(ROOT / "data" / "_dw_edit_split.png"))

    # ── 맞닿지 않은 실은 거절해야 한다 ────────────────────────────────
    page.click("#dw-tool-select")
    far = find_pair(False)
    print("떨어진 실 쌍:", far)
    click_room(far[0])
    click_room(far[1], shift=True)
    page.click("#dw-merge-btn")
    print("떨어진 실 합치기 →", wait_edit())
    print("  실 수(그대로여야 함):", len(json.loads(
        (root / "building.json").read_text(encoding="utf-8"))["rooms"]))

    # ── 지우기 ────────────────────────────────────────────────────────
    # 앞선 두 실이 골라진 채다. 그냥 누르면 **바꿔치기** 되어야 하고, 그러지 않으면
    # 지우기가 엉뚱한 실까지 지운다 — 실 수로 확인한다.
    reload_rooms()
    victim = next(rid for rid in STATE["rooms"] if on_screen(rid) and interior(rid))
    click_room(victim)
    print("고른 실 바꿔치기 →", page.inner_text("#dw-edit-note"))
    before = len(STATE["rooms"])
    page.click("#dw-delete-btn")
    print("지우기 →", wait_edit())
    building = reload_rooms()
    print("  실 수:", before, "→", len(STATE["rooms"]))

    print("gate.edits:")
    for rec in building["gate"]["edits"]:
        print("  ", json.dumps(rec, ensure_ascii=False)[:200])
    print("결손 진행:", page.inner_text("#dw-gate-progress"))
    page.screenshot(path=str(ROOT / "data" / "_dw_edit_end.png"))
    browser.close()

network = [e for e in errors if "Failed to load resource" in e]
js = [e for e in errors if e not in network]
print("\n네트워크 로그(예상됨):", network or "없음")
print("JS 오류:", js or "없음")
sys.exit(1 if js else 0)
