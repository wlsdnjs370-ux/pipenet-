# -*- coding: utf-8 -*-
"""PR-8c C4 헤드 배치 UI 런타임 검증 — 일회용.

보고 싶은 것은 넷이다. **구역 없이는 배치 버튼이 열리지 않는가**, **구역에 못 든
실이 화면에서 세어지는가**(0 으로 접히면 빠진 실을 아무도 못 본다), **미검증이
통과색으로 칠해지지 않는가**, **연면적을 넣으면 미검증이 실제로 풀리는가**.

세 번째가 핵심이다. 재지 못한 검사가 초록으로 나오면 검사표 전체가 장식이 된다.
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 5094
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

errors: list[str] = []
fails: list[str] = []


def check(label: str, ok: bool, detail: str = "") -> None:
    print(f"  [{'OK ' if ok else 'FAIL'}] {label}{(' — ' + detail) if detail else ''}")
    if not ok:
        fails.append(label)


# 후보 코어 한가운데의 클라이언트 좌표. 밸브는 실제 클릭으로 찍는다.
AIM = """(i) => {
  const cand = state.c3.candidates[i];
  if (!cand || !cand.center) return null;
  const c = toDrawingUnits([cand.center])[0];
  const rect = canvas.getBoundingClientRect();
  state.view.panX = rect.width / 2 - c[0] * state.view.zoom;
  state.view.panY = rect.height / 2 + c[1] * state.view.zoom;
  render();
  const x = rect.left + rect.width / 2, y = rect.top + rect.height / 2;
  const [wx, wy] = screenToWorld(x - rect.left, y - rect.top);
  const hit = candidateAt(wx, wy);
  return { x, y, hits: hit ? hit.core_id : null, core: cand.core_id,
           onCanvas: document.elementFromPoint(x, y) === canvas };
}"""

C4_STATE = """() => ({
  rooms: state.c4.rooms.length,
  heads: headCount(),
  skipped: state.c4.skipped.length,
  blocking: state.c4.blocking,
  flags: state.c4.flags.map((f) => f.code),
  checks: (state.c4.checks && state.c4.checks.checks || [])
            .map((r) => [r.code, r.status]),
  underObstacle: state.c4.rooms.reduce((n, r) => n
      + (r.heads || []).filter((h) => h.provenance === 'under_obstacle').length, 0),
  stage: state.stage,
  c3Flags: state.c3.flags.map((f) => f.code),
})"""

# 화면이 스스로 그린 단계는 증거가 아니다 — 세션에 실제로 무엇이 적혔는지 묻는다.
SERVER_STAGE = """async () => {
  const res = await fetch(`/api/design/session/${state.sessionId}`);
  return (await res.json()).meta.stage;
}"""

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
    check("구역 전에는 [헤드 배치 실행] 잠김", not page.is_enabled("#dw-run-c4"))
    check("배치 전에는 [검사표 보기] 잠김", not page.is_enabled("#dw-c4-open"))
    check("헤드 오버레이 항목은 0 개로 뜸",
          page.evaluate("() => headCount()") == 0)

    # ── C1 → GATE → C2 → C3 (앞 PR 에서 검증한 경로. 여기선 통과만 시킨다) ──
    page.set_input_files("#dw-dxf", str(DXF))
    page.wait_for_selector("#dw-layer-list .layer-row", timeout=120_000)
    page.select_option("#dw-wall-layers", "ARCHI")
    page.click("#dw-recognize-btn")
    page.wait_for_function(
        "() => !document.getElementById('dw-gate-open').disabled"
        " || /오류/.test(document.getElementById('dw-recognize-status').textContent)",
        timeout=180_000)
    print("인식:", page.inner_text("#dw-recognize-status"))

    page.click("#dw-gate-open")
    page.wait_for_selector("#dw-gate-tbody tr")

    def bulk(field_label, value, *, by_label=True):
        page.select_option("#dw-bulk-field", label=field_label)
        if by_label:
            page.select_option("#dw-bulk-value", label=value)
        else:
            page.fill("#dw-bulk-value", value)
        page.select_option("#dw-bulk-scope", "all")
        page.click("#dw-bulk-apply")

    bulk("용도", "공동주택")
    bulk("반자 유무", "있음")
    bulk("반자고(mm)", "2300", by_label=False)
    bulk("천장고(슬래브, mm)", "2800", by_label=False)

    floors = "#dw-gate-facts [data-field='building.floors_total']"
    page.click(floors)
    page.fill(floors, "15")
    page.press(floors, "Tab")
    page.select_option("#dw-gate-facts [data-field='building.structure']", label="내화구조")
    page.select_option("#dw-gate-facts [data-field='building.use']", label="공동주택")
    page.select_option("#dw-gate-facts [data-field='obstacles.status']", label="partial")
    cores = page.locator("#dw-gate-facts [data-field='confirmed']")
    for i in range(cores.count()):
        cores.nth(i).select_option(label="입상관으로 쓴다")
    page.fill("#dw-gate-operator", "jinwon")
    page.click("#dw-gate-confirm")
    page.wait_for_function(
        "() => !document.getElementById('dw-gate-panel').classList.contains('open')"
        " || /오류|남았|확정자/.test(document.getElementById('dw-gate-msg').textContent)",
        timeout=120_000)
    page.wait_for_timeout(300)
    check("게이트 통과", not page.is_visible("#dw-gate-panel"),
          page.inner_text("#dw-gate-msg"))

    page.click("#dw-run-c2")
    page.wait_for_function(
        "() => !/실행 중/.test(document.getElementById('dw-c2-status').textContent)",
        timeout=60_000)
    page.click("#dw-c2-close")
    check("C2 뒤에도 [헤드 배치 실행] 은 잠김", not page.is_enabled("#dw-run-c4"),
          "구역이 없으면 헤드에 물을 보낼 배관을 그릴 수 없다")

    page.click("#dw-c3-load")
    page.wait_for_function(
        "() => !/부르는 중/.test(document.getElementById('dw-c3-status').textContent)",
        timeout=60_000)
    reqs = page.evaluate("() => state.c3.requirements.map((r) => r.key)")
    systems = page.evaluate("() => state.c3.systemTypes")
    page.click("#dw-tool-valve")
    aim = page.evaluate(AIM, 0)
    check("후보 코어 조준", bool(aim and aim["hits"] == aim["core"] and aim["onCanvas"]),
          json.dumps(aim, ensure_ascii=False))
    page.mouse.click(aim["x"], aim["y"])
    page.wait_for_timeout(150)

    page.click("#dw-c3-open")
    page.wait_for_selector("#dw-c3-tbody tr")
    page.select_option("#dw-c3-tbody select[data-c3='system'][data-i='0']",
                       "습식" if "습식" in systems else systems[0])
    for key in reqs:
        page.select_option(
            f"#dw-c3-tbody select[data-c3='req'][data-i='0'][data-key='{key}']", "true")
    page.fill("#dw-c3-operator", "jinwon")
    page.click("#dw-c3-submit")
    page.wait_for_function(
        "() => !/확정 중/.test(document.getElementById('dw-c3-msg').textContent)",
        timeout=60_000)
    page.click("#dw-c3-zones")
    page.wait_for_function(
        "() => !/나누는 중/.test(document.getElementById('dw-c3-msg').textContent)",
        timeout=120_000)
    print("구역:", page.inner_text("#dw-c3-msg"))
    page.click("#dw-c3-close")
    zoned = page.evaluate("() => state.c3.zones.length")
    check("구역이 생김", zoned >= 1, f"{zoned}개")
    check("구역이 생기면 [헤드 배치 실행] 열림", page.is_enabled("#dw-run-c4"))
    check("아직 [검사표 보기] 는 잠김", not page.is_enabled("#dw-c4-open"))
    print("  C4 상태:", page.inner_text("#dw-c4-status"))

    # ── [1] 헤드 배치 — 연면적 없이 ─────────────────────────────────────────
    print("\n[1] 헤드 배치 (연면적 미입력)")
    t0 = time.time()
    page.click("#dw-run-c4")
    page.wait_for_function(
        "() => !/배치하는 중/.test(document.getElementById('dw-c4-status').textContent)",
        timeout=300_000)
    print(f"  {time.time() - t0:.1f}s — {page.inner_text('#dw-c4-status')}")
    c4 = page.evaluate(C4_STATE)
    print("  상태:", json.dumps(c4, ensure_ascii=False))
    check("헤드가 놓임", c4["heads"] > 0, f"{c4['heads']}개")
    check("배치한 실이 있음", c4["rooms"] > 0, f"{c4['rooms']}개")
    check("[검사표 보기] 열림", page.is_enabled("#dw-c4-open"))
    check("검사 5종이 전부 돔", len(c4["checks"]) >= 5,
          ", ".join(f"{c}={s}" for c, s in c4["checks"]))
    check("검사 상태는 셋 중 하나",
          all(s in ("pass", "flag", "unverified") for _c, s in c4["checks"]))
    by_code = dict(c4["checks"])
    check("연면적이 없으면 면적 대조는 미검증",
          by_code.get("ROOM_AREA_MISMATCH") == "unverified",
          str(by_code.get("ROOM_AREA_MISMATCH")))
    check("배관 총연장은 지어내지 않음",
          by_code.get("PIPE_LENGTH_PER_HEAD") == "unverified",
          str(by_code.get("PIPE_LENGTH_PER_HEAD")))
    # 실도면에는 문이 안 잡힌 실이 늘 남아 C3 플래그가 붙는다. 그때 배치는 하되
    # 단계는 c3 에 서야 한다 — 여기서 c5 로 밀면 닿지 않는 실이 그대로 딸려 간다.
    held = c4["blocking"] or c4["c3Flags"]
    server_stage = page.evaluate(SERVER_STAGE)
    check("화면 단계와 서버 단계가 같음", c4["stage"] == server_stage,
          f"화면={c4['stage']} 서버={server_stage}")
    if held:
        check("남은 문제가 있으면 C5 로 안 넘어감", server_stage != "c5",
              f"{server_stage} ← {', '.join(held)}")
    else:
        check("막는 문제가 없으면 C5 로 넘어감", server_stage == "c5", server_stage)

    # ── [2] 검사표 — 미검증은 통과색이 아니다 ───────────────────────────────
    print("\n[2] 검사표 패널")
    page.click("#dw-c4-open")
    page.wait_for_selector("#dw-c4-panel.open")
    cards = page.evaluate(
        "() => [...document.querySelectorAll('#dw-c4-checks .chk')]"
        ".map((el) => ({ st: el.querySelector('.st').className,"
        " label: el.querySelector('.st').textContent,"
        " code: el.querySelector('.code').textContent }))")
    for card in cards:
        print(f"    {card['code']:<26} {card['label']}")
    check("검사 카드가 검사 수만큼 렌더", len(cards) == len(c4["checks"]),
          f"{len(cards)}장")
    unver = [c for c in cards if "unverified" in c["st"]]
    check("미검증 카드가 있음", bool(unver), f"{len(unver)}장")
    check("미검증에 통과색을 쓰지 않음",
          all("pass" not in c["st"] for c in unver))
    check("미검증 라벨이 '미검증'",
          all(c["label"] == "미검증" for c in unver),
          ", ".join(c["label"] for c in unver))
    rows = page.locator("#dw-c4-tbody tr").count()
    check("실별 배치표가 렌더", rows == c4["rooms"], f"{rows}행")
    if c4["skipped"]:
        note = page.inner_text("#dw-c4-empty")
        print("  미배치:", note)
        check("구역에 못 든 실을 세어서 보임", "헤드를 놓지 않은 실" in note, note)
    check("헤드 수 배지", f"헤드 {c4['heads']}개" in page.inner_text("#dw-c4-remain")
          or c4["blocking"], page.inner_text("#dw-c4-remain"))
    page.screenshot(path=str(ROOT / "data" / "_dw_c4_checks.png"))
    page.click("#dw-c4-close")

    # ── [3] 연면적을 주면 미검증이 풀린다 ───────────────────────────────────
    print("\n[3] 연면적 입력 후 재배치")
    area = page.evaluate(
        "() => (state.c4.checks.checks.find((r) => r.code === 'ROOM_AREA_MISMATCH')"
        " || {}).room_area_m2")
    check("실 면적 합이 미검증에도 기록됨", bool(area), str(area))
    page.fill("#dw-c4-gross", str(area))
    page.click("#dw-run-c4")
    page.wait_for_function(
        "() => !/배치하는 중/.test(document.getElementById('dw-c4-status').textContent)",
        timeout=300_000)
    after = dict(page.evaluate(C4_STATE)["checks"])
    print("  검사:", json.dumps(after, ensure_ascii=False))
    check("연면적을 주면 면적 대조가 미검증에서 풀림",
          after.get("ROOM_AREA_MISMATCH") in ("pass", "flag"),
          str(after.get("ROOM_AREA_MISMATCH")))
    check("연면적을 줘도 배관 총연장은 미검증",
          after.get("PIPE_LENGTH_PER_HEAD") == "unverified")

    # ── [4] 화면 — 헤드가 실제로 그려지는가 ─────────────────────────────────
    print("\n[4] 헤드 오버레이")
    overlay = page.evaluate(
        "() => { const el = [...document.querySelectorAll('#dw-design-list input')]"
        ".find((i) => i.dataset.overlay === 'heads');"
        " return el ? { checked: el.checked, disabled: el.disabled,"
        " count: el.closest('label').querySelector('.count').textContent } : null; }")
    print("  산출물 항목:", json.dumps(overlay, ensure_ascii=False))
    check("헤드 항목이 산출물 목록에 있음", overlay is not None)
    check("헤드 항목이 켜져 있고 잠기지 않음",
          bool(overlay and overlay["checked"] and not overlay["disabled"]))
    check("헤드 개수가 목록에 나옴",
          bool(overlay) and overlay["count"] == str(c4["heads"]),
          f"{overlay['count'] if overlay else '?'} vs {c4['heads']}")
    print("  살수장애 아래 헤드:", c4["underObstacle"], "개")
    page.evaluate("() => { state.view.zoom *= 5; render(); }")
    aim2 = page.evaluate(AIM, 0)
    page.screenshot(path=str(ROOT / "data" / "_dw_c4_heads.png"),
                    clip={"x": aim2["x"] - 400, "y": aim2["y"] - 280,
                          "width": 800, "height": 560})

    # ── [5] 밸브를 건드리면 배치는 근거를 잃는다 ────────────────────────────
    print("\n[5] 밸브를 고치면 배치가 무효")
    page.click("#dw-c4-open")
    page.wait_for_selector("#dw-c4-panel.open")
    page.click("#dw-c4-close")
    page.click("#dw-c3-open")
    page.wait_for_selector("#dw-c3-tbody tr")
    page.select_option(
        f"#dw-c3-tbody select[data-c3='req'][data-i='0'][data-key='{reqs[-1]}']", "")
    page.click("#dw-c3-close")
    reset = page.evaluate(
        "() => ({ heads: headCount(), ran: state.c4.ran, zones: state.c3.zones.length })")
    check("헤드 배치 무효화", reset["heads"] == 0 and not reset["ran"],
          json.dumps(reset))
    check("[헤드 배치 실행] 다시 잠김", not page.is_enabled("#dw-run-c4"))
    check("[검사표 보기] 다시 잠김", not page.is_enabled("#dw-c4-open"))

    browser.close()

network = [e for e in errors if "Failed to load resource" in e]
js = [e for e in errors if e not in network]
print("\n네트워크 로그(예상됨):", network or "없음")
print("JS 오류:", js or "없음")
print("실패한 검사:", fails or "없음")
sys.exit(1 if (js or fails) else 0)
