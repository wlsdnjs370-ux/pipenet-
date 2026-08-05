# -*- coding: utf-8 -*-
"""PR-7b C3 밸브·구역 UI 런타임 검증 — 일회용.

여기서 보고 싶은 것은 넷이다. **밸브가 자동으로 놓이지 않는가**(후보 밖을 찍으면
끌어붙이지 않는가), **설비 종류와 요건 4가지를 다 답하기 전에 확정 버튼이 열리지
않는가**, **'아니오' 로 답한 요건이 확정 뒤에도 미충족으로 남는가**, **플래그가
남으면 단계가 C4 로 넘어가지 않는가**.

두 번째가 핵심이다. 버튼이 먼저 열리면 안 물어본 요건이 조용히 '아니오' 가 되고,
설비 종류를 비워 둔 밸브가 습식으로 확정된다.
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 5095
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


# 후보 코어를 화면 한가운데로 옮기고 그 지점의 클라이언트 좌표를 돌려준다.
# 뷰만 옮길 뿐 밸브는 실제 클릭 이벤트로 찍는다 — 찍는 경로가 검증 대상이다.
# `hits` 는 그 지점이 어느 후보로 읽히는지, `onCanvas` 는 그 지점이 정말 캔버스인지다.
# 떠 있는 도구막대 위를 찍으면 캔버스는 클릭을 받은 적이 없는데 검사만 통과한다.
AIM = """(i) => {
  const cand = state.c3.candidates[i];
  if (!cand || !cand.center) return null;
  const c = toDrawingUnits([cand.center])[0];
  const rect = canvas.getBoundingClientRect();
  state.view.panX = rect.width / 2 - c[0] * state.view.zoom;
  state.view.panY = rect.height / 2 + c[1] * state.view.zoom;
  render();
  const at = (dx, dy) => {
    const x = rect.left + rect.width / 2 + dx, y = rect.top + rect.height / 2 + dy;
    const [wx, wy] = screenToWorld(x - rect.left, y - rect.top);
    const hit = candidateAt(wx, wy);
    return { x, y, hits: hit ? hit.core_id : null,
             onCanvas: document.elementFromPoint(x, y) === canvas };
  };
  const out = at(0, 0);
  out.core = cand.core_id;
  // 같은 코어 안의 다른 점. 좁은 코어에는 없을 수 있어 `null` 로 남긴다.
  out.nudge = null;
  for (let r = 4; r <= 48 && !out.nudge; r += 4) {
    for (const [dx, dy] of [[r, 0], [0, r], [-r, 0], [0, -r]]) {
      const p = at(dx, dy);
      if (p.hits === cand.core_id && p.onCanvas) { out.nudge = p; break; }
    }
  }
  return out;
}"""

OUTSIDE = """() => {
  const rect = canvas.getBoundingClientRect();
  for (let fy = 0.1; fy <= 0.91; fy += 0.1) {
    for (let fx = 0.1; fx <= 0.91; fx += 0.1) {
      const x = rect.left + rect.width * fx, y = rect.top + rect.height * fy;
      const [wx, wy] = screenToWorld(x - rect.left, y - rect.top);
      if (candidateAt(wx, wy)) continue;
      if (document.elementFromPoint(x, y) !== canvas) continue;
      return { x, y };
    }
  }
  return null;
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
    page.set_input_files("#dw-dxf", str(DXF))
    page.wait_for_selector("#dw-layer-list .layer-row", timeout=120_000)
    page.select_option("#dw-wall-layers", "ARCHI")
    page.click("#dw-recognize-btn")
    page.wait_for_function(
        "() => !document.getElementById('dw-gate-open').disabled"
        " || /오류/.test(document.getElementById('dw-recognize-status').textContent)",
        timeout=180_000)
    print("인식:", page.inner_text("#dw-recognize-status"))
    check("C3 버튼은 C2 전에 잠김", not page.is_enabled("#dw-c3-load"))

    # ── 게이트 통과 → C2 (앞 PR 검증에서 이미 본 경로라 여기선 통과만 시킨다) ──
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
    check("C2 뒤 [밸브 후보 불러오기] 열림", page.is_enabled("#dw-c3-load"))
    check("후보 전에는 [밸브 찍기] 잠김", not page.is_enabled("#dw-tool-valve"))
    check("후보 전에는 [밸브 표] 잠김", not page.is_enabled("#dw-c3-open"))

    # ── [1] 밸브 후보 ───────────────────────────────────────────────────────
    print("\n[1] 밸브 후보 불러오기")
    page.click("#dw-c3-load")
    page.wait_for_function(
        "() => !/부르는 중/.test(document.getElementById('dw-c3-status').textContent)",
        timeout=60_000)
    print("  상태:", page.inner_text("#dw-c3-status"))
    meta = page.evaluate(
        "() => ({ n: state.c3.candidates.length, sys: state.c3.systemTypes,"
        " reqs: state.c3.requirements.map((r) => r.key) })")
    print("  후보/설비종류/요건:", meta["n"], meta["sys"], meta["reqs"])
    check("확정 코어가 후보로 나옴", meta["n"] >= 2, f"{meta['n']}개")
    check("설치 요건 4종", len(meta["reqs"]) == 4, str(meta["reqs"]))
    check("[밸브 찍기] 열림", page.is_enabled("#dw-tool-valve"))
    check("[밸브 표] 열림", page.is_enabled("#dw-c3-open"))
    check("아직 밸브 없음", page.evaluate("() => state.c3.drafts.length") == 0)

    page.click("#dw-c3-open")
    page.wait_for_selector("#dw-c3-panel.open")
    check("밸브 없으면 확정 잠김", not page.is_enabled("#dw-c3-submit"))
    check("밸브 없으면 구역 잠김", not page.is_enabled("#dw-c3-zones"))
    print("  안내:", page.inner_text("#dw-c3-empty"))
    page.click("#dw-c3-close")

    # ── [2] 사람이 찍는다 — 밖은 끌어붙이지 않는다 ──────────────────────────
    print("\n[2] 밸브 찍기 — 후보 밖은 끌어붙이지 않는다")
    page.click("#dw-tool-valve")
    check("밸브 도구 활성", page.evaluate("() => state.tool") == "valve")

    spot = page.evaluate(OUTSIDE)
    check("후보 밖 지점 확보", spot is not None)
    if spot:
        page.mouse.click(spot["x"], spot["y"])
        page.wait_for_timeout(120)
        status = page.inner_text("#dw-c3-status")
        print("  상태:", status)
        check("밖을 찍으면 거절", "후보 코어 안" in status, status)
        check("밖을 찍어도 밸브 안 생김",
              page.evaluate("() => state.c3.drafts.length") == 0)

    aim0 = page.evaluate(AIM, 0)
    check("후보 0 제안점이 코어 안", bool(aim0 and aim0["hits"] == aim0["core"]
                                    and aim0["onCanvas"]),
          json.dumps(aim0, ensure_ascii=False))
    page.mouse.click(aim0["x"], aim0["y"])
    page.wait_for_timeout(120)
    print("  상태:", page.inner_text("#dw-c3-status"))
    check("밸브 1개 찍힘", page.evaluate("() => state.c3.drafts.length") == 1)

    nudge = aim0["nudge"] or aim0
    before = page.evaluate("() => state.c3.drafts[0].point.join(',')")
    page.mouse.click(nudge["x"], nudge["y"])
    page.wait_for_timeout(120)
    moved = page.evaluate("() => state.c3.drafts[0].point.join(',')")
    check("같은 코어를 다시 찍으면 자리만 옮김",
          page.evaluate("() => state.c3.drafts.length") == 1
          and (moved != before or aim0["nudge"] is None),
          page.inner_text("#dw-c3-status"))

    aim1 = page.evaluate(AIM, 1)
    check("후보 1 제안점이 코어 안", aim1["hits"] == aim1["core"] and aim1["onCanvas"],
          json.dumps(aim1, ensure_ascii=False))
    page.mouse.click(aim1["x"], aim1["y"])
    page.wait_for_timeout(120)
    drafts = page.evaluate("() => state.c3.drafts.map((d) => d.core_id)")
    check("두 번째 코어에 밸브 추가", len(drafts) == 2, ", ".join(drafts))
    page.screenshot(path=str(ROOT / "data" / "_dw_c3_drafts.png"))

    # ── [3] 안 고른 것은 '아니오' 가 아니다 ─────────────────────────────────
    print("\n[3] 설비 종류·요건 4종 — 기본값 없음")
    page.click("#dw-c3-open")
    page.wait_for_selector("#dw-c3-tbody tr")
    blank = page.evaluate(
        "() => [...document.querySelectorAll('#dw-c3-tbody select')]"
        ".filter((s) => s.value === '').length")
    check("모든 칸이 빈 채로 뜸", blank == 2 * (1 + len(meta["reqs"])), f"{blank}칸")
    print("  남음:", page.inner_text("#dw-c3-remain"))
    check("미응답 10건", "10건 미응답" in page.inner_text("#dw-c3-remain"),
          page.inner_text("#dw-c3-remain"))
    check("확정 잠김", not page.is_enabled("#dw-c3-submit"))
    page.screenshot(path=str(ROOT / "data" / "_dw_c3_blank.png"))

    wet = "습식" if "습식" in meta["sys"] else meta["sys"][0]
    for i in (0, 1):
        page.select_option(f"#dw-c3-tbody select[data-c3='system'][data-i='{i}']", wet)
    check("설비 종류만 골라도 잠김", not page.is_enabled("#dw-c3-submit"),
          page.inner_text("#dw-c3-remain"))

    for i in (0, 1):
        for key in meta["reqs"]:
            sel = f"#dw-c3-tbody select[data-c3='req'][data-i='{i}'][data-key='{key}']"
            page.select_option(sel, "true")
    # 마지막 밸브의 마지막 요건만 '아니오' — 확정은 되되 미충족으로 남아야 한다.
    last = meta["reqs"][-1]
    page.select_option(
        f"#dw-c3-tbody select[data-c3='req'][data-i='1'][data-key='{last}']", "false")
    print("  남음:", page.inner_text("#dw-c3-remain"))
    check("요건까지 답하면 미응답 0", "응답 완료" in page.inner_text("#dw-c3-remain"),
          page.inner_text("#dw-c3-remain"))
    check("확정자 없으면 아직 잠김", not page.is_enabled("#dw-c3-submit"))

    # 하나를 다시 비우면 잠긴다 — '—' 는 '아니오' 가 아니다.
    page.select_option(
        f"#dw-c3-tbody select[data-c3='req'][data-i='0'][data-key='{last}']", "")
    page.fill("#dw-c3-operator", "jinwon")
    check("한 칸만 비어도 잠김", not page.is_enabled("#dw-c3-submit"),
          page.inner_text("#dw-c3-remain"))
    page.select_option(
        f"#dw-c3-tbody select[data-c3='req'][data-i='0'][data-key='{last}']", "true")
    check("다 채우고 확정자까지 있으면 열림", page.is_enabled("#dw-c3-submit"))
    page.screenshot(path=str(ROOT / "data" / "_dw_c3_filled.png"))

    # ── [4] 확정 — '아니오' 는 미충족으로 남는다 ────────────────────────────
    print("\n[4] 밸브 확정")
    page.click("#dw-c3-submit")
    page.wait_for_function(
        "() => !/확정 중/.test(document.getElementById('dw-c3-msg').textContent)",
        timeout=60_000)
    msg = page.inner_text("#dw-c3-msg")
    print("  결과:", msg)
    valves = page.evaluate("() => state.c3.valves.map((v) => v.id)")
    check("밸브 2개 확정", len(valves) == 2, ", ".join(valves))
    check("요건 미충족이 화면에 남음", "미충족" in msg, msg)
    check("[구역 나누기] 열림", page.is_enabled("#dw-c3-zones"))

    # ── [5] 구역 — 문을 지나서만 닿는다 ─────────────────────────────────────
    print("\n[5] 구역 나누기")
    page.click("#dw-c3-zones")
    page.wait_for_function(
        "() => !/나누는 중/.test(document.getElementById('dw-c3-msg').textContent)",
        timeout=120_000)
    summary = page.inner_text("#dw-c3-msg")
    print("  요약:", summary)
    zones = page.evaluate(
        "() => ({ zones: state.c3.zones.length, flags: state.c3.flags.map((f) => f.code),"
        " unreached: state.c3.unreached.length, stage: state.stage })")
    print("  구역/플래그/미도달/단계:", zones)
    check("구역이 생김", zones["zones"] >= 1, f"{zones['zones']}개")
    check("구역 요약 배지 렌더",
          page.locator("#dw-c3-zones-sum .z").count() == zones["zones"])
    check("플래그 카드 렌더",
          page.locator("#dw-c3-flags .flag").count() == len(zones["flags"]),
          ", ".join(zones["flags"]) or "없음")
    if zones["flags"]:
        check("플래그가 남으면 C4 로 안 넘어감", zones["stage"] == "c3", zones["stage"])
        check("미도달 실이 보고됨", zones["unreached"] > 0, f"{zones['unreached']}개")
    else:
        check("플래그 없으면 C4 로 넘어감", zones["stage"] == "c4", zones["stage"])
    step = page.locator("#dw-stepper li.current").inner_text()
    print("  스텝퍼:", step.replace("\n", " "))
    page.click("#dw-c3-close")
    page.wait_for_timeout(200)
    page.screenshot(path=str(ROOT / "data" / "_dw_c3_zones.png"))
    # 담당 실의 색칠과 미도달 실의 빗금은 화면에서만 구분된다. 눈으로 볼 배율까지 당긴다.
    page.evaluate("() => { state.view.zoom *= 6; }")
    hub = page.evaluate(AIM, 0)
    page.screenshot(path=str(ROOT / "data" / "_dw_c3_zones_zoom.png"),
                    clip={"x": hub["x"] - 340, "y": hub["y"] - 240,
                          "width": 680, "height": 480})
    print("  구역 안내:", page.inner_text("#dw-c3-zone-note"))

    # ── [6] 밸브를 건드리면 구역은 근거를 잃는다 ────────────────────────────
    print("\n[6] 밸브를 고치면 확정·구역이 무효")
    page.click("#dw-c3-open")
    page.wait_for_selector("#dw-c3-tbody tr")
    page.select_option(
        f"#dw-c3-tbody select[data-c3='req'][data-i='0'][data-key='{last}']", "false")
    after = page.evaluate(
        "() => ({ valves: state.c3.valves.length, zones: state.c3.zones.length,"
        " flags: state.c3.flags.length })")
    check("확정 밸브 무효화", after["valves"] == 0, json.dumps(after))
    check("구역 무효화", after["zones"] == 0 and after["flags"] == 0, json.dumps(after))
    check("[구역 나누기] 다시 잠김", not page.is_enabled("#dw-c3-zones"))
    page.click("#dw-c3-tbody button[data-c3='drop'][data-i='1']")
    check("밸브 지우기 동작", page.evaluate("() => state.c3.drafts.length") == 1)

    browser.close()

network = [e for e in errors if "Failed to load resource" in e]
js = [e for e in errors if e not in network]
print("\n네트워크 로그(예상됨):", network or "없음")
print("JS 오류:", js or "없음")
print("실패한 검사:", fails or "없음")
sys.exit(1 if (js or fails) else 0)
