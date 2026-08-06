# -*- coding: utf-8 -*-
"""PR-9f·PR-10 C5 배관망 + 방출 UI 런타임 검증 — 일회용.

보고 싶은 것은 다섯이다. **헤드 없이는 라우팅 버튼이 열리지 않는가**, **놓인 배관이
캔버스에 실제로 그려지는가**(표만 차고 화면이 비면 검수할 수 없다), **관경이
급수원 쪽에서 가늘어지지 않는가**, **네 포맷이 실제로 내려받아지는가**, **헤드를
다시 놓으면 배관도 방출물도 무효가 되는가**.

마지막이 핵심이다. 없어진 헤드를 먹이는 배관이 화면에 남아 있거나 그 망으로 구운
솔버 파일이 계속 내려받아지면, 그 그림도 그 파일도 어떤 설계가 아니다.

`node --check` 는 구문만 본다 — 함수-지역 헬퍼를 다른 스코프에서 부르는 회귀는
브라우저에서만 잡힌다. 그래서 진짜 크로미움으로 돈다.
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


def skip(label: str, why: str) -> None:
    """볼 근거가 없는 검사. 합격으로 세면 조용한 거짓 합격이 된다."""
    print(f"  [NA  ] {label} — {why}")


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

C5_STATE = """() => ({
  networks: state.c5.networks.length,
  pipes: pipeCount(),
  heads: state.c5.networks.reduce((n, x) => n + (x.metrics.heads || 0), 0),
  length: state.c5.networks.reduce((n, x) => n + (x.metrics.length_m || 0), 0),
  flags: state.c5.flags.map((f) => f.code),
  blocking: state.c5.blocking,
  roles: [...new Set(state.c5.networks.flatMap(
    (x) => (x.graph.pipes || []).map((p) => p.role)))].sort(),
  dias: [...new Set(state.c5.networks.flatMap(
    (x) => (x.graph.pipes || []).map((p) => p.dia)))].sort((a, b) => a - b),
  stage: state.stage,
})"""

# 관경 단조성은 화면 밖에서 깨진다 — 표를 보고는 알 수 없으므로 그래프로 직접 잰다.
MONOTONIC = """() => {
  const bad = [];
  for (const net of state.c5.networks) {
    const down = new Map();
    for (const p of net.graph.pipes) {
      if (!down.has(p.in)) down.set(p.in, []);
      down.get(p.in).push(p);
    }
    for (const p of net.graph.pipes) {
      for (const child of down.get(p.out) || []) {
        if (child.dia > p.dia) bad.push(`${p.label}(${p.dia}) -> ${child.label}(${child.dia})`);
      }
    }
  }
  return bad;
}"""

SERVER_STAGE = """async () => {
  const res = await fetch(`/api/design/session/${state.sessionId}`);
  return (await res.json()).meta.stage;
}"""

# 캔버스에 배관이 실제로 칠해졌는지는 픽셀로 센다. 오버레이를 끈 화면과 켠 화면의
# 차이가 0 이면 표만 차고 그림은 비어 있는 것이다.
PIPE_PIXELS = """() => {
  const w = canvas.width, h = canvas.height;
  const on = () => ctx.getImageData(0, 0, w, h).data;
  state.overlayVisible.pipes = false; render();
  const off = Array.from(on());
  state.overlayVisible.pipes = true; render();
  const now = on();
  let diff = 0;
  for (let i = 0; i < now.length; i += 4) {
    if (now[i] !== off[i] || now[i + 1] !== off[i + 1] || now[i + 2] !== off[i + 2]) diff++;
  }
  return diff;
}"""

def _keep_rooms(sid: str) -> list[str]:
    """남길 실 — 첫 밸브 후보가 앉을 실과 문으로 이어진 덩어리.

    화면이 나중에 고를 후보(`state.c3.candidates[0]`)와 같은 실을 골라야 하므로
    같은 규칙으로 뽑는다 — `draft.cores` 순서의 첫 SHAFT·STAIR(검증이 코어를 전부
    "입상관으로 쓴다" 로 확정하므로 필터 결과의 첫 항목이 같다).
    """
    from core.design.deterministic import zoning as Z
    from core.design.recognize.spatial import representative_point
    from core.design.schema import BuildingDraft

    raw = json.loads((mod.DESIGN_SESSION_DIR / sid / "building.json")
                     .read_text(encoding="utf-8"))
    draft = BuildingDraft.from_dict(raw)
    core = next(c for c in draft.cores if c.kind.upper() in Z.VALVE_CORE_KINDS)
    seed = Z.seed_room(draft, representative_point(core.polygon))
    if seed is None:
        return []

    graph = Z.build_room_graph(draft)
    keep, stack = {seed}, [seed]
    while stack:
        for nxt, _hop, _mid in graph.neighbors(stack.pop()):
            if nxt not in keep:
                keep.add(nxt)
                stack.append(nxt)
    return sorted(keep)


# 화면의 [지우기] 와 같은 라우트를 같은 쿠키로 부른다.
PRUNE_PLAN = """async () => {
  const keep = new Set(__KEEP__);
  let deleted = 0, failed = 0;
  for (const room of state.rooms.filter((r) => !keep.has(r.id))) {
    const res = await fetch("/api/design/gate/edit", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.sessionId, operator: "jinwon",
                             edit: { op: "delete", room: room.id } }),
    });
    const body = await res.json();
    if (res.ok) { deleted++; state.rooms = body.rooms; } else { failed++; }
  }
  render();
  return { deleted, failed, left: state.rooms.length };
}"""

EMIT_STATE = """() => ({
  bundles: state.emit.bundles.length,
  suffixes: state.emit.bundles.map(
    (b) => b.files.map((f) => f.name.slice(f.name.lastIndexOf('.'))).sort()),
  empty: state.emit.bundles.reduce(
    (n, b) => n + b.files.filter((f) => !(f.bytes > 0)).length, 0),
  rt_ok: state.emit.bundles.every((b) => b.round_trip && b.round_trip.ok),
  flags: state.emit.flags.map((f) => f.code),
  links: document.querySelectorAll('#dw-emit-files a').length,
})"""

# 링크 주소를 그대로 눌러 본다. SDF 는 XML 이므로 첫 글자가 '<' 여야 한다.
FETCH_FIRST = """async () => {
  const a = [...document.querySelectorAll('#dw-emit-files a')]
    .find((el) => el.getAttribute('href').endsWith('.sdf'));
  if (!a) return { status: 0, bytes: 0, sdf: false };
  const res = await fetch(a.href);
  const text = await res.text();
  return { status: res.status, bytes: text.length,
           sdf: text.trimStart().startsWith('<'), name: a.textContent };
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
    check("헤드 전에는 [배관망 라우팅 실행] 잠김", not page.is_enabled("#dw-run-c5"))
    check("라우팅 전에는 [배관망 표 보기] 잠김", not page.is_enabled("#dw-c5-open"))
    check("배관 오버레이 항목은 0 본으로 뜸", page.evaluate("() => pipeCount()") == 0)

    # ── C1 → GATE → C2 → C3 → C4 (앞 PR 에서 검증한 경로) ──
    page.set_input_files("#dw-dxf", str(DXF))
    page.wait_for_selector("#dw-layer-list .layer-row", timeout=120_000)
    page.select_option("#dw-wall-layers", "ARCHI")
    page.click("#dw-recognize-btn")
    page.wait_for_function(
        "() => !document.getElementById('dw-gate-open').disabled"
        " || /오류/.test(document.getElementById('dw-recognize-status').textContent)",
        timeout=180_000)
    print("인식:", page.inner_text("#dw-recognize-status"))

    # ── 방출까지 가려면 닿지 않는 실을 먼저 치워야 한다 ─────────────────────
    # 이 도면은 C1 이 채택한 실이 도면의 일부만 덮어, 문으로 이어진 실 쌍이 11 개뿐이다
    # (zoning.py 모듈 주석의 실측). 남은 실은 `ROOM_UNREACHABLE` 로 C3 를 막고, 막힌
    # 단계에서는 방출도 열리지 않는다 — 서버가 옳다. 그래서 사람이 화면에서 할 일을
    # 그대로 한다: 밸브가 문으로 닿을 수 있는 실만 남기고 나머지는 GATE 지우기로 뺀다.
    # 확정 전이라 편집이 열려 있고, 지우는 경로도 화면의 [지우기] 와 같은 라우트다.
    pruned = page.evaluate(PRUNE_PLAN.replace(
        "__KEEP__", json.dumps(_keep_rooms(page.evaluate("() => state.sessionId")))))
    print("실 정리:", json.dumps(pruned, ensure_ascii=False))
    check("닿지 않는 실을 GATE 지우기로 뺌", pruned["failed"] == 0 and pruned["left"] > 0,
          f"{pruned['deleted']}개 지움 / {pruned['left']}개 남음")

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
    check("구역 뒤에도 [배관망 라우팅 실행] 은 잠김", not page.is_enabled("#dw-run-c5"),
          "헤드가 없으면 어디로 물을 보낼지 모른다")

    page.click("#dw-run-c4")
    page.wait_for_function(
        "() => !/배치하는 중/.test(document.getElementById('dw-c4-status').textContent)",
        timeout=300_000)
    heads = page.evaluate("() => headCount()")
    print("헤드:", page.inner_text("#dw-c4-status"))
    check("헤드가 놓임", heads > 0, f"{heads}개")
    check("헤드가 놓이면 [배관망 라우팅 실행] 열림", page.is_enabled("#dw-run-c5"))
    check("아직 [배관망 표 보기] 는 잠김", not page.is_enabled("#dw-c5-open"))
    print("  C5 상태:", page.inner_text("#dw-c5-status"))

    # ── [1] 배관망 라우팅 ───────────────────────────────────────────────────
    print("\n[1] 배관망 라우팅")
    # 앞 단계가 문제를 안고 있으면 stage 는 거기 서 있고, 배관망이 깨끗해도 방출은
    # 열리지 않는다. 그 경우까지 "emit 이어야 한다" 고 세면 서버가 옳은데 검사가 운다.
    stage_before = page.evaluate(SERVER_STAGE)
    t0 = time.time()
    page.click("#dw-run-c5")
    page.wait_for_function(
        "() => !/놓는 중/.test(document.getElementById('dw-c5-status').textContent)",
        timeout=600_000)
    print(f"  {time.time() - t0:.1f}s — {page.inner_text('#dw-c5-status')}")
    c5 = page.evaluate(C5_STATE)
    print("  상태:", json.dumps(c5, ensure_ascii=False))
    check("망이 나옴", c5["networks"] > 0, f"{c5['networks']}개")
    check("배관이 놓임", c5["pipes"] > 0, f"{c5['pipes']}본")
    check("총연장이 0 이 아님", c5["length"] > 0, f"{c5['length']:.1f} m")
    check("[배관망 표 보기] 열림", page.is_enabled("#dw-c5-open"))
    roles = set(c5["roles"])
    check("급수원에서 헤드까지 주배관·가지배관이 이어짐",
          {"branch", "main"} <= roles, ", ".join(c5["roles"]))
    # 분기점이 하나뿐이면 교차배관은 이을 두 점이 없어 티 한 점으로 줄고, 모든 관이
    # 헤드 1개를 물어 별표1 같은 칸에 떨어진다. 없는 것이 정상인 검사를 불합격으로
    # 세면 거짓 경보고, 합격으로 세면 조용한 거짓 합격이다.
    if c5["heads"] >= 2:
        check("교차배관이 나옴", "cross" in roles, ", ".join(c5["roles"]))
        check("관경이 하나로 뭉개지지 않음", len(c5["dias"]) >= 2,
              ", ".join(f"{d}A" for d in c5["dias"]))
    else:
        why = f"헤드 {c5['heads']}개 — 분기점이 하나면 교차배관도 관경 차이도 없다"
        skip("교차배관이 나옴", why)
        skip("관경이 하나로 뭉개지지 않음", why)
    check("미방호 헤드 없음", "HEAD_UNREACHED" not in c5["flags"],
          ", ".join(c5["flags"]) or "플래그 없음")

    bad = page.evaluate(MONOTONIC)
    check("급수원 쪽 관경이 가늘어지지 않음", not bad, ", ".join(bad[:5]) or "역전 없음")

    server_stage = page.evaluate(SERVER_STAGE)
    check("화면 단계와 서버 단계가 같음", c5["stage"] == server_stage,
          f"화면={c5['stage']} 서버={server_stage}")
    if c5["blocking"]:
        check("막는 문제가 있으면 방출로 안 넘어감", server_stage != "emit",
              f"{server_stage} ← {', '.join(c5['blocking'])}")
    elif stage_before == "c5":
        check("막는 문제가 없으면 방출로 넘어감", server_stage == "emit", server_stage)
    else:
        check("앞 단계가 막고 있으면 방출로 안 넘어감", server_stage == stage_before,
              f"{stage_before} 에 서 있음")

    # ── [2] 배관망 표 ───────────────────────────────────────────────────────
    print("\n[2] 배관망 패널")
    page.click("#dw-c5-open")
    page.wait_for_selector("#dw-c5-panel.open")
    rows = page.locator("#dw-c5-tbody tr").count()
    check("망별 표가 렌더", rows == c5["networks"], f"{rows}행")
    dia_rows = page.locator("#dw-c5-dia tbody tr").count()
    check("호칭경별 연장표가 렌더", dia_rows > 0, f"{dia_rows}행")
    print("  호칭경표:\n   ", page.inner_text("#dw-c5-dia").replace("\n", " | "))
    check("배관 본수 배지", f"배관 {c5['pipes']}본" in page.inner_text("#dw-c5-remain")
          or bool(c5["blocking"]), page.inner_text("#dw-c5-remain"))
    page.screenshot(path=str(ROOT / "data" / "_dw_c5_panel.png"))
    page.click("#dw-c5-close")

    # ── [3] 화면 — 배관이 실제로 그려지는가 ─────────────────────────────────
    print("\n[3] 배관 오버레이")
    overlay = page.evaluate(
        "() => { const el = [...document.querySelectorAll('#dw-design-list input')]"
        ".find((i) => i.dataset.overlay === 'pipes');"
        " return el ? { checked: el.checked, disabled: el.disabled,"
        " count: el.closest('label').querySelector('.count').textContent } : null; }")
    print("  산출물 항목:", json.dumps(overlay, ensure_ascii=False))
    check("배관망 항목이 산출물 목록에 있음", overlay is not None)
    check("배관망 항목이 켜져 있고 잠기지 않음",
          bool(overlay and overlay["checked"] and not overlay["disabled"]))
    check("배관 본수가 목록에 나옴",
          bool(overlay) and overlay["count"] == str(c5["pipes"]),
          f"{overlay['count'] if overlay else '?'} vs {c5['pipes']}")

    page.evaluate("() => { state.view.zoom *= 4; render(); }")
    aim2 = page.evaluate(AIM, 0)
    painted = page.evaluate(PIPE_PIXELS)
    check("캔버스에 배관이 실제로 칠해짐", painted > 0, f"{painted:,} px")
    page.screenshot(path=str(ROOT / "data" / "_dw_c5_pipes.png"),
                    clip={"x": aim2["x"] - 500, "y": aim2["y"] - 320,
                          "width": 1000, "height": 640})

    # ── [4] 방출 — 솔버 파일 (§11 · §13.6) ──────────────────────────────────
    print("\n[4] 솔버 파일 방출")
    if server_stage != "emit":
        why = f"서버 단계가 {server_stage} — 방출 자체가 열리지 않는 상태다"
        check("방출로 넘어가지 못하면 버튼도 잠김", not page.is_enabled("#dw-emit-btn"), why)
        for label in ("망마다 네 포맷이 나옴", "왕복 검증이 일치", "파일이 실제로 내려받아짐"):
            skip(label, why)
        emitted = None
    else:
        check("방출 단계가 되면 [솔버 파일 방출] 열림", page.is_enabled("#dw-emit-btn"))
        t0 = time.time()
        page.click("#dw-emit-btn")
        page.wait_for_function(
            "() => !/쓰는 중/.test(document.getElementById('dw-emit-status').textContent)",
            timeout=600_000)
        print(f"  {time.time() - t0:.1f}s — {page.inner_text('#dw-emit-status')}")
        emitted = page.evaluate(EMIT_STATE)
        print("  상태:", json.dumps(emitted, ensure_ascii=False))

        check("망 수가 C5 와 같음", emitted["bundles"] == c5["networks"],
              f"방출 {emitted['bundles']} vs 라우팅 {c5['networks']}")
        check("망마다 네 포맷이 나옴", emitted["suffixes"] == [
            [".has", ".kfp", ".sdf", ".slf"]] * emitted["bundles"],
            json.dumps(emitted["suffixes"]))
        check("빈 파일이 없음", emitted["empty"] == 0, f"{emitted['empty']}개")
        check("왕복 검증이 일치", emitted["rt_ok"] and not emitted["flags"],
              ", ".join(emitted["flags"]) or "플래그 없음")
        # 링크가 그려진 것과 그 주소가 실제로 파일을 준다는 것은 다르다.
        got = page.evaluate(FETCH_FIRST)
        check("파일이 실제로 내려받아짐",
              got["status"] == 200 and got["bytes"] > 0 and got["sdf"],
              json.dumps(got, ensure_ascii=False))
        # 없는 이름은 `emit.json` 화이트리스트에서 걸러야 한다.
        bogus = page.evaluate(
            "async () => (await fetch("
            "`/api/design/emit/${state.sessionId}/meta.json`)).status")
        check("방출물이 아닌 파일은 거절", bogus == 404, str(bogus))
        page.screenshot(path=str(ROOT / "data" / "_dw_emit_panel.png"))

    # ── [5] 헤드를 다시 놓으면 배관은 근거를 잃는다 ─────────────────────────
    print("\n[5] 헤드를 다시 놓으면 배관이 무효")
    page.click("#dw-run-c4")
    page.wait_for_function(
        "() => !/배치하는 중/.test(document.getElementById('dw-c4-status').textContent)",
        timeout=300_000)
    reset = page.evaluate(
        "() => ({ pipes: pipeCount(), ran: state.c5.ran, nets: state.c5.networks.length })")
    check("배관망 무효화", reset["pipes"] == 0 and not reset["ran"], json.dumps(reset))
    check("[배관망 표 보기] 다시 잠김", not page.is_enabled("#dw-c5-open"))
    check("[배관망 라우팅 실행] 은 열려 있음", page.is_enabled("#dw-run-c5"),
          "헤드를 다시 놓았으니 다시 돌릴 수 있어야 한다")
    # 없어진 망으로 구운 파일을 계속 내려받게 두면 화면과 손에 쥔 파일이 갈린다.
    if emitted is None:
        skip("방출물도 함께 무효화", "방출을 돌린 적이 없다")
    else:
        after = page.evaluate(EMIT_STATE)
        check("방출물도 함께 무효화",
              after["bundles"] == 0 and after["links"] == 0
              and not page.is_enabled("#dw-emit-btn"), json.dumps(after))

    browser.close()

network = [e for e in errors if "Failed to load resource" in e]
js = [e for e in errors if e not in network]
print("\n네트워크 로그(예상됨):", network or "없음")
print("JS 오류:", js or "없음")
print("실패한 검사:", fails or "없음")
sys.exit(1 if (js or fails) else 0)
