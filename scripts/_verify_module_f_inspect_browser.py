# -*- coding: utf-8 -*-
"""[F-12] 수리계산 «속성 카드» — 실제 브라우저에서 확인한다.

구문 검사로는 못 잡는 것을 본다: 캔버스에서 배관·노드를 정말로 집는지,
카드가 **표와 같은 값**을 말하는지, 카드 안 라벨로 건너뛰는지.

★운영 JS 에 시험용 훅을 넣지 않는다. 좌표는 공개 UI 만으로 얻는다 —
  화면 아래 `#coord` 가 마우스 자리의 세계좌표를 적어 주므로, 픽셀 두 곳을
  훑어 «픽셀↔세계» 변환을 되풀어 쓴다. 시험을 위해 제품에 문을 내면 그
  문은 영영 남는다.
★값의 출처가 하나여야 한다. 카드가 제 손으로 다시 계산하면 표와 카드가
  다른 말을 하는 날이 온다 — 그래서 카드에 뜬 값을 **표의 그 행**과 직접
  맞대 본다.

실행:
    MF_BASE=http://127.0.0.1:5065 LOGIN_PASSWORD=… \
        python scripts/_verify_module_f_inspect_browser.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
SHOT = ROOT / "data" / "_mf_shots"
SHOT.mkdir(parents=True, exist_ok=True)
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5065")
PASSWORD = os.environ["LOGIN_PASSWORD"]
# ★렌더링이 성한 저장본이어야 «캔버스에서 집는» 검사가 뜻이 있다.
#   (일부 저장본은 좌표 이상치로 망이 몇 화소에 뭉쳐 그려진다 —
#    그 상태에서는 사람도 못 누른다.)
SAVED_KEY = os.environ.get("MF_KEY", "1. 입력도면 대명동 단위세대 평면도")

problems: list[str] = []


def bad(msg: str) -> None:
    problems.append(msg)
    print("   !!", msg)


def check(name: str, ok: bool, detail: str = "") -> bool:
    print(f"  [{'OK  ' if ok else 'FAIL'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        problems.append(name + (f" · {detail}" if detail else ""))
    return ok


def read_table(page) -> tuple[list[str], list[list[str]]]:
    """`#dg-grid` 표를 (머리, 행들) 로. 카드와 맞댈 «원본» 이다."""
    return page.evaluate("""() => {
      const t = document.querySelector('#dg-grid table');
      if (!t) return [[], []];
      const head = [...t.querySelectorAll('thead th')].map(e => e.textContent.trim());
      const rows = [...t.querySelectorAll('tbody tr')].map(
        tr => [...tr.querySelectorAll('td')].map(td => td.textContent.trim()));
      return [head, rows];
    }""")


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={"width": 1680, "height": 1000})
    page.on("console", lambda m: bad(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: bad(f"pageerror: {e}"))

    print("[0] 로그인 · 저장본으로 손질 진입")
    page.goto(f"{BASE}/module-f", wait_until="load")
    if "login" in page.url or page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PASSWORD)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    if "/module-f" not in page.url:
        page.goto(f"{BASE}/module-f", wait_until="load")
    page.wait_for_function(
        "document.querySelector('#saved').options.length > 0", timeout=30_000)
    opts = page.eval_on_selector_all("#saved option", "e => e.map(o => o.value)")
    if SAVED_KEY not in opts:
        bad(f"저장 목록에 {SAVED_KEY} 가 없다 — {opts[:4]}")
        browser.close()
        raise SystemExit(1)
    page.select_option("#saved", SAVED_KEY)
    page.click("#btn-reopen")
    page.wait_for_selector("#panel-edit:not(.hidden)", timeout=300_000)
    print("   손질:", page.inner_text("#status")[:70])

    print("[1] 급수 시작을 찍는다 — 저장본에는 안 들어 있다")
    # ★내부 상태를 훔쳐보지 않는다. 사람이 하는 대로 «배관이 그려진 화소» 를
    #   찾아 거기를 누른다. 캔버스는 검정 바탕이므로 검정 아닌 화소가 곧 망이다.
    # ★한 자리만 찍고 되기를 바라지 않는다. 검정 아닌 화소는 배관일 수도
    #   헤드 기호일 수도 있고, 급수 시작은 «배관» 위여야 잡힌다. 그려진
    #   자리를 여러 곳 모아 잡힐 때까지 눌러 본다 — 사람이 하는 그대로다.
    spots = page.evaluate("""() => {
      const cv = document.getElementById('cv');
      const g = cv.getContext('2d');
      const d = g.getImageData(0, 0, cv.width, cv.height).data;
      const r = cv.width / cv.getBoundingClientRect().width;
      const cx = (cv.width / 2) | 0, cy = (cv.height / 2) | 0;
      // 성기게 훑으면 얇은 선을 통째로 놓친다(실측: 화소를 하나도 못 찾았다).
      // 전면을 훑어 모은 뒤 «고르게» 솎는다.
      const hits = [];
      for (let y = 2; y < cv.height - 2; y += 2) {
        for (let x = 2; x < cv.width - 2; x += 2) {
          const i = (y * cv.width + x) * 4;
          if (d[i] + d[i + 1] + d[i + 2] > 200) hits.push([x, y]);
        }
      }
      if (!hits.length) return [];
      // 가운데에 가까운 것부터 — 가장자리는 잘려 못 누른다.
      hits.sort((p, q) => ((p[0]-cx)**2 + (p[1]-cy)**2) - ((q[0]-cx)**2 + (q[1]-cy)**2));
      const step = Math.max(1, (hits.length / 40) | 0);
      const out = [];
      for (let k = 0; k < hits.length && out.length < 40; k += step) {
        out.push({x: hits[k][0] / r, y: hits[k][1] / r});
      }
      return out;
    }""")
    if not spots:
        bad("캔버스에서 망 화소를 못 찾았다 — 아무것도 안 그려졌다")
        browser.close()
        raise SystemExit(1)
    page.click('.emode[data-mode="급수시작위치"]')
    page.wait_for_timeout(200)
    box0 = page.eval_on_selector("#cv", "e => { const r = e.getBoundingClientRect();"
                                        " return {x: r.x, y: r.y}; }")
    placed = False
    for i, sp in enumerate(spots[:20]):
        page.mouse.click(box0["x"] + sp["x"], box0["y"] + sp["y"])
        page.wait_for_timeout(400)
        if "잡히지" not in page.inner_text("#status"):
            placed = True
            print(f"   {i + 1}번째 자리에서 잡힘 · {page.inner_text('#status')[:60]}")
            break
    if not placed:
        bad("급수 시작을 어느 자리에서도 못 찍었다")
        browser.close()
        raise SystemExit(1)

    print("[2] 수리계산으로 · 표 확정")
    # 「수리계산 입력 →」 단추는 변환 패널에 산다(손질에서는 숨어 있다).
    # 단계바로 간다 — S.edit 만 있으면 갈 수 있는 단계다.
    page.click('.steps div:text-is("수리계산")')
    page.wait_for_selector("#panel-design:not(.hidden)", timeout=30_000)
    page.click("#dg-build")
    # 표가 서면 표 고르개가 채워지고 격자에 행이 생긴다.
    page.wait_for_function(
        "() => { const t = document.querySelector('#dg-grid table');"
        " return t && t.querySelectorAll('tbody tr').length > 0; }",
        timeout=900_000)
    print("   표 확정 완료 ·", page.inner_text("#status")[:70])

    print("[3] 화면에 «그려진 것» 을 눌러 집는다")
    # ★표의 x·y 로 클릭하면 안 된다 — 그것은 캔버스 좌표가 아니다.
    #   `display_tables` 가 아이소 투영·정규화를 거친 좌표를 view 에 따로
    #   내므로 둘은 다른 공간이다(실측으로 여기서 한 번 헛짚었다).
    #   사람이 하는 대로 «보이는 색» 을 눌러 집는다.
    box = page.eval_on_selector("#cv", """e => {
      const r = e.getBoundingClientRect();
      return {x: r.x, y: r.y, w: r.width, h: r.height};
    }""")

    def find_px(rgb, tol=40, want_dark=False):
        """그 색으로 칠해진 화소 하나(캔버스 CSS 좌표). 없으면 None."""
        return page.evaluate("""([rgb, tol, dark]) => {
          const cv = document.getElementById('cv');
          const g = cv.getContext('2d');
          const d = g.getImageData(0, 0, cv.width, cv.height).data;
          const r = cv.width / cv.getBoundingClientRect().width;
          const out = [];
          const st = dark ? 8 : 1;
          for (let y = 2; y < cv.height - 2; y += st)
            for (let x = 2; x < cv.width - 2; x += st) {
              const i = (y * cv.width + x) * 4;
              if (dark) {
                // ★«검은 화소» 로는 모자란다 — 배관 바로 옆도 검다. 둘레가
                //   통째로 비어 있어야 «빈 자리» 다(안 그러면 클릭이 잡힌다).
                let clear = true;
                for (let dy = -40; dy <= 40 && clear; dy += 8)
                  for (let dx = -40; dx <= 40; dx += 8) {
                    const yy = y + dy, xx = x + dx;
                    if (yy < 0 || xx < 0 || yy >= cv.height || xx >= cv.width) continue;
                    const k = (yy * cv.width + xx) * 4;
                    if (d[k] + d[k+1] + d[k+2] > 30) { clear = false; break; }
                  }
                if (clear) out.push([x, y]);
              } else if (Math.abs(d[i] - rgb[0]) <= tol
                      && Math.abs(d[i+1] - rgb[1]) <= tol
                      && Math.abs(d[i+2] - rgb[2]) <= tol) out.push([x, y]);
            }
          if (!out.length) return null;
          const p = out[(out.length / 2) | 0];
          return {x: p[0] / r, y: p[1] / r, n: out.length};
        }""", [list(rgb), tol, want_dark])

    def click_px(sp):
        page.mouse.click(box["x"] + sp["x"], box["y"] + sp["y"])
        page.wait_for_timeout(300)

    def card_open():
        return not page.eval_on_selector(
            "#dg-ins", "e => e.classList.contains('hidden')")

    print("[4] 급수원 노드 — 표의 Input 행에서 연다(색 추정에 안 매인다)")
    page.select_option("#dg-table", "nodes")
    page.wait_for_timeout(300)
    nhead, nrows = read_table(page)
    nci = {n: nhead.index(n) for n in nhead}
    src_lab = None
    if "입출력" in nci:
        for k, r in enumerate(nrows):
            if r[nci["입출력"]] == "Input":
                src_lab = r[nci.get("이름", 0)]
                page.eval_on_selector_all(
                    "#dg-grid tbody tr",
                    f"els => els[{k}].click()")
                page.wait_for_timeout(300)
                break
    if src_lab is None:
        bad("노드 표에 Input(급수원) 행이 없다")
    else:
        card = page.inner_text("#dg-ins")
        check("급수원 카드가 열렸다", card_open())
        check("역할이 «급수원» 으로 적힌다", "급수원" in card, src_lab)
        check("노즐·연결 배관이 함께 뜬다",
              "연결 배관" in page.inner_text("#dg-ins-body"))

    print("[5] 캔버스에서 «배관» 과 «노드» 를 각각 집는다")
    # ★한 화소만 눌러 보고 판단하지 않는다. 노드가 배관보다 먼저 잡히는 것이
    #   설계다 — 배관 화소라도 노드가 더 가까우면 노드가 잡힌다(옳다).
    #   그래서 «배관이 잡힐 때까지» 여러 자리를 눌러 본다. 사람이 하는 그대로다.
    cands = page.evaluate("""() => {
      const cv = document.getElementById('cv');
      const g = cv.getContext('2d');
      const d = g.getImageData(0, 0, cv.width, cv.height).data;
      const r = cv.width / cv.getBoundingClientRect().width;
      const hits = [];
      for (let y = 3; y < cv.height - 3; y += 3)
        for (let x = 3; x < cv.width - 3; x += 3) {
          const i = (y * cv.width + x) * 4;
          if (d[i] + d[i+1] + d[i+2] > 150) hits.push([x, y]);
        }
      const step = Math.max(1, (hits.length / 60) | 0);
      const out = [];
      for (let k = 0; k < hits.length && out.length < 60; k += step)
        out.push({x: hits[k][0] / r, y: hits[k][1] / r});
      return out;
    }""")
    got_pipe, got_node, pipe_lab = False, False, None
    for sp in cands:
        click_px(sp)
        if not card_open():
            continue
        kind = page.inner_text("#dg-ins-kind").strip()
        if kind == "노드":
            got_node = True
        elif kind == "배관" and not got_pipe:
            got_pipe = True
            pipe_lab = page.inner_text("#dg-ins-title").strip().split()[0]
        if got_pipe and got_node:
            break
    check("캔버스 클릭으로 노드가 잡힌다", got_node)
    check("캔버스 클릭으로 배관이 잡힌다", got_pipe, str(pipe_lab))

    print("[6] 카드 값이 «표의 그 행» 과 같은가")
    if not pipe_lab:
        bad("배관을 못 집어 값 대조를 건너뛴다")
    else:
        body = page.inner_text("#dg-ins-body")
        page.select_option("#dg-table", "pipes")
        page.wait_for_timeout(300)
        phead, prows = read_table(page)
        pci = {n: phead.index(n) for n in phead}
        row = next((r for r in prows
                    if "이름" in pci and r[pci["이름"]] == pipe_lab), None)
        if row is None:
            bad(f"카드가 말한 배관 {pipe_lab} 이 표에 없다")
        else:
            miss = [f"{n}={row[pci[n]]}" for n in
                    ("이름", "시작", "끝", "호칭경(mm)", "길이(m)")
                    if n in pci and row[pci[n]] and row[pci[n]] not in body]
            check("표 행의 값이 카드에 그대로", not miss, " · ".join(miss))
            check("부속 칸이 있다", "부속" in body)

    print("[7] 카드 안 라벨로 노드까지 건너뛴다")
    if pipe_lab:
        # 배관 카드로 되돌린 뒤 링크를 탄다.
        page.select_option("#dg-table", "pipes")
        page.wait_for_timeout(250)
        page.eval_on_selector_all(
            "#dg-grid tbody tr",
            f"els => {{ const e = els.find(t => t.dataset.label === '{pipe_lab}');"
            " if (e) e.click(); }")
        page.wait_for_timeout(300)
        page.eval_on_selector_all(
            "#dg-ins-body .ins-link",
            "els => { const e = els.find(x => x.dataset.insKind === 'node');"
            " if (e) e.click(); }")
        page.wait_for_timeout(300)
        check("노드 카드로 넘어갔다",
              page.inner_text("#dg-ins-kind").strip() == "노드",
              page.inner_text("#dg-ins-kind"))

    print("[8] 빈 자리 클릭 · Esc 로 닫힌다")
    blank = find_px((0, 0, 0), want_dark=True)
    if blank and card_open():
        click_px(blank)
        check("빈 자리를 누르면 닫힌다", not card_open())
    if cands:
        click_px(cands[0])
        page.keyboard.press("Escape")
        page.wait_for_timeout(250)
        check("Esc 로 닫힌다", not card_open())

    print("[9] 표 행 클릭도 같은 카드를 연다")
    page.select_option("#dg-table", "nodes")
    page.wait_for_timeout(300)
    page.eval_on_selector("#dg-grid tbody tr", "e => e.click()")
    page.wait_for_timeout(250)
    check("표에서도 카드가 열린다", card_open())

    page.screenshot(path=str(SHOT / "inspect.png"))
    check("콘솔 오류 0",
          not [p for p in problems if p.startswith(("console", "pageerror"))])
    browser.close()

print()
if problems:
    print(f"실패 {len(problems)}건")
    for p in problems:
        print("  -", p)
    raise SystemExit(1)
print("모두 통과.")
