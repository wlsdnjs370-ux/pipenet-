# -*- coding: utf-8 -*-
"""module_f 화면을 **실제 브라우저에서** 연다 — 단계바 리팩터 회귀 방벽.

`node --check` 도 스코프 검사도 «파일이 성립하는가» 까지다. 페이지가 뜰 때
정말 아무 것도 안 터지는지, 단계바가 슬롯에 맞게 그려지는지는 브라우저에서만
드러난다(실측 회귀: 함수-지역 헬퍼를 다른 스코프에서 불러 클릭하는 순간
ReferenceError 로 죽은 적이 있다).

운영 서버(5051)를 쓰지 않는다 — 로그인 실패가 쌓이면 IP 가 잠긴다. 여기서는
**임시 포트에 제 서버를 띄워** 그것만 본다.

    python scripts/_verify_module_f_ui.py
"""
from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# ASCII 로 둔다 — Windows 서브프로세스 env 는 비ASCII 를 코드페이지로 넘겨
# 서버가 받은 값과 여기서 타이핑한 값이 달라진다(실측으로 로그인이 실패했다).
PASSWORD = "ui-verify-pw-2026"
FAILS: list[str] = []
# 지시서 §6 이 청하는 화면 캡처가 여기 쌓인다.
SHOTS = ROOT / "data" / "_f8_shots"

# 캔버스를 «눈으로» 검사한다 — 색이 실제로 칠해졌는지는 픽셀만이 안다.
_CANVAS_HEAD = """() => {
  const c = document.getElementById('cv');
  const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
  let n = 0;
"""
_COUNT_LIT = _CANVAS_HEAD + """
  for (let i = 0; i < d.length; i += 4) {
    if (d[i] + d[i+1] + d[i+2] > 60) n++;
  }
  return n;
}"""
# 빨강 — 검출 헤드 표시(#ff3b30). 붉은 기가 확실히 우세한 픽셀만 센다.
_COUNT_RED = _CANVAS_HEAD + """
  for (let i = 0; i < d.length; i += 4) {
    if (d[i] > 150 && d[i] > d[i+1] * 2 && d[i] > d[i+2] * 2) n++;
  }
  return n;
}"""
# 청록 — 뽑아낸 배관망(#22d3ee).
_COUNT_CYAN = _CANVAS_HEAD + """
  for (let i = 0; i < d.length; i += 4) {
    if (d[i+2] > 140 && d[i+1] > 110 && d[i+2] > d[i] * 2) n++;
  }
  return n;
}"""
# 파랑 — 검출한 배관망(#60a5fa). 최불리(청록)와는 파랑이 더 세다는 점으로 가른다.
_COUNT_BLUE = _CANVAS_HEAD + """
  for (let i = 0; i < d.length; i += 4) {
    if (d[i+2] > 140 && d[i+2] > d[i+1] * 1.25 && d[i+2] > d[i] * 1.8) n++;
  }
  return n;
}"""
# 주황 — 분기(티) 표시(#f59e0b). 붉은 헤드(#ff3b30)와는 초록 성분으로 가른다.
_COUNT_AMBER = _CANVAS_HEAD + """
  for (let i = 0; i < d.length; i += 4) {
    const r = d[i], g = d[i+1], b = d[i+2];
    if (r > 150 && g > 90 && g < r * 0.92 && b < g * 0.6) n++;
  }
  return n;
}"""
# «환하게 남은 배경» — 뽑은 망(청록)도 헤드(빨강)도 아닌데 환한 픽셀.
#
# ★내림을 «칠해진 픽셀 총량» 으로 재면 안 된다. 추출 뒤에는 뽑은 자리로 화면을
#   확대하므로 도면이 커져 총량은 오히려 는다(실측 1,293 → 21,150). 재야 할
#   것은 양이 아니라 «밝기» 다: alpha 0.16 으로 내린 선은 최대 채널이 40 언저리라
#   120 을 못 넘는다. 그러므로 이 값이 작다 = 배경이 실제로 내려가 있다.
_COUNT_BRIGHT_BG = _CANVAS_HEAD + """
  for (let i = 0; i < d.length; i += 4) {
    const r = d[i], g = d[i+1], b = d[i+2];
    if (Math.max(r, g, b) <= 120) continue;
    if (b > 140 && g > 110 && b > r * 2) continue;          // 뽑은 망(청록)
    if (r > 150 && r > g * 2 && r > b * 2) continue;        // 검출 헤드(빨강)
    if (r > 150 && g > 90 && g < r * 0.92 && b < g * 0.6) continue;  // 분기(주황)
    if (b > 140 && b > r * 1.5) continue;                   // 검출망(파랑)
    n++;
  }
  return n;
}"""


def check(label: str, cond: bool, detail: str = "") -> bool:
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


def _free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    p = s.getsockname()[1]
    s.close()
    return p


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    from playwright.sync_api import sync_playwright

    port = _free_port()
    env = dict(os.environ)
    env.update({"PORT": str(port), "HOST": "127.0.0.1",
                "LOGIN_PASSWORD": PASSWORD,
                "DESIGN_WORKBENCH_ENABLED": "1",
                "PYTHONIOENCODING": "utf-8"})
    SHOTS.mkdir(parents=True, exist_ok=True)
    print(f"임시 서버 :{port} 기동 중… (캡처 → {SHOTS})")
    proc = subprocess.Popen([sys.executable, "serve.py"], cwd=str(ROOT),
                            env=env, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL)
    base = f"http://127.0.0.1:{port}"
    try:
        for _ in range(60):
            try:
                socket.create_connection(("127.0.0.1", port), 0.5).close()
                break
            except OSError:
                time.sleep(1)
        else:
            check("임시 서버 기동", False, "포트가 안 열림")
            return 1
        check("임시 서버 기동", True, f":{port}")

        errors: list[str] = []
        # sid 는 화면 JS 안(IIFE)에만 있다. 그것을 꺼내자고 프로덕션에 전역을
        # 심지 않는다 — 열기 응답에서 그대로 주우면 된다.
        seen: dict = {}
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.on("console", lambda m: errors.append(m.text)
                    if m.type == "error" else None)
            page.on("pageerror", lambda e: errors.append(str(e)))

            def _grab(resp):
                if "/api/module-f/slot/open" not in resp.url:
                    return
                try:
                    seen["sid"] = (resp.json() or {}).get("sid")
                except Exception:  # noqa: BLE001
                    pass
            page.on("response", _grab)

            page.goto(f"{base}/module-f", wait_until="load")
            if page.query_selector("input[type=password]"):
                page.fill("input[type=password]", PASSWORD)
                with page.expect_navigation(wait_until="load"):
                    page.click("button[type=submit]")
                page.goto(f"{base}/module-f", wait_until="load")
            if "login" in page.url:
                # 왜 막혔는지 화면 문구를 그대로 싣는다 — 비번인지 잠금인지.
                msg = (page.inner_text("body") or "")[:160].replace("\n", " ")
                check("로그인 통과", False, msg)
                browser.close()
                return 1
            check("로그인 통과", True, page.url)

            page.wait_for_timeout(900)

            # ── 단계바가 슬롯 흐름대로 그려졌나 (평면도가 기본)
            # 방식을 고르기 전에는 「도면 열기」 하나뿐 — 어느 길로 갈지 모른다.
            steps = page.eval_on_selector_all(
                "#steps div", "els => els.map(e => e.textContent.trim())")
            check("열기 전 단계바는 «도면 열기» 뿐",
                  steps == ["도면 열기"], " · ".join(steps))
            # [F-10a · D-F10-1] 방식 질문 자체가 없어졌다.
            check("방식 묻는 화면이 아예 없다",
                  page.query_selector("#panel-method") is None)
            check("올리기는 처음부터 열려 있다",
                  not page.is_disabled("#btn-open"))

            # ── 슬롯 탭이 그려졌나 (세션 전에는 안내문)
            slots = page.inner_text("#slots").strip()
            check("슬롯 자리 있음", bool(slots), slots[:50])

            # ── [D-F10-2] 자동 차선은 «고급» 안 한 줄로 남는다(엔드포인트 보존)
            check("고급 카드가 있다",
                  page.query_selector("#panel-advanced") is not None)
            check("자동 차선 입구가 고급에 있다",
                  page.query_selector("#adv-auto") is not None)
            check("자동 추출 패널이 있다",
                  page.query_selector("#panel-auto") is not None)
            for aid in ("au-anchor", "au-zone-arm", "au-heads", "au-run",
                        "au-k-preset", "au-k"):
                check(f"#{aid} 존재", page.query_selector(f"#{aid}") is not None)

            # ── 갈 수 없는 단계는 막힌다. ★도면을 열기 «전» 에 봐야 한다 —
            #    열고 나면 찍기가 실제로 도달 가능해져 이 검사가 뜻을 잃는다.
            #    (지금 단계바는 「도면 열기」 한 칸뿐이라 누를 다음 칸이 없다 —
            #     그 자체가 「갈 수 없다」의 표현이다.)
            steps_now = page.eval_on_selector_all(
                "#steps div", "els => els.map(e => e.textContent.trim())")
            check("열기 전에는 다음 단계 자체가 없다",
                  steps_now == ["도면 열기"], " · ".join(steps_now))
            page.click("#steps div:nth-child(1)")
            page.wait_for_timeout(300)
            after = page.eval_on_selector_all(
                "#steps div.on", "els => els.map(e => e.textContent.trim())")
            check("재료 없는 단계로는 안 넘어간다", after == ["도면 열기"],
                  " · ".join(after))

            # ── 새로 만든 패널들이 DOM 에 있나
            for pid in ("panel-sub", "panel-merge", "panel-design",
                        "ed-zone-arm", "ed-zones", "dg-bore-legend",
                        "ed-k-preset", "ed-k", "ed-worst-view"):
                check(f"#{pid} 존재", page.query_selector(f"#{pid}") is not None)

            # ── 기준개수 표가 서버에서 채워졌나 (화면이 표를 옮겨 적지 않는다)
            opts = page.eval_on_selector_all(
                "#ed-k-preset option",
                "els => els.map(e => e.textContent.trim())")
            check("기준개수 표가 채워진다", len(opts) > 1, f"{len(opts)}행")
            counts = set()
            for o in opts[1:]:
                head = o.split("개")[0].strip()
                if head.isdigit():
                    counts.add(int(head))
            check("표에 10 · 20 · 30 이 다 있다", {10, 20, 30} <= counts,
                  str(sorted(counts)))

            # ── 표에서 고르면 K 가 따라오나
            #    손질 패널은 지금 단계(도면 열기)에서 숨어 있다 — 검사 동안만
            #    펼친다. 숨긴 채로는 Playwright 가 select 를 못 건드린다.
            if len(opts) > 1:
                page.eval_on_selector(
                    "#panel-edit", "el => el.classList.remove('hidden')")
                page.select_option("#ed-k-preset", "0")
                page.wait_for_timeout(200)
                k = page.input_value("#ed-k")
                first_count = opts[1].split("개")[0].strip()
                check("표를 고르면 K 가 따라온다", k == first_count,
                      f"K={k} · 표={first_count}")
                why = page.inner_text("#ed-k-why").strip()
                check("고른 근거가 화면에 남는다", "NFTC-211" in why, why[:60])
                page.eval_on_selector(
                    "#panel-edit", "el => el.classList.add('hidden')")

            # ── 보기 모드 기본값
            wv = page.input_value("#ed-worst-view")
            check("나머지 배관망 기본 = 비활성 점선", wv == "dim", wv)

            # ── 걷어낸 것들이 정말 없나 (지운 뒤 JS 가 부르면 죽는다)
            for gone in ("pk-info", "panel-suggest"):
                check(f"#{gone} 없음", page.query_selector(f"#{gone}") is None)
            check("헤드 종류에 그림 없음",
                  page.eval_on_selector_all(
                      ".kinds img", "els => els.length") == 0)
            check("헤드 후보 제안이 찍기 카드 안으로",
                  page.query_selector("#panel-pick #pk-suggest") is not None)

            # ── 자동 추천은 접혀 있고, 제목을 누르면 열린다
            cls = page.get_attribute("#pk-auto-body", "class") or ""
            check("자동 추천은 처음엔 접혀 있다", "hidden" in cls, cls)
            page.eval_on_selector("#panel-pick",
                                  "el => el.classList.remove('hidden')")
            page.click("h2.fold[data-fold=pk-auto-body]")
            page.wait_for_timeout(200)
            cls = page.get_attribute("#pk-auto-body", "class") or ""
            check("제목을 누르면 펼쳐진다", "hidden" not in cls, cls)
            check("펼치면 배관 추천 일괄이 보인다",
                  page.is_visible("#pk-auto-pipe"))
            page.click("h2.fold[data-fold=pk-auto-body]")
            page.wait_for_timeout(200)
            check("다시 누르면 접힌다",
                  "hidden" in (page.get_attribute("#pk-auto-body", "class") or ""))
            page.eval_on_selector("#panel-pick",
                                  "el => el.classList.add('hidden')")

            # ── 진행·레이어도 접혀 있다
            for bid, label in (("log-body", "진행"), ("layers-body", "레이어")):
                cls = page.get_attribute(f"#{bid}", "class") or ""
                check(f"{label}는 처음엔 접혀 있다", "hidden" in cls, cls)
            check("진행 로그가 안 보인다", not page.is_visible("#log"))
            # 작업이 돌면 저절로 열려야 한다 — 큰 도면은 십 분을 넘긴다.
            page.evaluate(
                "() => document.querySelector('h2.fold[data-fold=\"log-body\"]')"
                ".click()")
            page.wait_for_timeout(200)
            check("눌러서 펼치면 로그가 보인다", page.is_visible("#log"))
            check("진행 표(job-chip)가 있다",
                  page.query_selector("#job-chip") is not None)
            page.evaluate(
                "() => document.querySelector('h2.fold[data-fold=\"log-body\"]')"
                ".click()")

            # ── 작업이 실제로 돌 때 저절로 열리나 — 작은 도면 한 장으로.
            #    큰 도면은 파싱이 십 분을 넘겨 검증에 못 쓴다.
            # ★도면이 너무 작으면 잡이 1초 안에 끝나 «도는 동안» 을 관측할 수
            #   없다(실측: 분기티.dxf 0.8s — 열렸다 닫히는 것을 못 잡았다).
            #   몇 초 걸리는 것을 고른다.
            small = next((p for p in [
                ROOT / "samples" / "dxf" / "LH306동_평면도.dxf",
                ROOT / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf",
                ROOT / "samples" / "dxf" / "분기티.dxf",
            ] if p.is_file()), None)
            if small is None:
                check("작업 중 진행 자동열림", False, "시험용 DXF 없음")
            else:
                # 불러오기 = 읽어서 «화면에 띄우기» 까지 (방식과 무관한 공통)
                page.set_input_files("#dxf", str(small))
                page.click("#btn-open")
                # ★진행 자동열림은 «불러오기» 잡에서 본다 — 방식 고르기(수동)는
                #   이제 잡이 없어(추가 0초) 열릴 일이 없다.
                opened = False
                for _ in range(600):          # 25ms × 600 = 15s
                    page.wait_for_timeout(25)
                    if "hidden" not in (page.get_attribute("#log-body", "class") or ""):
                        opened = True
                        break
                check("작업이 돌면 진행이 저절로 열린다", opened, small.name)
                if opened:
                    chip = page.inner_text("#job-chip").strip()
                    check("제목 옆에 도는 단계가 뜬다",
                          bool(chip) and chip != "—", chip)
                # [F-10a · D-F10-1] 질문 없이 흐른다 — 정찰이 성하면 손질까지,
                #   못 쓰겠으면 찍기까지. 둘 중 어디든 «사람 결정 0회» 다.
                landed = None
                for _ in range(900):          # 200ms × 900 = 180s
                    page.wait_for_timeout(200)
                    if page.is_visible("#panel-edit"):
                        landed = "edit"
                        break
                    if page.is_visible("#panel-pick"):
                        landed = "pick"
                        break
                check("불러오면 «묻지 않고» 흘러간다", landed is not None,
                      f"도착 = {landed}")
                check("방식 카드가 뜨지 않는다",
                      page.query_selector("#panel-method") is None)
                closed = False
                for _ in range(90):
                    page.wait_for_timeout(500)
                    if "hidden" in (page.get_attribute("#log-body", "class") or ""):
                        closed = True
                        break
                check("끝나면 다시 접힌다", closed,
                      page.inner_text("#job-line").strip()[:60])
                # ★도면이 실제로 화면에 있어야 한다.
                drawn = page.evaluate(
                    """() => {
                      const c = document.getElementById('cv');
                      const g = c.getContext('2d');
                      const d = g.getImageData(0, 0, c.width, c.height).data;
                      let lit = 0;
                      for (let i = 0; i < d.length; i += 4) {
                        if (d[i] || d[i+1] || d[i+2]) lit++;
                      }
                      return lit;
                    }""")
                check("도착한 화면에 도면이 보인다", drawn > 500,
                      f"칠해진 픽셀 {drawn:,}")
                note = page.inner_text("#start-note").strip().replace("\n", " ")
                check("무엇으로 시작했는지 배너가 말한다", bool(note) and note != "—",
                      note[:70])
                page.click('#panel-advanced h2.fold')
                check("올린 도면 이름이 고급에 뜬다",
                      "선분" in page.inner_text("#adv-file"),
                      page.inner_text("#adv-file").strip().replace("\n", " ")[:60])
                steps_m = page.eval_on_selector_all(
                    "#steps div", "els => els.map(e => e.textContent.trim())")
                check("기본 흐름은 수동 경로 그대로다",
                      steps_m == ["도면 열기", "찍기", "손질", "변환",
                                  "수리계산", "통합"],
                      " · ".join(steps_m))

                # ── [D-F10-2] 자동 차선은 고급 안 한 줄로 살아 있다
                page.reload(wait_until="load")
                page.wait_for_timeout(700)
                page.set_input_files("#dxf", str(small))
                page.click("#btn-open")
                # ★«도착했고 또 조용해질 때까지» 기다린다. 화면만 보고 누르면
                #   아직 도는 잡 위에 클릭을 얹게 되고, 그 클릭은 삼켜진다.
                for _ in range(1800):
                    page.wait_for_timeout(200)
                    if (page.is_visible("#panel-edit")
                            or page.is_visible("#panel-pick")) \
                            and page.is_hidden("#busy"):
                        break
                page.click('#panel-advanced h2.fold')
                page.click("#adv-auto")
                # ★단계바는 라디오를 고르는 순간 이미 바뀐다 — 그것으로 기다리면
                #   파싱이 끝나기 전에 다음으로 넘어간다(실측으로 헤드 0개가 났다).
                #   실제로 열렸는지는 자동 패널이 뜨는 것으로 본다.
                got_auto = False
                for _ in range(300):          # 100ms × 300 = 30s
                    page.wait_for_timeout(100)
                    if (page.is_visible("#panel-auto")
                            and page.is_hidden("#busy")):
                        got_auto = True
                        break
                steps = page.eval_on_selector_all(
                    "#steps div", "els => els.map(e => e.textContent.trim())")
                check("자동으로 열면 자동 화면으로 간다", got_auto, " · ".join(steps))
                if got_auto:
                    check("단계바가 자동 흐름이다",
                          steps == ["도면 열기", "자동 추출", "수리계산", "통합"],
                          " · ".join(steps))
                    # ★알람밸브만 없으면 막힌다. 영역은 «좁히는» 선택이라
                    #   그것 때문에 막히면 안 된다.
                    check("알람밸브 전에는 검출·추출이 막힌다",
                          page.is_disabled("#au-run")
                          and page.is_disabled("#au-heads"))
                    # ★범위 지정은 접이식이 아니라 «3단계» 다 — 어느 구역을
                    #   뽑을지가 결과를 가르므로 접어 두면 안 된다.
                    check("범위 지정이 3단계로 서 있다",
                          page.is_visible("#au-s3")
                          and page.is_visible("#au-zone-draw"),
                          page.inner_text("#au-s3").strip()
                          .replace("\n", " ")[:50])
                    check("접이식 잔재가 없다",
                          page.query_selector("#au-zone-body") is None)
                    steps_lbl = page.eval_on_selector_all(
                        "#panel-auto .step-h",
                        "els => els.map(e => e.textContent.trim())")
                    check("단계가 1~5 로 선다", len(steps_lbl) == 5,
                          " / ".join(s[:16] for s in steps_lbl))
                    want = ["알람밸브", "헤드 검출", "배관망 검출",
                            "범위 지정", "최불리 추출"]
                    for n, w in enumerate(want):
                        check(f"{n + 1}단계가 {w} 다",
                              len(steps_lbl) > n and w in steps_lbl[n],
                              steps_lbl[n][:24] if len(steps_lbl) > n else "")
                    check("안 그리면 «도면 전체» 라고 말한다",
                          "도면 전체" in page.inner_text("#au-s4"),
                          page.inner_text("#au-s4-mark").strip())
                    check("S270 가지치기가 기본으로 켜져 있다",
                          page.is_checked("#au-prune"))

                    # ★「알람밸브 지정 버튼이 어디 있는지 모르겠다」를 받고 세운
                    #   단계 제목 — 실제로 «읽을 수 있는 크기» 인지 잰다.
                    #   .card h2 는 10px·faint 라 눈에 안 들어왔다.
                    sz = page.evaluate(
                        """() => {
                          const h = document.querySelector('#au-s1 .step-h');
                          const cs = getComputedStyle(h);
                          return {px: parseFloat(cs.fontSize),
                                  txt: h.textContent.trim()};
                        }""")
                    check("① 단계 제목이 읽을 수 있는 크기다",
                          sz["px"] >= 12, f"{sz['px']}px · {sz['txt'][:24]}")
                    check("제목에 «알람밸브» 와 «시작 노드» 가 있다",
                          "알람밸브" in sz["txt"] and "시작 노드" in sz["txt"],
                          sz["txt"][:40])
                    check("단추가 무엇을 찍는지 말한다",
                          "알람밸브" in page.inner_text("#au-anchor"),
                          page.inner_text("#au-anchor").strip())
                    check("손질의 «급수 시작» 과 다름을 밝힌다",
                          "급수 시작" in page.inner_text("#au-s1"),
                          page.inner_text("#au-s1").strip()
                          .replace("\n", " ")[:60])
                    page.screenshot(path=str(SHOTS / "4_자동_단계.png"))

                    # ── 자동 경로를 실제로 끝까지 돌린다.
                    #    알람밸브는 «헤드 좌표» 를 쓴다 — bbox 모서리엔 배관이
                    #    없어 25m 결합 한도에 걸린다(실측).
                    ran = page.evaluate(
                        """async (sid) => {
                          const j = async (u, b) => (await fetch(u, {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify(b)})).json();
                          const hs = await j('/api/module-f/auto/heads', {sid});
                          if (!hs.ok || !hs.n) return {ok: false, why: 'heads 0'};
                          const xs = hs.heads.map(h => h.x);
                          const ys = hs.heads.map(h => h.y);
                          const cx = xs.reduce((a,b)=>a+b,0)/xs.length;
                          const cy = ys.reduce((a,b)=>a+b,0)/ys.length;
                          let best = hs.heads[0], bd = Infinity;
                          for (const h of hs.heads) {
                            const d = (h.x-cx)**2 + (h.y-cy)**2;
                            if (d < bd) { bd = d; best = h; }
                          }
                          await j('/api/module-f/auto/anchor',
                                  {sid, x: best.x, y: best.y});
                          return {ok: true, n: hs.n};
                        }""", seen.get("sid"))
                    check("헤드 검출·알람밸브 준비", bool(ran.get("ok")),
                          str(ran.get("why") or f"헤드 {ran.get('n')}개"))

                    # ★Ctrl+Z 가 자동 단계에서도 먹어야 한다 — 「배관망이 잘못
                    #   그려져서 되돌리려는데 못 되돌린다」를 받고 붙인 것.
                    #   영역을 그린 뒤 Ctrl+Z 로 사라지는지 실제로 눌러 본다.
                    # ★알람밸브는 위에서 «API 로» 찍었다 — 화면은 아직 모른다.
                    #   다른 단계를 거쳐 돌아와야 loadAuto 가 돌아 서버 상태를
                    #   화면으로 읽어 온다(같은 단계를 다시 누르면 그냥 돌아간다).
                    page.evaluate(
                        "() => document.querySelectorAll('#steps div')[0].click()")
                    page.wait_for_timeout(300)
                    page.evaluate(
                        "() => document.querySelectorAll('#steps div')[1].click()")
                    page.wait_for_timeout(900)
                    check("알람밸브가 화면에 실렸다",
                          "알람밸브" in page.inner_text("#au-anchor-info"),
                          page.inner_text("#au-anchor-info").strip()
                          .replace("\n", " ")[:40])
                    # 영역 그리기를 켜고 캔버스를 끌어 사각형 하나를 만든다.
                    page.click("#au-zone-draw")
                    box = page.eval_on_selector(
                        "#cv", "el => { const r = el.getBoundingClientRect();"
                               " return {x: r.x, y: r.y, w: r.width, h: r.height}; }")
                    page.mouse.move(box["x"] + box["w"] * 0.30,
                                    box["y"] + box["h"] * 0.30)
                    page.mouse.down()
                    page.mouse.move(box["x"] + box["w"] * 0.60,
                                    box["y"] + box["h"] * 0.60, steps=8)
                    page.mouse.up()
                    page.wait_for_timeout(500)
                    z1 = page.inner_text("#au-zones").strip()
                    check("영역이 그려진다", "영역 1곳" in z1, z1)
                    page.click("#au-zone-draw")          # 무장 해제
                    page.keyboard.press("Control+z")
                    page.wait_for_timeout(700)
                    z2 = page.inner_text("#au-zones").strip()
                    check("Ctrl+Z 로 영역이 되돌아간다",
                          "영역 없음" in z2, f"{z1} → {z2}")
                    srv = page.evaluate(
                        """async (sid) => (await (await fetch(
                            '/api/module-f/auto/state?sid=' + sid)).json())""",
                        seen.get("sid"))
                    zs = srv.get("zones")
                    check("서버의 영역도 같이 되돌아간다",
                          isinstance(zs, list) and len(zs) == 0,
                          f"서버 {len(zs) if isinstance(zs, list) else zs}곳")
                    # ★되돌리기가 «건드리지 않아야 할 것» — 알람밸브는 그대로여야
                    #   한다. 스냅샷이 비어 있으면 서버의 알람밸브까지 지운다.
                    check("되돌려도 알람밸브는 남는다", bool(srv.get("alarm")),
                          str(srv.get("alarm")))
                    if ran.get("ok"):
                        # ★같은 단계를 다시 누르면 gotoStage 가 그냥 돌아간다 —
                        #   다른 단계를 거쳐 와야 재적재가 돈다.
                        page.evaluate(
                            "() => document.querySelectorAll('#steps div')[0].click()")
                        page.wait_for_timeout(300)
                        page.evaluate(
                            "() => document.querySelectorAll('#steps div')[1].click()")
                        page.wait_for_timeout(700)
                        # 영역 없이도 추출이 열려야 한다 — 영역은 선택이다.
                        check("영역 없이도 추출이 열린다",
                              not page.is_disabled("#au-run"),
                              page.inner_text("#au-zones").strip())
                        page.click("#au-heads")
                        page.wait_for_timeout(2500)
                        check("헤드 검출이 화면에 뜬다",
                              "검출된 헤드" in page.inner_text("#au-heads-info"),
                              page.inner_text("#au-heads-info").strip()
                              .replace("\n", " ")[:50])

                        # ★[S270·S310] 배관망 검출 — 최불리 앞에 서야 한다.
                        page.click("#au-network")
                        netted = False
                        for _ in range(900):        # 100ms × 900 = 90s
                            page.wait_for_timeout(100)
                            if page.is_hidden("#busy") and \
                                    "도달 헤드" in page.inner_text("#au-net-info"):
                                netted = True
                                break
                        ni = page.inner_text("#au-net-info").strip()
                        check("배관망 검출이 끝난다", netted,
                              ni.replace("\n", " ")[:80])
                        check("거리 분포가 뜬다 (S310)",
                              "거리" in ni and "최원" in ni,
                              ni.replace("\n", " ")[:80])
                        check("3단계가 ✓ 로 닫힌다",
                              "done" in (page.get_attribute("#au-s3", "class")
                                         or ""),
                              page.inner_text("#au-s3-mark").strip())
                        blue = page.evaluate(_COUNT_BLUE)
                        check("검출한 망이 화면에 깔린다", blue > 200,
                              f"파랑 픽셀 {blue:,}")
                        page.screenshot(path=str(SHOTS / "5_배관망검출.png"))
                        # ★검출 표시가 «빨강» 인가 — 어두운 도면 위에서 티가
                        #   나야 한다. 캔버스 픽셀을 직접 센다.
                        red_lit = page.evaluate(_COUNT_RED)
                        check("검출 헤드가 빨강으로 찍힌다", red_lit > 20,
                              f"빨간 픽셀 {red_lit:,}")
                        page.screenshot(path=str(SHOTS / "1_헤드검출_빨강.png"))

                        bg_before = page.evaluate(_COUNT_BRIGHT_BG)
                        page.click("#au-run")
                        done = False
                        for _ in range(600):        # 100ms × 600 = 60s
                            page.wait_for_timeout(100)
                            if not page.is_disabled("#au-to-design"):
                                done = True
                                break
                        check("자동 추출이 끝난다", done,
                              page.inner_text("#au-summary").strip()[:70])
                        # 끝난 단계는 초록 ✓ 로 표시된다 — 순서가 눈에 보인다.
                        marks = page.evaluate(
                            """() => ['au-s1','au-s2','au-s3','au-s4','au-s5']
                                 .map(id => ({
                                   id, done: document.getElementById(id)
                                          .classList.contains('done'),
                                   m: document.getElementById(id + '-mark')
                                          .textContent.trim()}))""")
                        # 4단계(범위 지정)는 «선택» 이라 안 그렸으면 done 이 아니다.
                        need = [m for m in marks if m["id"] != "au-s4"]
                        check("끝낸 단계가 ✓ 로 표시된다",
                              all(m["done"] for m in need),
                              " · ".join(f"{m['id']}{m['m']}" for m in marks))
                        # ★추출이 끝나면 나머지 도면이 내려가야 한다. 총량이
                        #   아니라 «밝기» 로 잰다 — 화면을 뽑은 자리로 확대하므로
                        #   총량은 오히려 는다(실측 1,293 → 21,150).
                        page.wait_for_timeout(1200)
                        bg_after = page.evaluate(_COUNT_BRIGHT_BG)
                        lit_after = page.evaluate(_COUNT_LIT)
                        cyan = page.evaluate(_COUNT_CYAN)
                        check("추출 뒤 나머지 도면이 흐려진다",
                              bg_after < max(60, cyan * 0.25),
                              f"환한 배경 {bg_before:,} → {bg_after:,} "
                              f"(칠해진 픽셀 전체 {lit_after:,})")
                        check("뽑은 배관망이 살아 있다", cyan > 20,
                              f"청록 픽셀 {cyan:,}")
                        # ★흐리게 내리는 것만으로는 안 드러난다 — 도면이 971m
                        #   인데 설계면적은 25m 라, 화면을 도면 전체로 두면
                        #   결과가 점 하나로 남는다. 뽑은 자리로 맞춰야 한다.
                        check("뽑은 자리로 화면이 맞춰진다", cyan > 300,
                              f"청록 픽셀 {cyan:,} (점 하나면 100 미만)")
                        page.screenshot(path=str(SHOTS / "2_추출후_나머지흐림.png"))

                        # ★티와 교차가 화면에서 갈려야 한다 — 「그냥 봐선
                        #   구분이 안 간다」를 받고 붙인 것.
                        ji = page.inner_text("#au-junc-info").strip()
                        check("이음자리 수가 화면에 적힌다",
                              "분기(티)" in ji and "교차" in ji,
                              ji.replace("\n", " ")[:70])
                        amber = page.evaluate(_COUNT_AMBER)
                        check("분기(티)가 주황으로 찍힌다", amber > 30,
                              f"주황 픽셀 {amber:,}")
                        page.click("#au-junc")           # 꺼 보고
                        page.wait_for_timeout(400)
                        off = page.evaluate(_COUNT_AMBER)
                        check("이음자리 표시를 끌 수 있다", off < amber * 0.4,
                              f"켬 {amber:,} → 끔 {off:,}")
                        page.click("#au-junc")           # 다시 켠다
                        page.wait_for_timeout(400)
                        page.screenshot(path=str(SHOTS / "6_이음자리.png"))
                        page.click("#au-to-design")
                        page.wait_for_timeout(1200)
                    check("자동에서 「표 확정」이 감춰진다",
                          not page.is_visible("#dg-build-row"))
                    check("자동에서 K·규격 입력이 감춰진다",
                          not page.is_visible("#dg-build-inputs"))
                    check("자동에는 「← 자동 추출」 이 뜬다",
                          page.is_visible("#dg-back-auto-row"))
                    legend = page.inner_text("#dg-bore-legend").strip()
                    check("관경 근거를 정직하게 비운다",
                          "자동 경로는" in legend, legend[:60])
                    check("관경 근거 색칠이 꺼지고 잠긴다",
                          page.is_disabled("#dg-bore-color")
                          and not page.is_checked("#dg-bore-color"))
                    # ★검토에서 나온 것 — 자동 경로에서 산출 저장이 영영
                    #   잠겨 있었다(계통도 없이 뽑으면 저장할 길이 없었다).
                    check("자동에서도 «.sdf + .slf 저장» 이 열린다",
                          not page.is_disabled("#dg-emit"))
                    page.click("#dg-emit")
                    saved = False
                    for _ in range(400):        # 100ms × 400 = 40s
                        page.wait_for_timeout(100)
                        if not page.is_disabled("#dg-download"):
                            saved = True
                            break
                    check("자동 경로 산출이 실제로 저장된다", saved,
                          page.inner_text("#status").strip()[:70])
            # ── [F-10a] 기본 흐름 — 업로드 한 번으로 손질까지, 질문 0
            if small is not None:
                print("\n  ── [F-10a] 기본 흐름 (질문 없음)")
                page.reload(wait_until="load")
                page.wait_for_timeout(700)
                clicks = 0
                page.set_input_files("#dxf", str(small))   # 파일 고르기는 클릭 아님
                page.click("#btn-open"); clicks += 1
                # 채택·조립이 잇달아 도는 동안 기다린다. 도착지는 «도면이»
                # 정한다 — 정찰이 성하면 손질, 못 쓰겠으면 찍기. 어느 쪽이든
                # 사람 결정은 0회이고, 그 0회가 이 항목의 수용 기준이다.
                got_edit = False
                landed = None
                for _ in range(1800):       # 200ms × 1800 = 360s
                    page.wait_for_timeout(200)
                    if page.is_visible("#panel-edit") and page.is_hidden("#busy"):
                        got_edit, landed = True, "손질"
                        break
                    if page.is_visible("#panel-pick") and page.is_hidden("#busy"):
                        landed = "찍기"
                        break
                check("★올린 뒤 «사람 결정 0회» 로 도착한다",
                      landed is not None and clicks == 1,
                      f"클릭 {clicks}회(불러오기뿐) · 도착 {landed}")
                page.screenshot(path=str(SHOTS / "0_기본흐름_도착.png"))

                note = page.inner_text("#start-note").strip().replace("\n", " ")
                check("시작 배너가 무엇으로 왔는지 말한다",
                      "자동 인식" in note or "직접 찍어" in note, note[:80])
                if got_edit:
                    check("손질로 갔으면 배너가 되돌릴 길을 알린다",
                          "찍기" in note, note[:80])
                else:
                    # 폴백이면 «왜» 가 있어야 한다 — 묻지 않았으므로 화면이
                    # 대신 말해야 사람이 다음 수를 안다(D-F10-1).
                    check("찍기로 갔으면 배너가 사유를 말한다",
                          "직접 찍" in note or "찾지 못했" in note, note[:80])

                # [D-F10-2] 고급 — 정찰 수치와 채택 기준이 여기 산다.
                page.click('#panel-advanced h2.fold')
                rc = page.inner_text("#adv-recon").strip().replace("\n", " ")
                check("고급에 정찰 수치가 뜬다",
                      "헤드 후보" in rc and "배관 묶음" in rc, rc[:80])
                confs = page.eval_on_selector_all(
                    "#adv-conf option", "els => els.map(e => e.value)")
                check("채택 기준을 화면에서 고를 수 있다 (D-F8-4)",
                      confs and confs[0] == "0.9", " · ".join(confs))
                check("기본 기준은 0.9 다 (D-F8-4)",
                      page.input_value("#adv-conf") == "0.9")
                why0 = page.inner_text("#adv-conf-why").strip()
                if "후보가 없습니다" in why0:
                    check("맞는 후보가 0개면 잠기고 사유를 말한다",
                          page.is_disabled("#adv-readopt"), why0[:70])
                else:
                    check("기본 기준으로 찍을 것이 있다",
                          not page.is_disabled("#adv-readopt"), why0[:70])
                page.select_option("#adv-conf", "0.75")
                page.wait_for_timeout(200)
                why1 = page.inner_text("#adv-conf-why").strip()
                check("기준을 바꾸면 예정 수가 따라간다", why1 != why0, why1[:60])

                # [D-F10-3] 확정 지점은 손질이지만, 되돌리기로 찍기까지 내려간다.
                if got_edit:
                    page.click("#steps div:nth-child(2)")
                    page.wait_for_selector("#panel-pick:not(.hidden)",
                                           timeout=30000)
                check("단계바로 찍기까지 내려간다",
                      page.is_visible("#panel-pick"))

                # ★기준을 낮춰 «다시 채택» — 자동으로 낮추지 않는 것이 규약이라
                #   (D-F8-4 기본 0.9 유지), 낮추는 것은 사람의 몫이다. 이 길이
                #   기본 흐름과 같은 채택 경로(adoptRun)를 탄다.
                if not page.is_disabled("#adv-readopt"):
                    page.click("#adv-readopt")
                    for _ in range(1800):
                        page.wait_for_timeout(200)
                        if page.is_visible("#panel-pick") and page.is_hidden("#busy"):
                            break

                info = page.inner_text("#pk-adopt-info").strip().replace("\n", " ")
                check("채택 결과가 화면에 남는다",
                      "재료" in info and "헤드" in info, info[:90])
                page.screenshot(path=str(SHOTS / "3_채택직후_찍기화면.png"))
                st = page.evaluate(
                    """async (sid) => (await (await fetch(
                        '/api/module-f/world?sid=' + sid)).json()).state""",
                    seen.get("sid"))
                check("재료가 실제로 찍혀 있다",
                      len((st or {}).get("materials") or []) > 0,
                      f"{len((st or {}).get('materials') or [])}묶음")
                check("헤드가 실제로 찍혀 있다", (st or {}).get("n_heads", 0) > 0,
                      f"{(st or {}).get('n_heads')}픽")
                check("클릭 기록이 남는다 (사람 클릭과 같은 경로)",
                      (st or {}).get("n_clicks", 0) > 0,
                      f"{(st or {}).get('n_clicks')}회")

                # 유령 — 점선으로 남고, 그 자리를 누르면 정상 찍기로 간다.
                ghost_n = page.evaluate(
                    "() => document.getElementById('pk-adopt-info')"
                    ".textContent.match(/유령 (\\d+)/)")
                check("유령 수가 화면에 적힌다", ghost_n is not None,
                      str(ghost_n))
                check("낮은 신뢰도 토글이 있다",
                      page.query_selector("#pk-show-low") is not None)

                # ★유령 위의 클릭을 «후보 제외» 가 가로채면 안 된다 — 사람이
                #   직접 찍으려고 누르는 자리다. 코드로 확인한다.
                # 화면 JS 는 정적 파일로 떼어져 있다 — 인라인 textContent 만
                # 훑으면 «있는데 없다» 고 나온다. 붙어 있는 script 를 인라인·
                # 외부 가리지 않고 모아 본다.
                guard = page.evaluate(
                    """async () => {
                      const els = [...document.querySelectorAll('script')];
                      const parts = await Promise.all(els.map(async e => {
                        if (!e.src) return e.textContent || '';
                        try { return await (await fetch(e.src)).text(); }
                        catch (_) { return ''; }
                      }));
                      const s = parts.join('');
                      return s.includes('if (S.ghosts && S.ghosts.has(i)) continue;');
                    }""")
                check("유령 위 클릭을 가로채지 않는다", bool(guard))

                # 수동 차선 회귀 — 같은 화면에서 기존 단추가 그대로 동작한다.
                steps_x = page.eval_on_selector_all(
                    "#steps div", "els => els.map(e => e.textContent.trim())")
                check("혼합은 수동 흐름을 그대로 쓴다",
                      steps_x == ["도면 열기", "찍기", "손질", "변환",
                                  "수리계산", "통합"], " · ".join(steps_x))
                check("「배관망 구성」 으로 사람이 확정한다 (D-F8-5)",
                      page.is_visible("#pk-next"))

            # 표가 몇 개인지가 아니라 «한 판 안에서 겹치는가» 를 본다 — 패널이
            # 다르면 동시에 보이지 않으므로 중복이 아니다(손질 패널에 셋이
            # 몰려 있던 것이 문제였다).
            worst = page.evaluate(
                """() => {
                  let mx = 0, who = '';
                  for (const p of document.querySelectorAll('.side section')) {
                    const n = [...p.querySelectorAll('.tag')]
                      .filter(e => e.textContent.trim() === 'MODULE A').length;
                    if (n > mx) { mx = n; who = p.id; }
                  }
                  return {n: mx, id: who};
                }""")
            check("한 패널 안에서 MODULE A 표가 겹치지 않는다",
                  worst["n"] <= 1, f"{worst['id']} 에 {worst['n']}개")

            # ── 수직 전개 값이 «창» 으로 뜨나 (모듈 E 의 대화상자 자리)
            check("변환 칸이 옆판에 없다",
                  page.query_selector("#panel-conv #conv-fields") is None,
                  "옆판에 그대로 박혀 있다")
            check("변환 칸이 창 안에 있다",
                  page.query_selector("#conv-modal #conv-fields") is not None)
            check("창은 처음엔 닫혀 있다",
                  "hidden" in (page.get_attribute("#conv-modal", "class") or ""))
            # 칸을 채워 넣고 창을 열어 본다 — 서버에서 받은 칸이 그려지는지.
            page.evaluate(
                "() => document.getElementById('conv-fields').innerHTML ="
                " '<label class=\"f\"><span>① (m)</span>"
                "<input type=\"text\" data-key=\"t1\" value=\"1.5\"></label>'")
            page.eval_on_selector("#panel-conv",
                                  "el => el.classList.remove('hidden')")
            page.click("#btn-conv-fields")
            page.wait_for_timeout(200)
            check("단추를 누르면 창이 뜬다",
                  "hidden" not in (page.get_attribute("#conv-modal", "class") or ""))
            # 취소하면 고친 값이 되돌아가야 한다.
            page.fill("#conv-modal input[data-key=t1]", "9.9")
            page.click("#conv-cancel")
            page.wait_for_timeout(200)
            back = page.eval_on_selector(
                "#conv-fields input[data-key=t1]", "el => el.value")
            check("취소하면 값이 되돌아간다", back == "1.5", back)
            check("취소하면 창이 닫힌다",
                  "hidden" in (page.get_attribute("#conv-modal", "class") or ""))
            # 확인은 값을 남긴다.
            page.click("#btn-conv-fields")
            page.fill("#conv-modal input[data-key=t1]", "2.75")
            page.click("#conv-ok")
            page.wait_for_timeout(200)
            kept = page.eval_on_selector(
                "#conv-fields input[data-key=t1]", "el => el.value")
            check("확인하면 값이 남는다", kept == "2.75", kept)
            summ = page.inner_text("#conv-summary").strip()
            check("옆판에 채운 칸 요약이 뜬다", "1" in summ, summ[:40])
            page.eval_on_selector("#panel-conv",
                                  "el => el.classList.add('hidden')")

            # ── 콘솔 오류 0
            check("콘솔 오류 없음", not errors,
                  " | ".join(errors[:3]) if errors else "")

            page.screenshot(path=str(ROOT / "data" / "_ui_verify.png"),
                            full_page=False)
            # 변환 창 — 서버가 주는 진짜 칸으로 한 장.
            page.evaluate("() => { window.__f = null; }")
            fields = page.evaluate(
                "async () => (await (await fetch("
                "'/api/module-f/convert/fields')).json())")
            if fields and fields.get("groups"):
                page.evaluate(
                    """(d) => {
                      const box = document.getElementById('conv-fields');
                      box.innerHTML = '';
                      for (const g of d.groups) {
                        const h = document.createElement('div');
                        h.className = 'grp'; h.textContent = g.title;
                        box.appendChild(h);
                        if (g.diagram) {
                          const f = document.createElement('div');
                          f.className = 'grpfig';
                          const im = document.createElement('img');
                          im.className = 'diagram';
                          im.src = '/api/module-f/diagram/' + g.diagram;
                          f.appendChild(im); box.appendChild(f);
                        }
                        for (const fl of g.fields) {
                          const lb = document.createElement('label');
                          lb.className = 'f';
                          const sp = document.createElement('span');
                          sp.textContent = fl.label;
                          const inp = document.createElement('input');
                          inp.type = 'text'; inp.dataset.key = fl.key;
                          const dv = d.defaults[fl.key];
                          if (dv !== null && dv !== undefined) inp.value = String(dv);
                          lb.append(sp, inp); box.appendChild(lb);
                        }
                      }
                      document.getElementById('conv-modal')
                              .classList.remove('hidden');
                    }""", fields)
                page.wait_for_timeout(700)
                page.screenshot(path=str(ROOT / "data" / "_ui_conv_modal.png"))
                page.eval_on_selector(
                    "#conv-modal", "el => el.classList.add('hidden')")
            # 숨은 패널도 한 장 — 설명을 걷어낸 뒤 배치가 어색하지 않은지 눈으로.
            for pid, shot in (("panel-pick", "_ui_pick.png"),
                              ("panel-edit", "_ui_edit.png"),
                              ("panel-design", "_ui_design.png"),
                              ("panel-merge", "_ui_merge.png"),
                              ("panel-sub", "_ui_sub.png"),
                              ("panel-conv", "_ui_conv.png")):
                page.eval_on_selector(
                    ".side",
                    "el => [...el.children].forEach(c => c.classList.add('hidden'))")
                page.eval_on_selector(
                    f"#{pid}", "el => el.classList.remove('hidden')")
                page.wait_for_timeout(150)
                page.screenshot(path=str(ROOT / "data" / shot))
            browser.close()
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  - " + f)
        return 1
    print("PASS — 화면이 브라우저에서 정상으로 뜬다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
