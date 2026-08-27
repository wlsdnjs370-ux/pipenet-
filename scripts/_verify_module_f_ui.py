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
    print(f"임시 서버 :{port} 기동 중…")
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
            check("고르기 전 단계바는 «도면 열기» 뿐",
                  steps == ["도면 열기"], " · ".join(steps))
            check("고르기 전에는 열 수 없다", page.is_disabled("#btn-open"))

            page.check("#open-m-manual")
            page.wait_for_timeout(200)
            steps = page.eval_on_selector_all(
                "#steps div", "els => els.map(e => e.textContent.trim())")
            check("수동을 고르면 수동 흐름이 뜬다",
                  steps == ["도면 열기", "찍기", "손질", "변환", "수리계산", "통합"],
                  " · ".join(steps))
            page.check("#open-m-auto")
            page.wait_for_timeout(200)
            steps = page.eval_on_selector_all(
                "#steps div", "els => els.map(e => e.textContent.trim())")
            check("자동을 고르면 자동 흐름으로 바뀐다",
                  steps == ["도면 열기", "자동 추출", "수리계산", "통합"],
                  " · ".join(steps))
            check("고른 뒤에는 열 수 있다", not page.is_disabled("#btn-open"))
            page.check("#open-m-manual")     # 아래 검사는 수동 기준이다
            page.wait_for_timeout(200)

            # ── 슬롯 탭이 그려졌나 (세션 전에는 안내문)
            slots = page.inner_text("#slots").strip()
            check("슬롯 자리 있음", bool(slots), slots[:50])

            # ── 추출 방식 갈림 (A 자동 / E 수동)
            check("방식 고르는 자리가 있다",
                  page.query_selector("#open-method") is not None)
            check("자동도 고를 수 있다",
                  page.query_selector("#open-m-auto") is not None)
            check("자동 추출 패널이 있다",
                  page.query_selector("#panel-auto") is not None)
            for aid in ("au-anchor", "au-zone-arm", "au-heads", "au-run",
                        "au-k-preset", "au-k"):
                check(f"#{aid} 존재", page.query_selector(f"#{aid}") is not None)
            page.check("#open-m-manual")     # 아래 검사는 수동 기준이다

            # ── 갈 수 없는 단계는 막힌다. ★도면을 열기 «전» 에 봐야 한다 —
            #    열고 나면 찍기가 실제로 도달 가능해져 이 검사가 뜻을 잃는다.
            page.click("#steps div:nth-child(2)")     # 찍기 — 재료 없음
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
                page.set_input_files("#dxf", str(small))
                page.click("#btn-open")
                opened = False
                for _ in range(400):          # 25ms × 400 = 10s
                    page.wait_for_timeout(25)
                    if "hidden" not in (page.get_attribute("#log-body", "class") or ""):
                        opened = True
                        break
                check("작업이 돌면 진행이 저절로 열린다", opened, small.name)
                if opened:
                    chip = page.inner_text("#job-chip").strip()
                    check("제목 옆에 도는 단계가 뜬다",
                          bool(chip) and chip != "—", chip)
                closed = False
                for _ in range(90):
                    page.wait_for_timeout(500)
                    if "hidden" in (page.get_attribute("#log-body", "class") or ""):
                        closed = True
                        break
                check("끝나면 다시 접힌다", closed,
                      page.inner_text("#job-line").strip()[:60])

                # ── 자동 방식으로 같은 도면을 다시 연다 — 단계바가 갈리나
                page.reload(wait_until="load")
                page.wait_for_timeout(700)
                page.check("#open-m-auto")
                page.set_input_files("#dxf", str(small))
                page.click("#btn-open")
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
                    check("영역·알람밸브 전에는 실행이 막힌다",
                          page.is_disabled("#au-run"))

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
                          const pad = 1000;
                          await j('/api/module-f/auto/anchor',
                                  {sid, x: best.x, y: best.y});
                          await j('/api/module-f/auto/zones', {sid, zones: [[
                            Math.min(...xs)-pad, Math.min(...ys)-pad,
                            Math.max(...xs)+pad, Math.max(...ys)+pad]]});
                          return {ok: true, n: hs.n};
                        }""", seen.get("sid"))
                    check("헤드 후보·알람밸브·영역 준비", bool(ran.get("ok")),
                          str(ran.get("why") or f"후보 {ran.get('n')}개"))
                    if ran.get("ok"):
                        # 화면이 서버 상태를 되살리는가 — 영역이 되돌아와야 한다.
                        # ★같은 단계를 다시 누르면 gotoStage 가 그냥 돌아간다 —
                        #   다른 단계를 거쳐 와야 재적재가 돈다.
                        page.evaluate(
                            "() => document.querySelectorAll('#steps div')[0].click()")
                        page.wait_for_timeout(300)
                        page.evaluate(
                            "() => document.querySelectorAll('#steps div')[1].click()")
                        page.wait_for_timeout(700)
                        check("영역이 서버에서 되살아난다",
                              "영역 1곳" in page.inner_text("#au-zones"),
                              page.inner_text("#au-zones").strip())
                        check("준비되면 실행이 열린다",
                              not page.is_disabled("#au-run"))
                        page.click("#au-run")
                        done = False
                        for _ in range(600):        # 100ms × 600 = 60s
                            page.wait_for_timeout(100)
                            if not page.is_disabled("#au-to-design"):
                                done = True
                                break
                        check("자동 추출이 끝난다", done,
                              page.inner_text("#au-summary").strip()[:70])
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
