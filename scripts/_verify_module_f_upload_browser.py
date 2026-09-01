# -*- coding: utf-8 -*-
"""모듈 F 업로드 길 — 실제 브라우저에서 확인한다.

`node --check` 는 «함수-지역 헬퍼를 다른 스코프에서 부르는» 종류의 회귀를 못
잡는다(이 저장소가 이미 겪었다). 업로드 경로는 이번에 세 군데가 한꺼번에
바뀌었으므로 — 압축 · 진행률 · 이른 그리기 — 브라우저에서 직접 본다.

확인하는 것:
    ① 콘솔 오류 0
    ② 가림막 문구가 «압축 %» → «업로드 %» → «읽는 중» 으로 실제로 흐른다
    ③ 진행 막대가 셀 수 있는 구간에서 `det`(실측 채움)로 바뀐다
    ④ 도면이 실제로 그려진다(캔버스에 검정 아닌 화소가 있다)

실행:
    MF_BASE=http://127.0.0.1:5065 LOGIN_PASSWORD=… python scripts/_verify_module_f_upload_browser.py
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
# 8MB 를 넘어야 압축 길로 간다(GZIP_MIN_BYTES).
# ★«이른 그리기» 까지 보려면 A 의 파스 캐시가 **없는** 도면이어야 한다 —
#   캐시가 더우면 정찰이 0.5초에 끝나 창이 안 열린다. MF_DXF 로 갈아 준다.
BIG_DXF = Path(os.environ.get(
    "MF_DXF",
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"))

problems: list[str] = []


def bad(msg: str) -> None:
    problems.append(msg)
    print("   !!", msg)


def check(name: str, ok: bool, detail: str = "") -> bool:
    print(f"  [{'OK  ' if ok else 'FAIL'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        problems.append(name + (f" · {detail}" if detail else ""))
    return ok


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={"width": 1680, "height": 1000})
    page.on("console", lambda m: bad(f"console.{m.type}: {m.text}")
            if m.type == "error" else None)
    page.on("pageerror", lambda e: bad(f"pageerror: {e}"))

    # ★몸통 크기는 Playwright 로 못 잰다 — post_data_buffer 도 content-length 도
    #   sizes().requestBodySize 도 이 조합에서는 0/없음으로 온다. 그래서 XHR 이
    #   **실제로 건네받은 것**을 화면 안에서 붙잡는다. 그것이 곧 전선 바이트다.
    #   ★`add_init_script` 는 준 문자열을 **본문 그대로** 넣는다(evaluate 처럼
    #     함수로 알아보고 불러 주지 않는다). 화살표 함수를 주면 값만 만들어지고
    #     아무 일도 안 일어난다 — 실측으로 한 번 그렇게 조용히 지나갔다.
    page.add_init_script("""
      (() => {
        const send = XMLHttpRequest.prototype.send;
        window.__mfSent = null;
        XMLHttpRequest.prototype.send = function (body) {
          try {
            if (body instanceof FormData) {
              const f = body.get('dxf_file');
              if (f) window.__mfSent = { name: f.name, size: f.size };
            }
          } catch (e) {}
          return send.apply(this, arguments);
        };
      })();
    """)

    print("[0] 로그인")
    page.goto(f"{BASE}/module-f", wait_until="load")
    if "login" in page.url or page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PASSWORD)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    if "/module-f" not in page.url:
        page.goto(f"{BASE}/module-f", wait_until="load")
    check("모듈 F 화면", "/module-f" in page.url, page.url)
    check("upload_stream.js 실렸다",
          page.evaluate("typeof window.UploadStream === 'object'"))

    print("[1] 가림막 문구를 처음부터 지켜본다")
    # 클릭 «전에» 감시를 건다 — 압축 문구는 순식간에 지나간다.
    page.evaluate("""() => {
      window.__mfLog = [];
      const t = document.getElementById('busy-text');
      const bar = document.getElementById('busy-bar');
      const box = document.getElementById('busy');
      const snap = () => window.__mfLog.push({
        text: t.textContent,
        det: !!(bar && bar.classList.contains('det')),
        width: bar ? bar.style.width : '',
        peek: box.classList.contains('peek'),
        hidden: box.classList.contains('hidden'),
      });
      new MutationObserver(snap).observe(t, {childList: true, characterData: true, subtree: true});
      new MutationObserver(snap).observe(box, {attributes: true});
      snap();
    }""")

    print(f"[2] 업로드 — {BIG_DXF.name} ({BIG_DXF.stat().st_size / 1048576:.1f} MB)")
    page.set_input_files("#dxf", str(BIG_DXF))
    page.click("#btn-open")
    # 도면이 화면에 앉을 때까지. 상태줄이 «선분 …» 을 적으면 다 그린 것이다.
    page.wait_for_function(
        "document.querySelector('#status').textContent.includes('선분')",
        timeout=240_000)
    print("   상태줄:", page.inner_text("#status")[:110])

    log = page.evaluate("() => window.__mfLog")
    texts = [r["text"] for r in log]
    print("   가림막 문구 흐름:")
    seen = []
    for r in log:
        key = (r["text"][:22], r["det"], r["peek"])
        if key in seen:
            continue
        seen.append(key)
        print(f"     · {r['text'][:70]}  "
              f"[막대 {'채움 ' + r['width'] if r['det'] else '무한'}"
              f"{' · 비침' if r['peek'] else ''}]")

    check("압축 진행률을 보여줬다", any("압축 중" in t for t in texts),
          f"{sum('압축 중' in t for t in texts)}회")
    # ★루프백에서는 몸통이 한 번에 다 나가 중간 % 가 안 뜬다(XHR 이 곧바로
    #   loaded==total 을 준다). 그것은 결함이 아니라 «올리는 데 시간이 안
    #   걸렸다» 는 뜻이므로, 중간 % 든 «다 보냈다» 든 하나는 떠야 한다로 본다.
    check("업로드 구간을 말해 줬다",
          any("업로드 중" in t or "서버가 도면을 받는 중" in t for t in texts),
          f"중간% {sum('업로드 중' in t for t in texts)}회")
    check("셀 수 있는 구간은 실측 막대로 그렸다",
          any(r["det"] for r in log))
    check("퍼센트가 실제로 움직였다",
          len({r["width"] for r in log if r["det"]}) >= 3,
          f"서로 다른 너비 {len({r['width'] for r in log if r['det']})}개")

    sent = page.evaluate("() => window.__mfSent") or {}
    size = int(sent.get("size") or 0)
    orig = BIG_DXF.stat().st_size
    check("몸통이 압축돼 올라갔다", bool(size) and size < orig * 0.5,
          f"{size / 1048576:.1f} MB / 원본 {orig / 1048576:.1f} MB"
          + (f" · {orig / size:.1f}배 작아짐 · 보낸 이름 {sent.get('name')}"
             if size else " · 못 쟀다"))

    # ★이번 고침의 핵심 — 도면을 «잡이 끝나기 전에» 그렸나.
    #   정찰 캐시가 더운 도면에서는 창이 안 열리므로 실패로 세지 않고 알린다.
    peeked = any(r["peek"] for r in log)
    said = any("도면은 이미 화면에 있습니다" in t for t in texts)
    if peeked or said:
        check("정찰이 도는 동안 도면을 먼저 그렸다", True,
              f"가림막 비침 {peeked} · 문구 {said}")
    else:
        print("  [....] 이른 그리기 창이 안 열렸다 — 정찰이 캐시로 즉시 끝난 "
              "도면이면 정상이다(MF_DXF 로 캐시 없는 도면을 주면 보인다)")

    print("[3] 도면이 실제로 그려졌나")
    nonblack = page.evaluate("""() => {
      const cv = document.getElementById('cv');
      const g = cv.getContext('2d');
      const d = g.getImageData(0, 0, cv.width, cv.height).data;
      let n = 0;
      for (let i = 0; i < d.length; i += 4) if (d[i] || d[i+1] || d[i+2]) n++;
      return n;
    }""")
    check("캔버스에 도면이 있다", nonblack > 2000, f"검정 아닌 화소 {nonblack:,}")
    page.screenshot(path=str(SHOT / "upload_after.png"))

    check("콘솔 오류 0", not [p for p in problems if p.startswith(("console", "pageerror"))])
    browser.close()

print()
if problems:
    print(f"실패 {len(problems)}건")
    for p in problems:
        print("  -", p)
    raise SystemExit(1)
print("모두 통과.")
