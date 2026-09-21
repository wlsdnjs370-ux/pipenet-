# -*- coding: utf-8 -*-
"""Stage 2 헤드 reveal 소요시간 A/B — 현재 작업본 vs 직전 커밋(HEAD).

큰 도면(배관 ghost 8000 + 헤드 3000)을 주입해 stage2_switch~stage2_done
실측 시간과 render() 호출 횟수를 잰다.
"""
from __future__ import annotations

import functools
import http.server
import socketserver
import subprocess
import sys
import threading
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

from flask import Flask, render_template  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

TPL = "remote30_prototype.html"
BEFORE_TPL = "_bench_before.html"
(BASE / "templates" / BEFORE_TPL).write_bytes(subprocess.run(
    ["git", "show", f"HEAD:templates/{TPL}"], cwd=BASE,
    capture_output=True, check=True).stdout)

app = Flask(__name__, template_folder=str(BASE / "templates"),
            static_folder=str(BASE / "static"))
pages = {}
with app.test_request_context():
    for tag, name in (("after", TPL), ("before", BEFORE_TPL)):
        out = BASE / "data" / f"_bench_{tag}.html"
        out.write_text(render_template(name), encoding="utf-8")
        pages[tag] = out
(BASE / "templates" / BEFORE_TPL).unlink()

handler = functools.partial(http.server.SimpleHTTPRequestHandler,
                            directory=str(BASE))
httpd = socketserver.TCPServer(("127.0.0.1", 0), handler)
httpd.RequestHandlerClass.log_message = lambda *a, **k: None
port = httpd.server_address[1]
threading.Thread(target=httpd.serve_forever, daemon=True).start()

BENCH = """
async (N) => {
  // 배관망 ghost — 큰 도면 규모
  const ghost = [];
  for (let i = 0; i < 8000; i++) {
    const x = (i % 100) * 500, y = Math.floor(i / 100) * 500;
    ghost.push({t: "L", l: "PIPE", p: [x, y, x + 400, y]});
  }
  const heads = [];
  for (let i = 0; i < N; i++) {
    const x = (i % 60) * 800, y = Math.floor(i / 60) * 800;
    heads.push({t: "B", l: "_head_bbox", p: [x - 100, y - 100, x + 100, y + 100],
                k: "circle_signature", c: 0.8, n: "", i, pos: [x, y]});
  }
  state.bbox = {x_min: -1000, y_min: -1000, x_max: 50000, y_max: 45000};

  let renders = 0;
  const _r = window.render;
  window.render = function () { renders++; return _r.apply(this, arguments); };

  handleEvent({type: "entities", stage: 1, entities: ghost});
  handleEvent({type: "entities", stage: 2, entities: heads});
  const queued = animQueue.size();
  const t0 = performance.now();
  await new Promise((res) => {
    const poll = () => (animQueue.size() === 0 ? res() : setTimeout(poll, 20));
    poll();
  });
  const ms = performance.now() - t0;
  window.render = _r;
  return {ms: Math.round(ms), renders, queued,
          revealed: state.revealed_heads_count,
          heads: state.detected_heads.length};
}
"""

N = 3000
res = {}
errors: list[str] = []
with sync_playwright() as pw:
    br = pw.chromium.launch()
    for tag in ("before", "after"):
        page = br.new_page(viewport={"width": 1400, "height": 900})
        page.set_default_timeout(600000)
        page.on("pageerror", lambda e: errors.append(f"{e}"))
        page.goto(f"http://127.0.0.1:{port}/data/_bench_{tag}.html")
        page.wait_for_timeout(1000)
        res[tag] = page.evaluate(BENCH, N)
        print(f"{tag:7s} {res[tag]}")
        page.close()
    br.close()

httpd.shutdown()
for p in pages.values():
    p.unlink(missing_ok=True)

if errors:
    print("\nJS 오류:")
    for e in errors[:10]:
        print("  " + e)

b, a = res["before"], res["after"]
speed = b["ms"] / max(1, a["ms"])
print(f"\n헤드 {N}개 · reveal  {b['ms']/1000:.1f}s → {a['ms']/1000:.1f}s"
      f"  ({speed:.1f}배)   render {b['renders']} → {a['renders']}")
ok = (speed >= 5.0 and a["heads"] == N and a["revealed"] is None and not errors)
print("\n" + ("PASS" if ok else "FAIL"))
sys.exit(0 if ok else 1)
