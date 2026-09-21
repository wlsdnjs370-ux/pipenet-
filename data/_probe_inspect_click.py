# -*- coding: utf-8 -*-
"""급수 시작 클릭이 왜 «아무것도 안 잡히나» — 화소·좌표·모드를 함께 본다."""
from __future__ import annotations

import os
import sys

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5065")
PASSWORD = os.environ["LOGIN_PASSWORD"]
KEY = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도_도면정리(1)")

with sync_playwright() as pw:
    b = pw.chromium.launch()
    page = b.new_page(viewport={"width": 1680, "height": 1000})
    page.on("pageerror", lambda e: print("   pageerror:", e))
    page.goto(f"{BASE}/module-f", wait_until="load")
    if page.query_selector("input[type=password]"):
        page.fill("input[type=password]", PASSWORD)
        page.click("button[type=submit]")
        page.wait_for_load_state("load")
    if "/module-f" not in page.url:
        page.goto(f"{BASE}/module-f", wait_until="load")
    page.wait_for_function(
        "document.querySelector('#saved').options.length > 0", timeout=30_000)
    page.select_option("#saved", KEY)
    page.click("#btn-reopen")
    page.wait_for_selector("#panel-edit:not(.hidden)", timeout=300_000)
    page.wait_for_timeout(1500)

    info = page.evaluate("""() => {
      const cv = document.getElementById('cv');
      const rect = cv.getBoundingClientRect();
      const g = cv.getContext('2d');
      const d = g.getImageData(0, 0, cv.width, cv.height).data;
      const buckets = {};
      let n = 0;
      for (let y = 2; y < cv.height - 2; y += 2)
        for (let x = 2; x < cv.width - 2; x += 2) {
          const i = (y * cv.width + x) * 4;
          const s = d[i] + d[i+1] + d[i+2];
          if (s > 20) { n++; const k = (s / 60 | 0) * 60; buckets[k] = (buckets[k]||0)+1; }
        }
      return {cw: cv.width, ch: cv.height, rw: rect.width, rh: rect.height,
              dpr: window.devicePixelRatio, lit: n, buckets};
    }""")
    print("캔버스:", info["cw"], "x", info["ch"], "| CSS:",
          round(info["rw"]), "x", round(info["rh"]), "| DPR:", info["dpr"])
    print("밝은 화소:", info["lit"], "| 밝기 분포:", info["buckets"])

    print("배경 밑그림 체크:", page.evaluate(
        "() => { const e = document.getElementById('ed-bg');"
        " return e ? e.checked : 'ed-bg 없음'; }"))

    page.click('.emode[data-mode="급수시작위치"]')
    page.wait_for_timeout(300)
    print("모드 단추 on:", page.eval_on_selector_all(
        ".emode", "els => els.filter(e => e.classList.contains('on'))"
                  ".map(e => e.dataset.mode)"))

    spots = page.evaluate("""() => {
      const cv = document.getElementById('cv');
      const g = cv.getContext('2d');
      const d = g.getImageData(0, 0, cv.width, cv.height).data;
      const r = cv.width / cv.getBoundingClientRect().width;
      const hits = [];
      for (let y = 2; y < cv.height - 2; y += 2)
        for (let x = 2; x < cv.width - 2; x += 2) {
          const i = (y * cv.width + x) * 4;
          if (d[i] + d[i+1] + d[i+2] > 200) hits.push([x, y, d[i], d[i+1], d[i+2]]);
        }
      const step = Math.max(1, (hits.length / 8) | 0);
      const out = [];
      for (let k = 0; k < hits.length && out.length < 8; k += step)
        out.push({x: hits[k][0] / r, y: hits[k][1] / r,
                  rgb: [hits[k][2], hits[k][3], hits[k][4]]});
      return {n: hits.length, out};
    }""")
    print("밝은(>200) 화소:", spots["n"])
    box = page.eval_on_selector("#cv", "e => { const r = e.getBoundingClientRect();"
                                      " return {x: r.x, y: r.y}; }")
    for i, sp in enumerate(spots["out"]):
        page.mouse.move(box["x"] + sp["x"], box["y"] + sp["y"])
        page.wait_for_timeout(60)
        coord = page.inner_text("#coord")
        page.mouse.click(box["x"] + sp["x"], box["y"] + sp["y"])
        page.wait_for_timeout(500)
        print(f"  [{i}] css({sp['x']:.0f},{sp['y']:.0f}) rgb{sp['rgb']} "
              f"{coord} → {page.inner_text('#status')[:56]}")
    b.close()
