# -*- coding: utf-8 -*-
"""[오너 2026-09-22] 계통도 칸 화면 — 미리보기 힙이 종전과 같은 답을 내는가.

B1F 평면도를 계통도 칸에 올리면 경로 그래프가 노드 3만 개다. 종전 미리보기는
선형 탐색이라 마우스를 움직일 때마다 1.2초 멈췄다(주석: «절점이 수백 개라 선형
탐색으로 충분»). 힙으로 바꾸되 **답은 한 글자도 같아야** 한다 — 미리보기와 추출이
다른 길을 그리면 화면이 거짓말을 한다. 동률(같은 거리)이 많은 그래프에서 종전
선형 탐색과 경로를 대조한다.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 종전 선형 탐색 그대로(2026-09-22 이전 module_f.js) — 대조 기준.
_LINEAR = r"""
function subPathLinear(g, a, b) {
  if (!g || a < 0 || b < 0) return null;
  if (a === b) return [a];
  const pen = Number(g.forced_penalty_mm);
  if (!isFinite(pen)) return null;
  const n = g.nodes.length;
  const dist = new Float64Array(n).fill(Infinity);
  const prev = new Int32Array(n).fill(-1);
  const done = new Uint8Array(n);
  dist[a] = 0;
  for (;;) {
    let u = -1, bd = Infinity;
    for (let i = 0; i < n; i++) {
      if (!done[i] && dist[i] < bd) { bd = dist[i]; u = i; }
    }
    if (u < 0 || u === b) break;
    done[u] = 1;
    for (const [v, len, forced] of g.adj[u]) {
      const w = (len || 0) + (forced ? pen : 0);
      if (dist[u] + w < dist[v]) { dist[v] = dist[u] + w; prev[v] = u; }
    }
  }
  if (!isFinite(dist[b])) return null;
  const out = [];
  for (let k = b; k >= 0; k = prev[k]) out.push(k);
  return out.reverse();
}
"""

_DRIVER = r"""
let seed = 12345;
const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
let checked = 0, differ = [];
for (let t = 0; t < 60; t++) {
  const n = 20 + Math.floor(rnd() * 180);
  const nodes = []; for (let i = 0; i < n; i++) nodes.push([i, 0]);
  const adj = nodes.map(() => []);
  const m = n * 2;
  for (let e = 0; e < m; e++) {
    const a = Math.floor(rnd() * n), b = Math.floor(rnd() * n);
    if (a === b) continue;
    const len = 1 + Math.floor(rnd() * 3);          // 1~3 — 같은 거리(동률)가 많다
    const forced = rnd() < 0.1 ? 1 : 0;
    adj[a].push([b, len, forced]); adj[b].push([a, len, forced]);
  }
  S.subGraph = { nodes, adj, forced_penalty_mm: 7 };
  for (let q = 0; q < 25; q++) {
    const a = Math.floor(rnd() * n), b = Math.floor(rnd() * n);
    const p1 = JSON.stringify(subPath(a, b));
    const p2 = JSON.stringify(subPathLinear(S.subGraph, a, b));
    checked++;
    if (p1 !== p2) differ.push({ t, a, b, heap: p1, linear: p2 });
  }
}
console.log(JSON.stringify({ checked, differ: differ.slice(0, 5), n_differ: differ.length }));
"""


def test_heap_preview_path_is_identical_to_the_linear_one():
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    js = open(os.path.join(_ROOT, "static", "module_f.js"), encoding="utf-8").read()
    i = js.index("  function subPath(a, b) {")
    src = js[i:js.index("\n  }\n", i) + 4]
    assert "hpush" in src, "미리보기가 힙이 아니다"
    prog = "\n".join(["const S = {subGraph: null};", src, _LINEAR, _DRIVER])
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr
    got = json.loads(out.stdout.strip().splitlines()[-1])
    assert got["checked"] == 1500
    assert got["n_differ"] == 0, got["differ"]
