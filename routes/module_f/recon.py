# -*- coding: utf-8 -*-
"""[F-8a] 정찰 — A 가 도면에서 «무엇을 알아보는가» 를 한 번에 뽑는다.

정찰은 **제안일 뿐이다.** 여기서 나온 것은 board 에 닿지 않는다 — 채택은
언제나 E 의 확정 경로(`PickSession.click`)를 거친다(D-F8-3). 이 파일은 그
제안을 «만드는» 쪽만 맡고, 스펙으로 «번역하는» 쪽은 `adopt.py` 가 맡는다.

`/pick/suggest` 와 열기 잡이 이 함수 하나를 나눠 쓴다. 둘이 각자 A 를 부르면
같은 도면을 두 번 읽게 되고 — 더 나쁘게 — 언젠가 한쪽만 고쳐져 화면이 서로
다른 후보 수를 말하게 된다.

A 의 `parse_dxf_bundle_cached` 를 쓰므로 파스 캐시가 자동 차선과 공유된다.
정찰을 먼저 돌린 도면은 이후 자동 차선 진입이 그만큼 싸진다(실측 86배).
"""
from __future__ import annotations

import time
from pathlib import Path

from routes.module_f.common import _r1

# 정찰이 세는 카테고리 — 찍기가 실제로 쓰는 셋만.
RECON_CATS = ("PIPE", "HEAD", "ALARM")

# D-F8-4 — 채택 임계. 화면에서 조절하지만 기본은 여기서 나온다.
CONF_HIGH = 0.9
CONF_MID = 0.75

# 띠 이름은 `/pick/suggest` 가 종전에 쓰던 것 그대로다. 바꾸면 그 응답을
# 읽던 화면이 조용히 «0개» 를 그린다.
BAND_HIGH = f"높음(≥{CONF_HIGH})"
BAND_MID = f"중간(≥{CONF_MID})"
BAND_LOW = "낮음"


def band_of(conf) -> str:
    """신뢰도 하나가 어느 띠인가 — 세는 쪽과 고르는 쪽이 같은 자를 쓴다."""
    c = float(conf)
    if c >= CONF_HIGH:
        return BAND_HIGH
    return BAND_MID if c >= CONF_MID else BAND_LOW


def count_bands(cands) -> dict:
    out = {BAND_HIGH: 0, BAND_MID: 0, BAND_LOW: 0}
    for c in (cands or ()):
        out[band_of(c.get("conf", 0.0))] += 1
    return out


def bundle_counts(world) -> dict:
    """레이어 사전이 추천한 묶음 수 — 찍기 화면이 세는 것과 같은 값.

    `_world_payload` 가 이미 세어 둔 `cats` 를 읽는다. 여기서 다시 세면
    분류 규칙이 두 벌이 되어 화면과 카드가 다른 말을 하게 된다.
    """
    cats = ((world or {}).get("cats") or {})
    return {c: int(cats.get(c, 0) or 0) for c in RECON_CATS}


def run_recon(dxf, *, world=None, tag: str = "정찰") -> dict:
    """도면 한 장을 A 방식으로 훑어 «후보» 를 낸다. board 는 건드리지 않는다.

    반환:
        bundles     {PIPE, HEAD, ALARM} 추천 묶음 수 (world 를 줬을 때만 의미 있음)
        heads       [{x, y, conf, kind, layer}, …] 신뢰도 내림차순
        bands       {높음, 중간, 낮음} 개수
        elapsed_ms  걸린 시간

    실패는 올린다 — 부르는 쪽이 «열기는 성공» 으로 감쌀지 정한다(F-8a-4).
    """
    t0 = time.perf_counter()
    # 모듈 A 는 저장소 루트의 읽기 전용 참조다 — import 만 한다.
    import remote30_prototype as A

    print(f"[{tag}] 모듈 A 인식(R1~R5) 시작 — 도면을 A 방식으로 읽는 중…")
    bundle = A.parse_dxf_bundle_cached(Path(str(dxf)))
    layers = {ly.get("name"): A._categorize_layer(ly.get("name") or "")
              for ly in (bundle.layers or [])}
    heads = A.detect_heads(bundle.entities, layers)

    cands = [{"x": _r1(h.pos[0]), "y": _r1(h.pos[1]),
              "conf": round(float(h.confidence), 2),
              "kind": str(h.kind), "layer": str(h.layer or "")}
             for h in heads]
    cands.sort(key=lambda c: -c["conf"])

    bands = count_bands(cands)
    print(f"[{tag}] 후보 {len(cands)}개 · {bands}")
    return {"bundles": bundle_counts(world), "heads": cands, "bands": bands,
            "elapsed_ms": int((time.perf_counter() - t0) * 1000)}


def recon_view(rec) -> dict:
    """화면이 카드에 그릴 만큼만 — 후보 수천 개를 매번 내려보내지 않는다.

    후보 좌표 자체는 오버레이가 따로 받아 간다(`/api/module-f/recon?heads=1`).
    """
    if not rec:
        return {"state": "none"}
    if rec.get("error"):
        return {"state": "error", "error": rec["error"]}
    return {"state": "ok", "bundles": rec.get("bundles") or {},
            "bands": rec.get("bands") or {},
            "n": len(rec.get("heads") or ()),
            "elapsed_ms": rec.get("elapsed_ms")}
