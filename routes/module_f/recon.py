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


# ── [F-11a · D-F11-2] 지배 띠 채택 — 절대 임계를 «도면 분포» 로 ──────
#
# 절대 임계 0.9 는 퇴화한다. A 의 신뢰도가 사실상 이진값이기 때문이다 —
# `KNOWN_HEAD_BLOCKS` 블록 참조만 0.95 를 받고, 나머지 HEAD 도형은 0.70~0.89
# 에 몰린다. 그래서 블록을 쓴 도면과 직접 작도한 도면 사이에서 0.9 가
# «대부분» 과 «거의 없음» 으로 갈린다. 실측:
#
#     대명동 단위세대   높음  87 / 119   (73%)  → 정상
#     B1F 현장조사      높음  72 / 3,338 ( 2%)  → 최불리 2개로 퇴화
#     LH306동 평면도    높음   0 / 42    ( 0%)  → 조립 불가(§16 게이트로 막힘)
#
# 그래서 임계를 도면의 «분포» 가 정하게 한다. 높음이 지배적이면 높음만,
# 아니면 중간까지 — 규칙은 결정적이고, 발동한 규칙은 반드시 화면에 적는다
# (조용한 규칙 전환은 새 은닉 오류다).
DOMINANT_BAND_RATIO = 0.5


def dominant_band(bands: dict) -> dict:
    """채택 임계를 도면 분포로 정한다. 화면이 그대로 읽어 적을 수 있는 모양.

    반환:
        conf_min  채택 임계(0.9 / 0.75) 또는 None(채택할 것이 없음)
        rule      "high" | "high_mid" | "none"
        n         그 임계로 채택되는 후보 수
        why       사람이 읽을 한 줄 — 배너·카드가 이 문장을 그대로 쓴다
    """
    hi = int((bands or {}).get(BAND_HIGH) or 0)
    mid = int((bands or {}).get(BAND_MID) or 0)
    low = int((bands or {}).get(BAND_LOW) or 0)
    total = hi + mid + low
    if total and hi / total >= DOMINANT_BAND_RATIO:
        return {"conf_min": CONF_HIGH, "rule": "high", "n": hi,
                "why": f"후보 대부분이 높음 띠(블록 기반)라 높음만 "
                       f"채택했습니다 — {hi:,}개"}
    if hi + mid > 0:
        return {"conf_min": CONF_MID, "rule": "high_mid", "n": hi + mid,
                "why": f"이 도면은 후보 대부분이 중간 띠(블록 미사용)라 "
                       f"중간까지 채택했습니다 — {hi + mid:,}개"}
    return {"conf_min": None, "rule": "none", "n": 0,
            "why": "자동 인식이 찍을 만한 헤드 후보를 못 찾았습니다."}


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
    # ★번들이 이미 매긴 분류를 그대로 쓴다. 같은 dict 를 돌면서 이름으로 다시
    #   매기면 파서 끝의 레이어 승격이 사라진다 — 그 분류가 곧바로 detect_heads
    #   로 들어가므로 헤드 검출이 직접 손해를 본다(B1F 「6-소화-FLX (연결관
    #   승격)」 28 entity 가 OTHER 로 되돌아갔다).
    layers = {ly.get("name"): (ly.get("auto_category") or "OTHER")
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
    bands = rec.get("bands") or {}
    return {"state": "ok", "bundles": rec.get("bundles") or {},
            "bands": bands,
            "n": len(rec.get("heads") or ()),
            # [F-11a] 이 도면의 분포가 정한 채택 임계와 «그 이유». 화면이
            # 고르는 것이 아니라 여기서 정해 내려보낸다 — 규칙이 한 곳에만
            # 있어야 카드와 배너가 다른 말을 하지 않는다.
            "adopt": dominant_band(bands),
            "elapsed_ms": rec.get("elapsed_ms")}
