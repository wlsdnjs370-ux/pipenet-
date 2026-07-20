# -*- coding: utf-8 -*-
"""다이소 양주 패키지 — 답안 프로파일 + 점수 엔진 자기검증 드라이버.

실행:
    python calibration/daiso_yangju.py

단계
====
1) 답안 SDF 25개 프로파일 일람 (위상/연장/구경/헤드)
2) 점수 엔진 자기검증
   · round-trip(SDF→KFP→재파싱) 은 PASS·오차≈0 이어야  (자가 무결성)
   · 교란 3종은 각각 FAIL 로 떨어져야               (민감도)
       P1 연장 ×1.10   → 연장 FAIL
       P2 헤드 3개 제거 → 위상 FAIL
       P3 파이프 15% 제거 → 위상(연결성) FAIL
"""
from __future__ import annotations

import copy
import glob
import os
import sys
import tempfile
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

sys.path.insert(0, str(Path(__file__).resolve().parent))  # score_network 동일 폴더

from kfp_sdf_converter import parse_sdf, parse_kfp, emit_kfp  # noqa: E402
import score_network as sn  # noqa: E402

ANSWER_DIR = _ROOT / "수리계산 참고용 도서" / "1. 저수조_아성다이소 양주허브센터" / "수리계산 원본"


def _answer_files():
    return sorted(glob.glob(str(ANSWER_DIR / "**" / "*.sdf"), recursive=True))


def step1_profiles():
    print("=" * 100)
    print("STEP 1 — 답안 SDF 프로파일")
    print("=" * 100)
    files = _answer_files()
    print(f"답안 SDF: {len(files)}개\n")
    for f in files:
        name = os.path.relpath(f, ANSWER_DIR)
        try:
            net = parse_sdf(f)
            p = sn.profile(net)
            tag = "REMOTE" if "REMOTE" in name.upper() else ""
            print(f"  {name}")
            print(f"      {p.as_row()}  {tag}")
        except Exception as exc:  # noqa: BLE001
            print(f"  {name}\n      !! 파싱 실패: {exc}")
    print()


def _pick_reference():
    """자기검증 기준 — 헤드 30개짜리 REMOTE 1건."""
    for f in _answer_files():
        if "REMOTE" in f.upper() and "K-160" in f:
            return f
    return _answer_files()[0]


def _roundtrip(net):
    """SDF→KFP→재파싱 — 좌표는 표시배율로 바뀌지만 length_m/위상은 보존되어야."""
    with tempfile.TemporaryDirectory() as td:
        kfp = Path(td) / "rt.kfp"
        emit_kfp(net, kfp)
        return parse_kfp(kfp)


def _perturb_length(net, factor=1.10):
    n = copy.deepcopy(net)
    for p in n.pipes.values():
        p.length_m *= factor
    return n


def _perturb_drop_heads(net, k=3):
    n = copy.deepcopy(net)
    heads = [nid for nid, c in n.nodes.items() if c.kind in ("head", "nozzle")]
    drop = set(heads[:k])
    for nid in drop:
        n.nodes.pop(nid, None)
    n.pipes = {pid: p for pid, p in n.pipes.items()
               if p.start not in drop and p.end not in drop}
    return n


def _perturb_drop_pipes(net, frac=0.15):
    n = copy.deepcopy(net)
    pids = list(n.pipes.keys())
    cut = max(1, int(len(pids) * frac))
    for pid in pids[:cut]:
        n.pipes.pop(pid, None)
    return n


def step2_selfcheck():
    print("=" * 100)
    print("STEP 2 — 점수 엔진 자기검증")
    print("=" * 100)
    ref = _pick_reference()
    print(f"기준 답안: {os.path.relpath(ref, ANSWER_DIR)}\n")
    net = parse_sdf(ref)

    results = []

    # (A) round-trip — PASS 기대
    rt = _roundtrip(net)
    r = sn.score(rt, net)
    ok = r.overall_pass
    results.append(("round-trip 무결성 (PASS 기대)", ok, True, r))
    print("[A] round-trip (SDF→KFP→재파싱) — PASS 기대")
    print(r.report(), "\n")

    # (B) 교란 — FAIL 기대
    cases = [
        ("P1 연장 ×1.10 (연장 FAIL 기대)", _perturb_length(net, 1.10), "length"),
        ("P2 헤드 3개 제거 (위상 FAIL 기대)", _perturb_drop_heads(net, 3), "topology"),
        ("P3 파이프 15% 제거 (위상 FAIL 기대)", _perturb_drop_pipes(net, 0.15), "topology"),
    ]
    for label, mutated, _which in cases:
        r = sn.score(mutated, net)
        failed = not r.overall_pass
        results.append((label, failed, True, r))
        print(f"[B] {label}")
        print(r.report(), "\n")

    # 요약
    print("-" * 100)
    print("자기검증 요약")
    all_ok = True
    for label, got, want, _r in results:
        mark = "OK " if got == want else "WRONG"
        if got != want:
            all_ok = False
        print(f"  [{mark}] {label}")
    print("-" * 100)
    print("자기검증 전체:", "통과 — 점수 엔진 신뢰 가능" if all_ok else "실패 — 엔진 점검 필요")
    return all_ok


def _remote_files():
    return [f for f in _answer_files() if "REMOTE" in f.upper()]


def step3_bands():
    """prior 밴드 학습 + 검증기 자기검증.

    답안-키 1:1 채점이 불가한 코퍼스의 차선책. 정상 답안(REMOTE)에서 '배관망스러움'
    분포를 학습하고, (A) 학습에 쓴 정상 답안은 통과해야 하고 (B) 교란망은 경고/실패로
    잡혀야 한다.
    """
    print("=" * 100)
    print("STEP 3 — prior 밴드 학습 + 검증기 자기검증")
    print("=" * 100)

    # 밴드는 전 패키지 스프링클러 REMOTE 에서 학습(옥내소화전 제외) — validate_sdf 와 동일.
    import validate_sdf as vs
    remotes = vs._remote_answer_files()
    bands = vs.learn_default_bands()
    print(f"학습 표본: 스프링클러 REMOTE 답안 {len(remotes)}건 (전 패키지)")
    print(bands.report(), "\n")

    # (A) 정상 답안 — OK/WARN 기대 (FAIL 없어야)
    print("[A] 정상 답안 검증 — FAIL 없어야 정상")
    a_ok = True
    for f in remotes:
        rep = sn.validate(parse_sdf(f), bands)
        if rep.has_fail:
            a_ok = False
        print(f"  {rep.verdict:<4} {os.path.relpath(f, ANSWER_DIR)}")
        if rep.has_fail or rep.has_warn:
            for name, status, detail in rep.checks:
                if status != "OK":
                    print(f"         [{status}] {name}: {detail}")
    print()

    # (B) 교란망 — 적어도 WARN 으로 잡혀야
    ref = _pick_reference()
    net = parse_sdf(ref)
    print(f"[B] 교란망 검증 (기준 {os.path.relpath(ref, ANSWER_DIR)}) — 경고/실패로 잡혀야")
    cases = [
        ("헤드 절반 제거", _perturb_drop_heads(net, max(1, len(net.nodes) // 6))),
        ("파이프 15% 제거", _perturb_drop_pipes(net, 0.15)),
        ("연장 ×3 (축척 오류)", _perturb_length(net, 3.0)),
        ("루프 주입", _inject_loop(net)),
    ]
    b_ok = True
    for label, mutated in cases:
        rep = sn.validate(mutated, bands)
        caught = rep.has_fail or rep.has_warn
        if not caught:
            b_ok = False
        print(f"  [{'OK ' if caught else 'MISS'}] {label} → {rep.verdict}")
        for name, status, detail in rep.checks:
            if status != "OK":
                print(f"         [{status}] {name}: {detail}")
    print()

    print("-" * 100)
    all_ok = a_ok and b_ok
    print("검증기 자기검증:",
          "통과 — 밴드 검증기 신뢰 가능" if all_ok else "실패 — 검증기 점검 필요")
    return all_ok


def _inject_loop(net):
    """단말 두 개를 이어 루프를 만든다 — 트리성 위반 교란."""
    n = copy.deepcopy(net)
    from collections import defaultdict
    deg = defaultdict(int)
    for p in n.pipes.values():
        deg[p.start] += 1
        deg[p.end] += 1
    leaves = [nid for nid in n.nodes if deg[nid] == 1]
    if len(leaves) >= 2:
        a, b = leaves[0], leaves[1]
        pid = f"LOOP_{a}_{b}"
        proto = next(iter(n.pipes.values()))
        newp = copy.deepcopy(proto)
        newp.id = pid
        newp.start, newp.end = a, b
        newp.length_m = 1.0
        n.pipes[pid] = newp
    return n


if __name__ == "__main__":
    step1_profiles()
    ok2 = step2_selfcheck()
    print()
    ok3 = step3_bands()
    sys.exit(0 if (ok2 and ok3) else 1)
