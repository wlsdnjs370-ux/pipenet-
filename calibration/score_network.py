# -*- coding: utf-8 -*-
"""배관망 추출 검증용 점수 엔진 — "문제(도면) → 답안(수리계산)" 코퍼스 대조.

목적
====
``수리계산 참고용 도서`` 의 답안 SDF(사람이 PIPENET 으로 손수 만든 정답)와, 우리
추출기가 같은 도면에서 뽑은 망을 **같은 자(尺)** 로 재서 일반화 격차를 숫자로 본다.
ML 학습이 아니라 결정론적 추출기를 정답지에 맞춰 캘리브레이션하기 위한 측정 도구.

이 모듈은 두 ``CommonNetwork`` 를 받아
  1) profile(net)  — 위상/연장/구경/헤드 통계 프로파일
  2) score(a, b)   — 보수적 허용오차로 PASS/FAIL + 오차지표
를 낸다. 서버/변환기 코드는 건드리지 않는 독립 도구.

보수적(conservative) 허용오차 설계
==================================
정답 PIPENET 리포트는 정밀하므로, 작은 실측 오차도 FAIL 로 떨어뜨려 "추출기가
좋다"는 거짓 PASS 를 막는다. 즉 의심스러우면 FAIL.
  · 연결성(components): 답안과 동일(통상 단일망)            [엄격]
  · 헤드 수(head_count): 정확히 일치                        [엄격]
  · 단말(leaf) 수: 정확히 일치                              [엄격]
  · 노드 수: ±2%                                            [타이트]
  · 접합점(degree>=3) 수: ±1                                [타이트]
  · 총연장: ±3%                                             [타이트]
  · 정합 구간 개별연장: 90% 이상이 ±5% 이내, 단일 최대 <10% [타이트]
  · 구경(nominal) 집합: 동일, 빈별 개수 ±1                  [타이트]
"""
from __future__ import annotations

import math
from collections import Counter, defaultdict, deque
from dataclasses import dataclass, field
from pathlib import Path

# ── 보수적 허용오차 상수 (한곳에서 조정) ─────────────────────────────────────
TOL_TOTAL_LENGTH_PCT = 0.03      # 총연장 허용 ±3%
TOL_SEGMENT_LENGTH_PCT = 0.05    # 정합 구간 개별연장 밴드 ±5%
TOL_SEGMENT_PASS_FRAC = 0.90     # 구간 90% 이상이 밴드 내여야 PASS
TOL_SEGMENT_MAX_PCT = 0.10       # 단일 구간 최대 편차 < 10%
TOL_NODE_COUNT_PCT = 0.02        # 노드 수 ±2%
TOL_JUNCTION_COUNT_ABS = 1       # 접합점 수 ±1
TOL_DIAMETER_BIN_ABS = 1         # 구경 빈별 개수 ±1
COORD_MATCH_FACTOR = 0.5         # 좌표 정합 허용거리 = 중앙 구간연장 × 이 값


# ────────────────────────────────────────────────────────────────────────────
# 프로파일
# ────────────────────────────────────────────────────────────────────────────
@dataclass
class NetProfile:
    n_nodes: int
    n_edges: int
    n_components: int
    head_count: int
    leaf_count: int          # degree == 1
    junction_count: int      # degree >= 3
    total_length_m: float
    seg_min_m: float
    seg_med_m: float
    seg_max_m: float
    kind_counts: dict
    diameter_counts: dict    # nominal_mm -> 개수

    def as_row(self) -> str:
        return (f"nodes={self.n_nodes:<4} edges={self.n_edges:<4} "
                f"comp={self.n_components} head={self.head_count:<3} "
                f"leaf={self.leaf_count:<3} junc={self.junction_count:<3} "
                f"len={self.total_length_m:>8.2f}m "
                f"seg(min/med/max)={self.seg_min_m:.2f}/{self.seg_med_m:.2f}/{self.seg_max_m:.2f}")


def _degrees(net) -> dict:
    deg = defaultdict(int)
    for nid in net.nodes:
        deg[nid] = 0
    for p in net.pipes.values():
        if p.start in deg:
            deg[p.start] += 1
        if p.end in deg:
            deg[p.end] += 1
    return deg


def _components(net) -> int:
    adj = defaultdict(set)
    for p in net.pipes.values():
        adj[p.start].add(p.end)
        adj[p.end].add(p.start)
    seen = set()
    comps = 0
    # 파이프에 닿는 노드만 연결성 대상 — 고립 노드는 별도 component 로 셈
    all_nodes = set(net.nodes.keys())
    for start in all_nodes:
        if start in seen:
            continue
        comps += 1
        q = deque([start])
        seen.add(start)
        while q:
            cur = q.popleft()
            for nb in adj.get(cur, ()):  # 노드에 없을 수도 있는 끝점 방어
                if nb not in seen:
                    seen.add(nb)
                    q.append(nb)
    return comps


def profile(net) -> NetProfile:
    deg = _degrees(net)
    head_count = sum(1 for c in net.nodes.values() if c.kind in ("head", "nozzle"))
    leaf = sum(1 for nid in net.nodes if deg[nid] == 1)
    junc = sum(1 for nid in net.nodes if deg[nid] >= 3)
    lengths = sorted(p.length_m for p in net.pipes.values())
    total = sum(lengths)
    if lengths:
        seg_min, seg_max = lengths[0], lengths[-1]
        seg_med = lengths[len(lengths) // 2]
    else:
        seg_min = seg_med = seg_max = 0.0
    kinds = Counter(c.kind for c in net.nodes.values())
    diam = Counter(int(round(p.nominal_mm)) for p in net.pipes.values())
    return NetProfile(
        n_nodes=len(net.nodes), n_edges=len(net.pipes),
        n_components=_components(net),
        head_count=head_count, leaf_count=leaf, junction_count=junc,
        total_length_m=total, seg_min_m=seg_min, seg_med_m=seg_med, seg_max_m=seg_max,
        kind_counts=dict(kinds), diameter_counts=dict(sorted(diam.items())),
    )


# ────────────────────────────────────────────────────────────────────────────
# 정렬 (노드 대응) — 교차 출처 망을 비교하려면 노드 짝을 찾아야 한다.
# ────────────────────────────────────────────────────────────────────────────
def align_by_id(a, b) -> dict:
    """양쪽이 동일 노드 id 를 공유하는 경우(round-trip) — id 그대로 매칭."""
    return {nid: nid for nid in a.nodes if nid in b.nodes}


def align_by_coords(a, b) -> dict:
    """좌표 기반 탐욕 최근접 매칭(교차 출처용).

    각 망을 중심정렬·스케일정규화(중앙 구간연장=1) 후, a 의 노드를 b 의 가장
    가까운 미사용 노드에 짝짓는다. 허용거리(COORD_MATCH_FACTOR) 밖이면 미매칭.
    보수적: 못 찾으면 매칭하지 않고 unmatched 로 남긴다(거짓 짝 금지).

    주의 — 이건 동일 좌표계/유사 형상 전제의 1차 매처다. 회전·축척이 크게 다른
    실제 추출 vs 답안에는 강건한 등록(registration)이 따로 필요(Layer 2).
    """
    def _norm(net):
        xs = [c.x for c in net.nodes.values()]
        ys = [c.y for c in net.nodes.values()]
        if not xs:
            return {}, 1.0
        cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
        lengths = sorted(p.length_m for p in net.pipes.values() if p.length_m > 0)
        scale = lengths[len(lengths) // 2] if lengths else 1.0
        scale = scale or 1.0
        pts = {nid: ((c.x - cx) / scale, (c.y - cy) / scale)
               for nid, c in net.nodes.items()}
        return pts, scale

    pa, _ = _norm(a)
    pb, _ = _norm(b)
    used = set()
    mapping = {}
    thresh = COORD_MATCH_FACTOR  # 정규화 좌표계에서 중앙연장=1 기준
    for nid, (x, y) in pa.items():
        best = None
        best_d = thresh
        for mid, (mx, my) in pb.items():
            if mid in used:
                continue
            d = math.hypot(x - mx, y - my)
            if d < best_d:
                best_d = d
                best = mid
        if best is not None:
            mapping[nid] = best
            used.add(best)
    return mapping


# ────────────────────────────────────────────────────────────────────────────
# 점수
# ────────────────────────────────────────────────────────────────────────────
@dataclass
class ScoreResult:
    topology_pass: bool
    length_pass: bool
    diameter_pass: bool
    metrics: dict = field(default_factory=dict)
    failures: list = field(default_factory=list)

    @property
    def overall_pass(self) -> bool:
        return self.topology_pass and self.length_pass and self.diameter_pass

    def report(self) -> str:
        ov = "PASS" if self.overall_pass else "FAIL"
        lines = [f"[{ov}] topology={'O' if self.topology_pass else 'X'} "
                 f"length={'O' if self.length_pass else 'X'} "
                 f"diameter={'O' if self.diameter_pass else 'X'}"]
        for k, v in self.metrics.items():
            lines.append(f"    {k}: {v}")
        for f in self.failures:
            lines.append(f"    ✗ {f}")
        return "\n".join(lines)


def _matched_segments(extracted, answer, node_map):
    """양쪽 망에서 '같은 두 노드를 잇는 파이프' 쌍을 찾아 (len_ext, len_ans) 리스트."""
    def edge_index(net):
        idx = {}
        for p in net.pipes.values():
            key = frozenset((p.start, p.end))
            idx[key] = p.length_m
        return idx

    ext_idx = edge_index(extracted)
    ans_idx = edge_index(answer)
    pairs = []
    for ext_key, ext_len in ext_idx.items():
        a, b = tuple(ext_key) if len(ext_key) == 2 else (None, None)
        if a is None:
            continue
        ma, mb = node_map.get(a), node_map.get(b)
        if ma is None or mb is None:
            continue
        ans_key = frozenset((ma, mb))
        if ans_key in ans_idx:
            pairs.append((ext_len, ans_idx[ans_key]))
    return pairs


def score(extracted, answer, node_map=None) -> ScoreResult:
    """extracted(우리 추출) 를 answer(정답) 기준으로 보수적 채점."""
    if node_map is None:
        node_map = align_by_id(extracted, answer)
        if len(node_map) < 0.5 * len(answer.nodes):
            node_map = align_by_coords(extracted, answer)

    pe, pa = profile(extracted), profile(answer)
    failures = []
    metrics = {}

    # ── 위상 ──
    topo_ok = True
    metrics["components"] = f"{pe.n_components} vs {pa.n_components}"
    if pe.n_components != pa.n_components:
        topo_ok = False
        failures.append(f"연결성분 {pe.n_components} != 답안 {pa.n_components}")

    metrics["head_count"] = f"{pe.head_count} vs {pa.head_count}"
    if pe.head_count != pa.head_count:
        topo_ok = False
        failures.append(f"헤드수 {pe.head_count} != 답안 {pa.head_count} (엄격)")

    metrics["leaf_count"] = f"{pe.leaf_count} vs {pa.leaf_count}"
    if pe.leaf_count != pa.leaf_count:
        topo_ok = False
        failures.append(f"단말수 {pe.leaf_count} != 답안 {pa.leaf_count} (엄격)")

    node_tol = max(1, math.ceil(pa.n_nodes * TOL_NODE_COUNT_PCT))
    metrics["node_count"] = f"{pe.n_nodes} vs {pa.n_nodes} (±{node_tol})"
    if abs(pe.n_nodes - pa.n_nodes) > node_tol:
        topo_ok = False
        failures.append(f"노드수 {pe.n_nodes} vs 답안 {pa.n_nodes} (허용 ±{node_tol})")

    metrics["junction_count"] = f"{pe.junction_count} vs {pa.junction_count} (±{TOL_JUNCTION_COUNT_ABS})"
    if abs(pe.junction_count - pa.junction_count) > TOL_JUNCTION_COUNT_ABS:
        topo_ok = False
        failures.append(f"접합점수 {pe.junction_count} vs 답안 {pa.junction_count} (허용 ±{TOL_JUNCTION_COUNT_ABS})")

    # ── 연장 ──
    len_ok = True
    if pa.total_length_m > 0:
        tot_err = abs(pe.total_length_m - pa.total_length_m) / pa.total_length_m
    else:
        tot_err = 0.0 if pe.total_length_m == 0 else 1.0
    metrics["total_length"] = (f"{pe.total_length_m:.2f} vs {pa.total_length_m:.2f}m "
                               f"(오차 {tot_err*100:.2f}%, 허용 {TOL_TOTAL_LENGTH_PCT*100:.0f}%)")
    if tot_err > TOL_TOTAL_LENGTH_PCT:
        len_ok = False
        failures.append(f"총연장 오차 {tot_err*100:.2f}% > {TOL_TOTAL_LENGTH_PCT*100:.0f}%")

    pairs = _matched_segments(extracted, answer, node_map)
    metrics["matched_segments"] = f"{len(pairs)} / 답안 {pa.n_edges}"
    if pairs:
        within = 0
        max_dev = 0.0
        for le, la in pairs:
            dev = abs(le - la) / la if la > 0 else (0.0 if le == 0 else 1.0)
            max_dev = max(max_dev, dev)
            if dev <= TOL_SEGMENT_LENGTH_PCT:
                within += 1
        frac = within / len(pairs)
        metrics["segment_band"] = (f"{frac*100:.1f}% 내 (요구 {TOL_SEGMENT_PASS_FRAC*100:.0f}%), "
                                   f"최대편차 {max_dev*100:.1f}% (허용 <{TOL_SEGMENT_MAX_PCT*100:.0f}%)")
        if frac < TOL_SEGMENT_PASS_FRAC:
            len_ok = False
            failures.append(f"구간 밴드내 비율 {frac*100:.1f}% < {TOL_SEGMENT_PASS_FRAC*100:.0f}%")
        if max_dev >= TOL_SEGMENT_MAX_PCT:
            len_ok = False
            failures.append(f"구간 최대편차 {max_dev*100:.1f}% >= {TOL_SEGMENT_MAX_PCT*100:.0f}%")
    else:
        # 정합 구간 0 — 위상이 PASS 인데도 구간 매칭이 안 되면 보수적으로 FAIL
        if pa.n_edges > 0:
            len_ok = False
            failures.append("정합 구간 0개 — 노드 대응 실패(보수적 FAIL)")

    # ── 구경 ──
    diam_ok = True
    set_e = set(pe.diameter_counts)
    set_a = set(pa.diameter_counts)
    metrics["diameter_set"] = f"{sorted(set_e)} vs {sorted(set_a)}"
    if set_e != set_a:
        diam_ok = False
        failures.append(f"구경 집합 불일치: {sorted(set_e ^ set_a)}")
    for d in set_a & set_e:
        if abs(pe.diameter_counts[d] - pa.diameter_counts[d]) > TOL_DIAMETER_BIN_ABS:
            diam_ok = False
            failures.append(f"{d}mm 개수 {pe.diameter_counts[d]} vs {pa.diameter_counts[d]} (허용 ±{TOL_DIAMETER_BIN_ABS})")

    return ScoreResult(topology_pass=topo_ok, length_pass=len_ok,
                       diameter_pass=diam_ok, metrics=metrics, failures=failures)


# ════════════════════════════════════════════════════════════════════════════
# prior 밴드 검증기 — 답안-키 1:1 채점이 불가한 코퍼스용 차선책.
#
# 도면(전건물 합본)과 답안(추상 스키매틱)은 좌표계가 달라 score() 식 1:1 대조가
# 불가능하다. 대신 답안 다수에서 "정상 설계구역은 이런 모양"이라는 분포(prior)를
# 뽑아, 우리 추출 결과가 그 정상 범위(band) 안에 드는지 자동 점검한다.
#
# 정답을 외우는 게 아니라 "배관망스러움"을 학습한다 — 정답 없는 새 도면에도 적용.
# 한계: 범위 안이면 통과시킬 뿐 "정확히 맞다"는 보장은 아니다(정상처럼 생긴 오답은
# 통과 가능). score() 의 대체가 아니라, 그게 불가능할 때의 보조 점검.
# ════════════════════════════════════════════════════════════════════════════

# 표준 호칭경(mm) — KS 강관/배관 통용. 이 밖의 구경은 추출 잡음 의심.
STANDARD_NOMINAL_MM = frozenset({
    15, 20, 25, 32, 40, 50, 65, 80, 90, 100, 125, 150, 200, 250, 300, 350,
})
PRIOR_RATIO_MARGIN = 0.15   # 학습한 비율·길이 밴드에 주는 여유 ±15%


@dataclass
class BandSet:
    """답안 코퍼스에서 학습한 '정상 설계구역' 분포 밴드.

    구조 불변식(정수)은 관측 [min,max] 그대로, 비율·길이는 ±margin 여유를 둔다.
    """
    n_samples: int
    forest_excess: tuple      # n_edges-(n_nodes-comp)  — 트리면 0
    leaf_minus: tuple         # leaf-(head+comp)         — 헤드가 단말이면 0
    junc_minus: tuple         # junc-(head-comp)         — 트리+헤드단말이면 0
    node_per_head: tuple      # n_nodes / head
    total_per_head_m: tuple   # total_length / head
    seg_min_m: tuple
    seg_med_m: tuple
    diameters_seen: frozenset
    margin: float = PRIOR_RATIO_MARGIN

    def report(self) -> str:
        def r(t): return f"[{t[0]:.2f}, {t[1]:.2f}]"
        return "\n".join([
            f"prior 밴드 (표본 {self.n_samples}건, 여유 ±{self.margin*100:.0f}%)",
            f"  forest_excess(트리성)   {r(self.forest_excess)}  (0=순수 트리)",
            f"  leaf-(head+comp)        {r(self.leaf_minus)}  (0=헤드가 모두 단말)",
            f"  junc-(head-comp)        {r(self.junc_minus)}",
            f"  node/head               {r(self.node_per_head)}",
            f"  total_length/head (m)   {r(self.total_per_head_m)}",
            f"  seg_min (m)             {r(self.seg_min_m)}",
            f"  seg_med (m)             {r(self.seg_med_m)}",
            f"  구경 관측집합           {sorted(self.diameters_seen)}",
        ])


def _band(values, *, margin=0.0, integer=False, nonneg=False):
    """관측값 리스트 → (lo, hi). margin>0 이면 양끝에 비례 여유."""
    lo, hi = min(values), max(values)
    if margin > 0:
        span = hi - lo
        pad = max(abs(lo), abs(hi)) * margin + span * margin
        lo, hi = lo - pad, hi + pad
    if nonneg:
        lo = max(0.0, lo)
    if integer:
        lo, hi = math.floor(lo), math.ceil(hi)
    return (lo, hi)


def learn_bands(profiles, *, margin=PRIOR_RATIO_MARGIN) -> BandSet:
    """NetProfile 리스트(정상 답안)에서 밴드를 학습한다.

    헤드수가 위험등급마다 다르므로(K-160=30, ESFR=12, 인랙=39 …) 절대 개수가 아니라
    헤드수와 무관한 **구조 불변식·비율**로 밴드를 잡아 어떤 구역에도 적용 가능케 한다.
    """
    profs = [p for p in profiles if p.head_count > 0 and p.n_edges > 0]
    if not profs:
        raise ValueError("밴드 학습용 프로파일이 없음")
    fexc = [p.n_edges - (p.n_nodes - p.n_components) for p in profs]
    lminus = [p.leaf_count - (p.head_count + p.n_components) for p in profs]
    jminus = [p.junction_count - (p.head_count - p.n_components) for p in profs]
    nph = [p.n_nodes / p.head_count for p in profs]
    tph = [p.total_length_m / p.head_count for p in profs]
    smin = [p.seg_min_m for p in profs]
    smed = [p.seg_med_m for p in profs]
    diam = set()
    for p in profs:
        diam |= set(p.diameter_counts)
    return BandSet(
        n_samples=len(profs),
        forest_excess=_band(fexc, integer=True),
        leaf_minus=_band(lminus, integer=True),
        junc_minus=_band(jminus, integer=True),
        node_per_head=_band(nph, margin=margin, nonneg=True),
        total_per_head_m=_band(tph, margin=margin, nonneg=True),
        seg_min_m=_band(smin, margin=margin, nonneg=True),
        seg_med_m=_band(smed, margin=margin, nonneg=True),
        diameters_seen=frozenset(diam),
        margin=margin,
    )


@dataclass
class ValidationReport:
    """밴드 검증 결과 — 각 점검의 OK/WARN/FAIL.

    FAIL = 구조적으로 불가능(빈 망·0/음수 길이·헤드 없음) → 추출 실패.
    WARN = 정상 분포 밖(루프·헤드 비단말·비율 이탈·비표준 구경) → 점검 필요.
    OK   = 정상 범위.
    """
    checks: list = field(default_factory=list)  # (name, status, detail)

    @property
    def has_fail(self) -> bool:
        return any(s == "FAIL" for _, s, _ in self.checks)

    @property
    def has_warn(self) -> bool:
        return any(s == "WARN" for _, s, _ in self.checks)

    @property
    def verdict(self) -> str:
        if self.has_fail:
            return "FAIL"
        return "WARN" if self.has_warn else "OK"

    def report(self) -> str:
        lines = [f"[{self.verdict}] prior 밴드 검증 ({len(self.checks)}개 점검)"]
        for name, status, detail in self.checks:
            mark = {"OK": "O", "WARN": "!", "FAIL": "X"}[status]
            lines.append(f"    [{mark}] {name}: {detail}")
        return "\n".join(lines)


def _inband(v, band):
    return band[0] <= v <= band[1]


def validate(net, bands: BandSet) -> ValidationReport:
    """추출망 1건을 prior 밴드에 비춰 '배관망스러운가' 점검."""
    checks = []
    p = profile(net)

    # ── 구조적 FAIL 선별 (불가능한 형상) ──
    if p.head_count == 0:
        checks.append(("헤드 존재", "FAIL", "헤드 0개 — 구역 격리 실패"))
    if p.n_edges == 0:
        checks.append(("배관 존재", "FAIL", "배관 0개 — 추출 실패"))
        return ValidationReport(checks=checks)
    bad_len = [pid for pid, pp in net.pipes.items() if pp.length_m <= 0]
    if bad_len:
        checks.append(("배관 길이>0", "FAIL", f"0/음수 길이 {len(bad_len)}개 — 좌표 오류"))
    else:
        checks.append(("배관 길이>0", "OK", f"전 구간 양수 ({p.n_edges}개)"))

    # ── 트리성 (forest_excess) ──
    fexc = p.n_edges - (p.n_nodes - p.n_components)
    if _inband(fexc, bands.forest_excess):
        checks.append(("트리 구조", "OK", f"forest_excess={fexc} (정상 {list(bands.forest_excess)})"))
    else:
        checks.append(("트리 구조", "WARN",
                       f"forest_excess={fexc} — 루프 {fexc}개 추정 (격자배관 아니면 연결 오류 의심)"))

    # ── 헤드=단말 ──
    lminus = p.leaf_count - (p.head_count + p.n_components)
    if _inband(lminus, bands.leaf_minus):
        checks.append(("헤드↔단말", "OK", f"leaf-(head+comp)={lminus}"))
    else:
        sign = "비단말 헤드/누락 배관" if lminus < bands.leaf_minus[0] else "헤드 아닌 단말(끊긴 배관)"
        checks.append(("헤드↔단말", "WARN", f"leaf-(head+comp)={lminus} — {sign} 의심"))

    # ── 접합점 ──
    jminus = p.junction_count - (p.head_count - p.n_components)
    if _inband(jminus, bands.junc_minus):
        checks.append(("접합점 수", "OK", f"junc-(head-comp)={jminus}"))
    else:
        checks.append(("접합점 수", "WARN", f"junc-(head-comp)={jminus} — 분기 구조 이탈"))

    # ── 노드/헤드 비율 ──
    nph = p.n_nodes / p.head_count if p.head_count else 0.0
    if p.head_count and _inband(nph, bands.node_per_head):
        checks.append(("노드/헤드 비율", "OK", f"{nph:.2f} (정상 {bands.node_per_head[0]:.2f}~{bands.node_per_head[1]:.2f})"))
    elif p.head_count:
        why = "중간 분기 소실" if nph < bands.node_per_head[0] else "과분할/잡음 노드"
        checks.append(("노드/헤드 비율", "WARN", f"{nph:.2f} 정상범위 밖 — {why} 의심"))

    # ── 헤드당 연장 ──
    tph = p.total_length_m / p.head_count if p.head_count else 0.0
    if p.head_count and _inband(tph, bands.total_per_head_m):
        checks.append(("헤드당 연장", "OK", f"{tph:.2f}m"))
    elif p.head_count:
        checks.append(("헤드당 연장", "WARN",
                       f"{tph:.2f}m 정상범위 밖 ({bands.total_per_head_m[0]:.2f}~{bands.total_per_head_m[1]:.2f}) — 축척/누락 의심"))

    # ── 구간 길이 분포 ──
    if _inband(p.seg_min_m, bands.seg_min_m):
        checks.append(("최소 구간연장", "OK", f"{p.seg_min_m:.2f}m"))
    else:
        checks.append(("최소 구간연장", "WARN", f"{p.seg_min_m:.2f}m 비정상 (정상 {bands.seg_min_m[0]:.2f}~{bands.seg_min_m[1]:.2f})"))

    # ── 구경 ladder ──
    diam = set(p.diameter_counts)
    nonstd = sorted(d for d in diam if d not in STANDARD_NOMINAL_MM)
    if nonstd:
        checks.append(("표준 호칭경", "WARN", f"비표준 구경 {nonstd}mm — 추출 잡음 의심"))
    else:
        checks.append(("표준 호칭경", "OK", f"{sorted(diam)} 모두 표준"))
    unseen = sorted(d for d in diam if d not in bands.diameters_seen)
    if unseen and not nonstd:
        checks.append(("학습 구경범위", "WARN", f"답안 미관측 구경 {unseen}mm (표준이긴 함)"))

    return ValidationReport(checks=checks)
