"""
auto_design.py — 자동 설계 엔진 (② Zone, ③ 헤드 배치, ④ 배관 라우팅)

This module turns DXF-extracted geometry into a fully designed sprinkler
network: zones, heads, branches, cross-mains, risers, valves, pumps, hangers.

It implements the three middle stages of the v4 pipeline:
  ② Zone partition + system selection + head spec
  ③ Head auto-placement (grid sweep + (2R)² = S² + L² check + obstacle avoidance)
  ④ Pipe routing (branch clustering + cross-main MST + riser + Case topology)

Output is a dataclass network model that is consumed by:
  - SDF writer (PIPENET conversion in pipeline_orchestrator)
  - DXF writer (final drawing in stage ⑧)
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Iterable

from nftc_rules import (
    RuleDecision,
    TripleTrace,
    Verdict,
    decide_horizontal_distance,
    decide_temperature_rating,
    decide_reference_count,
    is_fast_response_required,
    decide_esfr_branch,
    validate_head_clearance,
    head_pressure_min_mpa,
    head_pressure_max_mpa,
    head_flow_min_lpm,
    hanger_max_spacing_m,
)
from hb_rules import (
    HBCase,
    HBCaseDecision,
    SystemType,
    ZonePartition,
    decide_system_type,
    decide_hb_case,
    decide_pipe_material,
    decide_zone_partition,
    hanger_positions_along_pipe,
    get_inner_diameter_mm,
)
from phd_rules import (
    PressureZone,
    FloorPressureZone,
    DiscretionaryVariables,
    classify_pressure_zones,
    decide_discretionary_variables,
)


# ---------------------------------------------------------------------------
# 1. Network model dataclasses
# ---------------------------------------------------------------------------


@dataclass
class HeadSpec:
    """Sprinkler head specification (output of ② head spec selection)."""

    zone_id: str
    horizontal_distance_m: float       # R
    k_factor_lpm_bar05: int            # K = 80 / 115 / 200 / 320 / 360
    temperature_rating_min_c: float
    temperature_rating_max_c: float | None
    rti_class: str                     # "fast" / "standard"
    head_type: str                     # "pendent" / "upright" / "sidewall" / "drypendent"
    corrosion_resistant: bool
    is_esfr: bool
    trace: TripleTrace


@dataclass
class HeadInstance:
    """A single placed sprinkler head."""

    head_id: str
    zone_id: str
    x: float
    y: float
    z: float
    spec: HeadSpec
    branch_axis: str                   # "EW" / "NS"
    cell_S: float                      # gap to adjacent head along branch axis
    cell_L: float                      # gap perpendicular to branch axis
    nftc_2773_pass: bool
    nftc_2771_pass: bool
    skipping_pass: bool                # cell_S, cell_L ≥ 1.8 m
    a_b_l_check: bool                  # (2R)² ≥ S² + L²


@dataclass
class PipeSegment:
    """A pipe segment between two nodes."""

    pipe_id: str
    role: str                          # "branch" / "cross_main" / "riser" / "main"
    n1: tuple[float, float, float]
    n2: tuple[float, float, float]
    length_m: float
    nominal: str                       # "25A" / "32A" / ... / "300A"
    material: str                      # "KSD 3507" / "KSD 3562"
    inner_diameter_mm: float
    c_factor: int                      # 120 (steel) / 150 (CPVC)
    head_count_downstream: int
    estimated_velocity_mps: float | None = None
    fittings: list[dict[str, Any]] = field(default_factory=list)
    hangers_m: list[float] = field(default_factory=list)


@dataclass
class ValveDevice:
    """A valve / equipment on the network."""

    valve_id: str
    type: str                          # "alarm" / "preaction" / "deluge" / "prv" / "os_y" / "check"
    location: tuple[float, float, float]
    spec: dict[str, Any] = field(default_factory=dict)
    equivalent_length_m: float = 0.0


@dataclass
class Zone:
    """A protection zone with its design output."""

    zone_id: str
    floor_label: str
    polygon: list[tuple[float, float]]
    area_m2: float
    use: str
    structure: str
    ceiling_h_m: float
    pressure_zone: PressureZone | None = None
    head_spec: HeadSpec | None = None
    heads: list[HeadInstance] = field(default_factory=list)
    branches: list[PipeSegment] = field(default_factory=list)
    cross_mains: list[PipeSegment] = field(default_factory=list)
    valves: list[ValveDevice] = field(default_factory=list)
    fire_compartment_id: str | None = None
    nftc_2755_fast_required: bool = False  # NFTC 2.7.5.5 strict mandate
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class DesignNetwork:
    """Complete designed network — output of ② + ③ + ④."""

    project_id: str
    building_height_m: float
    hb_case: HBCaseDecision | None
    system_type: str
    zones: list[Zone] = field(default_factory=list)
    floors_pressure: list[FloorPressureZone] = field(default_factory=list)
    risers: list[PipeSegment] = field(default_factory=list)
    pumps: list[dict[str, Any]] = field(default_factory=list)
    tanks: list[dict[str, Any]] = field(default_factory=list)
    discretionary: DiscretionaryVariables | None = None
    metadata: dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# 2. Stage ② — Zone + system + head spec
# ---------------------------------------------------------------------------


class AutoZonePlanner:
    """② Zone partition + system selection + head specification.

    Inputs come from ① object extraction (rooms with use/structure/ceiling_h).
    Outputs Zone instances ready for ③ head placement.
    """

    def __init__(self, building_meta: dict[str, Any]) -> None:
        self.building_meta = building_meta

    def plan(self, rooms: list[dict[str, Any]]) -> list[Zone]:
        zones: list[Zone] = []
        for room in rooms:
            if room.get("hb_excluded"):
                continue  # NFTC 2.4.18 exclusion masking
            partitions = decide_zone_partition(
                floor_area_m2=float(room.get("area_m2", 0.0)),
                estimated_head_count=int(room.get("estimated_head_count", 0)),
                floor_label=str(room.get("floor", "")),
                is_grid_layout=bool(room.get("is_grid_layout", False)),
                is_apartment_loft=bool(room.get("is_apartment_loft", False)),
                fire_compartment_id=room.get("fire_compartment_id"),
                system_type=str(room.get("system_type_hint", "wet")),
            )
            for part in partitions:
                z = Zone(
                    zone_id=part.zone_id,
                    floor_label=part.floor_label,
                    polygon=room.get("polygon", []),
                    area_m2=part.area_m2,
                    use=str(room.get("use", "other_low")),
                    structure=str(room.get("structure", "non_fire_resistant")),
                    ceiling_h_m=float(room.get("ceiling_h_m", 3.0)),
                    fire_compartment_id=part.fire_compartment_id,
                )
                z.head_spec = self._select_head_spec(z, room)
                # NFTC 2.7.5.5 mandate
                fr_dec = is_fast_response_required(z.use)
                z.nftc_2755_fast_required = bool(fr_dec.value)
                zones.append(z)
        return zones

    def _select_head_spec(self, zone: Zone, room: dict[str, Any]) -> HeadSpec:
        # NFTC 2.7.3 R
        r_dec = decide_horizontal_distance(
            room_use=zone.use,
            structure=zone.structure,
            has_special_combustible=bool(room.get("has_special_combustible", False)),
            is_rack_storage=(zone.use == "rack_storage"),
        )
        # NFTC 2.7.5.5 fast-response mandate
        fr_dec = is_fast_response_required(zone.use)
        # NFTC 2.7.6 temperature rating
        t_dec = decide_temperature_rating(
            ambient_temp_c=float(room.get("ambient_temp_c", 25.0)),
            is_factory_4m_high=(zone.use == "factory" and zone.ceiling_h_m >= 4.0),
            is_warehouse_4m_high=(zone.use == "warehouse" and zone.ceiling_h_m >= 4.0),
            is_rack_storage=(zone.use == "rack_storage"),
        )
        # NFTC 103B ESFR branch
        esfr_dec = decide_esfr_branch(room_use=zone.use, ceiling_h_m=zone.ceiling_h_m)
        # K-factor
        if esfr_dec.activated and esfr_dec.k_lpm_bar05:
            k = esfr_dec.k_lpm_bar05
        elif esfr_dec.k_lpm_bar05 == 115:
            k = 115
        else:
            k = 80
        # Head type
        sys_type = str(room.get("system_type_hint", "wet"))
        if sys_type in {"dry", "preaction_single", "preaction_none", "preaction_double"} \
                and room.get("head_orientation") == "pendent":
            head_type = "drypendent"
        elif sys_type == "wet" and room.get("head_orientation") == "pendent":
            head_type = "pendent"
        else:
            head_type = str(room.get("head_orientation", "pendent"))
        return HeadSpec(
            zone_id=zone.zone_id,
            horizontal_distance_m=float(r_dec.value),
            k_factor_lpm_bar05=k,
            temperature_rating_min_c=float(t_dec.value.get("min_c", 79.0)),
            temperature_rating_max_c=t_dec.value.get("max_c"),
            rti_class="fast" if fr_dec.value else "standard",
            head_type=head_type,
            corrosion_resistant=bool(room.get("corrosion_environment", False)),
            is_esfr=esfr_dec.activated,
            trace=TripleTrace(
                nftc=" + ".join(filter(None, [r_dec.trace.nftc, t_dec.trace.nftc, fr_dec.trace.nftc, esfr_dec.trace.nftc])),
                hb=" + ".join(filter(None, [r_dec.trace.hb, t_dec.trace.hb])),
                phd=None,
                note=f"R={r_dec.value} K={k} T={t_dec.value} fast={fr_dec.value} esfr={esfr_dec.activated}",
            ),
        )


# ---------------------------------------------------------------------------
# 3. Stage ③ — Auto head placement
# ---------------------------------------------------------------------------


@dataclass
class GridCandidate:
    """Candidate grid layout for one zone."""

    theta_deg: float
    s: float
    l: float
    offset_x: float
    offset_y: float
    head_count: int
    coverage_ratio: float
    skipping_violations: int
    obstacle_violations: int
    cost: float


class AutoHeadPlacer:
    """③ Auto-place heads on a grid with NFTC + (2R)² = S² + L² hard checks."""

    HEAD_SPACING_MIN_M = 1.8       # NFTC §2.7.7 — Skipping prevention
    HEAD_TO_WALL_MIN_M = 0.10      # NFTC 2.7.7.1 단서
    HEAD_CLEARANCE_RADIUS_M = 0.60 # NFTC 2.7.7.1
    THETA_STEP_DEG = 15            # search granularity (coarse, then refine)
    OFFSET_STEP_M = 0.3

    def __init__(self, *, zone: Zone, obstacles: list[dict[str, Any]] | None = None) -> None:
        self.zone = zone
        self.obstacles = obstacles or []

    def place(self) -> list[HeadInstance]:
        """Run grid sweep → score → refine → emit HeadInstance list."""
        if not self.zone.head_spec:
            return []
        r = self.zone.head_spec.horizontal_distance_m
        # Candidates: (S, L) pairs from HB §2.4.8 lookup, all with (2R)² ≥ S² + L²
        sl_pairs = self._candidate_sl_pairs(r)
        # Theta and offset sweep
        best: GridCandidate | None = None
        polygon_bbox = self._polygon_bbox(self.zone.polygon)
        for theta in range(0, 90, self.THETA_STEP_DEG):
            for s, l in sl_pairs:
                for ox in self._offset_range(polygon_bbox[0], polygon_bbox[2], s):
                    for oy in self._offset_range(polygon_bbox[1], polygon_bbox[3], l):
                        cand = self._evaluate_candidate(theta, s, l, ox, oy)
                        if best is None or cand.cost < best.cost:
                            best = cand
        if not best:
            return []
        return self._materialize_heads(best, r)

    def _candidate_sl_pairs(self, r: float) -> list[tuple[float, float]]:
        """Generate (S, L) pairs satisfying HB skipping rule + (2R)² ≥ S² + L².

        S, L iterate from 1.8 m to 4.5 m in 0.1 m steps.
        """
        pairs: list[tuple[float, float]] = []
        s = 1.8
        while s <= 2 * r + 0.01:
            l_max_sq = (2 * r) ** 2 - s * s
            if l_max_sq <= 0:
                s += 0.1
                continue
            l_max = math.sqrt(l_max_sq)
            l = max(1.8, 0.0)
            while l <= l_max + 0.01:
                pairs.append((round(s, 2), round(l, 2)))
                l += 0.2
            s += 0.2
        return pairs

    def _offset_range(self, lo: float, hi: float, step: float) -> Iterable[float]:
        """Generate offset values for grid origin within bbox."""
        if step <= 0:
            return [0.0]
        n = max(1, int((hi - lo) / step + 1))
        return [lo + i * step for i in range(min(n, 5))]  # limit to 5 offsets per axis

    def _evaluate_candidate(self, theta: float, s: float, l: float, ox: float, oy: float) -> GridCandidate:
        """Score one grid candidate."""
        heads = self._generate_grid_points(theta, s, l, ox, oy)
        skipping_v = sum(1 for h in heads if not self._skipping_ok(h, heads, s, l))
        obstacle_v = sum(1 for h in heads if not self._clearance_ok(h))
        coverage = min(1.0, len(heads) * s * l / max(self.zone.area_m2, 1.0))
        n = len(heads)
        # Cost: prefer high coverage, low head count, no violations
        cost = (
            (1.0 - coverage) * 100.0
            + skipping_v * 50.0
            + obstacle_v * 50.0
            + n * 0.5
        )
        return GridCandidate(theta, s, l, ox, oy, n, coverage, skipping_v, obstacle_v, cost)

    def _generate_grid_points(self, theta: float, s: float, l: float, ox: float, oy: float) -> list[tuple[float, float]]:
        """Generate grid points inside the zone polygon."""
        bbox = self._polygon_bbox(self.zone.polygon)
        if not bbox:
            return []
        x_lo, y_lo, x_hi, y_hi = bbox
        cos_t = math.cos(math.radians(theta))
        sin_t = math.sin(math.radians(theta))
        points: list[tuple[float, float]] = []
        for i in range(int((x_hi - x_lo) / s) + 1):
            for j in range(int((y_hi - y_lo) / l) + 1):
                x = x_lo + ox + i * s * cos_t - j * l * sin_t
                y = y_lo + oy + i * s * sin_t + j * l * cos_t
                if self._point_in_polygon((x, y), self.zone.polygon):
                    points.append((x, y))
        return points

    def _polygon_bbox(self, polygon: list[tuple[float, float]]) -> tuple[float, float, float, float]:
        if not polygon:
            return (0.0, 0.0, 0.0, 0.0)
        xs = [p[0] for p in polygon]
        ys = [p[1] for p in polygon]
        return (min(xs), min(ys), max(xs), max(ys))

    def _point_in_polygon(self, point: tuple[float, float], polygon: list[tuple[float, float]]) -> bool:
        """Ray-casting algorithm for point-in-polygon."""
        if not polygon:
            return False
        x, y = point
        n = len(polygon)
        inside = False
        j = n - 1
        for i in range(n):
            xi, yi = polygon[i]
            xj, yj = polygon[j]
            if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi):
                inside = not inside
            j = i
        return inside

    def _skipping_ok(self, head: tuple[float, float], all_heads: list[tuple[float, float]], s: float, l: float) -> bool:
        for other in all_heads:
            if other == head:
                continue
            d = math.hypot(head[0] - other[0], head[1] - other[1])
            if d < self.HEAD_SPACING_MIN_M:
                return False
        return True

    def _clearance_ok(self, head: tuple[float, float]) -> bool:
        if not self.obstacles:
            return True
        for obs in self.obstacles:
            poly = obs.get("polygon")
            if not poly:
                continue
            d = self._min_dist_to_polygon(head, poly)
            if obs.get("is_wall"):
                if d < self.HEAD_TO_WALL_MIN_M:
                    return False
            else:
                if d < self.HEAD_CLEARANCE_RADIUS_M:
                    return False
        return True

    def _min_dist_to_polygon(self, p: tuple[float, float], poly: list[tuple[float, float]]) -> float:
        """Minimum distance from a point to polygon edges (self-contained)."""
        if not poly:
            return float("inf")
        best = float("inf")
        n = len(poly)
        px, py = p
        for i in range(n):
            ax, ay = poly[i]
            bx, by = poly[(i + 1) % n]
            dx = bx - ax
            dy = by - ay
            if dx == 0 and dy == 0:
                d = math.hypot(px - ax, py - ay)
            else:
                t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)
                t = max(0.0, min(1.0, t))
                qx = ax + t * dx
                qy = ay + t * dy
                d = math.hypot(px - qx, py - qy)
            if d < best:
                best = d
        return best

    def _materialize_heads(self, cand: GridCandidate, r: float) -> list[HeadInstance]:
        """Convert grid candidate to HeadInstance list with hard checks recorded."""
        points = self._generate_grid_points(cand.theta_deg, cand.s, cand.l, cand.offset_x, cand.offset_y)
        spec = self.zone.head_spec
        heads: list[HeadInstance] = []
        all_pts = points
        for idx, (x, y) in enumerate(points, 1):
            skipping = self._skipping_ok((x, y), all_pts, cand.s, cand.l)
            clear = self._clearance_ok((x, y))
            a_b_l = (2 * r) ** 2 >= cand.s ** 2 + cand.l ** 2
            heads.append(HeadInstance(
                head_id=f"{self.zone.zone_id}-H{idx:03d}",
                zone_id=self.zone.zone_id,
                x=round(x, 3),
                y=round(y, 3),
                z=0.0,  # set later by elevation pass
                spec=spec,
                branch_axis="EW" if cand.theta_deg < 45 else "NS",
                cell_S=cand.s,
                cell_L=cand.l,
                nftc_2773_pass=True,  # R selected from NFTC table
                nftc_2771_pass=clear,
                skipping_pass=skipping,
                a_b_l_check=a_b_l,
            ))
        return heads


# ---------------------------------------------------------------------------
# 4. Stage ④ — Pipe routing
# ---------------------------------------------------------------------------


class AutoPipeRouter:
    """④ Pipe routing — branch clustering, cross-main MST, riser placement.

    Strategy:
      1. Cluster heads by row along branch_axis (DBSCAN-like, gap threshold)
      2. Each row → 1 branch pipe (with diameter sizing iterative)
      3. Branches → cross-main via Steiner-augmented MST
      4. Cross-mains → riser at given shaft location
    """

    BRANCH_HEAD_MAX = 8           # HB heuristic — heads per branch
    INITIAL_BRANCH_NOMINAL = "25A"
    INITIAL_CROSS_MAIN_NOMINAL = "65A"
    BRANCH_DIAMETER_LADDER = ["25A", "32A", "40A", "50A", "65A", "80A", "100A"]
    MAIN_DIAMETER_LADDER = ["65A", "80A", "100A", "125A", "150A", "200A"]

    def __init__(self, *, zone: Zone, riser_xy: tuple[float, float] | None = None) -> None:
        self.zone = zone
        self.riser_xy = riser_xy

    def route(self) -> tuple[list[PipeSegment], list[PipeSegment]]:
        """Return (branches, cross_mains)."""
        branches = self._cluster_branches(self.zone.heads)
        cross_mains = self._build_cross_mains(branches)
        return branches, cross_mains

    def _cluster_branches(self, heads: list[HeadInstance]) -> list[PipeSegment]:
        """Cluster heads into branch rows along their declared branch_axis."""
        if not heads:
            return []
        # Group by row coordinate (perpendicular to branch_axis)
        groups: dict[float, list[HeadInstance]] = {}
        for h in heads:
            key = round(h.y if h.branch_axis == "EW" else h.x, 1)
            groups.setdefault(key, []).append(h)
        branches: list[PipeSegment] = []
        for idx, (row_key, hs) in enumerate(sorted(groups.items()), 1):
            hs.sort(key=lambda h: h.x if h.branch_axis == "EW" else h.y)
            # Split if > BRANCH_HEAD_MAX
            chunks = [hs[i:i + self.BRANCH_HEAD_MAX] for i in range(0, len(hs), self.BRANCH_HEAD_MAX)]
            for cidx, chunk in enumerate(chunks, 1):
                if not chunk:
                    continue
                first = chunk[0]
                last = chunk[-1]
                length = math.hypot(last.x - first.x, last.y - first.y)
                # Initial diameter — will be iterated
                nominal = self._initial_branch_nominal(len(chunk))
                inner = get_inner_diameter_mm(nominal, "KSD 3507") or 27.5
                head_pos = [
                    math.hypot(h.x - first.x, h.y - first.y) for h in chunk
                ]
                hangers = hanger_positions_along_pipe(
                    pipe_length_m=length,
                    pipe_role="branch",
                    head_positions_m=head_pos,
                )
                pipe_id = f"{self.zone.zone_id}-B{idx:02d}-{cidx}"
                branches.append(PipeSegment(
                    pipe_id=pipe_id,
                    role="branch",
                    n1=(first.x, first.y, first.z),
                    n2=(last.x, last.y, last.z),
                    length_m=round(length, 3),
                    nominal=nominal,
                    material="KSD 3507",
                    inner_diameter_mm=inner,
                    c_factor=120,
                    head_count_downstream=len(chunk),
                    fittings=[
                        {"type": "tee_branch", "at": (first.x, first.y, first.z)},
                        {"type": "endcap", "at": (last.x, last.y, last.z)},
                    ],
                    hangers_m=hangers,
                ))
        return branches

    def _initial_branch_nominal(self, head_count: int) -> str:
        """Initial branch diameter from heuristic table."""
        if head_count <= 2:
            return "25A"
        if head_count <= 4:
            return "32A"
        if head_count <= 6:
            return "40A"
        return "50A"

    def _build_cross_mains(self, branches: list[PipeSegment]) -> list[PipeSegment]:
        """Build cross-main from MST connecting branch tee points to riser."""
        if not branches:
            return []
        riser = self.riser_xy or self._auto_riser_position(branches)
        # Connect each branch's tee point to riser via direct main (simple fallback)
        cross_mains: list[PipeSegment] = []
        for idx, br in enumerate(branches, 1):
            tee = br.n1
            length = math.hypot(tee[0] - riser[0], tee[1] - riser[1])
            if length < 1e-3:
                continue
            nominal = self._cross_main_nominal(idx, len(branches))
            inner = get_inner_diameter_mm(nominal, "KSD 3507") or 65.9
            hangers = hanger_positions_along_pipe(
                pipe_length_m=length,
                pipe_role="cross_main",
            )
            cross_mains.append(PipeSegment(
                pipe_id=f"{self.zone.zone_id}-CM{idx:02d}",
                role="cross_main",
                n1=tee,
                n2=(riser[0], riser[1], tee[2]),
                length_m=round(length, 3),
                nominal=nominal,
                material="KSD 3507",
                inner_diameter_mm=inner,
                c_factor=120,
                head_count_downstream=br.head_count_downstream,
                fittings=[
                    {"type": "tee_main", "at": tee},
                ],
                hangers_m=hangers,
            ))
        return cross_mains

    def _auto_riser_position(self, branches: list[PipeSegment]) -> tuple[float, float]:
        """If no shaft was specified, pick the centroid of branch start points."""
        if not branches:
            return (0.0, 0.0)
        xs = [b.n1[0] for b in branches]
        ys = [b.n1[1] for b in branches]
        return (sum(xs) / len(xs), sum(ys) / len(ys))

    def _cross_main_nominal(self, branch_idx: int, total_branches: int) -> str:
        """Cross-main diameter increases as more branches feed in."""
        # Simple heuristic — replace with hydraulic pre-sizing
        if total_branches <= 3:
            return "65A"
        if total_branches <= 6:
            return "80A"
        return "100A"


# ---------------------------------------------------------------------------
# 5. Auxiliary placement (valves, hangers already on pipes)
# ---------------------------------------------------------------------------


class AutoFittingPlacer:
    """Place valves and accessories per HB §2.4.7 + §2.4.13~15.

    - Each zone: 1 alarm valve / preaction valve / deluge valve per system_type
    - PRV at LSP zones
    - OS&Y or geared butterfly at riser
    - Auto air-bleed valve (wet only): top of riser + end of horizontal main
    - Water hammer arrestor: top of riser + end of horizontal main
    """

    @staticmethod
    def place_for_zone(
        zone: Zone,
        *,
        system_type: str,
        is_lsp: bool = False,
        riser_position: tuple[float, float, float] | None = None,
    ) -> list[ValveDevice]:
        valves: list[ValveDevice] = []
        if not riser_position:
            return valves
        # Alarm / preaction / deluge valve at riser entry to zone
        if system_type == "wet":
            valves.append(ValveDevice(
                valve_id=f"{zone.zone_id}-AV",
                type="alarm",
                location=riser_position,
                spec={"system": "wet", "model": "alarm_check"},
                equivalent_length_m=12.9,  # PIPENET standard
            ))
        elif system_type.startswith("preaction"):
            valves.append(ValveDevice(
                valve_id=f"{zone.zone_id}-PV",
                type="preaction",
                location=riser_position,
                spec={"system": system_type, "interlock": system_type.split("_")[1]},
                equivalent_length_m=10.1,  # PIPENET standard
            ))
        elif system_type == "dry":
            valves.append(ValveDevice(
                valve_id=f"{zone.zone_id}-DV",
                type="dry",
                location=riser_position,
                spec={"system": "dry"},
                equivalent_length_m=10.1,
            ))
        elif system_type == "deluge":
            valves.append(ValveDevice(
                valve_id=f"{zone.zone_id}-DLG",
                type="deluge",
                location=riser_position,
                spec={"system": "deluge"},
                equivalent_length_m=10.1,
            ))
        # PRV for LSP zones
        if is_lsp:
            valves.append(ValveDevice(
                valve_id=f"{zone.zone_id}-PRV",
                type="prv",
                location=riser_position,
                spec={"p1_bar": None, "p2_bar": 4.0},  # 2차측 4 bar (HB)
                equivalent_length_m=0.0,
            ))
        # OS&Y at riser (always)
        valves.append(ValveDevice(
            valve_id=f"{zone.zone_id}-OSY",
            type="os_y",
            location=riser_position,
            spec={"normally_open": True},
            equivalent_length_m=0.0,
        ))
        return valves


# ---------------------------------------------------------------------------
# 6. Top-level orchestration helper
# ---------------------------------------------------------------------------


def design_full_network(
    *,
    project_id: str,
    rooms: list[dict[str, Any]],
    obstacles: list[dict[str, Any]],
    floors: list[dict[str, Any]],
    building_meta: dict[str, Any],
    refuge_floor_interval_m: float | None = None,
    rooftop_tank_feasible: bool = True,
    pump_rated_q_lpm: float = 2400.0,
    pump_rated_h_m: float = 60.0,
    pump_churn_h_m: float = 70.0,
) -> DesignNetwork:
    """End-to-end ② → ③ → ④ pipeline.

    Returns a DesignNetwork that captures all decisions and outputs ready for
    PIPENET conversion (⑤) and validation (⑥).
    """
    # Stage ② — zones + system + head spec
    planner = AutoZonePlanner(building_meta)
    zones = planner.plan(rooms)
    # System type
    sys_dec = decide_system_type(
        has_freezing_risk=bool(building_meta.get("has_freezing_risk", False)),
        needs_open_heads=bool(building_meta.get("needs_open_heads", False)),
        detector_priority=bool(building_meta.get("detector_priority", False)),
        room_use=str(building_meta.get("primary_use", "")),
    )
    # HB Case
    hb_case = decide_hb_case(
        building_height_m=float(building_meta.get("height_m", 30.0)),
        refuge_floor_interval_m=refuge_floor_interval_m,
        rooftop_tank_feasible=rooftop_tank_feasible,
        water_source_type=str(building_meta.get("water_source", "fire_dedicated")),
    )
    # Pressure zones
    pressure_floors = classify_pressure_zones(
        floors=floors,
        hb_case=hb_case,
        elevated_tank_z_m=float(building_meta.get("elevated_tank_z_m", building_meta.get("height_m", 30.0))),
        pump_z_m=building_meta.get("pump_z_m"),
    )
    # Tag each zone with its pressure zone
    floor_to_pressure = {f.floor_label: f for f in pressure_floors}
    for z in zones:
        fp = floor_to_pressure.get(z.floor_label)
        if fp:
            z.pressure_zone = fp.zone
    # Stage ③ — auto place heads
    for z in zones:
        placer = AutoHeadPlacer(zone=z, obstacles=obstacles)
        z.heads = placer.place()
    # Stage ④ — auto route pipes + valves
    for z in zones:
        riser_xy = _zone_centroid(z.polygon)
        router = AutoPipeRouter(zone=z, riser_xy=riser_xy)
        z.branches, z.cross_mains = router.route()
        # Auxiliary: valves
        is_lsp = z.pressure_zone == PressureZone.LSP
        z.valves = AutoFittingPlacer.place_for_zone(
            z,
            system_type=str(sys_dec.value),
            is_lsp=is_lsp,
            riser_position=(riser_xy[0], riser_xy[1], 0.0),
        )
    # Discretionary variables
    discretionary = decide_discretionary_variables(
        floors=pressure_floors,
        hb_case=hb_case,
        elevated_tank_z_m=float(building_meta.get("elevated_tank_z_m", building_meta.get("height_m", 30.0))),
        pump_rated_q_lpm=pump_rated_q_lpm,
        pump_rated_h_m=pump_rated_h_m,
        pump_churn_h_m=pump_churn_h_m,
    )
    return DesignNetwork(
        project_id=project_id,
        building_height_m=float(building_meta.get("height_m", 30.0)),
        hb_case=hb_case,
        system_type=str(sys_dec.value),
        zones=zones,
        floors_pressure=pressure_floors,
        risers=[],  # filled by ⑤ converter (cross_zone risers)
        pumps=[{
            "rated_q_lpm": pump_rated_q_lpm,
            "rated_h_m": pump_rated_h_m,
            "churn_h_m": pump_churn_h_m,
            "location": hb_case.pump_location,
        }],
        tanks=[],
        discretionary=discretionary,
        metadata={
            "building_meta": building_meta,
            "system_decision_trace": sys_dec.trace.to_dict(),
            "hb_case_trace": hb_case.trace.to_dict(),
        },
    )


def _zone_centroid(polygon: list[tuple[float, float]]) -> tuple[float, float]:
    if not polygon:
        return (0.0, 0.0)
    return (
        sum(p[0] for p in polygon) / len(polygon),
        sum(p[1] for p in polygon) / len(polygon),
    )


__all__ = [
    "HeadSpec",
    "HeadInstance",
    "PipeSegment",
    "ValveDevice",
    "Zone",
    "DesignNetwork",
    "GridCandidate",
    "AutoZonePlanner",
    "AutoHeadPlacer",
    "AutoPipeRouter",
    "AutoFittingPlacer",
    "design_full_network",
]
