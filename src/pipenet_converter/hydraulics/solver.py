"""Sparse damped Newton solution of pipe energy and nodal continuity.

All edges survive (including chords and parallel pipes). Hazen-Williams uses
SI coefficient 10.67, flow exponent 1.852, bore exponent 4.871. Unknown flows
are L/s, heads metres. Nozzle outflows are pressure-dependent, zero when dry.
Both energy and mass residuals must pass; iteration count alone is not success.
"""
from __future__ import annotations

from dataclasses import dataclass
from math import pi
from typing import Callable
import warnings

import numpy as np
from scipy.sparse import bmat, coo_matrix, diags
from scipy.sparse.linalg import MatrixRankWarning, lsqr, spsolve

from .models import BAR_PER_M, Network


@dataclass
class Solution:
    """Residuals and signed physical values from one boundary solve."""
    converged: bool
    iterations: int
    mass_error_lps: float
    energy_error_m: float
    pressure_bar: dict[str, float]
    flow_lpm: dict[str, float]
    velocity_mps: dict[str, float]
    nozzle_flow_lpm: dict[str, float]
    source_flow_lpm: float
    source_pressure_bar: float


def resistance(length_m: float, inner_mm: float, c_factor: float) -> float:
    """Hazen-Williams head resistance for signed flow in L/s."""
    return 10.67 * length_m * .001**1.852 / (c_factor**1.852 * (inner_mm / 1000)**4.871)


def solve(network: Network, sizes: tuple[int, ...], source_pressure_bar: float,
          *, checkpoint: Callable[[], None] = lambda: None,
          max_iterations: int = 120) -> Solution:
    """Solve with a single fixed gauge-pressure boundary; never prune cycles."""
    network.validate()
    if not np.isfinite(source_pressure_bar) or source_pressure_bar < 0:
        raise ValueError("공급 압력이 유효하지 않습니다.")
    if len(sizes) != len(network.pipes) or any(not 0 <= j < len(p.sizes)
                                             for p, j in zip(network.pipes, sizes)):
        raise ValueError("관경 후보 선택 오류")
    nodes = {n.label: n for n in network.nodes}
    unknown = [n.label for n in network.nodes if n.label != network.source]
    index = {n: i for i, n in enumerate(unknown)}
    n, m = len(unknown), len(network.pipes)
    rows, cols, vals = [], [], []
    boundary = np.zeros(m)
    source_head = nodes[network.source].elevation_m + source_pressure_bar / BAR_PER_M
    for j, p in enumerate(network.pipes):
        for label, sign in ((p.a, 1.), (p.b, -1.)):
            if label in index:
                rows.append(index[label]); cols.append(j); vals.append(sign)
            else:
                boundary[j] += sign * source_head
    incidence = coo_matrix((vals, (rows, cols)), shape=(n, m)).tocsr()
    r = np.array([resistance(p.length_m + p.sizes[j].equivalent_m,
                            p.sizes[j].inner_mm, p.c_factor)
                  for p, j in zip(network.pipes, sizes)])
    z = np.array([nodes[label].elevation_m for label in unknown])
    k = np.zeros(n)
    for head in network.nozzles:
        if head.node in index:
            k[index[head.node]] += head.k_lpm_sqrt_bar * np.sqrt(BAR_PER_M) / 60

    def demand(h):
        return k * np.sqrt(np.maximum(h-z, 0))

    def residual(q, h):
        energy = r * q * np.abs(q)**.852 - incidence.T @ h - boundary
        mass = incidence @ q + demand(h)
        return energy, mass

    h = np.full(n, source_head - .01)
    q = lsqr(incidence, -demand(h), atol=1e-11, btol=1e-11)[0]
    converged = False
    for iteration in range(1, max_iterations + 1):
        checkpoint()
        energy, mass = residual(q, h)
        if np.max(np.abs(energy), initial=0) <= 1e-5 and np.max(np.abs(mass), initial=0) <= 1e-6:
            converged = True
            break
        gradient = np.maximum(1.852 * r * np.abs(q)**.852, 1e-7)
        d_demand = np.where(h-z > 0, .5*k / np.sqrt(np.maximum(h-z, 1e-8)), 0.)
        jac = bmat([[diags(gradient), -incidence.T], [incidence, diags(d_demand)]], format="csc")
        with warnings.catch_warnings():
            warnings.simplefilter("error", MatrixRankWarning)
            try:
                delta = spsolve(jac, -np.concatenate((energy, mass)))
            except (MatrixRankWarning, RuntimeError):
                break
        if not np.all(np.isfinite(delta)):
            break
        merit = float(energy @ energy + mass @ mass)
        alpha = 1.0
        accepted = False
        for _ in range(30):
            trial_q, trial_h = q + alpha*delta[:m], h + alpha*delta[m:]
            e2, m2 = residual(trial_q, trial_h)
            if float(e2 @ e2 + m2 @ m2) < merit:
                q, h = trial_q, trial_h
                accepted = True
                break
            alpha *= .5
        if not accepted:
            break
    energy, mass = residual(q, h)
    converged = bool(np.max(np.abs(energy), initial=0) <= 1e-5 and np.max(np.abs(mass), initial=0) <= 1e-6)
    pressure = {label: float((h[i]-z[i])*BAR_PER_M) for label, i in index.items()}
    pressure[network.source] = float(source_pressure_bar)
    nozzle_flows = {a.label: a.k_lpm_sqrt_bar * np.sqrt(max(pressure[a.node], 0.))
                   for a in network.nozzles}
    source_q = sum(float(q[j])*60 * (1 if p.a == network.source else -1)
                   for j, p in enumerate(network.pipes) if network.source in (p.a, p.b))
    source_q += sum(nozzle_flows[a.label] for a in network.nozzles if a.node == network.source)
    return Solution(converged, iteration, float(np.max(np.abs(mass), initial=0)),
                    float(np.max(np.abs(energy), initial=0)), pressure,
                    {p.label: float(q[j]*60) for j, p in enumerate(network.pipes)},
                    {p.label: float(abs(q[j])*.001 / (pi*(p.sizes[sizes[j]].inner_mm/1000)**2/4))
                     for j, p in enumerate(network.pipes)},
                    {key: float(v) for key, v in nozzle_flows.items()}, source_q, float(source_pressure_bar))
