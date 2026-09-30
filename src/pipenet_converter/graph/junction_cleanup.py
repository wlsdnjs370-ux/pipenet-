"""Remove evidenced duplicate symbol joins, never arbitrary small cycles."""
from __future__ import annotations

from collections import defaultdict, deque
import math
from typing import Iterable, Mapping, Sequence

from .flow import Edge, edge_key


def duplicate_symbol_joins(points: Sequence[Sequence[float]], edges: Iterable[Edge],
                           drawn_edges: Iterable[Edge], arcs: Sequence[Mapping], *,
                           protected_edges: Iterable[Edge] = ()) -> dict[Edge, dict]:
    """Identify off-centre inferred shortcuts to an already joined arc centre.

    Require a <=30 mm displaced seat, a precise (<0.2 mm) arc centre and
    an existing <=3-edge alternative containing original drawn pipe. The other
    endpoint must lie on a radial pipe at the centre. Original/user edges are
    immutable. Work on a private adjacency so every removal retains its path.
    These are duplicate contacts, not a diameter/length-based cycle filter.
    """
    drawn = {edge_key(*e) for e in drawn_edges}
    protected = drawn | {edge_key(*e) for e in protected_edges}
    adj = defaultdict(set)
    for a,b in edges:
        adj[a].add(b); adj[b].add(a)
    cell = 300.0
    grid = defaultdict(list)
    for n in adj:
        x,y=points[n][:2]
        grid[int(x//cell),int(y//cell)].append(n)
    removed = {}
    for arc in arcs:
        cx,cy,r=float(arc['cx']),float(arc['cy']),float(arc['r'])
        if r <= 0: continue
        gx,gy=int(cx//cell),int(cy//cell)
        nearby=[n for ix in range(gx-1,gx+2) for iy in range(gy-1,gy+2)
                for n in grid.get((ix,iy),())]
        centres=[n for n in nearby if math.dist(points[n][:2],(cx,cy))<=0.2]
        for seat in nearby:
            ds=math.dist(points[seat][:2],(cx,cy))
            if not 1.0 < ds <= 30: continue
            for end in sorted(adj[seat]):
                e=edge_key(seat,end)
                if e in protected or math.dist(points[end][:2],(cx,cy))>r: continue
                for centre in centres:
                    if centre not in adj[end]: continue
                    # Pipe runs radially through the real centre; the off-axis
                    # shortcut is not a distinct drawn branch.
                    v=[points[end][i]-points[centre][i] for i in (0,1)]
                    ln=math.hypot(*v)
                    if ln <= 1: continue
                    off=abs(v[0]*(points[seat][1]-points[centre][1])-
                            v[1]*(points[seat][0]-points[centre][0]))/ln
                    if off <= 1: continue
                    queue=deque([(centre,[end,centre])]); path=None
                    while queue:
                        at,walk=queue.popleft()
                        if at==seat:
                            path=walk; break
                        if len(walk)>=4: continue
                        for n in sorted(adj[at]):
                            if n not in walk and math.dist(points[n][:2],(cx,cy))<=30:
                                queue.append((n,walk+[n]))
                    if path is None: continue
                    via=[edge_key(a,b) for a,b in zip(path,path[1:])]
                    if not any(p in drawn for p in via): continue
                    adj[seat].remove(end); adj[end].remove(seat)
                    removed[e]={'edge':list(e),'via':path,'center':centre,
                                'arc_xy':[cx,cy], 'reason':'duplicate_offcentre_symbol_join'}
                    break
    return removed
