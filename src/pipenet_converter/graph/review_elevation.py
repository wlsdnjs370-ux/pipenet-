"""Arc-evidenced port splitting with cycle-wide elevation constraints.

No parent/child traversal is used. Each horizontal component has one elevation;
every symbol constrains high minus low to the configured rise. Contradictory
symbol groups retain source connections with explicit unconfirmed elevations.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import replace
import math
from typing import Mapping, Sequence, TYPE_CHECKING

import networkx as nx
from .arc_strokes import overlapping_arc_stroke
if TYPE_CHECKING:
    from .review_conversion import ReviewNetwork


def _open(arcs: Sequence[Mapping], vector: Sequence[float]) -> bool:
    angle = math.degrees(math.atan2(vector[1],vector[0])) % 360
    return all((angle-float(a['sa'])) % 360 > float(a['sweep'])+1e-8 for a in arcs)


def _unit(v):
    size=math.hypot(*v)
    return tuple(c/size for c in v) if size>1e-6 else None


def apply_arc_elevations(model: ReviewNetwork, points: Sequence[Sequence[float]],
                         source_edges: Sequence[tuple[int,int]],
                         arcs: Sequence[Mapping], rise_m: float, *,
                         drawn_edges: Sequence[tuple[int,int]] = (),
                         protected_nodes: Sequence[int] = (),
                         _excluded: frozenset[str] = frozenset()) -> None:
    """Split confirmed arc ports in a ReviewNetwork before grafting head stems.

    Coordinates are mm except elevations/rise (m). Tiny existing joint links
    are stood upright, never silently added to hydraulic length. Source port
    evidence includes unselected branches. A two-port arc with one open and
    one closed port is a vertical elbow connection even without a nearby
    return symbol. The same cycle-wide constraints validate single rises and
    paired jogs. Ambiguous symbols remain review items; no four-port 'cross'
    or guessed tee pair is generated.
    """
    from .review_conversion import ReviewNode, ReviewPipe
    if not math.isfinite(rise_m) or rise_m<0:
        raise ValueError('가지 상승 높이는 0 이상의 유한한 값이어야 합니다.')
    model.arc_report={'junctions':[], 'jogs':[], 'review':[], 'rise_m':rise_m,
                      'symbol_strokes':[]}
    if not arcs or rise_m<=1e-9: return
    original_nodes=dict(model.nodes);original_pipes=dict(model.pipes)
    full=defaultdict(set)
    for a,b in source_edges:
        full[a].add(b);full[b].add(a)
    drawn=defaultdict(set)
    for a,b in drawn_edges:
        drawn[a].add(b);drawn[b].add(a)
    by_source={n.source_node:nid for nid,n in model.nodes.items() if n.source_node is not None}
    selected_sources=set(by_source)
    protected_sources=set(protected_nodes)
    # A <=10 mm joined cluster is one symbolic site, but only near an arc.
    groups={}
    for arc in arcs:
        if arc.get('sa') is None or arc.get('sweep') is None: continue
        evidenced_contact = arc.get('connection_evidence') == 'arc_open_branch_unique_through'
        xy = (tuple(map(float, arc['connection_xy'])) if evidenced_contact
              else (float(arc['cx']),float(arc['cy'])))
        nearby=[(math.dist(points[n][:2],xy),n) for n in by_source
                if math.dist(points[n][:2],xy)<=max(30,float(arc.get('r',0)))]
        if not nearby: continue
        d,seed=min(nearby)
        # A symbol does not attach merely because a pipe passes through its rim.
        if d > (1e-3 if evidenced_contact else 30): continue
        cluster={seed};todo=[seed]
        for n in todo:
            for other in full[n]:
                if other not in cluster and math.dist(points[other][:2],xy)<=10:
                    cluster.add(other);todo.append(other)
        key=tuple(sorted(cluster))
        groups.setdefault(key,[]).append(arc)
    motifs=[]; candidates={}
    used=set()
    for cluster, symbols in sorted(groups.items()):
        present={by_source[n] for n in cluster if n in by_source}
        if not present or present & used: continue
        base=min(present,key=lambda n:(math.dist((model.nodes[n].x_mm,model.nodes[n].y_mm),
                       (symbols[0]['cx'],symbols[0]['cy'])),n))
        if base in _excluded:
            model.arc_report['review'].append({'node':base,'reason':'루프 복귀 경로와 상승 방향 충돌 — 높이 미확정, 원본 연결 유지'})
            continue
        if any(model.nodes[n].role=='pump' for n in present): continue
        center=model.nodes[base]
        stroke=overlapping_arc_stroke(points,full,drawn,cluster,symbols,
            selected_nodes=selected_sources,protected_nodes=protected_sources)
        ignored=stroke.edges if stroke else frozenset()
        ports=[]
        for n in cluster:
            for other in sorted(full[n]-set(cluster)):
                if tuple(sorted((n,other))) in ignored: continue
                v=_unit([points[other][i]-points[n][i] for i in (0,1)])
                if v and not any(sum(a*b for a,b in zip(v,p))>.996 for p in ports):
                    ports.append(v)
        opened=[v for v in ports if _open(symbols,v)]
        closed=[v for v in ports if not _open(symbols,v)]
        if not opened or not closed: continue
        internal=[p for p in model.pipes.values() if p.start in present and p.end in present]
        # Only a joined chain can collapse to a single port split (no cycle).
        if len(internal)!=len(present)-1 or any(p.length_m>.01 for p in internal):
            model.arc_report['review'].append({'node':base,'reason':'호 내부 연결·길이 확인 필요'});continue
        bindings={}
        for pid,p in model.pipes.items():
            for end,other in ((p.start,p.end),(p.end,p.start)):
                if end in present and other not in present:
                    q=model.nodes[other]
                    bindings[(pid,end)]=_open(symbols,(q.x_mm-center.x_mm,q.y_mm-center.y_mm))
        if not any(bindings.values()): continue
        motif={'base':base,'present':present,'open':opened,'closed':closed,
               'bindings':bindings,'internal':internal,'kind':'junction','stroke':stroke}
        if (len(ports) in (3,4) and len(closed)==2 and
                sum(a*b for a,b in zip(*closed))<-.996 and
                (len(opened)==1 or sum(a*b for a,b in zip(*opened))<-.996)):
            motifs.append(motif);used.update(present)
        elif len(ports)==2 and len(bindings)==2 and sum(bindings.values())==1:
            motif['kind']='jog'
            candidates[base]=motif
        else:
            model.arc_report['review'].append({'node':base,'reason':'호의 열린 포트·실제 접속 방향이 모호함'})
    # Recognize opposed pairs for the audit report, not as a prerequisite for
    # elevation. An independently evidenced two-port rise can feed a whole
    # elevated loop; that loop need not have a second height-change symbol.
    # Never infer elevation from a spanning-tree parent or traversal direction.
    adj=defaultdict(list)
    for pid,p in model.pipes.items():
        adj[p.start].append((pid,p.end));adj[p.end].append((pid,p.start))
    paired=set()
    for base,m in sorted(candidates.items()):
        if base in paired: continue
        first=next(pid for (pid,n),opened in m['bindings'].items() if opened)
        here=base;pid=first;seen={base};distance=0
        for _ in range(64):
            p=model.pipes[pid]; other=p.end if p.start==here else p.start
            distance+=p.length_m
            if distance>60 or other in seen: break
            if other in candidates:
                second=candidates[other]
                if other not in paired and second['bindings'].get((pid,other)):
                    motifs.extend((m,second));paired.update((base,other))
                    model.arc_report['jogs'].append([base,other])
                break
            if len(adj[other])!=2: break
            seen.add(other);here=other
            pid=next(k for k,_n in adj[here] if k!=pid)
    for base, motif in sorted(candidates.items()):
        if base not in paired:
            motif['kind'] = 'elbow_rise'
            motifs.append(motif)
    risers=[]
    for m in motifs:
        base=m['base'];old=model.nodes[base]
        top=f'V_{base}'
        # Remove only intra-symbol joint nodes; exterior source paths unchanged.
        for p in m['internal']: del model.pipes[p.id]
        for n in m['present']-{base}: del model.nodes[n]
        model.nodes[top]=ReviewNode(top,old.x_mm,old.y_mm,old.z_m)
        for (pid,end),opened in m['bindings'].items():
            p=model.pipes[pid];target=top if opened else base
            model.pipes[pid]=replace(p,**({'start':target} if p.start==end else {'end':target}))
        pid=f'V{len(risers)+1}'
        model.pipes[pid]=ReviewPipe(pid,base,top,rise_m)
        risers.append((pid,base,top))
        model.arc_ports[base]=[list(v)+[0.0] for v in m['closed']]+[[0.,0.,1.]]
        model.arc_ports[top]=[list(v)+[0.0] for v in m['open']]+[[0.,0.,-1.]]
        for n in m['present']-{base}: model.arc_ports[n]=[]
        model.arc_junctions[base]={'kind':'T' if len(m['open'])==2 else 'E',
                                 'top':top,'top_ports':len(m['open'])+1}
        model.arc_report['junctions'].append({'node':base,'top':top,'kind':m['kind']})
        if m['stroke']:
            stroke=m['stroke']
            model.arc_report['symbol_strokes'].append({'node':base,
                'ignored_port_edges':[list(e) for e in sorted(stroke.edges)],
                'drawn_path':list(stroke.drawn_path),'arc_xy':list(stroke.arc_xy),
                'radius_mm':stroke.radius_mm,'reason':'overlaid_arc_centre_tick',
                'source_preserved':True})
    if not risers: return
    horizontal=nx.Graph();horizontal.add_nodes_from(model.nodes)
    vertical_ids={pid for pid,_,_ in risers}
    horizontal.add_edges_from((p.start,p.end) for p in model.pipes.values() if p.id not in vertical_ids)
    comps=list(nx.connected_components(horizontal));component={n:i for i,c in enumerate(comps) for n in c}
    constraints=defaultdict(list)
    for pid,low,high in risers:
        a,b=component[low],component[high]
        constraints[a].append((b,rise_m,pid));constraints[b].append((a,-rise_m,pid))
    root=next(n for n,r in model.nodes.items() if r.role=='pump')
    elevations={component[root]:model.nodes[root].z_m};todo=[component[root]]
    parents={component[root]:(None,None)}
    conflicts=set()
    for i in todo:
        for j,delta,pid in constraints[i]:
            target=elevations[i]+delta
            if j in elevations:
                if abs(elevations[j]-target)>1e-7:
                    # Find just the inconsistent constraint cycle, not every
                    # symbol connected to this supply. Reject all its claims
                    # symmetrically; input order must not choose a 'winner'.
                    ancestors={};cur=i;trail=[]
                    while cur is not None:
                        ancestors[cur]=list(trail)
                        cur,edge=parents[cur]
                        if edge is not None: trail.append(edge)
                    cur=j;trail=[]
                    while cur not in ancestors:
                        cur,edge=parents[cur];trail.append(edge)
                    conflicts.update([pid]+ancestors[cur]+trail)
            else:
                elevations[j]=target;todo.append(j);parents[j]=(i,pid)
    if conflicts:
        excluded=_excluded|{low for pid,low,_ in risers if pid in conflicts}
        model.nodes=original_nodes;model.pipes=original_pipes
        model.arc_ports={};model.arc_junctions={}
        apply_arc_elevations(model,points,source_edges,arcs,rise_m,
                            drawn_edges=drawn_edges,protected_nodes=protected_nodes,
                            _excluded=frozenset(excluded))
        return
    for nid,n in list(model.nodes.items()):
        model.nodes[nid]=replace(n,z_m=elevations[component[nid]])
