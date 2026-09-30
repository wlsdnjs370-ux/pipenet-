"""Module F adapters for explicit, unverified loop/grid calculation inputs."""
from __future__ import annotations

from collections import Counter, defaultdict
import math

from src.pipenet_converter.graph.review_conversion import build_review_network


REVIEW_NOTICE = ("검토용 루프·그리드 입력입니다. 규약 관경은 적용하지 않습니다. "
                 "유량 방향·압력·최불리 여부는 PIPENET에서 계산해야 합니다. "
                 "확인된 호 기호는 입력한 가지 상승 높이로 전개하며, 복귀 경로의 표고를 검증합니다. "
                 "모호하거나 높이가 충돌하는 기호는 원본 연결·기준 표고로 남기고 미확정으로 표시합니다. 실제 표고와 티 손실, 공급 경계조건·노즐의 "
                 "요구 유량·최소 압력을 확인해야 계산할 수 있습니다.")


def expand_preserved(board, heads, *, selected_source=None, datum_m=0.0, dto=None,
                     source_edges=None, area_scope=None) -> dict:
    """Convert every selected path, with explicit local head stems and no tree routing."""
    from services.cad_import.design.flow import network_for_board, source_index
    from services.cad_import.dto import default_dto
    network = network_for_board(board, index=source_index(board,selected_source))
    dimensions = default_dto()
    dimensions.update({k:v for k,v in (dto or {}).items() if v is not None})
    protected=set(board.sources)|set(getattr(board,'valves',()) or ())
    for ns in board.hnodes:
        protected.update(ns)
    xy_nodes=defaultdict(set)
    for i,p in enumerate(board.pts):
        xy_nodes[tuple(p[:2])].add(i)
    for rec in getattr(board,'joins',()) or ():
        for a,b in rec.get('bridges',()):
            for p in (a,b):
                protected.update(xy_nodes.get(tuple(p[:2]),()))
    for key in (getattr(board,'edge_len_mm',None) or {}):
        protected.update(int(n) for n in (key.split('|') if isinstance(key,str) else key))
    model = build_review_network(network, board.pts, heads,
        dict(enumerate(board.disk_kinds)), datum_m=float(datum_m),
        upright_m=float(dimensions['upright_1_m']),
        pendant_rise_m=float(dimensions['pendant_1_m']),
        pendant_drop_m=float(dimensions['pendant_2_m']),
        combo_rise_m=float(dimensions['combo_1_m']),
        combo_up_m=float(dimensions['combo_up_m']),
        combo_drop_m=float(dimensions['combo_3_m']),
        combo_arm_m=float(dimensions['combo_2_m']),
        arcs=getattr(board,'ho',()) or (),branch_rise_m=float(dimensions['branch_rise_m']),
        drawn_edges=tuple(board.original_edges),protected_nodes=tuple(protected),
        source_edges=source_edges)
    ox = min(n.x_mm for n in model.nodes.values())
    oy = min(n.y_mm for n in model.nodes.values())
    nodes = {n.id: {'id':n.id, 'coords':[(n.x_mm-ox+1000)/1000,(n.y_mm-oy+1000)/1000,n.z_m],
        'elevation_m':n.z_m,'type_id':n.role,'type':'Head' if n.role=='head' else '기본',
        'k_factor_si':dimensions['head_k'] if n.role=='head' else None,
        'required_pressure_bar':dimensions.get('min_pressure_bar') or 0,
        'is_active':n.role=='head'} for n in model.nodes.values()}
    pipes = {p.id: {'start':p.start,'end':p.end,'length_m':p.length_m,
                   'C':120.0,'equivalent_length':0.0,'fittings':[]} for p in model.pipes.values()}
    node_ref = {n.id:n.source_node for n in model.nodes.values() if n.source_node is not None}
    edge_ref = {p.id:p.source_edge for p in model.pipes.values() if p.source_edge is not None}
    report = dict(network.report(), roots=list(network.reference.roots), revision=network.revision)
    if area_scope is not None:
        report.update(scope='region_and_supply', area_scope=dict(area_scope),
                      heads=len(heads), edges=area_scope['edges'],
                      cycle_rank=area_scope['cycle_rank'], length_m=area_scope['length_m'])
    from services.cad_import.design.cycle_source import cycle_source
    recovery = cycle_source(board)
    report['cycle_recovery'] = recovery.report()
    report['duplicate_joins'] = [r for r in recovery.records if r['status']=='duplicate_removed']
    kfp = {'nodes_meta_runtime':nodes,'pipe_data':pipes,'network_mode':network.mode,
           'review_only':True,'hydraulics_solved':False,'review_notice':REVIEW_NOTICE,
           'arc_junctions':model.arc_junctions,'arc_report':model.arc_report,
           'arc_ports':model.arc_ports}
    phys = Counter(n for e in network.reference.lengths_mm for n in e)
    ports = defaultdict(list)
    for a,b in network.reference.lengths_mm:
        for n,other in ((a,b),(b,a)):
            nid=f'N{model.source_aliases.get(n,n)}'
            if nid in node_ref:
                ports[nid].append([(board.pts[other][j]-board.pts[n][j])/1000 for j in range(2)]+[0.0])
    ports.update(model.arc_ports)
    return {'ok':True,'kfp':kfp,'edge_ref':edge_ref,'node_ref':node_ref,
            'origin_mm':[ox,oy], 'network_mode':network.mode,'cycle_rank':model.cycle_rank,
            'collapsed_zero_nodes':sum(a!=b for a,b in model.source_aliases.items()),
            'flow_report':report,'tree_loads':{},'source_paths':{p:[e] for p,e in edge_ref.items()},
            'phys':{nid:phys[i] for nid,i in node_ref.items()},
            'fitting_node_ref':dict(node_ref),'declared_pipes':[p for p,e in edge_ref.items()
                if e in getattr(board,'edge_len_mm',{}) or
                   f'{e[0]}|{e[1]}' in getattr(board,'edge_len_mm',{})],
            'physical_ports':dict(ports),
            'node_head_kinds':{n.id:n.head_kind for n in model.nodes.values() if n.role=='head'},
            'review_only':True,'review_notice':REVIEW_NOTICE,'head_dimensions':dimensions,
            'diagnostics_scope':'selected','excluded_heads':None,'candidate_heads':len(heads),
            'total_heads':len(board.disks),'chain_report':{}}


def review_bores(net, edge_ref, context, default_mm: int, overrides=None):
    """Explicit override > unambiguous annotation > user-chosen review diameter.

    A default is never NFPC sizing and remains unconfirmed even after export.
    Known competing annotations remain in the evidence, not silently accepted.
    """
    from services.cad_import.design.bore import Bores
    if context.summary.get('policy') == 'drawing_first_v1':
        from src.pipenet_converter.graph.diameter_definition import resolve_pipe_definitions
        values, evidence, changed = resolve_pipe_definitions(
            net, edge_ref, context, {}, tree=False, rule=None, overrides=overrides)
        out = Bores(values)
        out.evidence, out.overridden = evidence, changed
        return out
    out = Bores()
    for pid, pipe in net['pipe_data'].items():
        ref = edge_ref.get(pid)
        evidence = dict(context.for_path(ref)) if ref is not None else {}
        text = evidence.get('text_mm') if not evidence.get('block_export') else None
        dia, src = (int(text),'text') if text else (default_mm,'review_default')
        ov = (overrides or {}).get(tuple(sorted(ref))) if ref is not None else None
        if ov:
            out.overridden[pid] = {'dia':int(ov[0]),'note':str(ov[1]),'orig_dia':dia,
                                  'orig_src':src,'a':min(ref),'b':max(ref)}
            evidence['manual'] = {'previous_mm':dia,'value_mm':int(ov[0]),'note':str(ov[1])}
            dia, src = int(ov[0]), 'user'
        evidence.update(version=3,source=src,auto_mm=dia,rule_mm=None,head_count=None,
                        confirmed=src=='user',review_only=True,block_export=False,
                        review_default_mm=default_mm)
        if src=='review_default':
            evidence.setdefault('review_reasons',[]).append('사용자가 지정한 검토용 기본 관경 — 도면 관경·규약 관경이 아닙니다.')
        out[pid] = (dia,src)
        out.evidence[pid] = evidence
        # The exporter resolves this nominal size against the selected SLF.
        pipe.update(nominal_mm=dia,diameter=dia,type='KSD3507',bore_provenance=evidence)
    return out


def review_fittings(net, bores, physical_ports=None) -> dict:
    """Direction-independent elbows; geometric branch-port tees need solver review.

    A spanning-tree parent is deliberately never used. Four-way junctions are
    review items, not invented crosses or an invented pair of tees.
    """
    from services.cad_import.design.fitting import resolve_eq_len
    nodes, pipes = net['nodes_meta_runtime'], net['pipe_data']
    incident = defaultdict(list)
    for pid,p in pipes.items():
        incident[p['start']].append((pid,p['end']))
        incident[p['end']].append((pid,p['start']))
    per = {pid:{'fittings':[],'instances':[],'equivalent_length':0.0} for pid in pipes}
    issues, lengths = [], []
    counts = Counter()
    for nid,links in incident.items():
        links = list(links)
        if len(links)<2: continue
        center = nodes[nid]['coords']
        vectors = []
        for pid,other in links:
            v = tuple(nodes[other]['coords'][i]-center[i] for i in range(3))
            norm = math.sqrt(sum(c*c for c in v))
            vectors.append(tuple(c/norm for c in v) if norm>1e-9 else None)
        # A selected straight-through passage has no permitted fitting loss,
        # including when unselected physical side ports remain in the drawing.
        if len(vectors)==2 and all(vectors) and sum(a*b for a,b in zip(*vectors)) < -.996:
            continue
        # Removed operation-scenario branches are still physical fitting ports.
        # A pruned tee must not turn into an elbow.
        for v in (physical_ports or {}).get(nid,()):
            norm = math.sqrt(sum(c*c for c in v))
            if norm<=1e-9: continue
            unit=tuple(c/norm for c in v)
            if any(old and sum(a*b for a,b in zip(old,unit))>0.996 for old in vectors): continue
            vectors.append(unit); links.append((None,None))
        hit = None
        pairs = [(math.degrees(math.acos(max(-1,min(1,sum(a*b for a,b in zip(vectors[i],vectors[j])))))),i,j)
                 for i in range(len(links)) for j in range(i) if vectors[i] and vectors[j]]
        if len(links)==2 and pairs:
            bend = 180-pairs[0][0]
            if bend<5: continue
            # Stable loss owner, independent of dict order/reference arrows.
            # Prefer the smaller bore at a reducer; keep the physical node.
            owner = min((pid for pid,_ in links if pid is not None),
                        key=lambda pid:(bores[pid][0],str(pid)))
            if abs(bend-90)<5: hit=('elbow',owner)
            elif abs(bend-45)<5: hit=('elbow-45',owner)
        elif len(links)==3 and pairs:
            angle,i,j = max(pairs)
            if angle>175:
                branch = next(k for k in range(3) if k not in (i,j))
                hit=('tee',links[branch][0])
        if not hit:
            issues.append({'pipe':links[0][0],'node':nid,'n':1,'ports':len(links),
                           'reason':'루프 접속 형상·티 배치 또는 180도 헤드 접속관 확인 필요'})
            continue
        kind,pid = hit
        if pid is None: continue  # The branch is inactive; no straight-through tee loss.
        per[pid]['fittings'].append(kind); counts[kind]+=1
        per[pid]['instances'].append({'node':nid,'type':kind,
            'flow_direction':'solver_reference',
            'loss_status':'provisional_branch_port' if kind=='tee' else 'geometric_elbow'})
        length,_why = resolve_eq_len(kind,bores[pid][0])
        if length is None:
            lengths.append({'pipe':pid,'node':nid,'kind':kind,'dia':bores[pid][0]})
        else: per[pid]['equivalent_length']+=length
    for issue in net.get('arc_report',{}).get('review',()):
        nid=issue['node']
        existing=next((i for i in issues if i['node']==nid),None)
        if existing:
            existing['reason'] += ' · ' + issue['reason']
            continue
        if incident.get(nid):
            issues.append(dict(issue,pipe=incident[nid][0][0],n=1))
    return {'per_pipe':per,'counts':dict(counts),'unresolved_kind':len(issues),
            'unresolved_length':len(lengths),'unresolved_kind_items':issues,
            'unresolved_length_items':lengths,'unresolved_pairs':[]}
