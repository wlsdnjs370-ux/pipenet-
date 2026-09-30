"""HTTP-workflow adapter for cycle-preserving, explicitly provisional inputs."""
from __future__ import annotations


def assemble_review_tables(sess: dict, board, got: dict, context, cfg: dict):
    """Use explicit properties for both review file formats and fitting lengths."""
    from routes.module_f import overrides
    from routes.module_f.api_design import _bore_ov_map, schedule_bores_mm
    from services.cad_import.design.preserved import review_bores, review_fittings
    from services.cad_import.design.tables import build_design_tables, bfs_order
    from src.pipenet_converter.graph.preserved_properties import apply_review_properties, sync_review_kfp
    rows = overrides.ensure_loaded(sess)
    h_policy = cfg.get('diameter_policy') == 'drawing_first_v1'
    head_rows = [r for r in rows if h_policy and r.get('kind')=='head' and r.get('field') in ('k_factor_si','required_pressure_bar')]
    rows = [r for r in rows if r not in head_rows]
    if head_rows:
        index = overrides.build_index(got,board)
        applied_heads, missed_heads = overrides.resolve(head_rows,index)
        if missed_heads:
            raise ValueError('수정한 헤드를 현재 배관망에서 찾을 수 없습니다.')
        for key, row, nid in applied_heads:
            if nid is None:
                raise ValueError('수정한 헤드를 현재 배관망에서 찾을 수 없습니다.')
            meta = got['kfp']['nodes_meta_runtime'][nid]
            for field in overrides._META_HEAD[row['field']]:
                meta[field]=row['new']
    if overrides.ensure_ops_loaded(sess) or any(r.get('field') not in ('dia','c','type','elevation','length') for r in rows):
        raise ValueError('루프 검토 변환에서는 기존 위상·특수 요소 수정이 자동 이식되지 않습니다. 편집 기록을 확인해 정리해 주세요.')
    if any((sess.get('fitting_overrides') or {}).values()):
        raise ValueError('기존 트리 부속 직접 입력은 루프에 자동 이식하지 않습니다. 부속 직접 입력을 비운 뒤 재확정하고 검토 JSON의 미해결 항목을 확인하세요.')
    bores = review_bores(got['kfp'],got['edge_ref'],context,cfg.get('review_default_mm'),_bore_ov_map(sess))
    physical_ports = dict(got['physical_ports'])
    def make():
        table=build_design_tables(got['kfp'],got.get('worst') or {},got['edge_ref'],[],
            bores=bores,fittings=review_fittings(got['kfp'],bores,physical_ports),
            board_pts=board.pts,origin_mm=got['origin_mm'],default_schedule=cfg['schedule'],
            node_head_kinds=got['node_head_kinds'],declared_pipes=got['declared_pipes'],
            valve_nodes=[f'N{n}' for n in board.valves if f'N{n}' in got['kfp']['nodes_meta_runtime']])
        reverse={lab:pid for pid,lab in table.pipe_labels.items()}
        for row in table.pipes:
            row['length']=round(got['kfp']['pipe_data'][reverse[row['label']]]['length_m'],9)
        return table
    tbl=make()
    keys=overrides.label_keys(got,board,tbl)
    apply_review_properties(tbl,keys,rows)
    root = next(n for n,r in got['kfp']['nodes_meta_runtime'].items() if r['type_id']=='pump')
    order,_,_,_ = bfs_order(got['kfp'],root)
    before={n:tuple(r['coords']) for n,r in got['kfp']['nodes_meta_runtime'].items()}
    sync_review_kfp(tbl,got['kfp'],{str(i):n for i,n in enumerate(order,1)})
    changed={n for n,r in got['kfp']['nodes_meta_runtime'].items() if tuple(r['coords'])!=before[n]}
    affected=set(changed)
    for p in got['kfp']['pipe_data'].values():
        if p['start'] in changed or p['end'] in changed:
            affected.update((p['start'],p['end']))
    for n in affected:
        # Once edited, stale omitted-port/elevation evidence cannot override
        # the user's current geometry. Keep the active connections as truth.
        physical_ports.pop(n,None)
        got['kfp'].get('arc_ports',{}).pop(n,None)
    got['physical_ports'] = physical_ports
    reverse={lab:pid for pid,lab in tbl.pipe_labels.items()}
    for p in tbl.pipes:
        pending = cfg.get('diameter_policy') == 'drawing_first_v1' and not p['dia']
        if not pending and p['dia'] not in schedule_bores_mm(p['type']):
            raise ValueError(f"{p['label']}: 해당 배관 규격에 없는 호칭경입니다.")
        pid=reverse[p['label']]
        bores[pid]=(int(p['dia']),p['dia_src'])
        bores.evidence[pid]=p['bore_provenance']
    # Recompute geometry-dependent elbows and diameter-dependent loss values.
    tbl=make()
    apply_review_properties(tbl,keys,rows)
    if h_policy:
        from routes.module_h_attributes import apply_head_properties
        apply_head_properties(tbl,keys,head_rows)
        for pipe in tbl.pipes:
            pipe.setdefault('bore_provenance',{})['stable_key']=keys['pipe'].get(str(pipe['label']))
    return tbl,keys


def build_review_design(sess: dict, es, cfg: dict, selected_source=None) -> dict:
    """Build the selected operation scenario; never choose another set of heads."""
    from routes.module_f.api_design import _dia_texts, _bore_ov_map, _summary, schedule_bores_mm
    from routes.module_f.bore_context import build_bore_context
    from routes.module_f.selection import _selection_sig
    from routes.module_f import overrides
    from services.cad_import.design.flow import network_for_board, source_index
    from services.cad_import.design.preserved import expand_preserved, review_bores, review_fittings, REVIEW_NOTICE
    from services.cad_import.design.tables import build_design_tables
    default_mm = cfg.get('review_default_mm')
    if cfg.get('diameter_policy') != 'drawing_first_v1' and default_mm not in schedule_bores_mm(cfg['schedule']):
        return {'ok':False,'error':'루프·그리드 검토용 기본 관경을 직접 선택하세요. 규약 관경은 적용하지 않습니다.'}
    worst = sess.get('worst') or {}
    try:
        if sess.get('selection_mode') == 'area_all':
            from routes.module_f.area_selection import validate_area_selection
            validate_area_selection(sess)
        network = network_for_board(es.board, index=source_index(es.board,selected_source or worst.get('source_tag')))
        if not worst.get('heads') or worst.get('flow_revision') != network.revision:
            raise ValueError('물흐름 → 작동 헤드 후보 선정을 먼저 실행하세요. 오래된 경로는 출력하지 않습니다.')
        scoped = worst.get('selection_mode') == 'area_all'
        expected = (network.resolve_selection(worst['heads'], source_edges=worst['edges'])
                    if scoped else network.selected_edges(worst['heads']))
        if expected != set(worst['edges']):
            raise ValueError('선정 경로가 현재 연결 보존망과 다릅니다. 다시 선정하세요.')
        got = expand_preserved(es.board, worst['heads'], selected_source=selected_source or worst.get('source_tag'),
                               datum_m=cfg.get('review_datum_m',0),dto=cfg.get('review_dto'),
                               source_edges=tuple(expected) if scoped else None,
                               area_scope=worst.get('area_scope') if scoped else None)
        got['worst'] = dict(worst)
        context = build_bore_context(sess,es.board,network.reference,_dia_texts(sess))
        tbl,keys=assemble_review_tables(sess,es.board,got,context,cfg)
        if worst.get('selection_mode') == 'area_all':
            from routes.module_f.area_selection import validate_area_nozzles
            validate_area_nozzles(es.board,worst,tbl)
            tbl.meta = [('영역 내 전체 헤드 수' if name == '기준개수 K' else name, value)
                        for name,value in tbl.meta]
            tbl.meta.append(('헤드 선정 방식','지정 영역 전체 · 기준개수 제한 없음'))
            tbl.meta.append(('배관 추출 범위','지정 영역 내부 + 급수 연결 경로 · 영역 밖 우회 경로 제외'))
        tbl.meta.extend([('출력 상태',REVIEW_NOTICE),('배관망 방식',network.mode),
                         ('검토용 기본 관경(mm)',str(default_mm)),('수리계산 완료','아니오')])
        if cfg.get('diameter_policy') == 'drawing_first_v1':
            tbl.meta = [(k,v) for k,v in tbl.meta if k!='검토용 기본 관경(mm)']
            tbl.meta.append(('관경 적용 방식','도면 → 연속 관로 → 반복 구조 → 사용자 보정'))
        for p in tbl.pipes:
            p.update(network_mode=network.mode,review_only=True,flow_direction='solver_reference')
        got['review_default_mm'] = default_mm
        got['handoff'] = {'picked':len(worst['heads']),'from_edit':True,'messages':[REVIEW_NOTICE]}
        sess['design'] = {'got':got,'tables':tbl,'k':len(worst['heads']),
            'source':worst.get('source_tag'),'schedule':cfg['schedule'],'marks':{},'keys':keys,
            'bore_inference':context.summary,'sig':_selection_sig(sess)}
        from routes.module_f.network_edit import accept_rebuilt
        accept_rebuilt(sess)
        summary = _summary(got,tbl)
        summary.update(review_only=True,review_notice=REVIEW_NOTICE,cycle_rank=got['cycle_rank'],
                       review_default_mm=default_mm,
                       provisional_pipes=sum(p['dia_src']=='review_default' for p in tbl.pipes))
        return {'ok':True,'summary':summary}
    except (ValueError,TypeError) as exc:
        return {'ok':False,'error':str(exc)}


def emit_review_outputs(sess: dict, outputs: dict, upload_dir) -> dict:
    """Share the same tables for SDF and selected KFP; full KFP uses all paths."""
    import json
    from pathlib import Path
    from routes.module_f.api_design import emit_design_files, _dia_texts, _bore_ov_map
    from routes.module_f.bore_context import build_bore_context
    from routes.module_f.kfp_export import finish_kfp
    from services.cad_import.design.flow import network_for_board, source_index
    from services.cad_import.design.preserved import expand_preserved, review_bores, REVIEW_NOTICE
    from services.cad_import.design.emit import emit_design_kfp
    design = sess['design']
    cfg = sess['design_settings']
    out_dir = Path(upload_dir)/'module_f'/f"{sess['id']}_review"
    out_dir.mkdir(parents=True,exist_ok=True)
    summary = {'outputs':dict(outputs),'review_only':True,'review_notice':REVIEW_NOTICE}
    try:
        if outputs['full_kfp']:
            board = sess['edit'].board
            net = network_for_board(board,index=source_index(board,design.get('source')))
            heads = sorted(set(net.reference.representatives.values()))
            got = expand_preserved(board,heads,selected_source=design.get('source'),
                datum_m=cfg.get('review_datum_m',0),dto=cfg.get('review_dto'))
            context=build_bore_context(sess,board,net.reference,_dia_texts(sess))
            tbl,_=assemble_review_tables(sess,board,got,context,cfg)
            for n in got['kfp']['nodes_meta_runtime'].values(): n['is_active']=False
            path=out_dir/'전체망_검토용.kfp'
            res=emit_design_kfp(tbl,got,str(path))
            kfp,compatibility=finish_kfp(path,res['kfp'])
            sess['kfp'],sess['kfp_path']=kfp,str(path)
            summary['full']={'nodes':len(kfp['nodes_meta_runtime']),'pipes':len(kfp['pipe_data']),
                'bytes':path.stat().st_size,'filename':path.name,'compatibility':compatibility}
        if outputs['worst_kfp']:
            path=out_dir/'선정망_검토용.kfp'
            res=emit_design_kfp(design['tables'],design['got'],str(path))
            kfp,compatibility=finish_kfp(path,res['kfp'])
            sess['worst_kfp_path']=str(path)
            summary['worst']={'k':len(design['got']['worst']['heads']),
                'nodes':len(kfp['nodes_meta_runtime']),'pipes':len(kfp['pipe_data']),
                'bytes':path.stat().st_size,'filename':path.name,'compatibility':compatibility}
        if outputs['worst_sdf']:
            path,error=emit_design_files(sess,upload_dir)
            if error: raise ValueError(error)
            summary['design']={'sdf':path.name,'slf':path.with_suffix('.slf').name,'bytes':path.stat().st_size}
        return {'ok':True,'summary':summary,'stats':{},'diagnostics':[REVIEW_NOTICE]}
    except (ValueError,OSError) as exc:
        return {'ok':False,'blockers':[{'code':'review_export_failed','message':str(exc)}]}
