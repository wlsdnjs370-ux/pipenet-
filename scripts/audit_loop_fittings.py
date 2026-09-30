"""Read-only reconstruction of the B1F loop fitting report (no session writes)."""
import sys
from pathlib import Path
import math

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT/'cad_project_editor_g')]
from services.cad_import.pipeline.handoff import set_write_root
set_write_root(ROOT/'cad_project_editor_g/docs/import')


def live_board():
    from services.cad_import.edit.io import _board_from_data, load_edits
    from services.cad_import.pipeline.disp_cache import _disp_cache_load
    key = 'B1F 현장조사 소화설비 평면도_컨셉2-수정본 (1)'
    board = _board_from_data(key, _disp_cache_load(key))
    load_edits(board)
    board.network_mode = 'loop'
    board.sources = [board.snap_source(662681, -180629.4)]
    board.valves = list(board.sources)
    heads = [i for i,(x,y,*_) in enumerate(board.disks)
             if 751480 < x < 759297 and -145707 < y < -139033]
    return board, heads


if __name__ == '__main__':
    from services.cad_import.design.preserved import expand_preserved, review_fittings
    from services.cad_import.design.tables import build_design_tables
    from services.cad_import.pipeline.expand import stage1_body
    board, heads = live_board()
    got = expand_preserved(board, heads)
    bores = {p:(65,'review_default') for p in got['kfp']['pipe_data']}
    fits = review_fittings(got['kfp'], bores, got['physical_ports'])
    tbl = build_design_tables(got['kfp'],{},got['edge_ref'],[],bores=bores,fittings=fits)
    print('REPRO',len(heads),len(tbl.nodes),len(tbl.pipes),got['cycle_rank'])
    print('ISSUES',fits['unresolved_kind_items'])
    print('ELEVATION',got['kfp'].get('arc_report'))
    for i in (19624,36091,27699,27711,27760):
        xy=board.pts[i]
        print('CHECK',i,xy,'ARCS',[a for a in board.ho if math.dist((a['cx'],a['cy']),xy)<30],
              'PORTS',got['physical_ports'].get(f'N{i}'))
    c=(758463.6,-153051.5)
    print('ARCS', [a for a in board.ho if math.dist((a['cx'],a['cy']),c)<1000])
    for a,b in sorted(board.edges):
        if min(math.dist(board.pts[n],c) for n in (a,b))<100:
            print('BOARD',a,b,board.pts[a],board.pts[b], 'original', (a,b) in board.original_edges)
    if '--raw' in sys.argv:
        st = stage1_body(board.key)
        print('WORLD',vars(st['w']).keys())
        for l,col,a,b in st['mat_raw']:
            if min(math.dist(p[:2],c) for p in (a,b))<400:
                print('RAW MATERIAL',l,col,a,b)
        print('HO69', [s for s in st.get('sym_strokes',()) if str(s).find('7584')>=0][:12])
