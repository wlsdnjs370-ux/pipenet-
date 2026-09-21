"""Compare selected-head confirmation with whole-network diagnostics in an isolated copy."""
from __future__ import annotations

import argparse
import copy
import hashlib
import importlib
import json
import shutil
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT/'tests')]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--key', default='B1F 현장조사 소화설비 평면도_컨셉2')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    import pandas  # Initialise native dependencies on the main thread.
    from _workdir_iso import isolated_workdir
    from routes.module_f import jobs, attach
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import handoff
    source_world = Path(handoff.handoff_path(args.key))
    with isolated_workdir(prefix='design_speed_', copy_key=args.key) as work:
        if source_world.is_file():
            shutil.copy2(source_world, handoff.handoff_path(args.key))
        from services.cad_import.edit.session import EditSession
        srv = importlib.import_module('대조 서버')
        srv.UPLOAD_DIR = Path(work)/'uploads'
        srv.UPLOAD_DIR.mkdir()
        client = srv.app.test_client()
        with client.session_transaction() as session:
            session['authed'] = True
        started = time.perf_counter()
        es = EditSession.open(args.key)
        print('REOPEN', round(time.perf_counter()-started,2), flush=True)
        board = es.board
        source = min({n for edge in board.edges for n in edge},
                     key=lambda n: (board.pts[n][0]-662692)**2+(board.pts[n][1]-175774)**2)
        board.sources = [source]
        board.valves = [source]
        board._source_nodes_cache = None
        sess = jobs._new_session()
        sess.update(edit=es,key=args.key,method='manual')
        sid = sess['id']

        def run(endpoint, **body):
            start = time.perf_counter()
            r=client.post('/api/module-f/'+endpoint, json={'sid':sid,**body})
            assert r.status_code == 200, r.get_json()
            deadline = time.monotonic()+300
            while (sess.get('job') or {}).get('state') == 'run':
                if time.monotonic()>deadline:
                    raise TimeoutError(endpoint)
                time.sleep(.05)
            job=sess.get('job') or {}
            if job:
                assert job['state']=='done', job.get('error')
                assert (job.get('result') or {}).get('ok'), job.get('result')
            elapsed=round(time.perf_counter()-start,4)
            print('DONE',endpoint,elapsed,flush=True)
            return elapsed

        report={'key':args.key,'seconds':{}}
        report['seconds']['selection']=run('edit/worst',k=10,zones=[[655000,165000,675000,180000]])
        picked=copy.deepcopy(sess['worst'])
        report['seconds']['confirm_selected']=run('design/build',k=10)

        def rows():
            tbl=sess['design']['tables']
            return copy.deepcopy({name:getattr(tbl,name) for name in ('nodes','pipes','nozzles','fittings','equipment')})

        before=rows()
        report['summary']=copy.deepcopy(sess['job']['result']['summary'])
        # Report true crossing-check timing and preserve a reproducible geometry
        # fixture. Do not re-run the old 2-km index on this huge network.
        from services.cad_import.convert import planar
        original_crossings=planar._find_x_crossings
        crossing_stats=[]
        def crossing_probe(pos, edges):
            start=time.perf_counter()
            found=original_crossings(pos, edges)
            crossing_stats.append({'edges':len(edges),'crossings':len(found),
                                   'seconds':round(time.perf_counter()-start,4)})
            return found
        planar._find_x_crossings=crossing_probe
        try:
            report['seconds']['diagnose_all']=run('design/diagnose')
        finally:
            planar._find_x_crossings=original_crossings
        report['crossings']=crossing_stats
        assert rows()==before, 'Diagnostics changed hydraulic rows'
        report['diagnostic_summary']=copy.deepcopy(sess['job']['result']['summary'])
        original_probe=attach.design_probe
        attach.design_probe=lambda s,e,picked,**kw: attach.wet_heads(s,e,**kw)
        try:
            report['seconds']['confirm_previous_probe_cached']=run('design/build',k=10)
        finally:
            attach.design_probe=original_probe
        assert rows()==before, 'Whole-network and selected-head confirmation disagree'
        assert sess['worst']==picked, 'Confirmation changed selection'
        report['identical_hydraulic_rows']=True
        report['fingerprint']=hashlib.sha256(json.dumps(before,sort_keys=True,ensure_ascii=False).encode()).hexdigest()
        args.output.parent.mkdir(parents=True,exist_ok=True)
        args.output.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(report,ensure_ascii=False),flush=True)


if __name__=='__main__':
    main()
