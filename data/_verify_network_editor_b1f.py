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


        from routes.module_f import network_edit as ne
        from routes.module_f.api_network_edit import register
        ne.HISTORY_DIR=Path(work)/'editor_history'
        before=rows()
        def edit(command):
            st=client.get('/api/module-f/network-editor',query_string={'sid':sid}).get_json()
            rev=st['revision']
            p=client.post('/api/module-f/network-editor',json=dict(sid=sid,action='preview',revision=rev,command=command))
            if p.status_code != 200: return p
            r=client.post('/api/module-f/network-editor',json=dict(sid=sid,action='apply',revision=rev,command=command))
            return r
        failures=[]
        for p in sorted(sess['design']['tables'].pipes,key=lambda p:-p['length']):
            r=edit(dict(op='split',target=p['label'],distance=p['length']*.5))
            if r.status_code==200:
                split=r.get_json()['selection']['label'];break
            failures.append(r.get_json())
        else: raise AssertionError(failures)
        for axis in ('+Z','-Z','+Y','-Y','+X','-X'):
            r=edit(dict(op='extend',target=split,axis=axis,length=.75,schedule='CPVC2',dn=25))
            if r.status_code==200:
                end=r.get_json()['selection']['label'];break
            failures.append(r.get_json())
        else: raise AssertionError(failures)
        newpipe=next(p['label'] for p in sess['design']['tables'].pipes if p['out']==end)
        r=edit(dict(op='fitting',target=newpipe,fitting='ELBOW_90_STD',count=1,position=0))
        assert r.status_code==200,r.get_json()
        r=edit(dict(op='head',target=end,nozzle='SP-HEAD',flow=120))
        assert r.status_code==200,r.get_json()
        run('convert/run',dto={},outputs={'full_kfp':False,'worst_kfp':True,'worst_sdf':True})
        d=sess['design'];t=d['tables'];kf=json.loads(Path(sess['worst_kfp_path']).read_text(encoding='utf-8'))
        pid=next(pid for pid,lab in t.pipe_labels.items() if lab==newpipe)
        rec=kf['pipe_data'][pid]
        assert rec['diameter']==28.02 and rec['C']==150
        assert len(kf['pipe_data'])==len(t.pipes)
        assert len(kf['nodes_meta_runtime'])==len(t.nodes)
        assert len(t.nozzles)==len(before['nozzles'])+1
        assert sum(p['length'] for p in t.pipes)==__import__('pytest').approx(sum(p['length'] for p in before['pipes'])+.75)
        report['edited']={'split':split,'end':end,'pipe':newpipe,'axis':axis,'kfp_pipe':rec,
            'counts':{k:len(getattr(t,k)) for k in ('nodes','pipes','nozzles','equipment')},
            'source_counts':{k:len(before[k]) for k in before},'normal_rejections':failures}
        args.output.parent.mkdir(parents=True,exist_ok=True)
        for key in ('worst_kfp_path','design_sdf_path','design_slf_path','design_has_path'):
            if sess.get(key):shutil.copy2(sess[key],args.output.parent/Path(sess[key]).name)
        args.output.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(report,ensure_ascii=False),flush=True)

if __name__=='__main__':
    main()
