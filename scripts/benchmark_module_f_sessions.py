"""Reproducible session-copy microbenchmark; no server or user data writes."""
from __future__ import annotations

import argparse
import gc
import json
from pathlib import Path
import statistics
import sys
import time
import tracemalloc

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from routes.module_f.cancellation import _session_copy
from routes.module_f.session_policy import SELECTION_FIELDS
from routes.module_f.session_workspace import SessionWorkspace


def history_benchmark(nodes: int, edits: int, runs: int) -> dict:
    """Compare actual graph-editor replay; all generated data remains in memory."""
    from core.network_editor import Network, apply_edit
    from routes.module_f import network_edit as ne, editor_history as history
    ne._boot()
    from services.cad_import.design.tables import PipeTablesG
    tables = PipeTablesG(
        nodes=[dict(label=str(i),x=i*1000,y=0,elevation=0,io_node='No') for i in range(nodes)],
        pipes=[dict(label='P'+str(i),**{'in':str(i),'out':str(i+1)},length=1,
                    dia=25,type='KSD 3507',c=120,eq_len=0) for i in range(nodes-1)])
    graph = Network.from_tables(tables)
    editor = dict(base=graph,current=graph,base_hash=ne.fingerprint(graph),library='synthetic',
                  commands=[],cursor=0)
    library = ne.catalog()
    for i in range(edits):
        command = dict(op='pipe',target='P0',schedule='KSD 3507',dn=25 if i%2 else 32)
        graph,_ = apply_edit(graph, command, library)
        editor = dict(editor,current=graph,commands=editor['commands']+[command],cursor=i+1)
        history.remember(editor)
    def run(cache):
        return history.replay(dict(editor,_replay_cache=cache),editor['commands'],edits-1,library)[0]
    cache = editor['_replay_cache']
    assert ne.fingerprint(run(cache)) == ne.fingerprint(run(None))
    return dict(nodes=nodes,commands=edits,retained_checkpoints=cache.count,
                full_replay=benchmark(lambda:run(None),runs),
                checkpoint_replay=benchmark(lambda:run(cache),runs))


def sample_session(nodes: int) -> dict:
    """Create a labelled synthetic multi-drawing session, not a live workload."""
    def state() -> dict:
        return {'edit': {'points':[(i, i % 71) for i in range(nodes)],
                         'edges':[(i,i+1) for i in range(nodes-1)]},
                'design': {'pipes':[{'label':str(i),'length':6.2,'dia':65,
                            'evidence':{'text_xy':[i,0],'head_count':13}} for i in range(nodes)]}}
    live=state()
    live.update(id='synthetic-performance',active='plan',slots={'system':state(),'machineroom':state()},
                merged=state(),editor={'commands':[{'op':'pipe','target':str(i),'dn':65} for i in range(50)]})
    return live


def benchmark(fn, runs: int) -> dict:
    """Measure wall time separately from allocation tracing overhead."""
    times=[]
    for _ in range(runs):
        gc.collect();start=time.perf_counter();result=fn()
        times.append((time.perf_counter()-start)*1000);del result
    gc.collect();tracemalloc.start()
    result=fn();_,peak=tracemalloc.get_traced_memory();tracemalloc.stop();del result
    return {'median_ms':round(statistics.median(times),2),'peak_mib':round(peak/1048576,2)}


def main() -> None:
    """Print JSON suitable for before/after comparison on this machine."""
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--nodes',type=int,default=5000)
    parser.add_argument('--runs',type=int,default=3)
    parser.add_argument('--history-nodes',type=int,default=200)
    parser.add_argument('--edits',type=int,default=60)
    args=parser.parse_args()
    sess=sample_session(args.nodes)
    print(json.dumps({'synthetic':True,'nodes_per_state':args.nodes,
                      'session_snapshot':benchmark(lambda:_session_copy(sess),args.runs),
                      'selection_snapshot':benchmark(lambda:_session_copy(sess,SELECTION_FIELDS),args.runs),
                      'selection_workspace':benchmark(lambda:SessionWorkspace(sess,SELECTION_FIELDS),args.runs),
                      'history':history_benchmark(args.history_nodes,args.edits,args.runs)},indent=2))


if __name__=='__main__':
    main()
