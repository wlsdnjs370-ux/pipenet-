"""Read-only snapshot and serial-segment diagnostics for a known H session."""
from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import asdict
import html
import json
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT / 'core')]

from core.remote30_full_network import CombinedTables
from src.pipenet_converter.graph.export_compaction import protected_nodes, _straight


def main() -> None:
    """Inspect current tables without POSTing a build, edit, or export request."""
    parser = argparse.ArgumentParser()
    parser.add_argument('sid')
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    import requests
    client = requests.Session()
    base = 'http://127.0.0.1:5051'
    login = client.get(base + '/login', timeout=15)
    hint = re.search(r'<span class="pw-value">([^<]+)<', login.text)
    if hint:
        client.post(base + '/login', data={'password': html.unescape(hint[1]),
                    'next': '/module-h'}, timeout=15).raise_for_status()
    response = client.get(base + '/api/module-f/design/preview',
                          params={'sid': args.sid}, timeout=30)
    response.raise_for_status()
    snapshot = response.json()
    if not snapshot.get('tables') or snapshot.get('stale'):
        raise ValueError('Current non-stale plan tables are required.')
    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / 'snapshot.json').write_text(json.dumps(snapshot, ensure_ascii=False,
                                                     indent=2), encoding='utf-8')
    response = client.get(base + '/api/module-f/network-editor',
                          params={'sid': args.sid, 'scope': 'design'}, timeout=30)
    response.raise_for_status()
    editor = response.json()
    (args.out / 'editor.json').write_text(json.dumps(editor, ensure_ascii=False, indent=2), encoding='utf-8')
    print('editor', {k: editor.get(k) for k in ('undo', 'redo', 'conflict', 'notice', 'history')})
    data = snapshot['tables']
    tables = CombinedTables(**{k: data[k] for k in ('nodes', 'pipes', 'nozzles', 'fittings', 'equipment')})
    nodes = {str(n['label']): (n['x']/1000, n['y']/1000, n['elevation']) for n in tables.nodes}
    adj = {n: [] for n in nodes}
    for p in tables.pipes:
        for k in ('in', 'out'):
            adj[str(p[k])].append(p)
    stops = protected_nodes(tables)
    serial = []
    for node, pipes in adj.items():
        if len(pipes) != 2 or node in stops:
            continue
        a, b = pipes
        if str(a['out']) != node:
            a, b = b, a
        if str(a['out']) != node or str(b['in']) != node:
            continue
        if _straight(nodes[str(a['in'])], nodes[node], nodes[str(b['out'])]):
            serial.append((node, a, b))
    report = dict(nodes=len(nodes), pipes=len(tables.pipes), heads=len(tables.nozzles),
                  referenced=sum(bool(p.get('dia')) for p in tables.pipes),
                  unreferenced=sum(not p.get('dia') for p in tables.pipes),
                  unprotected_straight=len(serial),
                  unknown_pairs=sum(not a['dia'] and not b['dia'] for _, a, b in serial),
                  unknown_reasons=dict(Counter((p.get('bore_provenance') or {}).get('reason')
                                               for p in tables.pipes if not p['dia'])))
    (args.out / 'diagnosis.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
