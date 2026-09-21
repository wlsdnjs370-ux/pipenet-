"""Transactional sprinkler graph edits. Coordinates/lengths are metres.

Display coordinates never enter this module. Existing declared lengths are
preserved (system diagrams can be schematic). Every edit works on a copy.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from math import dist, isfinite
from typing import Any

EPS = 1e-7


class EditError(ValueError):
    """An edit cannot be applied without ambiguous or invalid geometry."""


def number(value: Any, name: str, low: float = 0, high: float = 10000) -> float:
    """Read a finite, bounded user number."""
    try:
        result = float(value)
    except (ValueError, TypeError):
        raise EditError(f"{name}: 숫자를 입력하세요.") from None
    if not isfinite(result) or not low <= result <= high:
        raise EditError(f"{name}: {low:g}~{high:g} 범위여야 합니다.")
    return result


@dataclass
class Node:
    """Real position plus the original table row."""
    xyz: tuple[float, float, float]
    row: dict


@dataclass
class Pipe:
    """Directed endpoints plus explicit hydraulic values."""
    a: str
    b: str
    row: dict


@dataclass
class Network:
    """Editable graph, retaining tables, equipment and export metadata."""
    nodes: dict[str, Node]
    pipes: dict[str, Pipe]
    tables: Any
    protected: set[str] = field(default_factory=set)

    @classmethod
    def from_tables(cls, tables: Any, *, positions: dict | None = None,
                    protected: set | None = None) -> Network:
        """Convert table mm/mm/m positions to metre coordinates."""
        tables = deepcopy(tables)
        positions = positions or {}
        nodes = {str(n['label']): Node(tuple(positions.get(str(n['label']),
                 (float(n.get('x', 0))/1000, float(n.get('y', 0))/1000,
                  float(n.get('elevation', 0))))), n) for n in tables.nodes}
        pipes = {str(p['label']): Pipe(str(p['in']), str(p['out']), p)
                 for p in tables.pipes}
        roots = {k for k, n in nodes.items() if n.row.get('io_node') == 'Input'}
        for attr in ('pumps', 'valves'):
            for row in getattr(tables, attr, []):
                roots.update(str(row[k]) for k in ('in', 'out') if k in row)
        return cls(nodes, pipes, tables, roots | set(protected or ()))

    def to_tables(self) -> Any:
        """Write graph changes back without rounding off-axis coordinates."""
        for lab, n in self.nodes.items():
            n.row.update(label=lab, x=n.xyz[0]*1000, y=n.xyz[1]*1000,
                         elevation=n.xyz[2])
        for lab, p in self.pipes.items():
            p.row.update(label=lab, **{'in': p.a, 'out': p.b},
                         elev=self.nodes[p.b].xyz[2]-self.nodes[p.a].xyz[2])
        self.tables.nodes = [n.row for n in self.nodes.values()]
        self.tables.pipes = [p.row for p in self.pipes.values()]
        for attr in ('fittings', 'equipment'):
            for row in getattr(self.tables, attr, []):
                p = self.pipes.get(str(row.get('pipe')))
                if p:
                    row.update({'in': p.a, 'out': p.b})
        return self.tables

    def incident(self, node: str) -> list[str]:
        return [k for k, p in self.pipes.items() if node in (p.a, p.b)]

    def heads(self) -> set[str]:
        return {str(r['in']) for r in self.tables.nozzles}

    def component(self, start: str, omit: str | None = None) -> set[str]:
        adjacency = {n: [] for n in self.nodes}
        for pid, p in self.pipes.items():
            if pid != omit:
                adjacency[p.a].append(p.b)
                adjacency[p.b].append(p.a)
        # Pumps/valves also connect hydraulic nodes.
        for attr in ('pumps', 'valves'):
            for r in getattr(self.tables, attr, []):
                a, b = str(r.get('in')), str(r.get('out'))
                if a in adjacency and b in adjacency:
                    adjacency[a].append(b)
                    adjacency[b].append(a)
        seen, todo = {start}, [start]
        while todo:
            for nxt in adjacency[todo.pop()]:
                if nxt not in seen:
                    seen.add(nxt)
                    todo.append(nxt)
        return seen

    def new_node(self, xyz: tuple) -> str:
        lab = str(max([int(n) for n in self.nodes if n.isdigit()] or [0])+1)
        self.nodes[lab] = Node(tuple(xyz), {'label': lab, 'io_node': 'No',
                                          'editor_added': True})
        return lab

    def new_pipe(self, a: str, b: str, length: float, spec: dict) -> str:
        i = 1
        while f'EP{i}' in self.pipes:
            i += 1
        lab = f'EP{i}'
        row = dict(label=lab, length=length, status='Normal', group='',
                   dia_src='직접 편집 · 라이브러리', editor_added=True, eq_len=0)
        row.update(spec)
        self.pipes[lab] = Pipe(a, b, row)
        return lab

    def split(self, pid: str, t: float) -> str:
        p = self.pipes[pid]
        a, b = self.nodes[p.a].xyz, self.nodes[p.b].xyz
        old_b, old_length = p.b, float(p.row['length'])
        node = self.new_node(tuple(a[i]+t*(b[i]-a[i]) for i in range(3)))
        self.nodes[node].row['editor_between'] = [p.a, old_b, t]
        spec = deepcopy(p.row)
        new = self.new_pipe(node, old_b, old_length*(1-t), {})
        spec.update(label=new, length=old_length*(1-t), eq_len=0, editor_added=True)
        self.pipes[new].row = spec
        p.b = node
        p.row['length'] = old_length*t
        # Existing fittings stay on the original piece. Point equipment follows
        # its position; no double counting or loss of an alarm valve.
        for e in self.tables.equipment:
            if str(e.get('pipe')) != pid:
                continue
            rel = float(e.get('rel_pos', .5))
            if rel <= t:
                e['rel_pos'] = rel/t
            else:
                e.update(pipe=new, rel_pos=(rel-t)/(1-t))
        return node


def axis(a: tuple, b: tuple) -> int:
    """Return the unique non-zero axis; reject diagonals and zero length."""
    axes = [i for i in range(3) if abs(a[i]-b[i]) > EPS]
    if len(axes) != 1:
        raise EditError('두 점이 X·Y·Z 중 한 축과 평행해야 합니다. 중간 노드를 추가하세요.')
    return axes[0]


def _point_on(point: tuple, a: tuple, b: tuple) -> float | None:
    d = tuple(b[i]-a[i] for i in range(3))
    dd = sum(x*x for x in d)
    if dd <= EPS**2:
        return None
    t = sum((point[i]-a[i])*d[i] for i in range(3))/dd
    if -EPS <= t <= 1+EPS and dist(point, tuple(a[i]+t*d[i] for i in range(3))) < EPS:
        return t
    return None


def _segment_contact(a: tuple, b: tuple, c: tuple, d: tuple) -> bool:
    """Detect 3D segment contacts, including collinear overlap, not projections."""
    if any(_point_on(p, u, v) is not None for p, u, v in
           ((a,c,d), (b,c,d), (c,a,b), (d,a,b))):
        return True
    u = tuple(b[i]-a[i] for i in range(3))
    v = tuple(d[i]-c[i] for i in range(3))
    w = tuple(a[i]-c[i] for i in range(3))
    dot = lambda x, y: sum(x[i]*y[i] for i in range(3))
    aa, bb, cc, dd, ee = dot(u,u), dot(u,v), dot(v,v), dot(u,w), dot(v,w)
    den = aa*cc-bb*bb
    if abs(den) < EPS**2:
        return False
    s, t = (bb*ee-cc*dd)/den, (aa*ee-bb*dd)/den
    return EPS < s < 1-EPS and EPS < t < 1-EPS and dist(
        tuple(a[i]+s*u[i] for i in range(3)), tuple(c[i]+t*v[i] for i in range(3))) < EPS


def validate_changes(before: Network, after: Network, *, allow_split: bool = False) -> None:
    """Reject new contacts/overlaps, orphaned heads and disconnected fragments.

    `allow_split` is only for the riser cut the owner confirms with 「아니오」:
    the pipe goes, and the merged view says the network cannot be output yet.
    """
    changed = {pid for pid, p in after.pipes.items() if pid not in before.pipes
               or (p.a, p.b) != (before.pipes[pid].a, before.pipes[pid].b)
               or any(after.nodes[n].xyz != before.nodes[n].xyz
                      for n in (p.a, p.b) if n in before.nodes)}
    for pid in changed:
        p = after.pipes[pid]
        a, b = after.nodes[p.a].xyz, after.nodes[p.b].xyz
        if dist(a,b) <= EPS:
            raise EditError('길이가 0인 배관은 만들 수 없습니다.')
        for nid, n in after.nodes.items():
            if nid not in (p.a,p.b) and _point_on(n.xyz,a,b) is not None:
                old = before.pipes.get(pid)
                if old and nid in before.nodes and nid not in (old.a,old.b) and _point_on(
                    before.nodes[nid].xyz,before.nodes[old.a].xyz,before.nodes[old.b].xyz) is not None:
                    continue
                raise EditError(f'배관 {pid}가 노드 {nid}를 통과합니다. 그 노드에서 나누어 연결하세요.')
        for other, q in after.pipes.items():
            if other == pid:
                continue
            c, d = after.nodes[q.a].xyz, after.nodes[q.b].xyz
            shared = {p.a,p.b} & {q.a,q.b}
            if shared:
                s = next(iter(shared))
                pn, qn = (p.b if p.a == s else p.a), (q.b if q.a == s else q.a)
                if _point_on(after.nodes[pn].xyz,c,d) is not None or _point_on(after.nodes[qn].xyz,a,b) is not None:
                    raise EditError(f'배관 {pid}와 {other}가 겹칩니다.')
            elif _segment_contact(a,b,c,d):
                old, other_old = before.pipes.get(pid),before.pipes.get(other)
                if old and other_old and not ({old.a,old.b} & {other_old.a,other_old.b}) and _segment_contact(
                    before.nodes[old.a].xyz,before.nodes[old.b].xyz,
                    before.nodes[other_old.a].xyz,before.nodes[other_old.b].xyz):
                    continue
                raise EditError(f'배관 {pid}와 {other}가 같은 높이에서 교차합니다. 분할 후 연결하세요.')
    roots = sorted(n for n in after.protected if n in after.nodes)
    if roots and not allow_split:
        # Judge from the water source. After a riser cut the graph is already in
        # two pieces, and a root in the other piece would judge the wrong one.
        root = next((n for n in roots if str(after.nodes[n].row.get('io_node', '')).lower() == 'input'),
                    roots[0])
        reached = after.component(root)
        # Do not hide pre-existing disconnected pieces; only reject new ones.
        old_reached = before.component(root)
        required = (old_reached & after.nodes.keys()) | (after.nodes.keys()-before.nodes.keys())
        if required-reached:
            raise EditError('수정하면 급수원에서 연결이 끊기는 노드가 생깁니다.')
    for h in after.heads():
        if h not in after.nodes or len(after.incident(h)) != 1:
            if h not in before.heads() or len(before.incident(h)) == 1:
                raise EditError('헤드는 관말에 한 배관으로 연결해야 합니다. 헤드를 일반 노드로 바꾼 후 연장하세요.')


def apply_edit(network: Network, command: dict, catalog: dict) -> tuple[Network, dict]:
    """Preview/apply one deterministic command on an isolated graph copy."""
    n = deepcopy(network)
    op, target = command.get('op'), str(command.get('target', ''))
    result = {'kind': 'pipe' if target in n.pipes else 'node', 'label': target}
    check_from, allow_split = network, False

    def spec() -> dict:
        schedule = str(command.get('schedule', ''))
        dn = int(number(command.get('dn'), '호칭경', 1, 1000))
        found = next((r for r in catalog['pipes'] if r['type'] == schedule and r['dia'] == dn), None)
        if found is None:
            raise EditError('라이브러리에 없는 재질·호칭경 조합입니다.')
        return deepcopy(found)

    if op in ('extend', 'connect', 'head', 'node', 'move_node', 'delete_node', 'merge_node', 'paste') and target not in n.nodes:
        raise EditError('선택한 노드가 없어졌습니다. 다시 선택하세요.')
    if op in ('split', 'resize', 'pipe', 'fitting', 'remove_fitting', 'delete',
              'delete_join', 'delete_cut') and target not in n.pipes:
        raise EditError('선택한 배관이 없어졌습니다. 다시 선택하세요.')
    if op == 'paste':
        source = str(command.get('source', ''))
        source_pipe = str(command.get('source_pipe', ''))
        if source not in n.nodes or source_pipe not in n.incident(source) or source == target:
            raise EditError('복사한 노드와 다른 붙여넣기 시작 노드를 선택하세요.')
        source_node = deepcopy(n.nodes[source])
        original_pipe = deepcopy(n.pipes[source_pipe])
        nozzle = next((deepcopy(r) for r in n.tables.nozzles if str(r['in']) == source), None)
        losses = {attr:[deepcopy(r) for r in getattr(n.tables,attr) if str(r.get('pipe')) == source_pipe]
                  for attr in ('equipment','fittings')}
        if command.get('cut'):
            _delete_terminal(n, source)
        extended = dict(command, op='extend')
        n, result = apply_edit(n, extended, catalog)
        end, pid = result['label'], result['pipe']
        if nozzle:
            flow = nozzle.get('flow_lmin') or float(nozzle.get('flow_m3s', 0))*60000
            n, _ = apply_edit(n, dict(op='head', target=end, nozzle=nozzle['lib'], flow=flow), catalog)
        # Copy the selected connecting pipe's hydraulic losses, with new IDs.
        # Never drop a valve during a cut, or count an old fitting twice.
        n.pipes[pid].row.update(eq_len=float(original_pipe.row.get('eq_len',0)),dia=original_pipe.row.get('dia'))
        for attr,rows in losses.items():
            existing = getattr(n.tables,attr)
            used = {str(r.get('label')) for r in existing}
            for row in rows:
                i=1
                while f'EC{i}' in used: i+=1
                label=f'EC{i}';used.add(label)
                row.update(label=label,pipe=pid)
                if attr=='equipment':
                    if source == original_pipe.a:
                        row['rel_pos']=1-float(row.get('rel_pos',.5))
                    if row.get('op_id'):
                        row['op_id']='editor:'+label
                    if row.get('editor_node'):
                        row['editor_node']=end if str(row['editor_node'])==source else target
                existing.append(row)
        n,_ = apply_edit(n,dict(command,op='pipe',target=pid),catalog)
        if source_node.row.get('editor_note'):
            n.nodes[end].row['editor_note'] = source_node.row['editor_note']
    elif op == 'delete_node':
        parent = _delete_terminal(n, target)
        result = {'kind': 'node', 'label': parent}
    elif op == 'merge_node':
        reason = merge_reason(n, target)
        if reason:
            raise EditError(reason)
        first = next(k for k in n.incident(target) if n.pipes[k].b == target)
        second = next(k for k in n.incident(target) if k != first)
        p, q = n.pipes[first], n.pipes[second]
        l1, l2 = float(p.row['length']), float(q.row['length'])
        for e in n.tables.equipment:
            if str(e.get('pipe')) in (first, second):
                offset = l1 if str(e['pipe']) == second else 0
                length = l2 if str(e['pipe']) == second else l1
                e.update(pipe=first, rel_pos=(offset+float(e.get('rel_pos', .5))*length)/(l1+l2))
        for f in n.tables.fittings:
            if str(f.get('pipe')) == second:
                f['pipe'] = first
        p.b = q.b
        for nd in n.nodes.values():
            if str(nd.row.get('editor_parent')) == target:
                nd.row['editor_parent'] = p.a
        p.row.update(length=l1+l2, eq_len=float(p.row.get('eq_len', 0))+float(q.row.get('eq_len', 0)))
        n.pipes.pop(second)
        n.nodes.pop(target)
        result = {'kind': 'pipe', 'label': first}
    elif op == 'move_node':
        if target in n.protected:
            raise EditError('급수원·펌프·이음매의 좌표는 이 메뉴에서 이동할 수 없습니다.')
        xyz = tuple(number(command.get(k), f'{k.upper()} 좌표(m)', -1e7, 1e7) for k in ('x', 'y', 'z'))
        old = n.nodes[target].xyz
        for pid in n.incident(target):
            p = n.pipes[pid]
            other = n.nodes[p.a if p.b == target else p.b].xyz
            axis(xyz, other)
            length = float(p.row['length']) + dist(xyz, other)-dist(old, other)
            p.row['length'] = number(length, '이동 후 배관 길이(m)', .001, 10000)
        n.nodes[target].xyz = xyz
        delta = n.nodes[target].row.get('editor_z_delta', 0)
        n.nodes[target].row['editor_z_delta'] = delta + xyz[2]-old[2]
    elif op in ('extend', 'connect'):
        a = n.nodes[target].xyz
        if target in n.heads():
            raise EditError('헤드를 일반 노드로 바꾼 후 연장하세요.')
        if op == 'extend':
            direction = str(command.get('axis'))
            if direction not in ('+X','-X','+Y','-Y','+Z','-Z'):
                raise EditError('배관 방향을 선택하세요.')
            length = number(command.get('length'), '길이(m)', .001, 10000)
            xyz = list(a)
            xyz['XYZ'.index(direction[1])] += length * (1 if direction[0] == '+' else -1)
            if any(dist(xyz, nd.xyz) < EPS for nd in n.nodes.values()):
                raise EditError('끝점에 기존 노드가 있습니다. 기존 노드 연결을 사용하세요.')
            end = n.new_node(tuple(xyz))
            n.nodes[end].row['editor_parent'] = target
        else:
            end = str(command.get('end'))
            if end not in n.nodes or end == target:
                raise EditError('연결할 다른 노드를 선택하세요.')
            axis(a, n.nodes[end].xyz)
            length = dist(a, n.nodes[end].xyz)
        pid = n.new_pipe(target,end,length,spec())
        result = {'kind':'node','label':end, 'pipe':pid}
    elif op == 'split':
        p = n.pipes[target]
        t = number(command.get('distance'), '시작점에서 거리(m)', .001, float(p.row['length'])-.001)/float(p.row['length'])
        result = {'kind':'node','label':n.split(target,t)}
    elif op == 'resize':
        p = n.pipes[target]
        length = number(command.get('length'), '길이(m)', .001, 10000)
        a, b = n.nodes[p.a].xyz, n.nodes[p.b].xyz
        i = axis(a,b)
        moving = n.component(p.b, target)
        if p.a in moving or moving & n.protected:
            raise EditError('루프·급수원·펌프·이음매가 함께 움직여 이 배관을 연장할 수 없습니다.')
        delta = (length-float(p.row['length'])) * (1 if b[i] > a[i] else -1)
        if (b[i]+delta-a[i])*(b[i]-a[i]) <= 0:
            raise EditError('도식 좌표의 길이보다 많이 줄일 수 없습니다. 분할을 사용하세요.')
        for nid in moving:
            xyz = list(n.nodes[nid].xyz)
            xyz[i] += delta
            n.nodes[nid].xyz = tuple(xyz)
        p.row['length'] = length
    elif op == 'pipe':
        new = spec()
        old_dn = n.pipes[target].row.get('dia')
        n.pipes[target].row.update(new)
        # Never retain a valve/fitting loss for the old bore.
        for e in n.tables.equipment:
            if str(e.get('pipe')) == target and e.get('editor_library'):
                e['eq_len'] = fitting_value(catalog, e['editor_library'], new['dia']) * e.get('count',1)
        if old_dn != new['dia']:
            native = [f for f in n.tables.fittings if str(f.get('pipe')) == target]
            if native:
                total = 0.0
                for f in native:
                    kind = str(f.get('type'))
                    aliases = catalog.get('fitting_aliases',{})
                    if kind in aliases and aliases[kind] is None:
                        continue
                    total += fitting_value(catalog,aliases.get(kind,kind),new['dia']) * int(f.get('count',1))
                n.pipes[target].row['eq_len'] = total
            for e in n.tables.equipment:
                if str(e.get('pipe')) == target and not e.get('editor_library'):
                    item = e.get('lib') or ('VALVE_ALARM' if e.get('desc')=='A/V' else None)
                    if item:
                        e['eq_len'] = fitting_value(catalog,item,new['dia'])
    elif op == 'fitting':
        p = n.pipes[target]
        item = str(command.get('fitting'))
        count = int(number(command.get('count',1), '개수', 1, 100))
        eq = fitting_value(catalog,item,int(p.row['dia'])) * count
        f = next(x for x in catalog['fittings'] if x['id']==item)
        eid = 'EF' + str(max([int(str(e.get('label',''))[2:]) for e in n.tables.equipment
                              if str(e.get('label','')).startswith('EF') and str(e['label'])[2:].isdigit()] or [0])+1)
        n.tables.equipment.append(dict(label=eid, pipe=target, desc=f['name'], eq_len=eq,
            rel_pos=number(command.get('position',.5),'설치 위치',0,1),
            editor_library=item, count=count, op_id='editor:'+eid))
        node = str(command.get('node', ''))
        if node:
            if node not in (p.a, p.b):
                raise EditError('부속을 넣을 노드에 연결된 배관을 선택하세요.')
            n.tables.equipment[-1].update(editor_node=node, rel_pos=0 if node == p.a else 1)
    elif op == 'remove_fitting':
        eid = str(command.get('equipment'))
        hit = [e for e in n.tables.equipment if str(e.get('label')) == eid
               and str(e.get('pipe')) == target and e.get('editor_library')]
        if not hit:
            raise EditError('이 편집기에서 추가한 부속을 선택하세요.')
        n.tables.equipment = [e for e in n.tables.equipment if e not in hit]
    elif op == 'head':
        if target in n.protected or len(n.incident(target)) != 1:
            raise EditError('헤드는 급수원·이음매가 아닌 관말 노드에만 설치할 수 있습니다.')
        lib = str(command.get('nozzle'))
        nozzle = next((r for r in catalog['nozzles'] if r['id']==lib),None)
        if nozzle is None:
            raise EditError('헤드 라이브러리를 선택하세요.')
        flow = number(command.get('flow'), '방수량(L/min)', .001, 100000)
        n.tables.nozzles = [r for r in n.tables.nozzles if str(r['in']) != target]
        used = {str(r['label']) for r in n.tables.nozzles}
        i = 1
        while f'EN{i}' in used:
            i += 1
        n.tables.nozzles.append(dict(label=f'EN{i}', **{'in':target, 'out':f'@/EN{i}'},
            status='1',lib=lib,flow_lmin=flow,flow_m3s=flow/60000))
        n.nodes[target].row['editor_head'] = dict(nozzle, flow_lmin=flow)
    elif op == 'node':
        n.tables.nozzles = [r for r in n.tables.nozzles if str(r['in']) != target]
        n.tables.equipment = [r for r in n.tables.equipment if str(r.get('editor_node')) != target]
        n.nodes[target].row['editor_head'] = None
    elif op == 'delete':
        p = n.pipes[target]
        if len(n.incident(p.b)) != 1 or p.b in n.protected:
            raise EditError('관말 배관만 삭제할 수 있습니다. 안쪽 배관을 지우면 연결이 끊깁니다.')
        n.pipes.pop(target)
        n.nodes.pop(p.b)
        n.tables.nozzles = [r for r in n.tables.nozzles if str(r['in']) != p.b]
        for attr in ('equipment','fittings'):
            setattr(n.tables,attr,[r for r in getattr(n.tables,attr) if str(r.get('pipe')) != target])
        result = {'kind':'node','label':p.a}
    elif op in ('delete_join', 'delete_cut'):
        # [오너 2026-09-21] 통합망 계통도의 세로관(층고 배관) 삭제.
        #   예(delete_join) — 위·아래 노드를 붙인다. 급수원(기계실) 쪽 표고는 그대로
        #     두고, 반대쪽(평면도 쪽)을 지운 배관의 높이만큼 옮긴다.
        #   아니오(delete_cut) — 배관만 지운다. 망이 끊긴 채 남고 산출은 막힌다.
        p = n.pipes[target]
        if abs(n.nodes[p.a].xyz[2]-n.nodes[p.b].xyz[2]) <= EPS:
            raise EditError('높이차가 없는 배관은 이 방법으로 지울 수 없습니다.')
        n.pipes.pop(target)
        for attr in ('equipment','fittings'):
            setattr(n.tables,attr,[r for r in getattr(n.tables,attr) if str(r.get('pipe')) != target])
        if op == 'delete_cut':
            allow_split = True
            result = {'kind':'node','label':p.a}
        else:
            check_from, survivor = _join_riser_gap(n, p, {str(k) for k in command.get('keep') or ()})
            result = {'kind':'node','label':survivor}
    else:
        raise EditError('지원하지 않는 편집 동작입니다.')
    validate_changes(check_from,n,allow_split=allow_split)
    n.to_tables()
    result['counts'] = dict(nodes=len(n.nodes),pipes=len(n.pipes),heads=len(n.tables.nozzles))
    return n,result


def terminal_pipe(network: Network, node: str) -> str | None:
    """Return the removable terminal pipe; never remove supply or seam nodes."""
    incident = network.incident(node)
    if node not in network.protected and len(incident) == 1:
        pid = incident[0]
        if network.pipes[pid].b == node:
            return pid
    return None


def _delete_terminal(network: Network, node: str) -> str:
    pid = terminal_pipe(network, node)
    if pid is None:
        raise EditError('삭제·자르기는 급수원과 이음매가 아닌 관말 노드에서 가능합니다.')
    parent = network.pipes.pop(pid).a
    network.nodes.pop(node)
    network.tables.nozzles = [r for r in network.tables.nozzles if str(r['in']) != node]
    for attr in ('equipment', 'fittings'):
        setattr(network.tables, attr, [r for r in getattr(network.tables, attr) if str(r.get('pipe')) != pid])
    return parent


def _join_riser_gap(network: Network, pipe: Pipe, keep: set) -> tuple[Network, str]:
    """Close the gap a removed riser pipe leaves; the water-source side keeps its Z.

    The piece without the water source moves by the removed pipe's height, then
    its end node and the other end node become one. Returns the moved-but-not-yet
    joined graph (the validation baseline) and the surviving label.
    """
    side_a = network.component(pipe.a)
    if pipe.b in side_a:
        raise EditError('이 배관을 지워도 위·아래가 다른 길로 이어져 있어 붙일 수 없습니다. 「아니오」로 지우세요.')
    side_b = network.component(pipe.b)
    supply = {k for k, nd in network.nodes.items() if str(nd.row.get('io_node', '')).lower() == 'input'}
    if (side_a & supply) and not (side_b & supply):
        fixed_end, moving_end, moving = pipe.a, pipe.b, side_b
    elif (side_b & supply) and not (side_a & supply):
        fixed_end, moving_end, moving = pipe.b, pipe.a, side_a
    else:
        raise EditError('급수원(기계실) 쪽을 정할 수 없어 붙일 수 없습니다. 「아니오」로 지우세요.')
    keep = set(keep) | network.protected | network.heads()
    if fixed_end in keep and moving_end in keep:
        raise EditError('위·아래 노드가 모두 급수원·기준점·펌프 노드라 붙일 수 없습니다. 「아니오」로 지우세요.')
    z_fixed = network.nodes[fixed_end].xyz[2]
    dz = z_fixed-network.nodes[moving_end].xyz[2]
    for nid in moving:
        x, y, z = network.nodes[nid].xyz
        network.nodes[nid].xyz = (x, y, round(z+dz, 9))
    x, y, _ = network.nodes[moving_end].xyz
    network.nodes[moving_end].xyz = (x, y, z_fixed)
    _shift_contacts(network, moving, fixed_end, moving_end)
    moved = deepcopy(network)
    survivor, victim = (moving_end, fixed_end) if moving_end in keep else (fixed_end, moving_end)
    for q in network.pipes.values():
        if q.a == victim:
            q.a = survivor
        if q.b == victim:
            q.b = survivor
    for attr in ('pumps', 'valves'):
        for row in getattr(network.tables, attr, []):
            for key in ('in', 'out'):
                if str(row.get(key)) == victim:
                    row[key] = survivor
    for row in network.tables.equipment:
        if str(row.get('editor_node')) == victim:
            row['editor_node'] = survivor
    for nd in network.nodes.values():
        if str(nd.row.get('editor_parent', '')) == victim:
            nd.row['editor_parent'] = survivor
        between = nd.row.get('editor_between')
        if between and victim in [str(k) for k in between[:2]]:
            nd.row['editor_between'] = [survivor if str(k) == victim else k for k in between[:2]] + list(between[2:])
    network.nodes.pop(victim)
    return moved, survivor


def _shift_contacts(network: Network, moving: set, fixed_end: str, moving_end: str) -> None:
    """A rigid shift keeps contacts inside the moved piece; check it against the rest."""
    def apart(a, b, c, d):
        return any(max(a[i], b[i]) < min(c[i], d[i])-EPS or max(c[i], d[i]) < min(a[i], b[i])-EPS
                   for i in range(3))
    inside = [(k, q) for k, q in network.pipes.items() if q.a in moving]
    outside = [(k, q) for k, q in network.pipes.items() if q.a not in moving]
    for k, q in inside:
        a, b = network.nodes[q.a].xyz, network.nodes[q.b].xyz
        at_joint = moving_end in (q.a, q.b)
        for m, r in outside:
            if at_joint and fixed_end in (r.a, r.b):
                continue  # They meet at the joint by design; overlaps are checked after joining.
            c, d = network.nodes[r.a].xyz, network.nodes[r.b].xyz
            if not apart(a, b, c, d) and _segment_contact(a, b, c, d):
                raise EditError(f'붙이면 배관 {k}와 {m}가 닿습니다. 「아니오」로 지운 뒤 따로 이으세요.')


def merge_reason(network: Network, node: str) -> str:
    """Explain why deleting a degree-two node would lose hydraulic meaning."""
    incident = network.incident(node)
    if node in network.protected or node in network.heads() or len(incident) != 2:
        return '급수원·헤드·분기점은 합칠 수 없습니다. 두 배관 사이의 일반 노드를 선택하세요.'
    incoming = [network.pipes[k] for k in incident if network.pipes[k].b == node]
    outgoing = [network.pipes[k] for k in incident if network.pipes[k].a == node]
    if len(incoming) != 1 or len(outgoing) != 1:
        return '두 배관의 유하 방향이 이어져야 합칠 수 있습니다.'
    p, q = incoming[0], outgoing[0]
    t = _point_on(network.nodes[node].xyz, network.nodes[p.a].xyz, network.nodes[q.b].xyz)
    if t is None or not EPS < t < 1-EPS:
        return '방향이 꺾이는 노드는 합칠 수 없습니다.'
    if any(p.row.get(k) != q.row.get(k) for k in
           ('type', 'dia', 'inner_mm', 'c', 'roughness_mm', 'status', 'group')):
        return '재질·관경·조도·상태가 같은 배관만 합칠 수 있습니다.'
    for row in network.tables.equipment:
        pid = str(row.get('pipe'))
        at_joint = pid in incident and abs(float(row.get('rel_pos', .5)) -
                                          (1 if network.pipes[pid].b == node else 0)) < EPS
        if str(row.get('editor_node')) == node or at_joint:
            return '해당 노드의 부속·밸브를 먼저 정리한 뒤 합치세요.'
    return ''


def fitting_value(catalog: dict, item: str, dn: int) -> float:
    """Missing loss is an error, never an implicit zero."""
    f = next((f for f in catalog['fittings'] if f['id']==item),None)
    value = (f or {}).get('lengths',{}).get(str(dn))
    if value is None:
        raise EditError(f'선택한 부속의 {dn}A 등가길이가 라이브러리에 없습니다.')
    return number(value,'등가길이(m)',0,100000)
