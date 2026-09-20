"""Drawing-slot reference layers for the combined canvas (display only)."""
from __future__ import annotations

import math

from routes.module_f.slots import SLOT_KINDS, SLOT_LABELS, _slot_active


def _slot(sess: dict, kind: str) -> dict:
    return sess if _slot_active(sess)==kind else (sess.get('slots') or {}).get(kind,{})


def _xy(node: dict) -> tuple[float, float]:
    return float(node.get('x',0)),float(node.get('y',0))


def _upright_reference(a: tuple, b: tuple, dest_a: tuple, dest_b: tuple) -> list[float] | None:
    """Align the AV and pump height without tilting the drawing's verticals.

    A is the AV. Preserve aspect ratio and horizontal offsets in the source;
    two horizontally offset source anchors cannot both match a vertical riser
    without distorting the drawing. Opposite source/supply directions retain
    the previous half-turn, but never introduce an oblique rotation.
    """
    x,y = b[0]-a[0],b[1]-a[1]
    u,v = dest_b[0]-dest_a[0],dest_b[1]-dest_a[1]
    d=x*x+y*y
    if d < 1e-9 or u*u+v*v < 1e-9:
        return None
    scale = v/y if abs(y) > 1e-9 and abs(v) > 1e-9 else math.sqrt((u*u+v*v)/d)
    return [scale,0,0,scale,dest_a[0]-scale*a[0],dest_a[1]-scale*a[1]]


def _compose(a: list, b: list) -> list:
    """Canvas affine a after b."""
    p,q,r,s,x,y=a; A,B,C,D,X,Y=b
    return [p*A+r*B,q*A+s*B,p*C+r*D,q*C+s*D,p*X+r*Y+x,q*X+s*Y+y]


def reference_layers(sess: dict, got: dict, *, iso: bool, geometry: bool=False) -> list[dict]:
    """Return per-slot availability/transforms; send cached drawing geometry on demand.

    Plan and machine-room layers share their network's transform. A system
    drawing stays upright at the AV, scaled to the pump's displayed height.
    Its horizontal offsets and intermediate vertices may differ from the
    physical pipe layout; this reference must not shear to fit them.
    No reference coordinate is ever written into the hydraulic graph.
    """
    # Edits move the network relative to its original drawing, not the drawing.
    editor=sess.get('merge_editor')
    base=editor['base_object'] if editor and editor.get('object') is got else got
    table=base['combined']
    at={str(n['label']):_xy(n) for n in table.nodes}
    ax,ay=at.get('10',(0,0))
    c,s=math.cos(math.pi/6),.5
    rot=[c,s,-c,s,0,0]
    a_iso=(c*(ax-ay),s*(ax+ay))
    pj=at.get(str(base.get('pump_junction')))
    physical = base.get('system_layout') == 'physical_xy'
    shown = at
    if physical and iso:
        from routes.module_f.merge import bake_combined_iso
        shown = {str(n['label']): _xy(n) for n in bake_combined_iso(base)[0]}
    result=[]
    for kind in SLOT_KINDS:
        slot=_slot(sess,kind)
        world=slot.get('world') or {}
        groups=world.get('bundles') or []
        row=dict(kind=kind,label=SLOT_LABELS[kind],available=False,
                 source=str(slot.get('key') or ''),token=f"{slot.get('key')}:{id(world)}",
                 reason='참조 도면이 열리지 않았습니다.')
        xf=None
        if kind=='plan':
            origin=((slot.get('design') or {}).get('got') or {}).get('origin_mm')
            if origin:
                xf=[1,0,0,1,1000-float(origin[0]),1000-float(origin[1])]
                if iso: xf=_compose(rot,xf)
            row['note']='원도면의 평면 위치를 기준점 높이에 겹칩니다.'
        elif kind=='system':
            source=slot.get('riser') or {}
            raw={str(n['label']):_xy(n) for n in source.get('nodes',[])}
            end=str(source.get('av_node_label') or '10')
            start=str(source.get('input_node_label') or '1')
            if start in raw and end in raw and start in at and '10' in at:
                target = shown if physical else at
                xf=_upright_reference(raw[end],raw[start],target['10'],target[start])
                if physical and xf is None and raw[end] != raw[start]:
                    # A true vertical pipe has coincident XY in plan view.
                    # Keep its reference upright at the AV, without changing XY.
                    xf=[1,0,0,1,target['10'][0]-raw[end][0],target['10'][1]-raw[end][1]]
                if iso and xf and not physical:
                    xf=_compose([1,0,0,1,a_iso[0]-ax,a_iso[1]-ay],xf)
            row['note']='알람밸브 기준으로 수직 정렬하고 펌프 높이에 맞춥니다. 원도면의 가로 간격·중간 경로는 유지됩니다.'
        else:
            source=slot.get('machineroom') or {}
            raw={str(n['label']):_xy(n) for n in source.get('nodes',[])}
            conn=source.get('conn_xy')
            if conn and pj:
                # Use the very same source bbox and scale as the merge layout.
                labels=set(map(str,(base.get('parts') or {}).get(kind,[])))
                points=[xy for label,xy in raw.items() if label in labels and label in at]
                for edge in source.get('plan_edges') or []:
                    points.extend((edge[:2],edge[2:4]))
                if points:
                    xs,ys=zip(*points)
                    diag=max(math.hypot(max(xs)-min(xs),max(ys)-min(ys)),1)
                    head=base.get('head_tables')
                    hy=[float(n.get('y',0)) for n in getattr(head,'nodes',[])]
                    span=max(hy)-min(hy) if hy else 0
                    scale=max(2000,span*.7)/diag
                    xf=[scale,0,0,scale,pj[0]-scale*float(conn[0]),pj[1]-scale*float(conn[1])]
                    if iso:
                        target_pj = shown[str(base['pump_junction'])] if physical else (a_iso[0]+pj[0]-ax,a_iso[1]+pj[1]-ay)
                        sx=target_pj[0]-c*(pj[0]-pj[1])
                        sy=target_pj[1]-s*(pj[0]+pj[1])
                        xf=_compose([c,s,-c,s,sx,sy],xf)
            row['note']='기계실 접속점을 기준으로 결합망과 같은 배율로 겹칩니다.'
        if not groups:
            row['reason']='참조 도면이 열리지 않았습니다.'
        elif xf is None:
            row['reason']='도면과 결합망을 맞출 추출 기준점이 없습니다.'
        else:
            row.update(available=True,matrix=xf,reason='')
            dropped=sum((world.get('dropped') or {}).values())
            if dropped:
                row['note']+=f' 대용량 미리보기 상한으로 도형 {dropped:,}개는 생략되어 있습니다.'
            if geometry:
                row['groups']=[{k:g.get(k,[]) for k in ('segs','circles','arcs')} for g in groups]
                row['bounds']=world.get('bounds')
                row['truncated']=bool(dropped)
        result.append(row)
    return result
