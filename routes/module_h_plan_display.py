"""H plan labels and CAD colours from the current (possibly cropped) world."""
from __future__ import annotations

from dataclasses import asdict
from typing import Any

from ezdxf.colors import aci2rgb
from ezdxf.tools.text import plain_mtext

from src.pipenet_converter.render.dxf_text import DisplayText


def source_color(color: int | str) -> str:
    """Native ACI palette on a dark background; colour 7 is white."""
    if isinstance(color, str):
        return color
    if not 0 < color < 256:
        color = 7
    r, g, b = aci2rgb(color)
    return f'#{r:02x}{g:02x}{b:02x}'


def add_plan_display(payload: dict, world: Any) -> dict:
    """Add upright native labels without rereading DXF or changing graph inputs.

    The pick reader already stores native text in millimetres. Using that same
    snapshot also ensures cropping cannot bring excluded annotations back.
    """
    bundles = {b['id']: b for b in payload['bundles']}
    for bundle in bundles.values():
        bundle['css'] = source_color(bundle['color'])
    texts = []
    for layer, color, x, y, height, raw in getattr(world, 'texts', ()):
        value = plain_mtext(str(raw)).strip()
        if not value or height <= 0:
            continue
        identity = f'{layer}\x1f{color}'
        if identity not in bundles:
            bundle = dict(id=identity, i=len(bundles), layer=layer, color=color,
                          name=f'색{color}', css=source_color(color), cat='TEXT',
                          segs=[], circles=[], arcs=[], n_seg=0, n_circle=0, n_arc=0,
                          n_all=0, n_circle_all=0, n_arc_all=0, len_m=0, len_mid=0,
                          display_partial=False)
            bundles[identity] = bundle
            payload['bundles'].append(bundle)
        label = asdict(DisplayText(x, y, height, value, 0, layer, color))
        label.update(bundle_id=identity, css=source_color(color))
        texts.append(label)
        bounds = payload['bounds']
        for axis, coordinate in (('x', x), ('y', y)):
            bounds[f'min{axis}'] = min(bounds[f'min{axis}'], coordinate)
            bounds[f'max{axis}'] = max(bounds[f'max{axis}'], coordinate)
    payload['texts'] = texts
    payload['h_source_display'] = True
    payload['counts']['texts'] = payload['shown']['texts'] = len(texts)
    payload['dropped']['texts'] = 0
    return payload
