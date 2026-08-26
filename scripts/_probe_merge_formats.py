# -*- coding: utf-8 -*-
"""결합망 산출 3종이 어긋나는 이유 — SDF 61절점 vs KFP/HAS 11절점.

특허 S760 은 「S750 의 결과 파일 자체를 원본으로 삼으므로 모든 형식이 항상 같은
배관망을 가리킨다」고 한다. 같은 파일에서 파생했는데 수가 다르면, 변환기가
그 파일의 일부만 읽고 있다는 뜻이다. 어느 부분을 버리는지 본다.

    python scripts/_probe_merge_formats.py
"""
from __future__ import annotations

import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SYSTEM_DXF = ROOT / "data" / "sample_problem" / "대명동201동 계통도.dxf"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    sys.path.insert(0, str(ROOT / "scripts"))
    from _verify_module_f_merge import _G          # 같은 최소 표를 쓴다

    from routes.module_f.emit import emit_merged
    from routes.module_f.merge import merge_network
    from routes.module_f.subdrawing import (entities_to_world, extract_system,
                                            parse_subdrawing)

    ents, _ = parse_subdrawing(SYSTEM_DXF)
    w = entities_to_world(ents)
    pts = [p for s in w.segs for p in (s[2], s[3])]
    lo = min(pts, key=lambda p: (p[0], p[1]))
    hi = max(pts, key=lambda p: (p[0], p[1]))
    riser = extract_system(ents, lo, hi, snap_tolerance_mm=5000)
    got = merge_network(_G(), riser=riser, mode="lsp_gravity")
    c = got["combined"]
    print(f"결합망(메모리): 절점 {len(c.nodes)} · 배관 {len(c.pipes)}")
    print(f"  노드 라벨 표본: {[n['label'] for n in c.nodes[:14]]} …")

    tmp = Path(tempfile.mkdtemp(prefix="mf_fmt_"))
    files = emit_merged(c, tmp, title="포맷 대조")
    for wmsg in files.get("warnings") or ():
        print(f"  ! {wmsg}")

    root = ET.parse(files["sdf"]).getroot()
    sdf_nodes = [n.attrib.get("label") for n in root.findall(".//Node")]
    sdf_pipes = [p.attrib.get("label") for p in root.findall(".//Pipe")]
    print(f"\nSDF : 절점 {len(sdf_nodes)} · 배관 {len(sdf_pipes)}")
    print(f"  Pipe-set 수: {len(root.findall('.//Pipe-set'))}")
    for i, ps in enumerate(root.findall(".//Pipe-set")):
        nm = ps.find("Pipe-type/Name")
        print(f"    [{i}] type={nm.text if nm is not None else None!r}"
              f" · Pipe {len(ps.findall('Pipe'))}")

    from kfp_sdf_converter import parse_kfp, parse_sdf
    net_sdf = parse_sdf(files["sdf"])
    print(f"\nparse_sdf (변환기의 눈): 절점 {len(net_sdf.nodes)} · "
          f"배관 {len(net_sdf.pipes)}")
    got_labels = {str(n) for n in net_sdf.nodes}
    missing = [l for l in sdf_nodes if l not in got_labels]
    print(f"  변환기가 못 본 절점 {len(missing)}개: {missing[:20]}")

    net_kfp = parse_kfp(files["kfp"])
    print(f"KFP : 절점 {len(net_kfp.nodes)} · 배관 {len(net_kfp.pipes)}")

    def _len(net):
        t = 0.0
        for p in net.pipes.values():
            t += float(getattr(p, "length_m", 0.0) or 0.0)
        return t

    def _nz(net):
        return sum(1 for n in net.nodes.values()
                   if str(getattr(n, "kind", "")).lower() in ("nozzle", "head"))

    print("\n길이·노즐 불변량 (특허 S443: 통합해도 길이는 합산 보존)")
    print(f"  SDF : 연장 {_len(net_sdf):10.3f} m · 노즐 {_nz(net_sdf)}")
    print(f"  KFP : 연장 {_len(net_kfp):10.3f} m · 노즐 {_nz(net_kfp)}")
    try:
        from has_converter import parse_has
        net_has = parse_has(files["has"])
        print(f"  HAS : 연장 {_len(net_has):10.3f} m · 노즐 {_nz(net_has)}"
              f" · 절점 {len(net_has.nodes)}")
    except Exception as exc:  # noqa: BLE001
        print(f"  HAS : 못 읽음 — {exc}")

    print(f"\n산출물: {tmp}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
