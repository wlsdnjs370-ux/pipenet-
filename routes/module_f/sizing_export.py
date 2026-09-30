"""Separate, explicitly unconfirmed SDFs matching the sizing loss model."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict
import json
from pathlib import Path
import xml.etree.ElementTree as ET

from routes.module_f.hydraulic_sizing import fingerprint


def emit_proposal(got: dict, report: dict, out_dir: Path) -> dict:
    """Write a supply-pressure review case; never invent a manufacturer pump."""
    from routes.module_f.common import _boot
    _boot()
    from core.remote30_full_network import ProjectContext, emit_full_sdf
    from routes.module_f.merge import ANCHOR_LABEL, bake_combined_iso
    from routes.module_f.export_compat import prepare_sdf_export
    from remote30_prototype import write_sdf_tree
    from src.pipenet_converter.graph.export_compaction import compact_export, protected_nodes
    from src.pipenet_converter.render.export_style import EXPORT_SYMBOL_SPREAD
    from src.pipenet_converter.render.framing import reference_display_scale

    original = got["combined"]
    if report["fingerprint"] != fingerprint(original):
        raise ValueError("결합망이 변경되었습니다. 역산을 다시 실행하세요.")
    if not report["feasible"]:
        raise ValueError("조건 미달/미수렴 검토안은 SDF로 출력하지 않습니다. 진단 JSON을 확인하세요.")
    edited = deepcopy(got)
    tables = edited["combined"]
    keep = protected_nodes(tables)  # Preserve real fittings before total-loss conversion.
    keep.update(str(k) for k in (ANCHOR_LABEL, got.get('pump_junction')) if k)
    valid_nodes = {n["label"] for n in report["model"]["nodes"]}
    tables.nodes = [n for n in tables.nodes if str(n["label"]) in valid_nodes]
    z = {str(n["label"]):float(n["elevation"]) for n in tables.nodes}
    source = report["source"]
    for n in tables.nodes:
        n["io_node"] = "Input" if str(n["label"]) == source else "No"
        n.pop("pressure_pa", None)
        if str(n["label"]) == source:
            n["pressure_pa"] = 101325 + report["solution"]["source_pressure_bar"]*100000
    tables.pumps = []  # This is a pressure boundary case, not a fake pump curve.
    tables.valves = []
    # Export exactly the solver's declared loss model, independent of PIPENET's
    # own fitting tables. Native fitting symbols remain in the original output.
    tables.fittings = []
    tables.equipment = []
    proposed = {p["label"]:p for p in report["pipes"]}
    for i, p in enumerate(tables.pipes):
        chosen = proposed[str(p["label"])]
        p.update(dia=chosen["after_dn"],inner_mm=chosen["inner_mm"],eq_len=0,
                 elev=z[str(p["out"])]-z[str(p["in"])],review_only=True,
                 bore_provenance=dict(method="hydraulic_review",review_only=True,block_export=False))
        if chosen["equivalent_m"]:
            tables.equipment.append(dict(pipe=str(p["label"]),**{"in":p["in"],"out":p["out"]},
                label=f"HR{i+1}",desc="Hydraulic review total equivalent length",
                eq_len=chosen["equivalent_m"],rel_pos=.5))
    active = {h["label"]:h for h in report["nozzles"]}
    tables.nozzles = [h for h in tables.nozzles if str(h["label"]) in active]
    for h in tables.nozzles:
        h.update(status="1",flow_lmin=active[str(h["label"])]["min_flow_lpm"],
                 flow_m3s=active[str(h["label"])]["min_flow_lpm"]/60000)
    for part, labels in (edited.get("parts") or {}).items():
        edited["parts"][part] = [x for x in labels if str(x) in valid_nodes]
    tables.meta = list(tables.meta) + [("hydraulic_review", "NOT VALIDATED; fixed discharge pressure; total equivalent losses")]
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True,exist_ok=True)
    iso_nodes, _ = bake_combined_iso(edited)
    references = [str(k) for k in (edited.get('parts') or {}).get('plan', [])]
    if not references:
        references = [str(n['label']) for n in tables.nodes]
    at = {str(n['label']):n for n in tables.nodes}
    display_scale = reference_display_scale((float(at[k]['x']),float(at[k]['y'])) for k in references)
    tables, audit = compact_export(tables, keep_nodes=keep, display_views=[iso_nodes])
    valid_nodes = {str(n['label']) for n in tables.nodes}
    references = [k for k in references if k in valid_nodes]
    # Each output pipe may now represent several equal-bore original pipes.
    output_targets = {str(p['label']):p for p in tables.pipes}
    paths = {}
    for suffix, positions in (("",tables.nodes),("_iso",iso_nodes)):
        candidate = deepcopy(tables)
        by_label = {str(n["label"]):n for n in positions}
        # Display transformation must not replace pressure boundary metadata.
        for n in candidate.nodes:
            pos = by_label[str(n["label"])]
            n.update(x=pos["x"],y=pos["y"])
        sdf = out_dir / f"module_f_hydraulic_review{suffix}.sdf"
        emit_full_sdf(candidate,sdf,ctx=ProjectContext.titled("관경·펌프 역산 검토안 / NOT VALIDATED"),
                      display_scale=display_scale*EXPORT_SYMBOL_SPREAD)
        tree = ET.parse(sdf)
        # The shared legacy writer uses six significant digits. Long merged
        # spans need more precision to preserve the solver's declared lengths.
        for pipe in tree.getroot().iter('Pipe'):
            row = output_targets.get(pipe.get('label'))
            if row is None:
                raise ValueError('역산 SDF에 계산 입력과 다른 배관이 생성되었습니다.')
            pipe.set('length',format(float(row['length']),'.12g'))
            pipe.set('rise',format(float(row['elev']),'.12g'))
        opts = tree.getroot().find(".//Network-spray/Attributes/Design-options")
        if opts is None:
            raise ValueError("SDF 계산 방식 설정이 없습니다.")
        opts.set("specification-type","user-defined")
        opts.set("pressure-equation","hazen-williams")
        opts.set("vp-model","ignore")
        # Do not inherit uninitialised fluid properties from the old template.
        # This engine explicitly uses water at 20 C (kinematic viscosity m2/s).
        fluid = tree.getroot().find(".//Network-spray/Attributes/Fluid-fixed-user")
        if fluid is None:
            raise ValueError("SDF 물성 설정 형식을 확인할 수 없습니다.")
        fluid.attrib.update(density="998.2",viscosity="1.004e-6",**{"vapour-pressure":"2339"})
        write_sdf_tree(tree,sdf)
        prepare_sdf_export(sdf,nozzle_reference_labels=references,display_spread=EXPORT_SYMBOL_SPREAD)
        # Round-trip checks fail closed before advertising downloads.
        root = ET.parse(sdf).getroot()
        pipes = list(root.iter("Pipe"))
        if {p.get("label") for p in pipes} != set(output_targets):
            raise ValueError("역산 SDF 배관 연결 개수 검증 실패")
        boundaries = [n for n in root.iter("Node") if n.get("io-node")=="Input"]
        if len(boundaries)!=1 or boundaries[0].get("label")!=source or boundaries[0].find("Calculation-spec") is None:
            raise ValueError("역산 SDF 공급 경계 검증 실패")
        slf = sdf.with_suffix(".slf")
        if not slf.is_file():
            raise ValueError("역산 SDF의 실제 내경 라이브러리가 없습니다.")
        library = ET.parse(slf).getroot()
        bores = {sch.findtext("Item-name"): {float(s.get("nominal")):float(s.get("internal"))
                    for s in sch.findall("./Metric-definition/Size-definition")
                    if s.get("internal") not in (None,"Unset")}
                 for sch in library.findall("./Schedule-section/Schedule")}
        original_pipes = {str(p['label']):p for p in tables.pipes}
        expected_losses = {label:sum(float(e['eq_len']) for e in tables.equipment
                                    if str(e['pipe'])==label) for label in output_targets}
        for p in pipes:
            lab = p.get('label'); row = original_pipes[lab]; target = output_targets[lab]
            if (p.get('input')!=str(row['in']) or p.get('output')!=str(row['out'])
                    or abs(float(p.get('length'))-float(row['length']))>1e-5
                    or abs(float(p.get('bore'))*1000-target['dia'])>1e-5
                    or abs(float(p.get('roughness-or-c'))-float(row['c']))>1e-5
                    or abs(float(p.get('rise'))-float(row['elev']))>1e-8
                    or abs(sum(float(e.get('equivalent-length')) for e in p.iter('Equipment'))
                           - expected_losses[lab])>1e-5
                    or abs(bores.get(row['type'],{}).get(target['dia'],-1)-target['inner_mm'])>.001):
                raise ValueError(f"배관 {lab}: SDF/SLF와 계산 입력 불일치")
        paths["sdf"+suffix],paths["slf"+suffix] = str(sdf),str(slf)
    report_path = out_dir / "hydraulic_review.json"
    exported_report = dict(report, export_compaction=asdict(audit))
    report_path.write_text(json.dumps(exported_report,ensure_ascii=False,indent=2,allow_nan=False),encoding="utf-8")
    print(f'[역산 SDF] {audit.message}')
    paths["report"] = str(report_path)
    return paths
