"""Recalculate a saved actual-model snapshot with the current solver, offline."""
from __future__ import annotations

import argparse
from dataclasses import asdict
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from src.pipenet_converter.hydraulics.models import Network, Node, Nozzle, Pipe, Settings, Size
from src.pipenet_converter.hydraulics.sizing import propose


def main() -> None:
    """Check reproducibility and water-volume fields without a live session."""
    parser = argparse.ArgumentParser()
    parser.add_argument("snapshot", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    saved = json.loads(args.snapshot.read_text(encoding="utf-8"))["report"]
    model = saved["model"]
    net = Network(tuple(Node(**n) for n in model["nodes"]),
        tuple(Pipe(**dict(p, sizes=tuple(Size(**s) for s in p["sizes"]))) for p in model["pipes"]),
        tuple(Nozzle(**n) for n in model["nozzles"]), model["source"])
    before = asdict(net)
    settings = Settings(**dict(saved["settings"], duration_minutes=20))
    report = propose(net, settings)
    assert asdict(net) == before
    assert report["feasible"] and not report["violations"]
    assert abs(report["pump"]["flow_lpm"] - saved["pump"]["flow_lpm"]) < .001
    assert abs(report["pump"]["differential_head_m"] - saved["pump"]["differential_head_m"]) < .001
    report.update(model=asdict(net),fingerprint=saved['fingerprint'],warnings=saved['warnings'],
                  verification_origin='offline_actual_model_snapshot',options=saved['options'])
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"feasible":report['feasible'],"pump":report['pump'],"water_supply":report['water_supply'],
                      "nodes":len(net.nodes),"pipes":len(net.pipes),"nozzles":len(net.nozzles)},ensure_ascii=False))


if __name__ == "__main__":
    main()
