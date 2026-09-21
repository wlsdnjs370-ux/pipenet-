"""Reproduce KFP axis-extension failures in the repository's editor clone.

This is a diagnostic, not a test of the installed Solver executable. It never
overwrites the input and runs each attempted extension in a fresh editor.
"""

from __future__ import annotations

import copy
import hashlib
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "cad_project_editor_g"))

from editor_core import PipeEditor  # noqa: E402


def attempts(payload: dict, nodes: list[str]) -> dict:
    """Try one metre in each signed axis without retaining trial edits."""
    results = []
    for node in nodes:
        for axis in ("+X", "-X", "+Y", "-Y", "+Z", "-Z"):
            editor = PipeEditor()
            editor.from_dict(copy.deepcopy(payload))
            target = "_AUDIT_NEW_NODE_"
            assert target not in payload["nodes_meta_runtime"]
            ok, info = editor.add_branch_from_node_axis(node, axis, 1.0, target)
            message = info.get("msg", "")
            outcome = (
                "success" if ok else
                "axis_error" if "축과 평행하지" in message else
                "other_rejection"
            )
            results.append({"node": node, "axis": axis, "outcome": outcome,
                            "message": message})
    return {"counts": dict(Counter(r["outcome"] for r in results)),
            "results": results}


def main() -> None:
    """Compare the original, precision grid, and save/reopen behaviour."""
    source = Path(sys.argv[1]).resolve()
    before = source.read_bytes()
    payload = json.loads(before)
    nodes = list(payload["nodes_meta_runtime"])
    precise = copy.deepcopy(payload)
    precise["grid_step"] = 1e-6
    assert precise["nodes_meta_runtime"] == payload["nodes_meta_runtime"]
    assert precise["pipe_data"] == payload["pipe_data"]

    report = {
        "engine": "repository cad_project_editor_g; not installed Solver v5.5",
        "input_sha256": hashlib.sha256(before).hexdigest(),
        "node_count": len(nodes),
        "original": attempts(payload, nodes),
        "precision_grid": attempts(precise, nodes),
    }
    editor = PipeEditor()
    editor.from_dict(copy.deepcopy(precise))
    saved = editor.to_dict()
    reopened = PipeEditor()
    reopened.from_dict(copy.deepcopy(saved))
    report["roundtrip"] = {
        "serialized_grid_step": saved.get("grid_step"),
        "loaded_grid_step": reopened.grid_step,
        "N152": attempts(saved, ["N152"]),
    }
    assert source.read_bytes() == before, "Input must remain unchanged"
    report["original_file_unchanged"] = True
    output = Path(__file__).with_name("all_nodes_probe.json")
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "original": report["original"]["counts"],
        "precision_grid": report["precision_grid"]["counts"],
        "reopened_grid": reopened.grid_step,
        "reopened_N152": report["roundtrip"]["N152"]["counts"],
        "input_unchanged": True,
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
