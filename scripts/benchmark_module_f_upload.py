"""Measure Module F upload processing in an isolated work directory.

Optional --baseline reads selected modules from Git into this process only;
it never replaces working files or removes a user's saved edits/caches.
Times exclude browser compression and network transfer. --profile adds overhead.
"""
from __future__ import annotations

import argparse
import cProfile
import hashlib
import importlib.abc
import importlib.util
import json
from pathlib import Path
import pstats
import subprocess
import sys
import tempfile
import time

ROOT = Path(__file__).resolve().parents[1]
MODULES = {
    "services.cad_import.pipeline.heads": "cad_project_editor_g/services/cad_import/pipeline/heads.py",
    "services.cad_import.pipeline.flow": "cad_project_editor_g/services/cad_import/pipeline/flow.py",
    "services.cad_import.pipeline.stage45": "cad_project_editor_g/services/cad_import/pipeline/stage45.py",
    "services.cad_import.pick.board": "cad_project_editor_g/services/cad_import/pick/board.py",
    "routes.module_f.adopt": "routes/module_f/adopt.py",
}


class GitModules(importlib.abc.MetaPathFinder, importlib.abc.Loader):
    """Load benchmark baseline modules without changing files on disk."""

    def __init__(self, ref: str):
        self.sources = {
            name: subprocess.check_output(["git", "show", f"{ref}:{path}"], cwd=ROOT)
            for name, path in MODULES.items()
        }

    def find_spec(self, fullname, path=None, target=None):
        if fullname in self.sources:
            return importlib.util.spec_from_loader(fullname, self)
        return None

    def create_module(self, spec):
        return None

    def exec_module(self, module):
        module.__file__ = str(ROOT / MODULES[module.__name__])
        exec(compile(self.sources[module.__name__], module.__file__, "exec"), module.__dict__)


def fingerprint(value: object) -> str:
    """Hash complete structured output, retaining order and numeric precision."""
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("dxf", type=Path)
    parser.add_argument("--baseline")
    parser.add_argument("--profile", action="store_true")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--graph", action="store_true")
    args = parser.parse_args()
    sys.path.insert(0, str(ROOT))
    if args.baseline:
        sys.meta_path.insert(0, GitModules(args.baseline))
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import handoff
    from services.cad_import.pick.session import PickSession
    from routes.module_f.world import _world_payload
    from routes.module_f.recon import dominant_band, run_recon
    from routes.module_f.adopt import adopt_bundles, adopt_heads, select_heads, ADOPT_MAX_D_MM
    import remote30_prototype as module_a

    report = {"source": str(args.dxf.resolve()), "bytes": args.dxf.stat().st_size,
              "baseline": args.baseline, "profile": args.profile, "seconds": {}}
    profiler = cProfile.Profile()

    def measure(name, fn):
        print(f"START {name}", flush=True)
        start = time.perf_counter()
        result = fn()
        report["seconds"][name] = round(time.perf_counter() - start, 4)
        print(f"DONE {name}: {report['seconds'][name]}s", flush=True)
        return result

    old_root, old_cache = handoff.import_write_root(), module_a._PARSE_CACHE_DIR
    output_dir = ROOT / "cad_project_editor_g" / "docs" / "benchmarks"
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        with tempfile.TemporaryDirectory(prefix="upload_", dir=output_dir) as work:
            handoff.set_write_root(work)
            module_a._PARSE_CACHE_DIR = Path(work) / "parse_cache"
            if args.profile:
                profiler.enable()
            ps = measure("read_and_pick_index", lambda: PickSession.open(str(args.dxf.resolve())))
            world = measure("display_payload", lambda: _world_payload(ps.world))
            recon = measure("recognition_cold", lambda: run_recon(args.dxf.resolve(), world=world))
            ps.select_pipe()
            materials = measure("adopt_materials", lambda: adopt_bundles(ps, world, "PIPE"))
            if ps.complete_pipe():
                ps.set_slot(ps.head_label)
                lo = dominant_band(recon["bands"])["conf_min"]
                picks = select_heads(recon["heads"], conf_min=lo) if lo is not None else []
                adopted = measure("adopt_heads", lambda: adopt_heads(ps, picks, max_d=ADOPT_MAX_D_MM))
                report["adopted"] = adopted
            report["counts"] = world["counts"]
            report["fingerprints"] = {
                "display": fingerprint(world), "heads": fingerprint(recon["heads"]),
                "pick_spec": fingerprint(ps.spec()), "materials": fingerprint(materials),
            }
            if args.graph and ps.mat_done:
                measure("save_pick", ps.commit)
                from services.cad_import.edit.session import EditSession
                es = measure("build_network", lambda: EditSession.open(ps.key, load_saved=True, use_cache=False))
                b = es.board
                report["network_counts"] = {"nodes": len(b.pts), "pipes": len(b.edges), "heads": len(b.disks)}
                report["fingerprints"]["network"] = fingerprint({
                    "pts": b.pts, "edges": sorted(tuple(sorted(e)) for e in b.edges),
                    "disks": b.disks, "kinds": b.disk_kinds,
                })
            if args.profile:
                profiler.disable()
                pstats.Stats(profiler).strip_dirs().sort_stats("cumulative").print_stats(35)
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            print(json.dumps({k: v for k, v in report.items() if k != "adopted"}, ensure_ascii=False), flush=True)
    finally:
        handoff.set_write_root(old_root)
        module_a._PARSE_CACHE_DIR = old_cache


if __name__ == "__main__":
    main()
