"""KFP geometry preservation, legacy editing, and actual download boundaries."""
from __future__ import annotations

import copy
import json
import math
from pathlib import Path

import pytest
from flask import Flask

from core.kfp_editability import prepare_kfp
from routes.module_f import api_convert, api_merge, jobs
from routes.module_f.kfp_export import finish_kfp


def network(points=None, edges=None):
    points = points or {"N152": (91.16033980168356, 37.1540934468307, 0),
                        "N72": (94.6706296230139, 37.1540934468307, 0),
                        "N94": (91.16033980168356, 37.1540934468307, .3)}
    if edges is None:
        edges = {"P211": ("N72", "N152"), "P270": ("N152", "N94")}
    return {"version": "4.0-NFPA13-EQL", "nodes_meta_runtime": {
        n: {"coords": list(p), "elevation_m": p[2], "type_id": "base", "type": "기본"}
        for n, p in points.items()}, "pipe_data": {
            pid: dict(start=a, end=b, length_m=math.dist(points[a], points[b]),
                      nominal_mm=25, diameter=27.6, C=120, equivalent_length=3.1,
                      fittings=["Elbow"], flow_lpm=80, velocity_mps=2, headloss_m=.2)
            for pid, (a, b) in edges.items()}}


def test_precise_preserves_all_calculation_fields_and_input():
    original = network()
    untouched = copy.deepcopy(original)
    result, report = prepare_kfp(original)
    assert original == untouched
    assert result == {**original, "grid_step": 1e-6}
    assert set(report.off_grid_nodes) == {"N152", "N72", "N94"}
    assert report.warnings and not report.errors
    assert prepare_kfp(result)[0] == result


def test_edit_copy_reports_changes_and_keeps_pipe_attributes():
    original = network()
    result, report = prepare_kfp(original, mode="solver-edit")
    assert not report.errors
    assert result["grid_step"] == .01
    assert result["nodes_meta_runtime"]["N152"]["coords"] == pytest.approx([91.16, 37.15, 0])
    assert report.changed_nodes == 3
    changes = {r["pipe"]: r for r in report.length_changes}
    assert changes["P211"]["after_m"] == pytest.approx(3.51)
    assert report.max_length_change_mm == pytest.approx(abs(3.51 - original["pipe_data"]["P211"]["length_m"]) * 1000)
    for pid, p in original["pipe_data"].items():
        for key in ("start", "end", "C", "diameter", "nominal_mm", "equivalent_length", "fittings"):
            assert result["pipe_data"][pid][key] == p[key]
        assert result["pipe_data"][pid]["flow_lpm"] == 0
    assert original["pipe_data"]["P211"]["flow_lpm"] == 80


@pytest.mark.parametrize("points,edges,message", [
    ({"a": (0, 0, 0), "b": (.004, 0, 0)}, {"p": ("a", "b")}, "겹칩니다"),
    ({"a": (0, 0, 0), "b": (1, 1, 0)}, {"p": ("a", "b")}, "사선"),
    ({"a": (0, 0, 0), "b": (2, 0, 0), "c": (1, .004, 0), "d": (3, .004, 0)},
     {"p": ("a", "b"), "q": ("c", "d")}, "새 겹침"),
])
def test_unsafe_rounding_never_emits(points, edges, message):
    source = network(points, edges)
    before = copy.deepcopy(source)
    result, report = prepare_kfp(source, mode="solver-edit")
    assert result is None
    assert any(message in m for m in report.errors)
    assert source == before


@pytest.mark.parametrize("fault,message", [("length", "계산 길이"), ("elevation", "표고")])
def test_schematic_coordinates_are_not_calculation_geometry(fault, message):
    source = network()
    if fault == "length":
        source["pipe_data"]["P211"]["length_m"] = 90
    else:
        source["nodes_meta_runtime"]["N152"]["elevation_m"] = 100
    assert prepare_kfp(source)[0]["pipe_data"] == source["pipe_data"]
    result, report = prepare_kfp(source, mode="solver-edit")
    assert result is None and any(message in m for m in report.errors)


@pytest.mark.parametrize("fault", ["nan", "missing_node", "zero", "coords"])
def test_invalid_network_is_reported(fault):
    source = network()
    if fault == "nan":
        source["nodes_meta_runtime"]["N152"]["coords"][0] = float("nan")
    elif fault == "missing_node":
        source["pipe_data"]["P211"]["start"] = "absent"
    elif fault == "zero":
        source["pipe_data"]["P211"]["start"] = "N152"
    else:
        source["nodes_meta_runtime"]["N152"]["coords"] = None
    result, report = prepare_kfp(source)
    assert result is None and report.errors


def test_contact_check_limit_fails_closed(monkeypatch):
    import core.kfp_editability as compat
    monkeypatch.setattr(compat, "MAX_CONTACT_CHECKS", 0)
    result, report = prepare_kfp(network(), mode="solver-edit")
    assert result is None and any("검사 범위" in m for m in report.errors)


def test_legacy_coordinates_supported():
    source = network()
    source["version"] = "3.0"
    source["nodes"] = {n: v["coords"] for n, v in source.pop("nodes_meta_runtime").items()}
    result, report = prepare_kfp(source, mode="solver-edit")
    assert not report.errors
    assert result["nodes"]["N152"] == pytest.approx([91.16, 37.15, 0])


@pytest.fixture(scope="module")
def editor_type():
    from routes.module_f.common import _boot
    _boot()
    from editor_core import PipeEditor
    return PipeEditor


@pytest.mark.parametrize("point", [(91.16033980168356, 37.1540934468307, 0), (-.0037, -.0041, -.1234)])
@pytest.mark.parametrize("axis", ["+X", "-X", "+Y", "-Y", "+Z", "-Z"])
def test_axis_extension_after_save_and_reopen(editor_type, point, axis):
    source = editor_type().to_dict()
    source.update(network({"N152": point}, {}))
    old = editor_type()
    old.from_dict(copy.deepcopy(source))
    ok, info = old.add_branch_from_node_axis("N152", axis, 1, "N166")
    assert not ok and "축과 평행하지" in info["msg"]
    assert "N166" not in old.to_dict()["nodes_meta_runtime"]
    # Precision export fixes initial editing without changing existing geometry.
    precise, _ = prepare_kfp(source)
    ed = editor_type()
    ed.from_dict(precise)
    assert ed.add_branch_from_node_axis("N152", axis, 1, "N166")[0]
    # The distinct cm-grid copy survives a legacy serializer losing grid_step.
    editing, report = prepare_kfp(source, mode="solver-edit")
    assert not report.errors
    for _ in range(2):
        ed = editor_type()
        ed.from_dict(editing)
        editing = ed.to_dict()
        editing.pop("grid_step", None)  # Emulate a Solver that omits metadata.
    ed = editor_type()
    ed.from_dict(editing)
    assert ed.add_branch_from_node_axis("N152", axis, 1, "N166")[0]


def test_real_editor_preserves_overlap_rejection(editor_type):
    source = editor_type().to_dict()
    source.update(network())
    editing, _ = prepare_kfp(source, mode="solver-edit")
    ed = editor_type()
    ed.from_dict(editing)
    before = ed.to_dict()
    ok, info = ed.add_branch_from_node_axis("N152", "+Z", .3, "N166")
    assert not ok and "이미 배관" in info["msg"]
    assert ed.to_dict()["nodes_meta_runtime"] == before["nodes_meta_runtime"]


@pytest.fixture
def download_app(tmp_path):
    app = Flask(__name__)
    app.config["TESTING"] = True
    api_convert.register(app, UPLOAD_DIR=tmp_path)
    api_merge.register(app, UPLOAD_DIR=tmp_path)
    sess = jobs._new_session()
    path = tmp_path / "original.kfp"
    path.write_text(json.dumps(network()), encoding="utf8")
    sess.update(key="시험", kfp_path=str(path), worst_kfp_path=str(path),
                merge_files={"kfp": str(path), "kfp_iso": str(path), "warnings": ["test"]})
    yield app.test_client(), sess, path
    jobs._SESSIONS.pop(sess["id"], None)


@pytest.mark.parametrize("route,what", [("download", "kfp"), ("download", "worst-kfp"),
                                       ("merge/download", "kfp"), ("merge/download", "kfp_iso")])
def test_all_download_paths_audit_and_preserve_original(download_app, route, what):
    client, sess, path = download_app
    before = path.read_bytes()
    url = f"/api/module-f/{route}?sid={sess['id']}&what={what}"
    response = client.get(url)
    assert response.status_code == 200
    assert response.json["grid_step"] == 1e-6
    assert response.json["pipe_data"] == json.loads(before)["pipe_data"]
    assert client.get(url + "&variant=solver-edit").status_code == 409
    checked = client.get(url + "&variant=solver-edit&inspect=1").json
    assert checked["report"]["ok"] and checked["report"]["length_changes"]
    edit = client.get(url + "&variant=solver-edit&source_sha256=" + checked["source_sha256"])
    assert edit.status_code == 200
    assert edit.json["grid_step"] == .01
    assert "1cm.kfp" in edit.headers["Content-Disposition"]
    assert path.read_bytes() == before
    path.write_text(json.dumps(network({"N152": (1.123, 2.456, 0)}, {})), encoding="utf8")
    assert client.get(url + "&variant=solver-edit&source_sha256=" + checked["source_sha256"]).status_code == 409


def test_invalid_download_is_clear_and_not_a_file(download_app):
    client, sess, path = download_app
    url = f"/api/module-f/download?sid={sess['id']}&what=kfp"
    assert client.get(url + "&variant=invalid").status_code == 400
    p = network({"a": (0, 0, 0), "b": (.004, 0, 0)}, {"p": ("a", "b")})
    path.write_text(json.dumps(p), encoding="utf8")
    report = client.get(url + "&variant=solver-edit&inspect=1").json["report"]
    assert not report["ok"] and report["errors"]
    assert client.get(url + "&variant=solver-edit").status_code == 422
    assert client.get(f"/api/module-f/merge/download?sid={sess['id']}&what=warnings").status_code == 404


def test_new_output_file_receives_grid(tmp_path):
    path = tmp_path / "export.kfp"
    source = network()
    result, report = finish_kfp(path, source)
    assert json.loads(path.read_text(encoding="utf8")) == result
    assert result["nodes_meta_runtime"] == source["nodes_meta_runtime"]
    assert not report["errors"]
    assert len(list(tmp_path.iterdir())) == 1


def test_convert_full_and_worst_use_audited_output(download_app, editor_type, monkeypatch):
    from types import SimpleNamespace
    from services.cad_import.convert import engine, preflight
    from services.cad_import.design import emit
    client, sess, path = download_app
    source = network()

    def generate(*args, **kwargs):
        # File generation belongs to the existing converter; exercise the real
        # F route's subsequent audit, file replacement and returned summary.
        out = args[-1]
        Path(out).write_text(json.dumps(source), encoding="utf8")
        return {"ok": True, "kfp": copy.deepcopy(source), "stats": {}}

    monkeypatch.setattr(engine, "convert_to_kfp", generate)
    monkeypatch.setattr(engine, "ensure_planar", lambda data: {**data, "kfp": source})
    monkeypatch.setattr(preflight, "preflight_kfp_convert", lambda _: {"ok": True})
    monkeypatch.setattr(emit, "emit_design_kfp", generate)
    monkeypatch.setattr(api_convert, "_run_job", lambda s, _title, fn: s.update(job={"result": fn()}))
    sess.update(edit=SimpleNamespace(convert_payload=lambda: {}), worst={"heads": [1]},
                design={"tables": None, "got": {}})
    response = client.post("/api/module-f/convert/run", json={"sid": sess["id"],
        "outputs": {"full_kfp": True, "worst_kfp": True}})
    assert response.status_code == 200
    summary = sess["job"]["result"]["summary"]
    for key, part in (("kfp_path", "full"), ("worst_kfp_path", "worst")):
        written = json.loads(Path(sess[key]).read_text(encoding="utf8"))
        assert written == {**source, "grid_step": 1e-6}
        assert summary[part]["compatibility"]["off_grid_nodes"]
