"""Profile selection is explicit, persistent and isolated to its drawing."""
import json
from types import SimpleNamespace

from flask import Flask

from routes.module_f import api_pick, jobs
from routes.module_f.common import _boot


def test_profile_endpoint_checks_id_and_updates_only_pick_board(monkeypatch):
    app=Flask(__name__)
    app.testing=True
    api_pick.register(app)
    board=SimpleNamespace(head_symbol_profile="")
    sess=jobs._new_session(key="profile-test")
    sess["pick"]=SimpleNamespace(board=board)
    monkeypatch.setattr(api_pick,"_pick_state",lambda s:{"head_symbol_profile":s["pick"].board.head_symbol_profile})
    try:
        client=app.test_client()
        bad=client.post("/api/module-f/pick/head-profile",json={"sid":sess["id"],"profile":"unknown"})
        assert bad.status_code==400 and board.head_symbol_profile==""
        response=client.post("/api/module-f/pick/head-profile",json={"sid":sess["id"],"profile":"company_260404"})
        assert response.get_json()["state"]["head_symbol_profile"]=="company_260404"
        assert sess.get("edit") is None
        monkeypatch.setattr(api_pick,"_job_running",lambda _:True)
        assert client.post("/api/module-f/pick/head-profile",json={"sid":sess["id"],"profile":""}).status_code==409
    finally:
        jobs._SESSIONS.pop(sess["id"],None)


def test_profile_survives_saved_pick_spec(tmp_path):
    _boot()
    from services.cad_import.pick.board import Board
    from services.cad_import.pick.io import load_existing
    from services.cad_import.pipeline.stage1 import World, DEFAULT_KNOBS
    board=Board(World(),DEFAULT_KNOBS)
    board.head_symbol_profile="company_260404"
    spec=board.spec()
    assert spec["head_symbol_profile"]=="company_260404"
    (tmp_path/"profile-test_찍은스펙.json").write_text(json.dumps(spec),encoding="utf-8")
    new=Board(World(),DEFAULT_KNOBS)
    load_existing("profile-test",new,str(tmp_path))
    assert new.spec()["head_symbol_profile"]=="company_260404"
