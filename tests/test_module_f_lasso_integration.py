"""HTTP selection/save/reopen use the exact freehand boundary."""
import json

from test_module_f_flow_integration import flow_client
from services.cad_import.edit.board import EditBoard
from services.cad_import.edit.io import load_edits


ZONE={"type":"polygon","points":[[1500,1500],[2500,1500],[2500,2500],[1500,2500]]}


def test_http_polygon_worst_save_restore_and_clear(flow_client):
    client,sess=flow_client
    sid=sess["id"]
    original=sess["edit"].board
    response=client.post("/api/module-f/edit/zones",json={"sid":sid,"zones":[ZONE]})
    assert response.status_code==200,response.get_json()
    worst=client.post("/api/module-f/edit/worst",json={"sid":sid,"k":1,"zones":[ZONE]})
    assert worst.status_code==200,worst.get_json()
    assert sess["worst"]["heads"]==[0]
    assert sess["worst"]["zones"]==[ZONE]
    assert client.get("/api/module-f/edit/state",query_string={"sid":sid}).get_json()["state"]["worst"]["zones"]==[ZONE]
    saved=client.post("/api/module-f/edit/save",json={"sid":sid,"zones":[ZONE]})
    assert saved.status_code==200,saved.get_json()
    b=EditBoard(original.key,original.pts,original.edges,original.disks)
    assert load_edits(b,sess["edit"].out_dir)
    assert b.selection_zones==[ZONE]
    cleared=client.post("/api/module-f/edit/zones",json={"sid":sid,"zones":[]})
    assert cleared.status_code==200 and sess["worst"] is None
    assert cleared.get_json()["state"]["selection_zones"]==[]


def test_invalid_save_does_not_overwrite_previous_handwork(flow_client):
    client,sess=flow_client
    sid=sess["id"]
    client.post("/api/module-f/edit/save",json={"sid":sid,"zones":[ZONE]})
    from pathlib import Path
    paths=list(Path(sess["edit"].out_dir).glob("*.json"))
    before={p:p.read_bytes() for p in paths}
    bad=client.post("/api/module-f/edit/save",json={"sid":sid,"zones":[{"type":"polygon","points":[[0,0]]}]})
    assert bad.status_code==400
    assert all(p.read_bytes()==value for p,value in before.items())
