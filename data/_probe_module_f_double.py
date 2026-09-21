# -*- coding: utf-8 -*-
"""잡이 도는 중에 같은 작업을 또 밀어넣으면 어떻게 되나 — 실측."""
from __future__ import annotations
import importlib.util, os, sys, time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
spec = importlib.util.spec_from_file_location("daejo", os.path.join(ROOT, "대조 서버.py"))
srv = importlib.util.module_from_spec(spec); sys.modules["daejo"] = srv
spec.loader.exec_module(srv)
app = srv.app; app.config["TESTING"] = True

with app.test_client() as c:
    with c.session_transaction() as s:
        s["authed"] = True
    r = c.post("/api/module-f/reopen", json={"key": "B1F 현장조사 소화설비 평면도"})
    sid = r.get_json()["sid"]
    while c.get(f"/api/module-f/job?sid={sid}").get_json()["state"] == "run":
        time.sleep(0.5)
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    e0 = st["counts"]["edges"]
    print(f"시작 간선 {e0} · 덩이 {st['counts']['bodies']}")

    c.post("/api/module-f/edit/autojoin/scan", json={"sid": sid})
    print("후보 준비됨")

    # 가림막이 캔버스만 덮으므로 사람이 실제로 두 번 누를 수 있다.
    a1 = c.post("/api/module-f/edit/autojoin/apply", json={"sid": sid}).get_json()
    a2 = c.post("/api/module-f/edit/autojoin/apply", json={"sid": sid}).get_json()
    print("1차:", a1)
    print("2차:", a2, "  ← 거절되어야 정상")

    while c.get(f"/api/module-f/job?sid={sid}").get_json()["state"] == "run":
        time.sleep(1.0)
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    print(f"끝 간선 {st['counts']['edges']} (증가 {st['counts']['edges']-e0})"
          f" · 덩이 {st['counts']['bodies']}")
    rep = st.get("autojoin_report")
    print("보고:", rep)

    r = c.post("/api/module-f/edit/undo", json={"sid": sid}).get_json()
    back = r["state"]["counts"]["edges"]
    print(f"되돌리기 한 번 → 간선 {back} (원래 {e0}) "
          f"{'✔ 복구' if back == e0 else '✘ 안 돌아옴'}")
