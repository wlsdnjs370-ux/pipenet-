# -*- coding: utf-8 -*-
"""슬롯 «범위» 계약 — 무엇이 도면에 딸리고 무엇이 안 딸리나.

`slots.py` 의 규약은 「SESSION_KEYS 밖은 전부 슬롯 상태」다. 도면별 키를
열거하지 않는 것이 요점이라 그 설계는 옳다 — 열거하면 늘 때마다 빠뜨린다.

그런데 그 규약에는 반대쪽 짝이 있다: **도면에 안 딸리는 것은 반드시
SESSION_KEYS 에 적혀 있어야 한다.** 안 적으면 슬롯을 바꾸는 것만으로 조용히
사라진다. 실측으로 제5국면(S700 통합) 상태가 통째로 그랬다 — 급수방식을
고르고 계통도를 고치러 다녀오면 그 선택이 없어졌다.

이 시험이 지키는 것은 둘이다.
  ① 통합 상태는 슬롯을 건너 살아남는다.
  ② 도면 상태는 여전히 슬롯에 갇힌다(막는 자를 너무 넓게 잡지 않았나).
"""
from __future__ import annotations

import importlib
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _app():
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    for p in (_ROOT, os.path.join(_ROOT, "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    return srv.app


def _client(app):
    c = app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


def test_통합_상태는_슬롯을_건너_살아남는다():
    from routes.module_f.jobs import _new_session
    c = _client(_app())
    sess = _new_session(slot="plan")
    sid = sess["id"]

    rv = c.post("/api/module-f/merge/mode",
                json={"sid": sid, "mode": "hsp_pump", "source_drop_m": 3.5,
                      "pump": {"head_m": 80}})
    assert rv.status_code == 200, rv.get_data(as_text=True)[:200]
    # 결합 결과가 나온 뒤를 흉내낸다 — 실제 결합은 도면 세 장이 필요하다.
    sess["merged"] = {"fake": True}
    sess["merge_summary"] = {"nodes": 123}
    sess["merge_files"] = ["a.sdf"]

    assert c.post("/api/module-f/slot/switch",
                  json={"sid": sid, "kind": "system"}).status_code == 200
    for key, want in (("supply_mode", "hsp_pump"), ("source_drop_m", 3.5),
                      ("pump_spec", {"head_m": 80}), ("merged", {"fake": True}),
                      ("merge_summary", {"nodes": 123}),
                      ("merge_files", ["a.sdf"])):
        assert sess.get(key) == want, \
            f"슬롯을 바꾸자 통합 상태 {key} 가 사라졌다: {sess.get(key)!r}"
    # 되돌아와도 그대로 — 한 번 더 바꿔 본다(복원이 덮어쓰지 않는지).
    assert c.post("/api/module-f/slot/switch",
                  json={"sid": sid, "kind": "plan"}).status_code == 200
    assert sess.get("supply_mode") == "hsp_pump"
    assert sess.get("merged") == {"fake": True}


def test_도면_상태는_여전히_슬롯에_갇힌다():
    """막는 자를 넓게 잡아 도면별 값까지 새면 «남의 산출물» 을 제 것으로 본다."""
    from routes.module_f.jobs import _new_session
    c = _client(_app())
    sess = _new_session(slot="plan")
    sid = sess["id"]
    sess["key"] = "평면도키"
    sess["design_sdf_path"] = "/tmp/plan.sdf"

    assert c.post("/api/module-f/slot/switch",
                  json={"sid": sid, "kind": "machineroom"}).status_code == 200
    assert sess.get("key") is None, "도면 키가 슬롯을 넘어 샜다"
    assert sess.get("design_sdf_path") is None, "남의 산출물을 제 것으로 본다"

    assert c.post("/api/module-f/slot/switch",
                  json={"sid": sid, "kind": "plan"}).status_code == 200
    assert sess.get("key") == "평면도키", "돌아왔는데 도면 상태가 안 살아났다"
    assert sess.get("design_sdf_path") == "/tmp/plan.sdf"


def test_통합키는_슬롯_빈칸에_들어_있지_않다():
    """`_slot_blank` 가 통합 키를 들면 전환 때 None 으로 덮어써 버린다."""
    from routes.module_f.slots import SESSION_KEYS, _MERGE_KEYS, _slot_blank
    blank = set(_slot_blank())
    overlap = sorted(blank & _MERGE_KEYS)
    assert not overlap, f"슬롯 빈칸이 통합 키를 든다: {overlap}"
    assert _MERGE_KEYS <= SESSION_KEYS
