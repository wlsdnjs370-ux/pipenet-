# -*- coding: utf-8 -*-
"""KFP 변환 직전 엔진 검사 — UI 없음.

계획 API: extract_pick_spec → build_planar_graph →
  apply_user_edits / confirm_head_kinds → preflight_kfp_convert → convert_to_kfp

이 모듈은 preflight 만. convert_to_kfp · 변환 폼 · Node.head_kind 는 여기 없다.
종류 권위: 1-1 classify → kind_overrides. 찍기 heads[].kind 는 권위 아님.
미지정(또는 kind 없음) → 변환 차단. 상향 가정 금지.
"""
from services.cad_import.kinds import normalize_head_kind, require_head_kinds
from services.cad_import.pipeline.user_net import apply_kind_overrides

# 확정된 종류만 변환 진행. 미지정·빈값·그 밖 = 미확정.
_CONFIRMED_KINDS = ("상향식", "하향식", "상하향식")
BLOCKER_UNCONFIRMED_HEADS = "unconfirmed_heads"


def _topology_revalidate(_payload):
    """물길·토폴로지 재검증 훅.

    합격/불합격 재검증 함수는 아직 없다. wet_from_sources 는 변환 시
    물닿음 필터이지 게이트가 아니다. 새 잇기 엔진을 만들지 않는다.
    반환: (status, blockers, diagnostics) — 지금은 not_run · 빈 목록.
    """
    return ("not_run", [],
            [{"code": "topology", "status": "not_run"}])


def _unconfirmed_heads(head_kinds):
    """kind 가 상향식|하향식|상하향식 이 아닌 레코드. 분류기를 새로 돌리지 않는다."""
    bad = []
    for rec in head_kinds or ():
        if not isinstance(rec, dict):
            rec = {}
        kind = normalize_head_kind(rec.get("kind"))
        if kind in _CONFIRMED_KINDS:
            continue
        item = {"kind": "미지정"}
        if "c" in rec:
            item["c"] = list(rec["c"])
        if "head_r" in rec:
            item["head_r"] = rec["head_r"]
        bad.append(item)
    return bad


def preflight_kfp_convert(payload):
    """현재 편집 그래프/리비전으로 변환 가능 여부를 돌려준다.

    payload 키 (있는 것만 씀 · 없는 키는 건너뜀):
      head_kinds, kind_overrides, hcov|disks
    종류는 이미 보존된 head_kinds 를 읽는다. KFP Node.head_kind 불필요.

    반환:
      {ok: bool,
       blockers: [{code, message, ...}, ...],
       diagnostics: [...]}
    ok 는 blockers 가 비었을 때만 True. diagnostics 는 차단이 아니다.
    """
    payload = payload or {}
    head_kinds = list(payload.get("head_kinds") or [])
    ovs = payload.get("kind_overrides") or []
    if ovs:
        head_kinds = apply_kind_overrides(head_kinds, ovs)
    hcov = payload.get("hcov")
    if hcov is None:
        hcov = payload.get("disks")
    if hcov is not None:
        head_kinds = require_head_kinds(hcov, head_kinds)

    blockers = []
    bad = _unconfirmed_heads(head_kinds)
    if bad:
        blockers.append({
            "code": BLOCKER_UNCONFIRMED_HEADS,
            "message": "미지정 헤드가 있습니다. 편집으로 돌아가 헤드 종류를 확정하세요.",
            "count": len(bad),
            "heads": bad,
        })

    _status, topo_blockers, topo_diag = _topology_revalidate(payload)
    blockers.extend(topo_blockers)
    return {
        "ok": not blockers,
        "blockers": blockers,
        "diagnostics": list(topo_diag),
    }


def smoke():
    """합성 그래프만. 이음 규칙을 돌리지 않는다."""
    ok = preflight_kfp_convert({
        "hcov": [(100.0, 0.0, 5.0), (200.0, 0.0, 5.0), (300.0, 0.0, 5.0)],
        "head_kinds": [
            {"c": (100.0, 0.0), "head_r": 5.0, "kind": "상향식"},
            {"c": (200.0, 0.0), "head_r": 5.0, "kind": "하향식"},
            {"c": (300.0, 0.0), "head_r": 5.0, "kind": "상하향식"},
        ],
    })
    assert ok["ok"] is True
    assert ok["blockers"] == []
    assert ok["diagnostics"] == [{"code": "topology", "status": "not_run"}]

    blocked = preflight_kfp_convert({
        "hcov": [(100.0, 0.0, 5.0), (200.0, 0.0, 5.0)],
        "head_kinds": [
            {"c": (100.0, 0.0), "head_r": 5.0, "kind": "상향식"},
            {"c": (200.0, 0.0), "head_r": 5.0, "kind": "미지정"},
        ],
    })
    assert blocked["ok"] is False
    assert len(blocked["blockers"]) == 1
    b = blocked["blockers"][0]
    assert b["code"] == BLOCKER_UNCONFIRMED_HEADS
    assert b["count"] == 1
    assert "편집" in b["message"]
    assert b["heads"][0]["kind"] == "미지정"

    missing = preflight_kfp_convert({
        "hcov": [(100.0, 0.0, 5.0)],
        "head_kinds": [{"c": (100.0, 0.0), "head_r": 5.0}],
    })
    assert missing["ok"] is False
    assert missing["blockers"][0]["code"] == BLOCKER_UNCONFIRMED_HEADS

    filled = preflight_kfp_convert({
        "hcov": [(100.0, 0.0, 5.0), (200.0, 0.0, 5.0)],
        "head_kinds": [
            {"c": (100.0, 0.0), "head_r": 5.0, "kind": "하향식"},
        ],
    })
    assert filled["ok"] is False
    assert filled["blockers"][0]["count"] == 1

    ov = preflight_kfp_convert({
        "hcov": [(100.0, 0.0, 5.0)],
        "head_kinds": [
            {"c": (100.0, 0.0), "head_r": 5.0, "kind": "미지정"},
        ],
        "kind_overrides": [
            {"c": [100.0, 0.0], "r": 5.0, "kind": "상향식"},
        ],
    })
    assert ov["ok"] is True
    print("gate: preflight ok=all_confirmed "
          "block=미지정 block=missing_kind block=hcov_gap "
          "override_confirms=ok topology=not_run")


if __name__ == "__main__":
    smoke()
