# -*- coding: utf-8 -*-
"""Remote30 특성화(characterization) golden 케이스 등록부.

리팩토링 안전망 — 현재 동작을 golden 스냅샷으로 동결해두고, 리팩토링 단계마다
출력이 변하지 않는지 자동 대조한다(동작 보존 검증). 외부 기능을 바꾸지 않는
리팩토링이라면 모든 케이스가 green 이어야 한다.

각 케이스는 zero-arg callable 이며 JSON 직렬화 가능한 canonical dict 를 돌려준다.
부동소수/그래프 등은 정렬·반올림·해시로 안정화한다.

추가/실행은 test_golden.py 참조.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
_CAL = REPO_ROOT / "calibration"
_CORE = REPO_ROOT / "core"
for _p in (str(REPO_ROOT), str(_CAL), str(_CORE)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

# ── 픽스처 (작고 대표성 있는 실제 도면) ──────────────────────────────────
FIXTURES = {
    "plane_daemyeong": "samples/dxf/대명동201동 단위세대_layer정리.dxf",
    "system_daemyeong_min": "data/sample_problem/대명동201동 계통도_최소.dxf",
    "system_lh306": "samples/dxf/계통도_LH_306.dxf",
    "fitting_tee": "samples/dxf/분기티.dxf",
    "machineroom_tank": "data/sample_problem/옥상수조.dxf",
}


def fixture_path(key: str) -> Path:
    return REPO_ROOT / FIXTURES[key]


# ── 모듈 로딩 (1회 캐시) ──────────────────────────────────────────────────
_CACHE: dict = {}


def _rp():
    if "rp" not in _CACHE:
        import remote30_prototype as rp  # noqa: E402
        _CACHE["rp"] = rp
    return _CACHE["rp"]


def _li():
    if "li" not in _CACHE:
        import linkpred_integrate as li  # noqa: E402
        _CACHE["li"] = li
    return _CACHE["li"]


def _raw_graph_fn():
    if "raw_graph" not in _CACHE:
        from clean_candidate_survey import _raw_graph  # noqa: E402
        _CACHE["raw_graph"] = _raw_graph
    return _CACHE["raw_graph"]


def _server_mod():
    if "server_mod" not in _CACHE:
        spec = importlib.util.spec_from_file_location(
            "server_app_golden", str(REPO_ROOT / "대조 서버.py"))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        _CACHE["server_mod"] = mod
    return _CACHE["server_mod"]


def _app():
    return _server_mod().app


# ── canonical 직렬화 헬퍼 ─────────────────────────────────────────────────
def _sha(obj) -> str:
    return hashlib.sha1(
        json.dumps(obj, sort_keys=True, ensure_ascii=False).encode("utf-8")
    ).hexdigest()[:16]


def _n_components(graph: dict, edge_len: dict) -> int:
    parent: dict = {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for n in graph:
        find(n)
    for a, b in edge_len:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra
    return len({find(n) for n in parent})


def _canon_graph(graph: dict, edge_len: dict) -> dict:
    """노드/엣지 수 + 정렬된 엣지(정수반올림·길이0.1) 해시로 그래프 동결.

    edge_sig 는 "달라졌다"만 알려주고 무엇이 달라졌는지는 못 알려준다. 총연장과
    연결성분 수를 함께 동결해 회귀가 났을 때 방향(간선 소실 vs 분단)을 읽는다.
    """
    edges = []
    for (a, b), L in edge_len.items():
        edges.append((int(round(a[0])), int(round(a[1])),
                      int(round(b[0])), int(round(b[1])), round(float(L), 1)))
    edges.sort()
    return {
        "nodes": len(graph),
        "edges": len(edge_len),
        "total_len": round(sum(float(v) for v in edge_len.values()), 1),
        "components": _n_components(graph, edge_len),
        "edge_sig": _sha(edges),
    }


def _canon_entities(parsed: dict) -> dict:
    ents = parsed.get("entities", [])
    by_t: dict = {}
    for e in ents:
        by_t[e.get("t", "?")] = by_t.get(e.get("t", "?"), 0) + 1
    layers = parsed.get("layers", [])
    layer_names = sorted(
        (l.get("name") if isinstance(l, dict) else str(l)) for l in layers
    )
    return {
        "n_ent": len(ents),
        "by_type": dict(sorted(by_t.items())),
        "n_layers": len(layers),
        "layers": layer_names,
        "entities_sig": _sha(ents),
    }


# 실행마다 달라지는 휘발성 필드 — golden 동결 전 제거 (캐시 플래그/타임스탬프/
# run-dir 해시/다운로드 URL). 재귀적으로 키 기준 삭제하므로 부동소수를 깨지 않는다.
_VOLATILE_KEYS = {
    "cached", "from_cache", "run_id", "job_id", "elapsed", "elapsed_ms",
    "ms", "duration", "timestamp", "ts",
    "download_url", "download_url_zip", "download_url_sdf",
    "download_url_slf", "download_url_kfp", "url",
}


def _normalize_response(obj):
    """엔드포인트 응답에서 휘발성 필드 제거 + 부동소수 반올림."""
    if isinstance(obj, dict):
        return {k: _normalize_response(v) for k, v in obj.items()
                if k not in _VOLATILE_KEYS}
    if isinstance(obj, list):
        return [_normalize_response(v) for v in obj]
    if isinstance(obj, float):
        return round(obj, 3)
    return obj


def _post_dxf(route: str, fixture_key: str, field: str):
    app = _app()
    c = app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    p = fixture_path(fixture_key)
    with open(p, "rb") as fh:
        data = {field: (fh, p.name)}
        r = c.post(route, data=data, content_type="multipart/form-data")
    return r


# ── 케이스 정의 ────────────────────────────────────────────────────────────
def _case_parse(fixture_key):
    def run():
        rp = _rp()
        parsed = rp.parse_dxf_for_view(str(fixture_path(fixture_key)),
                                       include_hidden_layers=True)
        return _canon_entities(parsed)
    return run


def _case_raw_graph(fixture_key):
    def run():
        rp = _rp()
        parsed = rp.parse_dxf_for_view(str(fixture_path(fixture_key)),
                                       include_hidden_layers=True)
        g, el, _scale = _raw_graph_fn()(parsed["entities"])
        return _canon_graph(g, el)
    return run


def _case_build_graph(fixture_key):
    def run():
        rp = _rp()
        parsed = rp.parse_dxf_for_view(str(fixture_path(fixture_key)),
                                       include_hidden_layers=True)
        g, el, stats = rp.build_system_graph(parsed["entities"])
        out = _canon_graph(g, el)
        out["stats"] = {k: stats[k] for k in sorted(stats)
                        if isinstance(stats[k], (int, float, str, bool))}
        return out
    return run


def _case_reconcile(fixture_key):
    def run():
        rp = _rp()
        li = _li()
        pair = li.load_model("allt")
        if pair is None:
            return {"skipped": "model_unavailable"}
        model, feats = pair
        parsed = rp.parse_dxf_for_view(str(fixture_path(fixture_key)),
                                       include_hidden_layers=True)
        res = li.reconcile_entities(parsed["entities"], model, feats)
        return li.serialize_result(res)
    return run


def _case_endpoint_connection_review(fixture_key):
    def run():
        r = _post_dxf("/api/remote30/system/connection_review",
                      fixture_key, "system_dxf_file")
        return {"status": r.status_code, "body": _normalize_response(r.json)}
    return run


# 안정적(git-tracked) PIPENET 3-1형 레퍼런스 SDF — _analyze_sdf_sprinkler_network 동결용
REF_SDF_31 = "3-1형_자연낙차_LSP_4F_OA_지하층포함_120m~200m미만_6.6K로 감압_알람밸브.sdf"


def _case_analyze_sdf():
    def run():
        mod = _server_mod()
        res = mod._analyze_sdf_sprinkler_network(REPO_ROOT / "assets" / REF_SDF_31)
        return _normalize_response(res)
    return run


def _seed_plane_job(mod, plane_key, job_id):
    """run_stages_0_2 를 소비해 평면 job 을 결정적으로 시드 (stream 엔드포인트와 동일 로직).

    combined_build 가 요구하는 detected_heads/pipe_ents/layer_cat/dxf_path/alarm_xy 를
    채워 mod._PROTOTYPE_JOBS[job_id] 에 심는다. alarm_xy 는 build_input_tables 가 AV("10")
    노드를 만들도록 첫 탐지 헤드 좌표로 고정(결정적).
    """
    rp = _rp()
    p = fixture_path(plane_key)
    job = {"dxf_path": str(p)}
    for evt in rp.run_stages_0_2(p, job_id, alarm_xy=None):
        if evt.get("type") != "entities":
            continue
        stage = evt.get("stage")
        if stage == 0:
            job["layers"] = evt["layers"]
            job["bbox"] = evt["bbox"]
        elif stage == 1:
            job["pipe_ents"] = evt["entities"]
        elif stage == 2:
            detected = []
            for be in evt["entities"]:
                if be.get("t") == "B":
                    q = be["p"]
                    detected.append({"pos": [(q[0] + q[2]) / 2, (q[1] + q[3]) / 2],
                                     "bbox": q, "k": be.get("k", ""),
                                     "c": be.get("c", 0), "i": be.get("i", 0)})
            job["detected_heads"] = detected
            job["layer_cat"] = {l["name"]: l["auto_category"]
                                for l in job.get("layers", [])}
    if job.get("detected_heads"):
        job["alarm_xy"] = tuple(job["detected_heads"][0]["pos"])
    mod._PROTOTYPE_JOBS[job_id] = job
    return job


def _normalize_combined(body):
    """combined/build 응답에서 휘발성 plumbing(job_id·다운로드 URL·파일명) 제거.

    동결 대상은 geometry(노드/배관 좌표·io)와 구조 카운트. job_id 는 secrets 로 매번
    달라지고 sdf/zip/download_url* 파일명에 박혀 있어 모두 제거한다.
    """
    if not isinstance(body, dict):
        return _normalize_response(body)
    drop = {k for k in body if k.startswith("download_url")}
    drop |= {"job_id", "sdf", "zip"}
    out = {k: _normalize_response(v) for k, v in body.items() if k not in drop}
    # geometry.riser_labels 는 엔드포인트에서 list(set(...)) 이라 프로세스마다 순서가
    # 달라진다(hash seed). 집합 순서는 무의미하므로 정렬해 스냅샷을 안정화한다.
    geom = out.get("geometry")
    if isinstance(geom, dict) and isinstance(geom.get("riser_labels"), list):
        geom["riser_labels"] = sorted(geom["riser_labels"])
    return out


def _case_combined_build(plane_key):
    def run():
        mod = _server_mod()
        rp = _rp()
        job_id = "golden_combined_plane"
        _seed_plane_job(mod, plane_key, job_id)
        riser = rp.extract_riser_msp_28f((0.0, 0.0), (0.0, -3000.0))
        c = mod.app.test_client()
        with c.session_transaction() as s:
            s["authed"] = True
        r = c.post("/api/remote30/combined/build",
                   json={"plane_job_id": job_id, "system_riser": riser})
        return {"status": r.status_code, "body": _normalize_combined(r.json)}
    return run


def _case_endpoint_inspect(fixture_key):
    def run():
        r = _post_dxf("/api/remote30/inspect", fixture_key, "dxf_file")
        # progress 로 흘러오는 entities 전부를 모아 정규 해시로 동결한다 — 렌더러
        # (_render_entity 등) 변경 시 회귀를 잡는다(result 메타만으로는 못 잡음).
        text = r.get_data(as_text=True)
        result = None
        ents: list = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except Exception:
                continue
            if msg.get("type") == "progress":
                ents.extend(msg.get("entities", []))
            elif msg.get("type") == "result":
                result = msg
        meta = ({k: v for k, v in result.items() if k != "entities"}
                if result is not None else None)
        return {
            "status": r.status_code,
            "result": _normalize_response(meta),
            "n_entities": len(ents),
            "entities_sig": _sha(ents),
        }
    return run


# name -> callable
CASES: dict = {}


def _register():
    core_targets = [
        ("plane_daemyeong", True),
        ("system_daemyeong_min", True),
        ("system_lh306", True),
        ("fitting_tee", True),
        ("machineroom_tank", True),
    ]
    for key, _ in core_targets:
        CASES[f"parse__{key}"] = _case_parse(key)
        CASES[f"raw_graph__{key}"] = _case_raw_graph(key)
        CASES[f"build_graph__{key}"] = _case_build_graph(key)
    # reconcile 는 계통도/평면도 대표만 (모델 필요)
    for key in ("system_lh306", "system_daemyeong_min", "plane_daemyeong"):
        CASES[f"reconcile__{key}"] = _case_reconcile(key)
    # 엔드포인트 (라우트 스캐폴딩 통과)
    CASES["ep_connection_review__system_lh306"] = \
        _case_endpoint_connection_review("system_lh306")
    CASES["ep_inspect__system_lh306"] = _case_endpoint_inspect("system_lh306")
    # SDF 분석기 (대조 서버._analyze_sdf_sprinkler_network) — Phase 3 분해 안전망
    CASES["analyze_sdf__ref31"] = _case_analyze_sdf()
    # combined_build (stateful) — 평면 job 시드 + legacy 라이저 → Phase 3 분해 안전망
    CASES["combined_build__plane_daemyeong"] = \
        _case_combined_build("plane_daemyeong")


_register()
