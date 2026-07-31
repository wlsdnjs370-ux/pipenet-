# -*- coding: utf-8 -*-
"""통합 배관망(평면도+계통도+기계실) → 각 포맷 충실도 감사.

브라우저 없이 실제 Flask 라우트(/api/remote30/combined/build)를 태워 통합망을
만들고, 방출된 SDF/SLF/KFP/HAS 를 다시 파싱해 서버가 들고 있는 진실
(응답 geometry)과 대조한다.

감사 대상 (사용자 요구: "각 포맷에 적절히 제대로 들어가는지"):
  · SDF  — 참조무결성/노드·배관·노즐 누락/bore·length·rise·C 값/부속·등가길이/
           기계실 구간 포함/연결성(단일 컴포넌트)
  · SLF  — SDF 가 참조하는 schedule 이 전부 라이브러리에 있는지(호칭경↔내경 lookup)
  · KFP  — SDF 대비 위상 보존 (passthrough 병합분 제외), 헤드 수/유량
  · HAS  — 노드·배관 수, round-trip 무손실

실행: python scripts/_format_audit.py
"""
from __future__ import annotations

import importlib.util
import math
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict, deque
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
for _p in (BASE, BASE / "core", BASE / "calibration"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

PLANE = BASE / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"
MROOM = BASE / "data" / "sample_problem" / "옥상수조.dxf"

FAILS: list[str] = []
WARNS: list[str] = []


def check(label, got, want=True, tol=None):
    if tol is not None:
        ok = got is not None and want is not None and abs(float(got) - float(want)) <= tol
    else:
        ok = (got == want)
    print(f"   {'PASS' if ok else 'FAIL'}  {label}"
          + ("" if ok else f"   (got={got!r} want={want!r})"))
    if not ok:
        FAILS.append(label)
    return ok


def warn(label, msg):
    print(f"   WARN  {label} — {msg}")
    WARNS.append(f"{label}: {msg}")


def _server_mod():
    spec = importlib.util.spec_from_file_location(
        "server_app_audit", str(BASE / "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _seed_plane_job(mod, rp, job_id):
    """golden_cases._seed_plane_job 과 동일 — 평면 job 결정적 시드."""
    job = {"dxf_path": str(PLANE)}
    for evt in rp.run_stages_0_2(PLANE, job_id, alarm_xy=None):
        if evt.get("type") != "entities":
            continue
        if evt.get("stage") == 0:
            job["layers"] = evt["layers"]
            job["bbox"] = evt["bbox"]
        elif evt.get("stage") == 1:
            job["pipe_ents"] = evt["entities"]
        elif evt.get("stage") == 2:
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


def _machine_room(rp):
    """기계실 추출 — bbox 대각 두 끝을 탱크/연결점 클릭으로 사용(snap 무제한)."""
    ents = rp.parse_dxf_for_view(MROOM, include_hidden_layers=True)["entities"]
    def _flat(v, out):
        if isinstance(v, (list, tuple)):
            for it in v:
                _flat(it, out)
        elif isinstance(v, (int, float)):
            out.append(float(v))

    xs, ys = [], []
    for e in ents:
        if e.get("t") not in ("L", "P"):
            continue
        flat: list[float] = []
        _flat(e.get("p") or [], flat)
        xs.extend(flat[0::2]); ys.extend(flat[1::2])
    src = (min(xs), max(ys))
    conn = (max(xs), min(ys))
    return rp.extract_machine_room_path(ents, src, conn, snap_tolerance_mm=1e9)


def _components(adj, labels):
    seen, comps = set(), 0
    for s in labels:
        if s in seen:
            continue
        comps += 1
        q = deque([s]); seen.add(s)
        while q:
            u = q.popleft()
            for v in adj.get(u, ()):
                if v not in seen:
                    seen.add(v); q.append(v)
    return comps


# ══════════════════════════════════════════════════════════════════════════
def main():
    print("=" * 74)
    print("통합 배관망 포맷 충실도 감사 — 평면도 + 계통도 + 기계실")
    print("=" * 74)
    for p in (PLANE, MROOM):
        if not p.is_file():
            print(f"[중단] 도면 없음: {p}")
            return 2

    mod = _server_mod()
    import remote30_prototype as rp

    print("\n① 소스 준비")
    job_id = "audit_combined"
    job = _seed_plane_job(mod, rp, job_id)
    print(f"   평면도 헤드 {len(job.get('detected_heads') or [])}개, "
          f"배관 entity {len(job.get('pipe_ents') or [])}개")
    riser = rp.extract_riser_msp_28f((0.0, 0.0), (0.0, -3000.0))
    print(f"   계통도 라이저 노드 {len(riser['nodes'])} / 배관 {len(riser['pipes'])}")
    mr = _machine_room(rp)
    print(f"   기계실 노드 {len(mr['nodes'])} / 배관 {len(mr['pipes'])} "
          f"(source={mr.get('source_node_label')}, conn={mr.get('conn_node_label')})")

    print("\n② 통합 빌드 (/api/remote30/combined/build)")
    c = mod.app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    r = c.post("/api/remote30/combined/build",
               json={"plane_job_id": job_id, "system_riser": riser,
                     "machine_room": mr})
    check("HTTP 200", r.status_code, 200)
    if r.status_code != 200:
        print(r.get_data(as_text=True)[:800])
        return 1
    body = r.json
    check("ok=True", body.get("ok"), True)
    geom = body.get("geometry") or {}
    g_nodes = geom.get("nodes") or []
    g_pipes = geom.get("pipes") or []
    g_nozzles = geom.get("nozzles") or []
    g_fittings = geom.get("fittings") or []
    g_equip = geom.get("equipment") or []
    mr_labels = set(map(str, geom.get("machine_room_labels") or []))
    print(f"   통합 노드 {len(g_nodes)} / 배관 {len(g_pipes)} / 노즐 {len(g_nozzles)}"
          f" / 부속 {len(g_fittings)} / 등가길이 {len(g_equip)}")
    print(f"   기계실 라벨 {len(mr_labels)}개 편입: {sorted(mr_labels)[:8]}")
    check("기계실이 통합망에 붙었는가", len(mr_labels) > 0, True)

    # 다운로드
    files = {}
    for key in ("sdf", "slf", "kfp", "has"):
        url = body.get(f"download_url_{key}")
        if not url:
            warn(f"{key.upper()} 다운로드 URL", "응답에 없음 (emit 실패 가능)")
            continue
        rr = c.get(url)
        if rr.status_code != 200:
            warn(f"{key.upper()} 다운로드", f"HTTP {rr.status_code}")
            continue
        out = BASE / "data" / f"_audit_out.{key}"
        out.write_bytes(rr.data)
        files[key] = out
    check("SDF/SLF/KFP/HAS 4종 모두 방출", sorted(files), ["has", "kfp", "sdf", "slf"])

    # ══════════════════════════════════════════════════════════════════
    print("\n③ SDF 감사 (PIPENET 입력 원본)")
    if "sdf" not in files:
        FAILS.append("SDF 없음")
        return 1
    root = ET.parse(files["sdf"]).getroot()
    sdf_nodes = {n.get("label"): n for n in root.iter("Node")}
    sdf_pipes = {p.get("label"): p for p in root.iter("Pipe")}
    sdf_noz = list(root.iter("Nozzle"))
    print(f"   SDF 노드 {len(sdf_nodes)} / 배관 {len(sdf_pipes)} / 노즐 {len(sdf_noz)}")

    # (1) 참조무결성 — 모든 Pipe/Nozzle 의 input·output 이 선언된 Node 인가
    dangling = [(p.get("label"), p.get("input"), p.get("output"))
                for p in sdf_pipes.values()
                if p.get("input") not in sdf_nodes or p.get("output") not in sdf_nodes]
    check(f"배관 端點 전부 선언된 Node (미선언 {len(dangling)})", len(dangling), 0)
    if dangling:
        print(f"        예: {dangling[:5]}")
    noz_dangling = [(z.get("label"), z.get("input"), z.get("output"))
                    for z in sdf_noz
                    if z.get("input") not in sdf_nodes or z.get("output") not in sdf_nodes]
    check(f"노즐 端點 전부 선언된 Node (미선언 {len(noz_dangling)})", len(noz_dangling), 0)

    # (2) 누락 — geometry 의 노드/배관/노즐이 전부 SDF 에 있는가
    miss_n = [str(n["label"]) for n in g_nodes if str(n["label"]) not in sdf_nodes]
    check(f"노드 누락 0 ({len(g_nodes)}개 중)", len(miss_n), 0)
    if miss_n:
        print(f"        누락: {miss_n[:10]}")
    miss_p = [str(p["label"]) for p in g_pipes if str(p["label"]) not in sdf_pipes]
    check(f"배관 누락 0 ({len(g_pipes)}개 중)", len(miss_p), 0)
    if miss_p:
        print(f"        누락: {miss_p[:10]}")
    check(f"노즐 수 일치 ({len(sdf_noz)} vs geometry {len(g_nozzles)})",
          len(sdf_noz), len(g_nozzles))

    # (3) FX materialize 로 늘어난 배관 (등가길이 Equipment → 실배관 확장)
    extra_p = len(sdf_pipes) - len(g_pipes)
    print(f"   FX 실배관 확장분 = SDF 배관 - geometry 배관 = {extra_p}")

    # (4) 수치 전달 — bore(mm→m) / length / rise / roughness-or-c
    bad_bore, bad_len, bad_rise, bad_c = [], [], [], []
    for p in g_pipes:
        el = sdf_pipes.get(str(p["label"]))
        if el is None:
            continue
        try:
            if abs(float(el.get("bore")) - float(p["dia"]) / 1000.0) > 1e-6:
                bad_bore.append((p["label"], el.get("bore"), p["dia"]))
            if abs(float(el.get("length")) - float(p["length"])) > 1e-3:
                bad_len.append((p["label"], el.get("length"), p["length"]))
            if abs(float(el.get("rise")) - float(p.get("elev", 0.0) or 0.0)) > 1e-3:
                bad_rise.append((p["label"], el.get("rise"), p.get("elev")))
            if abs(float(el.get("roughness-or-c")) - float(p["c"])) > 1e-6:
                bad_c.append((p["label"], el.get("roughness-or-c"), p["c"]))
        except (TypeError, ValueError) as exc:
            bad_bore.append((p["label"], f"파싱실패 {exc}", None))
    check(f"내경 bore = dia/1000 (불일치 {len(bad_bore)})", len(bad_bore), 0)
    if bad_bore:
        print(f"        예: {bad_bore[:5]}")
    check(f"길이 length 일치 (불일치 {len(bad_len)})", len(bad_len), 0)
    if bad_len:
        print(f"        예: {bad_len[:5]}")
    check(f"상승고 rise = elev (불일치 {len(bad_rise)})", len(bad_rise), 0)
    if bad_rise:
        print(f"        예: {bad_rise[:5]}")
    check(f"C계수 roughness-or-c = c (불일치 {len(bad_c)})", len(bad_c), 0)
    if bad_c:
        print(f"        예: {bad_c[:5]}")

    # (5) 노즐 유량 (L/min → m3/s)
    noz_by_in = {}
    for z in sdf_noz:
        fd = z.find("Flow-define")
        if fd is not None:
            noz_by_in[z.get("input")] = float(fd.get("flow"))
    bad_q = []
    for z in g_nozzles:
        got = noz_by_in.get(str(z["in"]))
        want = float(z.get("flow_lmin", 0.0)) / 60000.0
        # SDF 는 유효숫자 6자리로 직렬화 → 상대오차 허용
        if got is None or abs(got - want) > max(1e-12, abs(want) * 1e-5):
            bad_q.append((z.get("label"), got, want))
    check(f"노즐 유량 = L/min ÷ 60000 (불일치 {len(bad_q)})", len(bad_q), 0)
    if bad_q:
        print(f"        예: {bad_q[:5]}")

    # (6) 부속류 <Fitting> — geometry 의 (배관,종류,개수) 가 전부 실렸는가
    sdf_fit = defaultdict(list)
    for lbl, el in sdf_pipes.items():
        for f in el.iter("Fitting"):
            sdf_fit[lbl].append((f.get("type"), int(float(f.get("count", 1)))))
    miss_f = [(f["pipe"], f["type"], f["count"]) for f in g_fittings
              if (str(f["type"]), int(f["count"])) not in sdf_fit.get(str(f["pipe"]), [])]
    check(f"부속류 전달 (미전달 {len(miss_f)} / 전체 {len(g_fittings)})", len(miss_f), 0)
    if miss_f:
        print(f"        예: {miss_f[:5]}")

    # (7) 등가길이 <Equipment> — FX 는 실배관으로 이관되므로 파이프 무관 합계로 대조
    sdf_eq_total = sum(float(e.get("equivalent-length", 0) or 0)
                       for e in root.iter("Equipment"))
    g_eq_total = sum(float(e.get("eq_len", 0) or 0) for e in g_equip)
    check(f"등가길이 총합 보존 (SDF {sdf_eq_total:.2f} vs geometry {g_eq_total:.2f})",
          sdf_eq_total, g_eq_total, tol=0.05)
    n_sdf_eq = len(list(root.iter("Equipment")))
    check(f"등가길이 항목 수 일치 ({n_sdf_eq} vs {len(g_equip)})", n_sdf_eq, len(g_equip))

    # (8) 기계실 구간이 SDF 안에 실제로 존재하는가
    mr_in_sdf = [lb for lb in mr_labels if lb in sdf_nodes]
    check(f"기계실 노드 전원 SDF 등재 ({len(mr_in_sdf)}/{len(mr_labels)})",
          len(mr_in_sdf), len(mr_labels))
    mr_pipes = [lb for lb, el in sdf_pipes.items()
                if el.get("input") in mr_labels or el.get("output") in mr_labels]
    check(f"기계실 배관 SDF 등재 ({len(mr_pipes)}개)", len(mr_pipes) > 0, True)

    # (9) 경계조건 — Input 노드 + 압력
    inputs = [lb for lb, el in sdf_nodes.items() if el.get("io-node") == "Input"]
    check(f"Input(수원) 노드 정확히 1개 ({inputs})", len(inputs), 1)
    if inputs:
        cs = sdf_nodes[inputs[0]].find("Calculation-spec")
        check("Input 노드에 Calculation-spec 압력",
              cs is not None and bool(cs.get("pressure")), True)

    # (10) 연결성 — 배관망이 하나로 이어져 있는가
    adj = defaultdict(set)
    for el in sdf_pipes.values():
        a, b = el.get("input"), el.get("output")
        adj[a].add(b); adj[b].add(a)
    for z in sdf_noz:
        adj[z.get("input")].add(z.get("output"))
        adj[z.get("output")].add(z.get("input"))
    comps = _components(adj, list(sdf_nodes))
    check(f"단일 연결 컴포넌트 (={comps})", comps, 1)

    # (11) Pipe-set / schedule 참조
    ptypes = {pt.findtext("Name", "").strip() for pt in root.iter("Pipe-type")}
    ptypes.discard("")
    print(f"   SDF 참조 schedule: {sorted(ptypes)}")

    # ══════════════════════════════════════════════════════════════════
    print("\n④ SLF 감사 (호칭경↔내경 라이브러리)")
    if "slf" in files:
        slf_root = ET.parse(files["slf"]).getroot()
        # SLF 항목명은 <Item-name> 텍스트에 있다 (Schedule/Nozzle/Pump 공통)
        slf_names = {(el.text or "").strip() for el in slf_root.iter("Item-name")}
        slf_names.discard("")
        missing_sched = sorted(ptypes - slf_names)
        check(f"SDF 참조 schedule 전부 SLF 에 존재 (누락 {len(missing_sched)})",
              len(missing_sched), 0)
        if missing_sched:
            print(f"        누락: {missing_sched}")
        libs = [n for n in {str(z["lib"]) for z in g_nozzles} if n]
        miss_lib = [n for n in libs if n not in slf_names]
        check(f"노즐 라이브러리 항목 존재 ({libs} 누락 {len(miss_lib)})", len(miss_lib), 0)
        # 펌프 곡선 (가압 방식이 펌프일 때만)
        n_pump = len(list(root.iter("Pump-fan")))
        print(f"   SDF <Pump-fan> {n_pump}개 (자연낙차면 0 정상)")
        # PIPENET 은 .sdf 와 .slf 가 같은 폴더에 있어야 호칭경↔내경 lookup 이 된다
        import io as _io, zipfile as _zf
        rr = c.get(body.get("download_url_sdf_zip") or "")
        if rr.status_code == 200:
            names = _zf.ZipFile(_io.BytesIO(rr.data)).namelist()
            exts = sorted(Path(n).suffix for n in names)
            check(f"PIPENET ZIP = .sdf + .slf 쌍 ({names})", exts, [".sdf", ".slf"])
            stems = {Path(n).stem for n in names}
            check("ZIP 안 두 파일의 stem 동일 (PIPENET lookup 조건)", len(stems), 1)
        else:
            FAILS.append("PIPENET ZIP 다운로드 실패")
    else:
        FAILS.append("SLF 없음")

    # ══════════════════════════════════════════════════════════════════
    print("\n⑤ KFP 감사 (K-Fire Solver)")
    import kfp_sdf_converter as kc
    net_sdf = kc.parse_sdf(files["sdf"])
    # 수리 권위값은 SDF 의 length attribute (= geometry 의 실배관장). parse_sdf 는
    # 노드 좌표거리로 length 를 덮어쓰므로 총연장 기준으로 쓰면 안 된다.
    L_true = sum(float(el.get("length", 0) or 0) for el in sdf_pipes.values())
    print(f"   실배관장(SDF length attr) 총연장 {L_true:.2f}m")
    if "kfp" in files:
        net_kfp = kc.parse_kfp(files["kfp"])
        print(f"   SDF→Common 노드 {len(net_sdf.nodes)} / 배관 {len(net_sdf.pipes)}")
        print(f"   KFP→Common 노드 {len(net_kfp.nodes)} / 배관 {len(net_kfp.pipes)}")
        merged = len(net_sdf.nodes) - len(net_kfp.nodes)
        print(f"   passthrough 병합 {merged}노드 / {len(net_sdf.pipes) - len(net_kfp.pipes)}배관")
        check("KFP 노드 수 ≤ SDF (병합만 허용, 증식 없음)",
              len(net_kfp.nodes) <= len(net_sdf.nodes), True)
        # 헤드(노즐) 는 절대 병합 대상이 아니다 → 수가 보존돼야 한다
        n_noz_sdf = sum(1 for n in net_sdf.nodes.values() if n.kind in ("nozzle", "head"))
        n_noz_kfp = sum(1 for n in net_kfp.nodes.values() if n.kind in ("nozzle", "head"))
        check(f"헤드 수 보존 (KFP {n_noz_kfp} / SDF {n_noz_sdf} / geometry {len(g_nozzles)})",
              (n_noz_kfp, n_noz_sdf), (len(g_nozzles), len(g_nozzles)))
        # 배관 총연장 — 병합은 일직선 노드만이므로 총연장은 보존돼야 한다
        L_kfp = sum(p.length_m for p in net_kfp.pipes.values())
        check(f"총연장 보존 (KFP {L_kfp:.2f}m vs 실배관장 {L_true:.2f}m)",
              L_kfp, L_true, tol=max(0.5, L_true * 0.002))
        # 연결성
        adj_k = defaultdict(set)
        for p in net_kfp.pipes.values():
            adj_k[p.start].add(p.end); adj_k[p.end].add(p.start)
        ck = _components(adj_k, list(net_kfp.nodes))
        check(f"KFP 단일 연결 컴포넌트 (={ck})", ck, 1)
        # 내경 전달
        bad_kd = [p.id for p in net_kfp.pipes.values() if not p.diameter_inner_mm]
        check(f"KFP 내경 0 인 배관 없음 ({len(bad_kd)})", len(bad_kd), 0)
        # 수원 경계
        n_wt = sum(1 for n in net_kfp.nodes.values() if n.kind == "wt")
        check(f"KFP 수원(wt) 노드 존재 ({n_wt})", n_wt >= 1, True)
        # ★ 좌표 규약 — K-Fire_Solver 는 노드 3D 좌표거리에서 배관장을 역산한다
        # (사용자 수작업 .kfp 실측: 좌표거리/length_m 중앙값 1.0000, n=139/122).
        # 통합망도 이 규약을 지켜야 솔버 표·자동수리가 실배관장을 쓴다.
        import json as _json
        _kj = _json.loads(Path(files["kfp"]).read_text(encoding="utf-8"))
        _kn = {k: (v.get("coords") or [])
               for k, v in (_kj.get("nodes_meta_runtime") or {}).items()}
        _rat = []
        _coord_total = 0.0
        for _p in (_kj.get("pipe_data") or {}).values():
            _s, _e = _kn.get(_p.get("start")), _kn.get(_p.get("end"))
            _lm = float(_p.get("length_m") or 0.0)
            if not _s or not _e or _lm <= 0:
                continue
            _d = math.dist(_s[:3], _e[:3])
            _coord_total += _d
            if _d > 1e-9:
                _rat.append(_d / _lm)
        _rat.sort()
        if _rat:
            _p10 = _rat[int(len(_rat) * 0.10)]
            _med = _rat[len(_rat) // 2]
            _p90 = _rat[int(len(_rat) * 0.90)]
            print(f"   좌표거리/length_m — p10 {_p10:.3f} / 중앙값 {_med:.4f} / p90 {_p90:.3f}")
            print(f"   좌표 총연장 {_coord_total:.2f}m vs 실배관장 {L_true:.2f}m")
            check(f"좌표거리==length_m 규약 (중앙값 {_med:.4f})", _med, 1.0, tol=0.02)
    else:
        FAILS.append("KFP 없음")

    # ══════════════════════════════════════════════════════════════════
    print("\n⑥ HAS 감사 (HASS)")
    if "has" in files:
        import core.has_converter as hc
        net_has = hc.parse_has(files["has"])
        print(f"   HAS→Common 노드 {len(net_has.nodes)} / 배관 {len(net_has.pipes)}")
        check("HAS 배관 수 = KFP 배관 수",
              len(net_has.pipes), len(net_kfp.pipes) if "kfp" in files else len(net_has.pipes))
        L_has = sum(p.length_m for p in net_has.pipes.values())
        check(f"HAS 총연장 보존 ({L_has:.2f}m vs 실배관장 {L_true:.2f}m)",
              L_has, L_true, tol=max(0.5, L_true * 0.002))
        n_noz_has = sum(1 for n in net_has.nodes.values() if n.kind == "nozzle")
        check(f"HAS 헤드 수 보존 ({n_noz_has} vs geometry {len(g_nozzles)})",
              n_noz_has, len(g_nozzles))
        rt = hc.round_trip_check(files["has"])
        check(f"HAS round-trip 노드 무손실 ({rt['nodes_in']}→{rt['nodes_out']})",
              rt["nodes_out"], rt["nodes_in"])
        check(f"HAS round-trip 배관 무손실 ({rt['pipes_in']}→{rt['pipes_out']})",
              rt["pipes_out"], rt["pipes_in"])
        rmse = rt.get("length_rmse_m")
        if rmse is not None and not math.isnan(rmse):
            check(f"HAS round-trip 길이 RMSE {rmse:.4f}m", rmse < 0.01, True)
    else:
        FAILS.append("HAS 없음")

    # ══════════════════════════════════════════════════════════════════
    print("\n" + "=" * 74)
    if FAILS:
        print(f"결과: {len(FAILS)} FAIL")
        for f in FAILS:
            print(f"  ✗ {f}")
    else:
        print("결과: 전 항목 PASS")
    if WARNS:
        print(f"경고 {len(WARNS)}건")
        for w in WARNS:
            print(f"  ! {w}")
    print("=" * 74)
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
