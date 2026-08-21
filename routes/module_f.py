# -*- coding: utf-8 -*-
"""모듈 F — CAD 임포트 웹 워크벤치. 모듈 E 의 소스를 브라우저에서 돌린다.

모듈 E(`cad_project_editor/`)는 PySide6 데스크톱 앱이라 `/module-e-cad-editor`
는 서버 PC 화면에 Qt 창을 띄울 뿐 브라우저 안에서는 아무것도 보이지 않는다.
그런데 E 는 `ui/` 밖이 전부 순수 Python 이다 — 찍기·손질·변환 세 파사드가
모두 "화면 없음" 을 계약으로 달고 있고, 판정은 전부 그 아래 board 에 있다.

그래서 이 모듈은 **E 의 소스를 한 줄도 고치지 않고** 그 파사드만 HTTP 로
열고, 캔버스·버튼은 브라우저에서 새로 그린다. Qt 는 import 하지 않는다
(`cad_project_editor/_smoke_headless.py` 로 3단 전부 Qt 없이 도는 것을 확인).

`대조 서버.py` 에서 `register(app, ...)` 로 등록. 다른 도메인 라우트와 같은
패턴이라 엔드포인트명에 접두사가 붙지 않는다.
"""
from __future__ import annotations

import heapq
import json
import math
import os
import sys
import threading
import time
import traceback
import uuid
import zipfile
from pathlib import Path

from flask import jsonify, render_template, request, send_file

EDITOR_ROOT = Path(__file__).resolve().parent.parent / "cad_project_editor"
# 찍은스펙·표시캐시·유저손질이 쌓이는 곳. 데스크톱 E 는 cwd 가 편집기 폴더라
# 상대경로 "docs/import" 로 여기를 가리킨다. 웹서버는 cwd 가 프로젝트 루트라
# 같은 상대경로가 엉뚱한 곳을 가리키므로, 부팅 때 절대경로로 고정한다.
# 같은 폴더를 쓰므로 데스크톱에서 찍은 도면이 웹에서 그대로 이어진다.
IMPORT_WORK_ROOT = EDITOR_ROOT / "docs" / "import"

# 캔버스로 내려보내는 도형 상한. B1F 실도면이 선분 69,384 개라 이 위로는
# 브라우저가 아니라 JSON 직렬화에서 먼저 막힌다. 조용히 자르지 않고 몇 개를
# 뺐는지 응답에 실어 화면에 그대로 띄운다.
MAX_SEGS = 150_000
MAX_CIRCLES = 40_000
MAX_ARCS = 40_000

SESSION_TTL_SECONDS = 3 * 3600
LOG_TAIL = 40

# 모듈 A 에서 빌려오는 것 — 레이어 이름 사전(찍기 추천)과 Remote 30 개념.
# NFPC 103 이 요구하는 것은 "가장 불리한 헤드 30개" 다. 모듈 E 는 물 닿은
# 헤드를 전부 변환하므로(실측 264개) 법정 계산 대상보다 넓다.
REMOTE_K_DEFAULT = 30

_boot_lock = threading.Lock()
_booted = False

_SESSIONS: dict[str, dict] = {}
_SESSIONS_LOCK = threading.Lock()
# 무거운 단계(도면 파싱·망 구성·평면 그래프)는 한 번에 하나만 돈다.
# docs/import 캐시와 stdout 을 공유하므로 겹치면 로그가 섞이고 캐시가 깨진다.
_HEAVY_LOCK = threading.Lock()


# ─────────────────────────────────────────────────────────── 부팅
def _boot() -> None:
    """편집기 소스를 import 가능하게 하고 쓰기 루트를 절대경로로 못박는다."""
    global _booted
    with _boot_lock:
        if _booted:
            return
        if not (EDITOR_ROOT / "main.py").exists():
            raise RuntimeError(f"모듈 E 소스를 찾을 수 없습니다: {EDITOR_ROOT}")
        root = str(EDITOR_ROOT)
        # append 다 — insert(0) 로 앞에 두면 편집기의 services/domain 이 본
        # 프로젝트의 같은 이름 패키지를 가릴 수 있다. 지금은 겹치는 이름이
        # 없지만, 나중에 생겨도 본 서버가 먼저 이기게 둔다.
        if root not in sys.path:
            sys.path.append(root)
        from services.cad_import.pipeline import disp_cache, handoff
        work = str(IMPORT_WORK_ROOT)
        handoff.import_write_root = lambda: work
        # OUT_DIR·_DISP_CACHE_DIR 은 import 때 이미 상대경로로 굳었다.
        # 함수만 갈아끼우면 이 둘은 안 따라오므로 직접 덮는다.
        handoff.OUT_DIR = handoff.pick_out_dir()
        disp_cache._DISP_CACHE_DIR = work
        os.makedirs(handoff.pick_out_dir(), exist_ok=True)
        os.makedirs(handoff.default_edits_dir(), exist_ok=True)
        _booted = True


class _Tee:
    """파이프라인이 print 로 뱉는 단계 문구를 잡아 화면 진행표시로 쓴다.

    지어낸 퍼센트를 그리지 않기 위해서다 — 실제로 찍힌 줄만 보여준다.
    서버 로그도 그대로 유지해야 하므로 원본으로도 흘려보낸다.

    sys.stdout 은 프로세스 전역이라, 잡이 도는 동안 다른 요청이 찍은 줄까지
    이 잡의 진행표시로 새어 들어간다. 그래서 **작업 스레드가 쓴 줄만** 담는다.
    """

    def __init__(self, real, sink, owner):
        self._real = real
        self._sink = sink
        self._owner = owner
        self._buf = ""

    def write(self, s):
        if self._real is not None:
            try:
                self._real.write(s)
            except Exception:  # noqa: BLE001 — 로그 실패가 작업을 죽이면 안 된다
                pass
        if threading.current_thread() is not self._owner:
            return len(s)
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            line = line.strip()
            if line:
                self._sink(line)
        return len(s)

    def flush(self):
        if self._real is not None:
            try:
                self._real.flush()
            except Exception:  # noqa: BLE001
                pass

    def isatty(self):
        return False


# ─────────────────────────────────────────────────────────── 세션
def _sweep() -> None:
    now = time.time()
    with _SESSIONS_LOCK:
        dead = [k for k, s in _SESSIONS.items()
                if now - s.get("touched", 0) > SESSION_TTL_SECONDS]
        for k in dead:
            _SESSIONS.pop(k, None)


def _new_session(**kw) -> dict:
    _sweep()
    sid = uuid.uuid4().hex[:16]
    sess = {
        "id": sid, "created": time.time(), "touched": time.time(),
        "dxf": None, "key": None, "pick": None, "edit": None,
        "world": None, "kfp": None, "kfp_path": None,
        "water_path": None, "worst": None,
        "sdf_path": None, "slf_path": None,
        "job": None, "log": [],
    }
    sess.update(kw)
    with _SESSIONS_LOCK:
        _SESSIONS[sid] = sess
    return sess


def _sess(sid: str) -> dict:
    with _SESSIONS_LOCK:
        found = _SESSIONS.get(str(sid or ""))
    if found is None:
        raise ValueError("작업이 만료되었습니다. 도면을 다시 여세요.")
    found["touched"] = time.time()
    return found


def _run_job(sess: dict, phase: str, fn) -> dict:
    """무거운 단계 하나를 백그라운드로 돌린다. 진행은 실제 출력 줄로만 보고."""
    job = {"state": "run", "phase": phase, "started": time.time(),
           "ended": None, "error": None, "result": None}
    sess["job"] = job
    sess["log"] = []

    def sink(line: str) -> None:
        log = sess["log"]
        log.append(line)
        if len(log) > 400:
            del log[:200]

    def worker() -> None:
        me = threading.current_thread()
        with _HEAVY_LOCK:
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = _Tee(old_out, sink, me)
            sys.stderr = _Tee(old_err, sink, me)
            try:
                job["result"] = fn()
                job["state"] = "done"
            except Exception as exc:  # noqa: BLE001 — 무엇이 나든 화면에 알린다
                job["state"] = "error"
                job["error"] = f"{type(exc).__name__}: {exc}"
                sink("!! " + job["error"])
                for ln in traceback.format_exc().splitlines()[-6:]:
                    sink("   " + ln)
            finally:
                sys.stdout, sys.stderr = old_out, old_err
                job["ended"] = time.time()
                sess["touched"] = time.time()

    threading.Thread(target=worker, daemon=True).start()
    return job


def _job_view(sess: dict) -> dict:
    job = sess.get("job")
    if job is None:
        return {"state": "idle", "phase": "", "elapsed": 0.0, "lines": []}
    end = job["ended"] or time.time()
    return {
        "state": job["state"], "phase": job["phase"],
        "elapsed": round(end - job["started"], 1),
        "error": job["error"],
        "lines": sess["log"][-LOG_TAIL:],
        "queued": _HEAVY_LOCK.locked() and job["state"] == "run",
    }


# ─────────────────────────────────────────────────────────── 도형 직렬화
def _r1(v) -> float:
    return round(float(v), 1)


def _layer_category(name: str) -> str:
    """모듈 A 의 레이어 이름 사전으로 분류한다(PIPE/HEAD/ALARM/ARCH/…).

    **추천이지 결정이 아니다.** 사전은 이름만 보므로 실도면에서 절반 넘게
    OTHER 로 떨어진다(B1F 51묶음 중 35). 그래도 51개를 맨눈으로 훑는 것보다는
    출발점이 되고, 확정은 모듈 E 의 찍기(사람 클릭)가 그대로 맡는다.
    """
    try:
        from remote30_prototype import _categorize_layer
    except Exception:  # noqa: BLE001 — 모듈 A 를 못 불러도 찍기는 돌아야 한다
        return "OTHER"
    try:
        return _categorize_layer(str(name))
    except Exception:  # noqa: BLE001
        return "OTHER"


def _world_payload(world) -> dict:
    """DXF 세계 → 캔버스가 그릴 수 있는 묶음별 좌표 다발.

    레이어×색(bundle) 단위로 접는다. 찍기가 재료를 그 단위로 고르므로
    화면 토글·강조도 같은 단위여야 손으로 맞출 필요가 없다.
    """
    from services.cad_import.colors import cname, rgb_dark

    bundles: dict[tuple, dict] = {}
    cat_cache: dict[str, str] = {}

    def slot(ly, c) -> dict:
        k = (ly, int(c) if isinstance(c, int) else c)
        b = bundles.get(k)
        if b is None:
            name = str(ly)
            cat = cat_cache.get(name)
            if cat is None:
                cat = _layer_category(name)
                cat_cache[name] = cat
            b = {"layer": name, "color": c, "name": cname(c),
                 "css": rgb_dark(c), "cat": cat,
                 "segs": [], "circles": [], "arcs": [],
                 "n_seg": 0, "n_circle": 0, "n_arc": 0}
            bundles[k] = b
        return b

    minx = miny = float("inf")
    maxx = maxy = float("-inf")

    def grow(x, y):
        nonlocal minx, miny, maxx, maxy
        if x < minx:
            minx = x
        if x > maxx:
            maxx = x
        if y < miny:
            miny = y
        if y > maxy:
            maxy = y

    n_seg = n_cir = n_arc = 0
    shown_seg = shown_cir = shown_arc = 0

    for ly, c, a, b in world.segs:
        n_seg += 1
        grow(a[0], a[1])
        grow(b[0], b[1])
        if shown_seg >= MAX_SEGS:
            continue
        s = slot(ly, c)
        s["segs"] += [_r1(a[0]), _r1(a[1]), _r1(b[0]), _r1(b[1])]
        s["n_seg"] += 1
        shown_seg += 1

    for ly, c, cx, cy, r in world.circles:
        n_cir += 1
        grow(cx - r, cy - r)
        grow(cx + r, cy + r)
        if shown_cir >= MAX_CIRCLES:
            continue
        s = slot(ly, c)
        s["circles"] += [_r1(cx), _r1(cy), _r1(r)]
        s["n_circle"] += 1
        shown_cir += 1

    angs = list(getattr(world, "arc_ang", ()) or ())
    for i, (ly, c, cx, cy, r) in enumerate(world.arcs):
        n_arc += 1
        grow(cx - r, cy - r)
        grow(cx + r, cy + r)
        if shown_arc >= MAX_ARCS:
            continue
        ang = angs[i] if i < len(angs) else None
        sa, sweep = (float(ang[0]), float(ang[1])) if ang else (0.0, 360.0)
        s = slot(ly, c)
        s["arcs"] += [_r1(cx), _r1(cy), _r1(r), round(sa, 2), round(sweep, 2)]
        s["n_arc"] += 1
        shown_arc += 1

    if minx == float("inf"):
        minx = miny = 0.0
        maxx = maxy = 1.0

    ordered = sorted(bundles.items(),
                     key=lambda kv: -(kv[1]["n_seg"] + kv[1]["n_circle"]))
    out = []
    for i, ((ly, c), b) in enumerate(ordered):
        b = dict(b)
        b["i"] = i
        b["id"] = f"{ly}{c}"
        out.append(b)

    cats: dict[str, int] = {}
    for b in out:
        cats[b["cat"]] = cats.get(b["cat"], 0) + 1

    return {
        "bounds": {"minx": minx, "miny": miny, "maxx": maxx, "maxy": maxy},
        "bundles": out,
        "cats": cats,
        "counts": {"segs": n_seg, "circles": n_cir, "arcs": n_arc},
        "shown": {"segs": shown_seg, "circles": shown_cir, "arcs": shown_arc},
        "dropped": {"segs": n_seg - shown_seg, "circles": n_cir - shown_cir,
                    "arcs": n_arc - shown_arc},
    }


def _worst_k_heads(pts, edges, hnodes, sources, k=REMOTE_K_DEFAULT) -> dict:
    """급수원에서 가장 먼 헤드 K개 + 그 최단경로 부분망.

    모듈 A 의 Remote 30(`select_worst30_heads`)은 제 DXF 엔티티 위에서 도는
    별개 파이프라인이라 그대로 가져올 수 없다. 그래서 **규칙만** 옮겼다 —
    급수원 기점 최단경로, 길이는 도면 실측, 가장 먼 K개. 사람이 손질로 고친
    망 위에서 도는 것이 모듈 F 의 차이다(모듈 A 는 자동 추출망 위에서 돈다).
    """
    adj: dict[int, list[tuple[int, float]]] = {}
    for a, b in edges:
        d = math.dist(pts[a], pts[b])
        adj.setdefault(a, []).append((b, d))
        adj.setdefault(b, []).append((a, d))

    INF = float("inf")
    dist: dict[int, float] = {}
    prev: dict[int, int] = {}
    pq: list[tuple[float, int]] = []
    for s in sources:
        if isinstance(s, int) and 0 <= s < len(pts):
            dist[s] = 0.0
            heapq.heappush(pq, (0.0, s))
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist.get(u, INF):
            continue
        for v, w in adj.get(u, ()):
            nd = d + w
            if nd < dist.get(v, INF):
                dist[v] = nd
                prev[v] = u
                heapq.heappush(pq, (nd, v))

    scored = []
    for hi, nodes in enumerate(hnodes):
        reach = [n for n in nodes if n in dist]
        if not reach:
            continue
        node = min(reach, key=lambda n: dist[n])
        scored.append((dist[node], hi, node))
    scored.sort(reverse=True)
    picked = scored[:max(1, int(k))]

    keep_edges: set[tuple[int, int]] = set()
    keep_nodes: set[int] = set()
    for _d, _hi, node in picked:
        cur = node
        keep_nodes.add(cur)
        while cur in prev:
            nxt = prev[cur]
            keep_edges.add((min(cur, nxt), max(cur, nxt)))
            keep_nodes.add(nxt)
            cur = nxt
    return {
        "heads": [hi for _d, hi, _n in picked],
        "dists": {hi: d for d, hi, _n in picked},
        "edges": keep_edges,
        "nodes": keep_nodes,
        "reachable": len(scored),
        "far_m": round((picked[0][0] if picked else 0.0) / 1000.0, 2),
        "near_m": round((picked[-1][0] if picked else 0.0) / 1000.0, 2),
    }


def _worst_view(sess: dict) -> dict | None:
    """화면용 — 최불리 헤드 원과 경로 간선 좌표."""
    w = sess.get("worst")
    if not w:
        return None
    b = sess["edit"].board
    pts = b.pts
    return {
        "k": len(w["heads"]),
        "reachable": w["reachable"],
        "far_m": w["far_m"],
        "near_m": w["near_m"],
        "heads": [[_r1(b.disks[hi][0]), _r1(b.disks[hi][1]), _r1(b.disks[hi][2])]
                  for hi in w["heads"] if hi < len(b.disks)],
        "path": [[_r1(pts[a][0]), _r1(pts[a][1]), _r1(pts[c][0]), _r1(pts[c][1])]
                 for a, c in w["edges"]],
    }


def _restrict_to_worst(payload: dict, board, worst: dict) -> dict:
    """변환 대상을 최불리 K 헤드로 좁힌다 — 헤드만 지우고 배관은 안 자른다.

    간선을 직접 잘라내고 싶은 유혹이 있지만 그러면 안 된다. 모듈 E 의
    `build_planar_graph` 는 이미 «급수원에서 물 닿는 간선만 남기고, 헤드로
    가지 않는 막다른관을 쳐내는» 단계를 갖고 있다(실측 로그: 물길 필터 →
    막다른관 삭제). 그러니 남길 헤드만 남겨 두면 그 배관은 E 가 제 규칙으로
    정리한다. 손으로 자르면 E 가 지키는 불변식(티 겹침·노드정리)을 깬다.

    hcov / disk_kinds / head_kinds 는 같은 디스크 집합을 가리키므로 함께 건다.
    ups 는 좌표 집합으로만 쓰여 남아 있어도 해가 없다.
    """
    from services.cad_import.kinds import disk_key

    keep_idx = {int(i) for i in (worst or {}).get("heads") or ()}
    disks = list(board.disks)
    kept = [disks[i] for i in sorted(keep_idx) if 0 <= i < len(disks)]
    if not kept:
        return payload

    keys = {disk_key(d[0], d[1], d[2]) for d in kept}
    out = dict(payload)
    out["hcov"] = [list(d) for d in kept]
    dk = payload.get("disk_kinds") or []
    out["disk_kinds"] = [dk[i] for i in sorted(keep_idx) if 0 <= i < len(dk)]

    fresh = []
    for rec in payload.get("head_kinds") or ():
        if not isinstance(rec, dict) or "c" not in rec:
            continue
        c = rec["c"]
        r = rec.get("head_r")
        if r is None and rec.get("tri_side"):
            r = float(rec["tri_side"]) / math.sqrt(3.0)
        if r is None:
            continue
        if disk_key(c[0], c[1], r) in keys:
            fresh.append(dict(rec))
    out["head_kinds"] = fresh
    print(f"[변환] 최불리 {len(kept)} 헤드로 범위를 좁힘 "
          f"(도면 헤드 {len(disks)} · 종류표 {len(fresh)}행)")
    return out


def _emit_pipenet(sess: dict, kfp: dict, out_dir: Path) -> dict:
    """.kfp → PIPENET .sdf(+표준 .slf). 11번 모듈의 변환기를 그대로 쓴다.

    모듈 A 는 제 표(`PipeTables`)에서 SDF 를 직접 찍지만, 모듈 F 의 산출물은
    모듈 E 계열의 .kfp 다. 여기서 SDF 를 새로 짜면 규약이 셋으로 갈라지므로,
    이미 있는 KFP↔SDF 변환기(`kfp_sdf_converter`)를 태운다.
    SDF 는 SLF(라이브러리) 없이 열면 PIPENET 이 "pipe bore must be given" 을
    내므로 항상 한 세트로 묶는다.
    """
    info: dict = {"ok": False}
    try:
        from kfp_sdf_converter import emit_sdf_xml, kfp_dict_to_network
        net = kfp_dict_to_network(kfp)
        xml = emit_sdf_xml(net)
    except Exception as exc:  # noqa: BLE001 — SDF 실패가 .kfp 를 무르게 하면 안 된다
        print(f"[변환] SDF 생성 실패 — .kfp 는 정상입니다: {exc}")
        info["error"] = f"{type(exc).__name__}: {exc}"
        return info

    sdf_path = out_dir / f"{sess['id']}.sdf"
    sdf_path.write_text(xml, encoding="utf-8")
    sess["sdf_path"] = str(sdf_path)
    info.update({
        "ok": True, "bytes": sdf_path.stat().st_size,
        "nodes": xml.count("<Node "), "pipes": xml.count("<Pipe "),
        "nozzles": xml.count("<Nozzle "),
    })

    try:
        from kfp_sdf_converter import _resolve_standard_slf
        slf = _resolve_standard_slf()
    except Exception:  # noqa: BLE001
        slf = None
    if slf and os.path.isfile(str(slf)):
        sess["slf_path"] = str(slf)
        info["slf"] = os.path.basename(str(slf))
    else:
        info["slf"] = None
        print("[변환] 표준 SLF 를 찾지 못했습니다 — SDF 만 담습니다.")
    print(f"[변환] PIPENET SDF · 노드 {info['nodes']} · 배관 {info['pipes']} · "
          f"노즐 {info['nozzles']} · {info['bytes']:,} bytes")
    return info


def _pick_state(sess: dict) -> dict:
    ps = sess["pick"]
    hl = ps.highlight_geom()
    spec = ps.spec()
    return {
        "mode": ps.mode,
        "armed": ps.armed,
        "mat_done": bool(ps.mat_done),
        "head_label": ps.head_label,
        "materials": [{"layer": ly, "color": c} for ly, c in ps.board.mat],
        "n_heads": len(ps.board.heads),
        "n_clicks": len(ps.board.clicks),
        "clicks": ps.board.clicks[-12:],
        "highlight": {
            "pipe_bundles": [f"{ly}{c}" for ly, c in hl["pipe_bundles"]],
            "pipe_segs": [[_r1(a[0]), _r1(a[1]), _r1(b[0]), _r1(b[1])]
                          for a, b in hl["pipe_segs"]],
            "head_circles": [[_r1(x), _r1(y), _r1(r)]
                             for x, y, r in hl["head_circles"]],
            "tri_segs": [[_r1(a[0]), _r1(a[1]), _r1(b[0]), _r1(b[1])]
                         for a, b in hl["tri_segs"]],
            "last_click": hl["last_click"],
        },
        "n_spec_heads": len(spec.get("heads") or []),
        "n_ho": len(spec.get("ho") or []),
    }


def _edit_state(sess: dict, net: bool = True) -> dict:
    from services.cad_import.colors import (
        EDIT_SOURCE, EDIT_VALVE, EDIT_WET_PIPE, KIND_COLORS)
    es = sess["edit"]
    g = es.display_geom(net=net)
    b = es.board
    kinds: dict[str, int] = {}
    for k in b.disk_kinds:
        kinds[k] = kinds.get(k, 0) + 1
    return {
        "mode": es.mode,
        "counts": {"pts": len(b.pts), "edges": len(b.edges),
                   "heads": len(b.disks), "bodies": len(g["body_groups"]),
                   "joins": len(b.joins), "deletes": len(b.deletes)},
        "kinds": kinds,
        "body_groups": [
            {"css": color,
             "segs": [[_r1(a[0]), _r1(a[1]), _r1(c[0]), _r1(c[1])]
                      for a, c in segs]}
            for segs, color in g["body_groups"]],
        "heads": [[_r1(d[0]), _r1(d[1]), _r1(d[2]), css]
                  for d, css in g["heads"]],
        "multi_heads": [[_r1(d[0]), _r1(d[1]), _r1(d[2])]
                        for d in g["multi_heads"]],
        "sources": [[_r1(x), _r1(y)] for x, y in g["sources"]],
        "valves": [[_r1(x), _r1(y)] for x, y in g["valves"]],
        "pending": ([[_r1(g["pending"][0][0]), _r1(g["pending"][0][1])],
                     [_r1(g["pending"][1][0]), _r1(g["pending"][1][1])]]
                    if g["pending"] else None),
        "selected_head": ([_r1(v) for v in g["selected_head"]]
                          if g["selected_head"] else None),
        # E 의 wet_pipes 는 연출 프레임 전용이라 연출이 끝나면 빈다(헤드 색으로
        # 결과가 남는 구조). 브라우저는 연출을 돌리지 않으므로, 물이 닿은 간선
        # 전체를 따로 붙잡아 두고 그 위에 겹쳐 그린다. 손질을 건드리면 지운다.
        "wet_pipes": [[_r1(a[0]), _r1(a[1]), _r1(c[0]), _r1(c[1])]
                      for a, c in g["wet_pipes"]] or sess.get("water_path") or [],
        "wet_counts": es.wet_kind_counts(),
        "flowed": bool(es._flowed),
        "worst": _worst_view(sess),
        "palette": {"source": EDIT_SOURCE, "valve": EDIT_VALVE,
                    "wet": EDIT_WET_PIPE, "kinds": dict(KIND_COLORS)},
        "bounds": _pts_bounds(b.pts),
    }


def _pts_bounds(pts) -> dict:
    if not pts:
        return {"minx": 0.0, "miny": 0.0, "maxx": 1.0, "maxy": 1.0}
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return {"minx": min(xs), "miny": min(ys),
            "maxx": max(xs), "maxy": max(ys)}


def _saved_keys() -> list[dict]:
    """이미 찍어 둔 도면들 — 데스크톱 E 로 찍은 것도 여기 그대로 보인다."""
    from services.cad_import.pipeline import handoff
    out = []
    d = handoff.pick_out_dir()
    if not os.path.isdir(d):
        return out
    for name in sorted(os.listdir(d)):
        if not name.endswith("_찍은스펙.json") or "자동백업" in name:
            continue
        key = name[: -len("_찍은스펙.json")]
        path = os.path.join(d, name)
        src = ""
        try:
            with open(path, encoding="utf-8") as f:
                src = json.load(f).get("source_dxf") or ""
        except Exception:  # noqa: BLE001 — 목록이므로 한 건 실패로 멈추지 않는다
            pass
        out.append({
            "key": key,
            "source_dxf": src,
            "source_exists": bool(src) and os.path.isfile(src),
            "picked_at": time.strftime("%Y-%m-%d %H:%M",
                                       time.localtime(os.path.getmtime(path))),
        })
    return out


def register(app, *, _save_upload, UPLOAD_DIR):

    def _fail(msg, code=400):
        return jsonify({"ok": False, "message": msg}), code

    @app.get("/module-f")
    def module_f_page():
        return render_template("module_f.html")

    # ─────────────────────────────────────────── 0. 열기
    @app.get("/api/module-f/saved")
    def module_f_saved():
        try:
            _boot()
            return jsonify({"ok": True, "items": _saved_keys()})
        except Exception as exc:  # noqa: BLE001
            return _fail(f"저장된 찍기 목록을 읽지 못했습니다: {exc}", 500)

    @app.post("/api/module-f/open")
    def module_f_open():
        """DXF 를 올려 찍기 세션을 연다. 파싱이 길어 잡으로 돌린다."""
        try:
            _boot()
            dxf = _save_upload("dxf_file", {".dxf"}, required=True)
        except ValueError as exc:
            return _fail(str(exc))
        except Exception as exc:  # noqa: BLE001
            return _fail(f"도면을 저장하지 못했습니다: {exc}", 500)

        sess = _new_session(dxf=str(dxf))

        def job():
            from services.cad_import.pick.session import PickSession
            t0 = time.perf_counter()
            print(f"[찍기] DXF 읽는 중 — {os.path.basename(str(dxf))}")
            ps = PickSession.open(str(dxf))
            # E 의 찍기판은 열자마자 armed 가 아니다 — Qt 대화상자도 "배관 선택"
            # 을 눌러야 클릭이 먹는다. 웹에서는 첫 할 일이 어차피 그것뿐이라
            # 같은 호출을 미리 해 둔다(엔진 상태는 단추를 누른 것과 동일).
            ps.select_pipe()
            sess["pick"] = ps
            sess["key"] = ps.key
            payload = _world_payload(ps.world)
            sess["world"] = payload
            print(f"[찍기] 완료 {time.perf_counter() - t0:.1f}s · "
                  f"선분 {payload['counts']['segs']} · "
                  f"원 {payload['counts']['circles']} · "
                  f"호 {payload['counts']['arcs']}")
            return {"key": ps.key}

        _run_job(sess, "도면 읽기", job)
        return jsonify({"ok": True, "sid": sess["id"],
                        "filename": os.path.basename(str(dxf))})

    @app.post("/api/module-f/reopen")
    def module_f_reopen():
        """이미 찍어 둔 키로 손질부터 시작한다. 찍기 단계를 건너뛴다."""
        body = request.get_json(silent=True) or {}
        key = str(body.get("key") or "").strip()
        if not key:
            return _fail("이어서 열 도면 키가 필요합니다.")
        try:
            _boot()
        except Exception as exc:  # noqa: BLE001
            return _fail(str(exc), 500)
        sess = _new_session(key=key)

        def job():
            from services.cad_import.edit.session import EditSession
            t0 = time.perf_counter()
            print(f"[손질] 저장본으로 배관망을 여는 중 — {key}")
            es = EditSession.open(key, out_dir=None, load_saved=True,
                                  use_cache=True)
            sess["edit"] = es
            print(f"[손질] 완료 {time.perf_counter() - t0:.1f}s · "
                  f"노드 {len(es.board.pts)} · 간선 {len(es.board.edges)} · "
                  f"헤드 {len(es.board.disks)}")
            return {"key": key}

        _run_job(sess, "배관망 열기", job)
        return jsonify({"ok": True, "sid": sess["id"], "key": key})

    @app.get("/api/module-f/job")
    def module_f_job():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        view = _job_view(sess)
        view["ok"] = True
        view["stage"] = ("edit" if sess.get("edit") is not None
                         else ("pick" if sess.get("pick") is not None else ""))
        view["key"] = sess.get("key")
        return jsonify(view)

    @app.get("/api/module-f/world")
    def module_f_world():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("world") is None:
            return _fail("도면이 아직 준비되지 않았습니다.")
        return jsonify({"ok": True, "world": sess["world"],
                        "key": sess["key"], "state": _pick_state(sess)})

    # ─────────────────────────────────────────── 1. 찍기
    @app.post("/api/module-f/pick/mode")
    def module_f_pick_mode():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        action = str(body.get("action") or "")
        if action == "pipe":
            ok = ps.select_pipe()
            msg = "배관(재료)을 찍으세요. 레이어×색 단위로 잡힙니다."
        elif action == "complete":
            ok = ps.complete_pipe()
            msg = ("재료 선택 완료 — 이제 헤드를 찍습니다."
                   if ok else "재료를 하나 이상 찍어야 완료할 수 있습니다.")
        elif action == "slot":
            ok = ps.set_slot(body.get("slot"))
            msg = (f"헤드 칸 = {ps.head_label}" if ok
                   else "재료 선택을 먼저 완료하세요.")
        else:
            return _fail(f"모르는 동작입니다: {action}")
        return jsonify({"ok": True, "applied": bool(ok), "message": msg,
                        "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/auto")
    def module_f_pick_auto():
        """모듈 A 의 레이어 사전이 고른 묶음을 한 번에 찍는다.

        `board.mat` 에 직접 밀어넣지 않고 **그 묶음의 실제 선분 중점**으로
        정상 클릭 경로(`PickSession.click`)를 태운다. 그래야 클릭 기록·되돌리기
        ·스펙 저장이 사람이 찍은 것과 완전히 같은 상태가 된다.
        """
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        want = str(body.get("cat") or "PIPE").upper()
        if want not in {"PIPE", "HEAD", "ALARM"}:
            return _fail(f"추천 카테고리가 아닙니다: {want}")

        world = sess.get("world") or {}
        targets = [b for b in (world.get("bundles") or []) if b.get("cat") == want]
        if not targets:
            return _fail(f"{want} 로 추천된 레이어가 없습니다. 직접 찍어 주세요.")

        if want == "PIPE":
            ps.select_pipe()
        else:
            if not ps.mat_done:
                return _fail("재료 선택을 먼저 완료해야 헤드를 찍을 수 있습니다.")
            ps.set_slot(ps.head_label)

        applied, skipped = [], []
        for b in targets:
            segs = ps.board.by_bundle.get((b["layer"], b["color"])) or []
            if not segs:
                skipped.append(b["layer"])
                continue
            a, c = segs[0]
            rep = ps.click((a[0] + c[0]) / 2.0, (a[1] + c[1]) / 2.0)
            if rep is None or rep.get("동작") != "추가":
                skipped.append(b["layer"])
            else:
                applied.append(b["layer"])
        return jsonify({
            "ok": True, "applied": applied, "skipped": skipped,
            "message": (f"{want} 추천 {len(applied)}묶음을 찍었습니다."
                        + (f" ({len(skipped)}묶음은 이미 찍혀 있거나 건너뜀)"
                           if skipped else "")),
            "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/click")
    def module_f_pick_click():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        try:
            x = float(body.get("x"))
            y = float(body.get("y"))
        except (TypeError, ValueError):
            return _fail("클릭 좌표가 올바르지 않습니다.")
        max_d = body.get("max_d")
        max_d = float(max_d) if max_d is not None else None
        rep = ps.click(x, y, max_d=max_d)
        return jsonify({"ok": True, "report": rep,
                        "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/undo")
    def module_f_pick_undo():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        undone = ps.undo()
        return jsonify({"ok": True, "undone": undone,
                        "state": _pick_state(sess)})

    @app.post("/api/module-f/pick/commit")
    def module_f_pick_commit():
        """찍은 스펙을 저장하고, 그 스펙으로 1~6단계를 다시 돌려 손질망을 만든다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        ps = sess.get("pick")
        if ps is None:
            return _fail("찍기 세션이 없습니다.")
        if not ps.mat_done:
            return _fail("재료(배관) 선택을 완료해야 다음으로 넘어갈 수 있습니다.")

        def job():
            from services.cad_import.edit.session import EditSession
            t0 = time.perf_counter()
            spec_path = ps.commit()
            print(f"[찍기] 스펙 저장 — {spec_path}")
            print("[손질] 찍은 스펙으로 배관망을 다시 구성하는 중…")
            es = EditSession.open(ps.key, out_dir=None, load_saved=False,
                                  use_cache=False)
            sess["edit"] = es
            print(f"[손질] 완료 {time.perf_counter() - t0:.1f}s · "
                  f"노드 {len(es.board.pts)} · 간선 {len(es.board.edges)} · "
                  f"헤드 {len(es.board.disks)}")
            return {"spec_path": spec_path}

        _run_job(sess, "배관망 구성", job)
        return jsonify({"ok": True})

    # ─────────────────────────────────────────── 2. 손질
    @app.get("/api/module-f/edit/state")
    def module_f_edit_state():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("edit") is None:
            return _fail("손질 세션이 없습니다.")
        return jsonify({"ok": True, "key": sess["key"],
                        "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/mode")
    def module_f_edit_mode():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        from services.cad_import.edit.session import (
            MODE_DELETE, MODE_JOIN, MODE_SOURCE, MODE_VALVE)
        allowed = {MODE_JOIN, MODE_DELETE, MODE_SOURCE, MODE_VALVE}
        mode = str(body.get("mode") or "")
        if mode not in allowed:
            return _fail(f"모르는 손질 모드입니다: {mode}")
        es.set_mode(mode)
        return jsonify({"ok": True, "state": _edit_state(sess, net=False)})

    @app.post("/api/module-f/edit/click")
    def module_f_edit_click():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        try:
            x = float(body.get("x"))
            y = float(body.get("y"))
            max_d = float(body.get("max_d"))
        except (TypeError, ValueError):
            return _fail("클릭 좌표가 올바르지 않습니다.")
        rep = es.click(x, y, max_d)
        if rep and rep.get("동작") not in ("헤드선택",):
            # 망이 바뀌면 앞서 잡아 둔 물길·최불리 선정은 더 이상 사실이 아니다.
            sess["water_path"] = None
            sess["worst"] = None
        return jsonify({"ok": True, "report": rep,
                        "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/kind")
    def module_f_edit_kind():
        """고른 헤드의 종류를 덮는다. 미지정이 남으면 변환이 막힌다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        from services.cad_import.kinds import CONFIRMED_KINDS
        kind = str(body.get("kind") or "")
        if kind not in CONFIRMED_KINDS:
            return _fail(f"헤드 종류가 아닙니다: {kind}")
        applied = es.set_kind(kind)
        if applied is None:
            return _fail("먼저 헤드를 하나 고르세요.")
        return jsonify({"ok": True, "applied": applied,
                        "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/undo")
    def module_f_edit_undo():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        ok = es.undo()
        if ok:
            sess["water_path"] = None
            sess["worst"] = None
        return jsonify({"ok": True, "undone": bool(ok),
                        "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/flow")
    def module_f_edit_flow():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        state = es.flow()
        if state is None:
            return _fail("급수 시작 위치를 먼저 찍어야 물흐름을 볼 수 있습니다.")
        # 브라우저에는 연출 프레임을 돌리지 않는다 — 끝까지 돌려 최종 상태로 둔다.
        while es.flow_tick():
            pass
        pts = es.board.pts
        sess["water_path"] = [
            [_r1(pts[a][0]), _r1(pts[a][1]), _r1(pts[b][0]), _r1(pts[b][1])]
            for a, b in state["wet_edges"]]
        return jsonify({
            "ok": True,
            "water": {"wet_heads": len(state["wet_heads"]),
                      "total_heads": state["total_heads"],
                      "wet_edges": len(state["wet_edges"]),
                      "reach": len(state["reach"])},
            "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/worst")
    def module_f_edit_worst():
        """Remote 30 — 급수원에서 가장 불리한 K 헤드와 그 경로를 고른다."""
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        b = es.board
        if not b.sources:
            return _fail("급수 시작 위치를 먼저 찍어야 최불리 헤드를 고를 수 있습니다.")
        try:
            k = int(body.get("k") or REMOTE_K_DEFAULT)
        except (TypeError, ValueError):
            k = REMOTE_K_DEFAULT
        k = max(1, min(k, 200))
        w = _worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=k)
        if not w["heads"]:
            sess["worst"] = None
            return _fail("급수원에서 닿는 헤드가 없습니다. 이음·급수 위치를 확인하세요.")
        sess["worst"] = w
        return jsonify({
            "ok": True,
            "summary": {"k": len(w["heads"]), "reachable": w["reachable"],
                        "far_m": w["far_m"], "near_m": w["near_m"],
                        "path_edges": len(w["edges"])},
            "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/worst-clear")
    def module_f_edit_worst_clear():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("edit") is None:
            return _fail("손질 세션이 없습니다.")
        sess["worst"] = None
        return jsonify({"ok": True, "state": _edit_state(sess)})

    @app.post("/api/module-f/edit/save")
    def module_f_edit_save():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        path = es.commit()
        return jsonify({"ok": True, "path": path,
                        "message": f"유저손질을 저장했습니다: {os.path.basename(path)}"})

    # ─────────────────────────────────────────── 3. 변환
    @app.get("/api/module-f/convert/fields")
    def module_f_convert_fields():
        try:
            _boot()
        except Exception as exc:  # noqa: BLE001
            return _fail(str(exc), 500)
        from services.cad_import.dto import (
            BRANCH_FIELDS, COMBO_FIELDS, FLEX_FIELDS, PENDANT_FIELDS,
            SHARED_FIELDS, UPRIGHT_FIELDS, VALVE_FIELDS, default_dto)
        groups = [
            ("메인 → 가지", BRANCH_FIELDS),
            ("상향식", UPRIGHT_FIELDS),
            ("하향식", PENDANT_FIELDS),
            ("상하향식", COMBO_FIELDS),
            ("후렉시블", FLEX_FIELDS),
            ("공통", SHARED_FIELDS),
            ("알람밸브", VALVE_FIELDS),
        ]
        return jsonify({
            "ok": True,
            "defaults": default_dto(),
            "groups": [{"title": t,
                        "fields": [{"key": k, "label": lb, "placeholder": ph}
                                   for k, lb, ph, _d in fs]}
                       for t, fs in groups],
        })

    def _src_view(src, i):
        tag = src.get("tag") if isinstance(src, dict) else None
        xy = src.get("xy") if isinstance(src, dict) else src
        return {"tag": tag or f"Z{i + 1}", "index": i + 1,
                "xy": [round(float(v), 1) for v in (xy or [0, 0])[:2]]}

    @app.post("/api/module-f/convert/run")
    def module_f_convert_run():
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        es = sess.get("edit")
        if es is None:
            return _fail("손질 세션이 없습니다.")
        dto = body.get("dto") or {}
        selected = body.get("selected_source")
        remote_only = bool(body.get("remote_only"))
        want_sdf = body.get("emit_sdf", True)
        if remote_only and not sess.get("worst"):
            return _fail("최불리 헤드를 먼저 선정해야 그 범위로 변환할 수 있습니다.")

        def job():
            from services.cad_import.convert.engine import (
                convert_to_kfp, ensure_planar)
            from services.cad_import.convert.planar import pick_convert_sources
            from services.cad_import.convert.preflight import (
                preflight_kfp_convert)
            from services.cad_import.dto import (
                default_dto, dto_to_convert_kwargs)

            payload = es.convert_payload()
            if remote_only:
                payload = _restrict_to_worst(payload, es.board, sess["worst"])
            if selected is not None:
                payload["selected_source"] = selected
            srcs = payload.get("sources") or ()
            if len(srcs) > 1:
                picked, err = pick_convert_sources(srcs, selected)
                if err:
                    return {"ok": False, "blockers": [
                        {"code": err[0], "message": err[1]}],
                        "sources": [_src_view(s, i)
                                    for i, s in enumerate(srcs)]}
                payload["sources"] = picked

            pf = preflight_kfp_convert(payload)
            if not pf["ok"]:
                print(f"[변환] 사전검사 막힘 {len(pf['blockers'])}건")
                return {"ok": False, "blockers": list(pf["blockers"]),
                        "diagnostics": list(pf.get("diagnostics") or [])}

            print("[변환] 평면 그래프를 만드는 중…")
            payload = ensure_planar(payload)
            if payload.get("kfp") is None and not payload.get("kfp_path"):
                return {"ok": False, "blockers": [{
                    "code": payload.get("_planar_code") or "planar_kfp_missing",
                    "message": payload.get("_planar_error")
                    or "평면 그래프 .kfp 가 없습니다."}]}

            merged = default_dto()
            for k, v in (dto or {}).items():
                if k in merged:
                    merged[k] = v
            print("[변환] 수직 전개 후 .kfp 를 씁니다…")
            out_dir = Path(UPLOAD_DIR) / "module_f"
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / f"{sess['id']}.kfp"
            res = convert_to_kfp(payload, str(out_path),
                                 **dto_to_convert_kwargs(merged))
            if not res["ok"]:
                return {"ok": False, "blockers": list(res["blockers"]),
                        "diagnostics": list(res.get("diagnostics") or [])}
            kfp = res["kfp"]
            sess["kfp"] = kfp
            sess["kfp_path"] = str(out_path)
            sess["sdf_path"] = None
            sess["slf_path"] = None
            stats = dict(res.get("stats") or {})
            summary = {
                "nodes": len(kfp.get("nodes_meta_runtime") or {}),
                "pipes": len(kfp.get("pipe_data") or {}),
                "bytes": out_path.stat().st_size,
                "filename": f"{sess['key'] or 'cad'}_변환.kfp",
                "remote_only": remote_only,
                "heads": len(payload.get("hcov") or []),
            }
            print(f"[변환] 완료 · 노드 {summary['nodes']} · "
                  f"배관 {summary['pipes']} · {summary['bytes']:,} bytes")

            if want_sdf:
                summary["sdf"] = _emit_pipenet(sess, kfp, out_dir)
            return {"ok": True, "stats": stats, "summary": summary,
                    "diagnostics": list(res.get("diagnostics") or [])}

        _run_job(sess, "KFP 변환", job)
        return jsonify({"ok": True})

    @app.get("/api/module-f/convert/result")
    def module_f_convert_result():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        job = sess.get("job") or {}
        return jsonify({"ok": True, "job": _job_view(sess),
                        "result": job.get("result")})

    @app.get("/api/module-f/download")
    def module_f_download():
        """`what=kfp|sdf|set` — 낱개 또는 한 벌(zip)."""
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        what = (request.args.get("what") or "kfp").lower()
        stem = sess.get("key") or "cad"
        kfp = sess.get("kfp_path")
        sdf = sess.get("sdf_path")
        slf = sess.get("slf_path")

        if what == "kfp":
            if not kfp or not os.path.isfile(kfp):
                return _fail("아직 변환된 .kfp 가 없습니다.", 404)
            return send_file(kfp, as_attachment=True,
                             download_name=f"{stem}_변환.kfp",
                             mimetype="application/json")
        if what == "sdf":
            if not sdf or not os.path.isfile(sdf):
                return _fail("아직 생성된 .sdf 가 없습니다.", 404)
            return send_file(sdf, as_attachment=True,
                             download_name=f"{stem}.sdf",
                             mimetype="application/xml")
        if what != "set":
            return _fail(f"내려받을 대상이 아닙니다: {what}")

        if not kfp or not os.path.isfile(kfp):
            return _fail("아직 변환 결과가 없습니다.", 404)
        out_dir = Path(UPLOAD_DIR) / "module_f"
        zip_path = out_dir / f"{sess['id']}_set.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(kfp, f"{stem}_변환.kfp")
            if sdf and os.path.isfile(sdf):
                z.write(sdf, f"{stem}.sdf")
            # SDF 는 라이브러리(.slf) 없이는 PIPENET 이 열지 못한다 — 같이 담는다.
            if slf and os.path.isfile(slf):
                z.write(slf, f"{stem}.slf")
        return send_file(str(zip_path), as_attachment=True,
                         download_name=f"{stem}_수리계산입력.zip",
                         mimetype="application/zip")
