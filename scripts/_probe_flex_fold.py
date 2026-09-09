# -*- coding: utf-8 -*-
"""[신축배관 접기] 조치 전/후 같은 표 — 지시서 `ModuleF_신축배관_접기_지시서.md` §4.

어제의 「빼기」는 철회됐다. 신축배관은 **헤드를 가지관에 잇는 유일한 경로**이고
(빼면 물닿음 111 → 5), 도면의 후렉시블 선과 변환이 만드는 수직 접속관은
서로 다른 구간이다. 그래서 빼는 대신 **한 가닥을 직선 하나로 접는다.**

여기서 재는 것:

    [가닥]  신축배관 묶음의 선분을 사슬로 복원 — 가닥 수 · 선분 분포
            총연장(유클리드) · 현 길이 · **L1(맨해튼) 합** ← 표 length 의 근거
            한 끝이 헤드에 · 다른 끝이 가지관에 닿는가
            사슬이 아닌 것(차수 3 이상 · 고리)은 세기만
    [산출]  등각 격자 벗어남 · 평면 비직각 · X 교차 · 접속관 겹침
    [표]    물닿음 헤드 · 45° 엘보 계상 · 신축배관 표 총연장

★L1 을 함께 재는 이유 — `planar.py:875` 는 배관 길이를 **유클리드가 아니라
  L1(|Δx|+|Δy|+|Δz|)** 로 잡는다. 가닥이 x·y 로 단조롭게 흐르면 그 합은 현의
  L1 과 **같다**. 그러면 접어도 표 length 가 안 변한다 — 새 통로 없이 §5 기준 3
  이 성립하는지가 여기서 갈린다. 아니면 그 자리를 보고하고 멈춘다(§2).

    python scripts/_probe_flex_fold.py [--tag 조치전] [--k 30]
"""
from __future__ import annotations

import argparse
import math
import os
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
# 자동 사전 — **추천만** 한다(지시서 §3). 대명동에서는 과매칭 0 이었다.
FLEX_WORDS = ("후렉시블", "후렉", "플렉", "flex", "fx")
SNAP = 10.0          # mm — 끝점 스냅. 「헤드에 10 mm 이내」와 같은 자
NEAR = 10.0          # mm — 「닿는다」의 자


def is_flex(layer) -> bool:
    s = str(layer or "").lower()
    return any(w.lower() in s for w in FLEX_WORDS)


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


# ───────────────────────────────────────────── 가닥 복원
class _Cluster:
    """끝점을 **eps 로 뭉친다** — 격자 반올림이 아니다.

    ★`round(x / SNAP)` 로 버킷을 나누면 4.9 mm 와 5.1 mm 가 «다른 칸» 이 된다.
      한 폴리라인의 이어진 두 선분이 그렇게 갈라져 가닥이 토막 난다 — 실측으로
      320 선분이 85 가닥이 아니라 **196 토막**으로 나왔다(가닥당 1~4 선분 ·
      중앙값 200 mm). 이 저장소가 이미 값을 치른 함정이다(DXF 끝점 → 그래프
      노드 매핑은 격자 snap 금지 · NodeIndex 로 raw 좌표 보존).
    """

    def __init__(self, eps=0.5):
        self.eps = float(eps)
        self.cell = max(self.eps, 1e-9)
        self.grid: dict = {}
        self.at: list = []

    def key(self, p):
        gx, gy = int(p[0] // self.cell), int(p[1] // self.cell)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for i in self.grid.get((gx + dx, gy + dy), ()):
                    q = self.at[i]
                    if abs(q[0] - p[0]) <= self.eps \
                            and abs(q[1] - p[1]) <= self.eps:
                        return i
        i = len(self.at)
        self.at.append((float(p[0]), float(p[1])))
        self.grid.setdefault((gx, gy), []).append(i)
        return i


def strands(segs, eps=0.5):
    """선분 목록 → 사슬(가닥). 차수 1 끝점 둘을 가진 것만 «가닥» 이다.

    ★폴리라인 정체성은 `stage1` 이 선분으로 쪼갠 뒤라 이미 없다(지시서 §2).
      끝점 스냅으로 이어 붙이는 이 방식이 오히려 낫다 — 도면이 후렉시블을 여러
      LINE 으로 그렸어도 같은 결과가 나온다.
    """
    cl = _Cluster(eps)
    adj: dict = {}
    at: dict = {}
    for i, (a, b) in enumerate(segs):
        ka, kb = cl.key(a), cl.key(b)
        if ka == kb:                      # 길이 0 — 사슬에 안 넣는다
            continue
        at.setdefault(ka, a)
        at.setdefault(kb, b)
        adj.setdefault(ka, []).append((kb, i))
        adj.setdefault(kb, []).append((ka, i))

    seen_seg = set()
    chains, tangled = [], []
    for start, links in adj.items():
        if len(links) != 1:
            continue
        for _nxt, si in links:
            if si in seen_seg:
                continue
            path = [start]
            cur, prev_seg = start, None
            ok = True
            while True:
                nxts = [(n, s) for n, s in adj[cur] if s != prev_seg]
                if not nxts:
                    break
                if len(nxts) > 1:          # 갈래 — 사슬이 아니다
                    ok = False
                    break
                n, s = nxts[0]
                seen_seg.add(s)
                path.append(n)
                prev_seg, cur = s, n
                if len(path) > 400:        # 고리 방어
                    ok = False
                    break
            (chains if ok else tangled).append([at[k] for k in path])
    rest = [i for i in range(len(segs)) if i not in seen_seg]
    return chains, tangled, rest


def _eu(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def _l1(a, b):
    return abs(b[0] - a[0]) + abs(b[1] - a[1])


def _run(path, tot):
    """가닥의 총연장 · 현 · L1 합 · 현의 L1."""
    eu = sum(_eu(path[i], path[i + 1]) for i in range(len(path) - 1))
    l1 = sum(_l1(path[i], path[i + 1]) for i in range(len(path) - 1))
    tot["eu"] += eu
    tot["l1"] += l1
    tot["chord_eu"] += _eu(path[0], path[-1])
    tot["chord_l1"] += _l1(path[0], path[-1])
    return eu


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag", default="")
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--fold", action="store_true",
                    help="신축배관 레이어를 «접기» 로 지정하고 잰다")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    if not PLAN.is_file():
        print(f"표본 없음: {PLAN}")
        return 0
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    print(f"\n■ 신축배관 접기 계측 {args.tag}  (대명동 평면도 · K={args.k})")

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(PLAN, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        ad = rec.get("adopt") or {}
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": ad.get("conf_min")}})
        wait(c, sid)

        from routes.module_f.jobs import _sess
        ps = _sess(sid)["pick"]
        board = ps.board

        # ── ① 가닥 기하 (찍기 판 — 레이어가 살아 있는 마지막 지점)
        flex_keys = [k for k in board.mat if is_flex(k[0])]
        flex_segs = [s for k in flex_keys
                     for s in board.by_bundle.get(k, ())]
        other_segs = [s for k in board.mat if not is_flex(k[0])
                      for s in board.by_bundle.get(k, ())]
        print(f"  [가닥] 신축배관 묶음 {len(flex_keys)} {flex_keys}"
              f" · 선분 {len(flex_segs)}")
        chains, tangled, rest = strands(flex_segs)
        tot = {"eu": 0.0, "l1": 0.0, "chord_eu": 0.0, "chord_l1": 0.0}
        lens = [_run(p, tot) for p in chains]
        nseg = Counter(len(p) - 1 for p in chains)
        print(f"         가닥 {len(chains)} · 사슬 아님 {len(tangled)}"
              f" · 어디에도 안 붙은 선분 {len(rest)}")
        print(f"         가닥당 선분 {dict(sorted(nseg.items()))}")
        if lens:
            lens_s = sorted(lens)
            print(f"         가닥 총연장 {min(lens_s):.0f} ~ {max(lens_s):.0f} mm"
                  f" (중앙값 {lens_s[len(lens_s) // 2]:.0f})")
        seg_eu = sum(_eu(a, b) for a, b in flex_segs)
        seg_l1 = sum(_l1(a, b) for a, b in flex_segs)
        print(f"         묶음 전체 선분 합 — 유클리드 {seg_eu:,.0f} mm"
              f" · L1 {seg_l1:,.0f} mm")
        print(f"         ★총연장 유클리드 {tot['eu']:,.0f} mm"
              f" · 현 {tot['chord_eu']:,.0f} mm"
              f" ({(tot['chord_eu'] / tot['eu'] - 1) * 100:+.1f}%)"
              if tot["eu"] else "         ★총연장 0")
        print(f"         ★★L1(맨해튼) 합 {tot['l1']:,.0f} mm"
              f" · 현의 L1 {tot['chord_l1']:,.0f} mm"
              f" · 차 {tot['chord_l1'] - tot['l1']:+,.0f} mm")
        print("         ← `planar.py:875` 가 L1 로 재므로, 이 둘이 같으면"
              " 접어도 표 length 가 안 변한다")

        # ── ①-b 접기 지정 (있으면) — 실제 접기는 commit 뒤 손질판에서 돈다
        if args.fold:
            for ly in sorted({str(k[0]) for k in flex_keys}):
                r = c.post("/api/module-f/pick/fold-layer",
                           json={"sid": sid, "layer": ly, "on": True})
                print(f"  [접기 지정] {ly} → {r.get_json()}")

        # ── ② 표 지표 — 물닿음 · 45° 엘보 · 신축배관 표 총연장
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        es = _sess(sid)["edit"]
        # 끝점이 헤드/다른 배관에 닿는가 — 헤드 좌표는 손질판의 disk 가 권위다
        # (찍기 heads 는 bundle·지문 기록이라 좌표를 안 들고 있다).
        heads = [(float(d[0]), float(d[1]))
                 for d in (getattr(es.board, "disks", None) or ())]
        n_head_end = n_pipe_end = n_both = 0
        for p in chains:
            e0, e1 = p[0], p[-1]
            hit_h = [any(_eu(e, h) <= NEAR for h in heads) for e in (e0, e1)]
            hit_p = [any(_seg_d(e, a, b) <= NEAR for a, b in other_segs)
                     for e in (e0, e1)]
            n_head_end += 1 if any(hit_h) else 0
            n_pipe_end += 1 if any(hit_p) else 0
            n_both += 1 if (any(hit_h) and any(hit_p)) else 0
        print(f"         한 끝이 헤드에 {NEAR:.0f} mm 이내 {n_head_end}/{len(chains)}"
              f" · 한 끝이 다른 배관에 {n_pipe_end}/{len(chains)}"
              f" · 양쪽 다 {n_both}")

        from services.cad_import.design.restrict import attachable_heads
        wet = attachable_heads(es.convert_payload())
        print(f"  [표] 물닿음 헤드 {len(wet['wet'])} / 전개가 본 헤드"
              f" {wet['total']} · 떨어짐 {wet['dropped']}")

        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": args.k})
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"  ★표 확정 실패 — {j}")
            return 1
        d = _sess(sid)["design"]
        tbl = d["tables"]
        fit = Counter()
        for f in tbl.fittings:
            fit[str(f.get("type") or f.get("kind") or "?")] += int(
                f.get("count") or 1)
        print(f"  [표] 부속 {dict(fit)}")
        # ★부속표의 종류 이름은 `elbow-45` 다(`ELBOW_45` 는 KFP 쪽 이름).
        #   엉뚱한 키를 세면 늘 0 이 나와 「좋아졌다」로 읽힌다.
        n45 = fit.get("elbow-45", 0) + fit.get("ELBOW_45", 0)
        print(f"       ★45° 엘보 계상 {n45} (수작업 기준 1)")

        fold = _sess(sid).get("fold")
        if fold:
            print(f"  [접기] 가닥 {fold['strands']} → 접힌 구간"
                  f" {fold['folded_runs']} · 못 접음 {fold['skipped']}"
                  f" {fold['why']} · 총연장 {fold['total_mm']:,.0f} mm"
                  f" 중 {fold['declared_mm']:,.0f} 선언")
        got = d.get("got") or {}
        # ★표 전체 총연장 — §5 기준 3 의 실질이다. 접기로 «중간 노드가 사라져»
        #   배관 수는 줄지만 길이 합은 그대로여야 한다.
        all_m = sum(float(p.get("length") or 0) for p in tbl.pipes)
        dec_rows = [p for p in tbl.pipes if p.get("len_src")]
        dec_m = sum(float(p.get("length") or 0) for p in dec_rows)
        print(f"       ★표 전체 총연장 {all_m * 1000:,.0f} mm"
              f" · 배관 {len(tbl.pipes)}")
        # ★그 배관들의 «좌표 길이» 와 나란히 — 선언이 좌표를 이겼는지가 여기서
        #   보인다. 좌표가 그대로면 통로가 안 이어진 것이다(§2 «끊기는 자리»).
        nodes = {str(n.get("label")): (float(n.get("x", 0) or 0),
                                       float(n.get("y", 0) or 0))
                 for n in tbl.nodes}
        eu_m = l1_m = 0.0
        for p in dec_rows:
            a, b = nodes.get(str(p.get("in"))), nodes.get(str(p.get("out")))
            if a and b:
                eu_m += math.dist(a, b) / 1000.0
                l1_m += (abs(a[0] - b[0]) + abs(a[1] - b[1])) / 1000.0
        # ★표의 길이 규약은 **유클리드** 다(`compute_length` → `dist_3d`).
        #   선언이 실렸다면 표 > 좌표 유클리드 여야 한다 — 접은 간선의 좌표는
        #   «현» 이고 선언은 «접기 전 총연장» 이기 때문이다.
        print(f"       ★길이를 «선언» 에서 받은 배관 {len(dec_rows)}"
              f" · 표 {dec_m * 1000:,.0f} mm"
              f" vs 좌표 유클리드 {eu_m * 1000:,.0f}"
              f" ({(dec_m - eu_m) * 1000:+,.0f})"
              f" · (참고 L1 {l1_m * 1000:,.0f})")
        # 좌표↔length 불일치를 부류로 갈라 센다(§5 기준 8)
        _mismatch(tbl, got)

        # ── ③ 산출 지표 — 등각·비직각·X교차·겹침
        sdf = _emit(c, sid)
        if sdf:
            _angles(sdf)
    return 0


def _seg_d(p, a, b):
    """점–선분 거리."""
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    if L2 <= 1e-12:
        return _eu(p, a)
    t = max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2))
    return _eu(p, (a[0] + t * vx, a[1] + t * vy))


def _mismatch(tbl, got):
    """좌표 거리 ÷ 1000 이 표 length 와 5 % 넘게 다른 배관 — **부류로 갈라** 센다.

    ★한 덩어리로 세면 「불일치 77건」이 늘 경고처럼 보인다. 정상인 부류가 둘
      있다: «표고차뿐인 접속관»(평면 좌표는 2D 다)과 «신축배관 — 접었음»
      (좌표는 현이고 길이는 선언이다). 나머지만이 진짜 봐야 할 것이다.
    """
    nodes = {str(n.get("label")): (float(n.get("x", 0) or 0),
                                   float(n.get("y", 0) or 0),
                                   float(n.get("elevation", 0) or 0))
             for n in tbl.nodes}
    kinds = Counter()
    for p in tbl.pipes:
        a, b = nodes.get(str(p.get("in"))), nodes.get(str(p.get("out")))
        ln = float(p.get("length") or 0)
        if a is None or b is None or ln <= 0.01:
            continue
        # 표 좌표는 mm(§T3) — 표고만 m 다.
        d = math.dist(a[:2], b[:2]) / 1000.0
        if abs(d - ln) <= max(0.05 * ln, 0.05):
            continue
        if p.get("len_src"):
            kinds["신축배관 — 접었음 (정상)"] += 1
        elif abs(abs(a[2] - b[2]) - ln) <= 0.05:
            kinds["표고차뿐인 접속관 (정상)"] += 1
        else:
            kinds["★그 밖"] += 1
    print(f"       [좌표↔length] 5% 초과 {sum(kinds.values())}"
          f" / {len(tbl.pipes)} · {dict(kinds)}")


def _flex_table_len(tbl, got, flex_segs, es):
    """표의 배관 중 «신축배관에서 온» 것의 length 합 (mm).

    `edge_ref` 는 kfp 배관 id → board 노드쌍이다. board 좌표로 되짚어, 그 간선의
    중점이 신축배관 선분의 중점과 겹치면 신축배관에서 온 것으로 센다.
    """
    ref = (got or {}).get("edge_ref") or {}
    pts = list(getattr(es.board, "pts", None) or ())
    if not ref or not pts:
        return 0.0
    mids = {(round((a[0] + b[0]) / 2 / SNAP), round((a[1] + b[1]) / 2 / SNAP))
            for a, b in flex_segs}
    flex_pid = set()
    for pid, ij in ref.items():
        if not ij or len(ij) < 2:
            continue
        try:
            a, b = pts[int(ij[0])], pts[int(ij[1])]
        except (IndexError, TypeError, ValueError):
            continue
        k = (round((a[0] + b[0]) / 2 / SNAP), round((a[1] + b[1]) / 2 / SNAP))
        if k in mids:
            flex_pid.add(str(pid))
    tot = 0.0
    for p in tbl.pipes:
        if str(p.get("label")) in flex_pid:
            tot += float(p.get("length") or 0.0) * 1000.0
    print(f"       (신축배관에서 온 표 배관 {len(flex_pid)}개)")
    return tot


def _emit(c, sid):
    """산출을 한 번 내고 그 SDF 경로를 돌려준다."""
    r = c.post("/api/module-f/convert/run",
               json={"sid": sid, "dto": {}, "outputs": {"worst_sdf": True}})
    j = r.get_json() or {}
    wait(c, sid)
    from routes.module_f.jobs import _sess
    sess = _sess(sid)
    for key in ("worst_sdf_path", "design_sdf_path", "sdf_path"):
        p = sess.get(key)
        if p and Path(p).is_file():
            return Path(p)
    if not j.get("ok"):
        print(f"  [산출] 못 냈다 — {str(j)[:120]}")
    return None


def _angles(sdf: Path):
    """등각·평면 지표는 이미 있는 자를 그대로 쓴다(같은 판정이어야 비교가 선다)."""
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "_iso_probe", str(ROOT / "scripts" / "_iso_angle_probe.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    print(f"  [산출] {sdf.name}")
    sys.argv = ["_iso_angle_probe.py", str(sdf)]
    mod.main()


if __name__ == "__main__":
    raise SystemExit(main())
