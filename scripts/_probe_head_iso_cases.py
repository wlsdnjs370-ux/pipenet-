# -*- coding: utf-8 -*-
"""[헤드 4개] 아이소를 켤 때 이 헤드들에 무엇이 일어나는가 — 지목 사례 추적.

사용자가 표 값으로 **정확히 지목한** 네 헤드를 그대로 따라간다::

    118  (13050,  5000) 표고 0.3 · 접속관 P749 ↔ 101 · 부속 «(1) 인데 내용 없음»
    119  (32900, 14950) 표고 0.3 · 접속관 P772 ↔ 102 · 부속 «(1) 인데 내용 없음»
     59  (33000, 15800) 표고 0.3 · 접속관 P770 ↔  49 · 부속 elbow 1 (P770)
     50  (33700, 17900) 표고 0.3 · 접속관 P771 ↔  42 · 부속 elbow 1 (P771)

배관 번호가 P749~P772 라 K=30(배관 240) 범위가 아니다 — **전체 헤드**로 만든
표다. 그래서 여기서도 K 를 헤드 수까지 올려 같은 자리를 만든다.

재는 것:

    · 평면(iso 끔) 좌표와 등각(iso 켬) 좌표를 **나란히**
    · 접속관의 화면 각도 — 등각에서 헤드 스텁은 **수직**이어야 한다
    · 그 스텁과 화면에서 겹치는 배관 (있으면 무엇과)
    · 부속표에서 그 헤드·그 배관에 달린 것
    · 같은 도면의 «성한» 헤드와 견주기

    python scripts/_probe_head_iso_cases.py [--k 120]
"""
from __future__ import annotations

import argparse
import math
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
# 사용자가 지목한 네 헤드 — 라벨과 «표에 적힌» 좌표(mm).
CASES = {"118": (13050, 5000), "119": (32900, 14950),
         "59": (33000, 15800), "50": (33700, 17900)}


def wait(c, sid, limit=40000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _ang(dx, dy):
    """화면 각도(0~180). 수평 0 · 수직 90."""
    return math.degrees(math.atan2(dy, dx)) % 180.0


def _cross(p1, p2, p3, p4):
    def d(o, a, b):
        return ((a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]))
    d1, d2 = d(p3, p4, p1), d(p3, p4, p2)
    d3, d4 = d(p1, p2, p3), d(p1, p2, p4)
    return ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0))


def main() -> int:
    ap = argparse.ArgumentParser()
    # ★기본은 30 이다 — 사용자 표의 노즐 번호가 1·2·6·7 로 작다(K개만큼만
    #   매긴다). 배관 라벨이 P749~P772 로 큰 것은 노드정리·세로처리가 번호를
    #   새로 매기기 때문이지 배관이 772개라서가 아니다.
    ap.add_argument("--k", type=int, default=30)
    # 저장본을 «이어서 열기» 로 태운다 — 사용자가 매일 쓰는 도면이 따로 있다.
    ap.add_argument("--key", default="")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    if not PLAN.is_file():
        print(f"표본 없음: {PLAN}")
        return 0
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        if args.key:
            r = c.post("/api/module-f/reopen", json={"key": args.key})
            j0 = r.get_json() or {}
            if not j0.get("ok"):
                print(f"★이어서 열기 실패 — {str(j0)[:200]}")
                return 1
            sid = j0["sid"]
            if wait(c, sid).get("state") != "done":
                print("★이어서 열기 잡 실패")
                return 1
        else:
            with open(PLAN, "rb") as fh:
                r = c.post("/api/module-f/slot/open",
                           data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
                           content_type="multipart/form-data")
            sid = (r.get_json() or {})["sid"]
            wait(c, sid)
            c.post("/api/module-f/slot/read",
                   json={"sid": sid, "method": "manual"})
            rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
                   .get("recon") or {})
            c.post("/api/module-f/pick/adopt",
                   json={"sid": sid, "materials": True,
                         "heads": {"conf_min": (rec.get("adopt") or {})
                                   .get("conf_min")}})
            wait(c, sid)
        if not args.key:
            c.post("/api/module-f/pick/commit", json={"sid": sid})
            wait(c, sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": args.k})
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {j}")
            return 1

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        tbl = sess["design"]["tables"]
        cfg = dict(sess.get("design_settings") or {})
        print(f"\n■ 대명동 평면도 · K={args.k}"
              f" · 노드 {len(tbl.nodes)} · 배관 {len(tbl.pipes)}"
              f" · 노즐 {len(tbl.nozzles)}")

        from routes.module_f.api_design import _valve_label, _view_opts
        from services.cad_import.design.emit import display_tables

        opts = dict(_view_opts(cfg))
        opts["iso_ref_label"] = (_valve_label(tbl)
                                 if cfg.get("lift_ref") == "valve" else None)
        plan_opts = dict(opts)
        plan_opts["iso"] = False
        flat, _n1 = display_tables(tbl, **plan_opts)
        iso_opts = dict(opts)
        iso_opts["iso"] = True
        iso, stood = display_tables(tbl, **iso_opts)
        print(f"  iso 옵션: { {k: v for k, v in iso_opts.items() if k != 'iso_no_lift_labels'} }")
        print(f"  헤드 {stood.get('heads')} · 세움 {stood.get('vertical')}"
              f" · 관말 아님 {stood.get('not_terminal')}"
              f" · 스텁 폴백 {stood.get('stub')}"
              f" · 겹침 {stood.get('crossings')}")

        _report(tbl, flat.nodes, iso.nodes)
        _origin(sess, tbl)
    return 0


def _origin(sess, tbl):
    """★그 «붙어 있는 쌍» 이 도면의 어느 기호에서 왔는가.

    스프링클러 간격이 0.86 m 일 수는 없다. 두 쌍이 **정확히 같은 오프셋**
    (x 100 · y 850)이라는 것은 우연이 아니라 도면 기호의 구조 탓이다 —
    한 헤드를 두 번 세고 있다는 뜻이다. 손질판의 disk(원)로 되짚는다.
    """
    es = sess.get("edit")
    if es is None:
        return
    disks = list(getattr(es.board, "disks", None) or ())
    kinds = list(getattr(es.board, "disk_kinds", None) or ())
    print(f"\n  ── 손질판 헤드 원(disk) {len(disks)}")
    close = []
    for i in range(len(disks)):
        for j in range(i + 1, len(disks)):
            d = math.dist(disks[i][:2], disks[j][:2])
            if d < 1000.0:
                close.append((i, j, round(d, 1)))
    close.sort(key=lambda r: r[2])
    print(f"    ★1 m 안에 붙어 있는 원 쌍 {len(close)}")
    for i, j, d in close[:10]:
        a, b = disks[i], disks[j]
        ka = kinds[i] if i < len(kinds) else "?"
        kb = kinds[j] if j < len(kinds) else "?"
        print(f"        #{i}({a[0]:.0f},{a[1]:.0f}) r={a[2]:.1f} {ka}"
              f"  ↔  #{j}({b[0]:.0f},{b[1]:.0f}) r={b[2]:.1f} {kb}"
              f"  · {d} mm · Δ({b[0] - a[0]:+.0f},{b[1] - a[1]:+.0f})")
    # 반지름별로 세어 본다 — 두 기호가 «다른 크기» 면 서로 다른 블록이다.
    from collections import Counter
    rs = Counter(round(float(t[2]), 1) for t in disks)
    print(f"    원 반지름 분포 {dict(sorted(rs.items()))}")


def _report(tbl, flat, iso):
    at_t = {str(n["label"]): (float(n.get("x", 0) or 0),
                              float(n.get("y", 0) or 0),
                              float(n.get("elevation", 0) or 0))
            for n in tbl.nodes}
    at_f = {str(n["label"]): (float(n.get("x", 0) or 0),
                              float(n.get("y", 0) or 0)) for n in flat}
    at_i = {str(n["label"]): (float(n.get("x", 0) or 0),
                              float(n.get("y", 0) or 0)) for n in iso}
    noz = {str(z.get("in")) for z in tbl.nozzles}
    inc: dict = {}
    for p in tbl.pipes:
        inc.setdefault(str(p.get("in")), []).append(p)
        inc.setdefault(str(p.get("out")), []).append(p)
    fit_by_pipe: dict = {}
    fit_by_node: dict = {}
    for f in tbl.fittings:
        fit_by_pipe.setdefault(str(f.get("pipe")), []).append(f)
        # ★부속이 «어느 절점» 에 달렸는지도 본다 — 헤드 절점은 차수 1 이라
        #   엘보가 나올 수 없다. 거기 달려 있으면 그것부터가 이상하다.
        for key in ("in", "out"):
            if f.get(key):
                fit_by_node.setdefault(str(f.get(key)), []).append(f)

    # 라벨로 못 찾으면 «표에 적힌 좌표» 로 찾는다 — 라벨은 K 에 따라 달라진다.
    found = {}
    for lab, xy in CASES.items():
        hit = [n for n in at_t
               if abs(at_t[n][0] - xy[0]) < 2 and abs(at_t[n][1] - xy[1]) < 2]
        if hit:
            found[lab] = hit[0]
            continue
        if lab in at_t:
            # 라벨은 있는데 좌표가 다르다 — K 가 달라 번호가 옮겨간 것이다.
            print(f"    [{lab}] 라벨은 있으나 좌표가 다르다 —"
                  f" 표 ({at_t[lab][0]:.0f},{at_t[lab][1]:.0f})"
                  f" vs 지목 {xy}")
        found[lab] = None

    # ── 헤드 전수 — «표 값만으로 드러나는» 결함을 찾는다.
    #
    #   ★스텁 길이는 `|Δ표고| × units_per_m` 인데, **표고차가 0 이면** 그 정보가
    #     없다고 보고 고정 길이(캔버스 2.5 %)로 세운다. 표는 0.3 m 라는데 그림은
    #     그보다 긴 토막을 그린다 — `sdf_post` 주석이 스스로 경고한 상황이다.
    bad_de, bad_len, bad_deg = [], [], []
    for h in sorted(noz):
        ps = inc.get(h, [])
        if len(ps) != 1:
            bad_deg.append((h, len(ps)))
            continue
        p = ps[0]
        other = str(p.get("out")) if str(p.get("in")) == h else str(p.get("in"))
        de = at_t[h][2] - at_t.get(other, (0, 0, 0))[2]
        ln = float(p.get("length") or 0)
        if abs(de) <= 1e-9:
            bad_de.append((h, other, at_t[h][2], ln))
        elif abs(abs(de) - ln) > 0.005:
            bad_len.append((h, other, round(de, 3), ln))
    # ★상하향이 표고의 **부호**를 가른다(상향 +0.3 · 하향 −0.3). 등각에서는
    #   그 부호가 스텁이 서는 방향이 된다 — `de >= 0` 이면 위, 아니면 아래.
    from collections import Counter as _C
    ez = _C(round(at_t[h][2], 3) for h in noz)
    up = down = 0
    for h in noz:
        ps = inc.get(h, [])
        if len(ps) != 1:
            continue
        p = ps[0]
        o = str(p.get("out")) if str(p.get("in")) == h else str(p.get("in"))
        a_i, b_i = at_i.get(h), at_i.get(o)
        if a_i and b_i:
            up += 1 if a_i[1] > b_i[1] else 0
            down += 1 if a_i[1] < b_i[1] else 0
    print(f"\n  ── 헤드 표고 분포 {dict(sorted(ez.items()))}")
    print(f"    등각에서 스텁이 **위로** {up} · **아래로** {down}")

    print(f"\n  ── 헤드 전수 {len(noz)}")
    print(f"    ★부모와 표고가 같아 «고정 스텁» 으로 세워지는 헤드"
          f" {len(bad_de)} {[b[0] for b in bad_de][:12]}")
    for b in bad_de[:6]:
        print(f"        {b[0]} ↔ {b[1]} · 표고 {b[2]} · 접속관 길이 {b[3]} m"
              f"  ← 표는 {b[3]} m 인데 그림은 고정 길이로 선다")
    print(f"    ★표고차와 접속관 길이가 어긋나는 헤드 {len(bad_len)}"
          f" {[(b[0], b[2], b[3]) for b in bad_len][:8]}")
    print(f"    ★붙은 배관이 1개가 아닌 헤드 {len(bad_deg)} {bad_deg[:8]}")

    # ★서로 붙어 있는 헤드 쌍 — 스프링클러 간격은 보통 2~3 m 다. 수백 mm 안에
    #   둘이 있으면 같은 헤드를 두 번 찍었거나 도면에 기호가 겹쳐 있는 것이고,
    #   등각에서 두 스텁이 같은 자리에 서서 하나로 뭉쳐 보인다.
    close = []
    hs = sorted(noz)
    for i in range(len(hs)):
        for j in range(i + 1, len(hs)):
            a, b = at_t.get(hs[i]), at_t.get(hs[j])
            if not a or not b:
                continue
            d = math.dist(a[:2], b[:2])
            if d < 1000.0:
                close.append((hs[i], hs[j], round(d, 1)))
    close.sort(key=lambda r: r[2])
    print(f"    ★1 m 안에 붙어 있는 헤드 쌍 {len(close)}")
    for r in close[:8]:
        pa = inc.get(r[0], [{}])[0]
        pb = inc.get(r[1], [{}])[0]
        oa = (str(pa.get("out")) if str(pa.get("in")) == r[0]
              else str(pa.get("in")))
        ob = (str(pb.get("out")) if str(pb.get("in")) == r[1]
              else str(pb.get("in")))
        da = math.dist(at_i.get(r[0], (0.0, 0.0)), at_i.get(r[1], (0.0, 0.0)))
        print(f"        {r[0]}({at_t[r[0]][0]:.0f},{at_t[r[0]][1]:.0f})"
              f" ↔ {r[1]}({at_t[r[1]][0]:.0f},{at_t[r[1]][1]:.0f})"
              f" · 평면 {r[2]} mm · 부모 {oa} / {ob}"
              f" · **등각 화면 거리 {da:.1f}**")

    print("\n  ── 지목한 네 헤드")
    stubs = []
    for lab, real in found.items():
        if real is None:
            print(f"    [{lab}] ★이 K 에서는 그 좌표의 절점이 없다"
                  f" (표 좌표 {CASES[lab]})")
            continue
        tag = f"{lab}" + ("" if real == lab else f" → 지금 라벨 {real}")
        e = at_t[real][2]
        pipes = inc.get(real, [])
        here = fit_by_node.get(real, [])
        print(f"    [{tag}] 표 ({at_t[real][0]:.0f},{at_t[real][1]:.0f})"
              f" 표고 {e} · 노즐 {'있음' if real in noz else '★없음'}"
              f" · 붙은 배관 {len(pipes)}"
              f" · 이 절점의 부속 {[(f.get('type'), f.get('count'), f.get('pipe')) for f in here]}")
        for p in pipes:
            other = (str(p.get("out")) if str(p.get("in")) == real
                     else str(p.get("in")))
            fits = fit_by_pipe.get(str(p.get("label")), [])
            fl = " · ".join(f"{f.get('type')}×{f.get('count')}" for f in fits)
            a_f, b_f = at_f.get(real), at_f.get(other)
            a_i, b_i = at_i.get(real), at_i.get(other)
            print(f"        {p.get('label')} ↔ {other}"
                  f" · 길이 {p.get('length')} m · 호칭경 {p.get('dia')}"
                  + (f" · 부속 {fl}" if fl else " · 부속 없음"))
            if a_f and b_f:
                d_f = math.dist(a_f, b_f)
                print(f"          평면: ({a_f[0]:.0f},{a_f[1]:.0f})–"
                      f"({b_f[0]:.0f},{b_f[1]:.0f}) · 화면거리 {d_f:.1f}"
                      + (f" · 각 {_ang(b_f[0] - a_f[0], b_f[1] - a_f[1]):.1f}°"
                         if d_f > 1e-9 else " · **길이 0**"))
            if a_i and b_i:
                d_i = math.dist(a_i, b_i)
                ang = (_ang(b_i[0] - a_i[0], b_i[1] - a_i[1])
                       if d_i > 1e-9 else None)
                mark = ""
                if ang is not None and abs(ang - 90.0) > 1.5:
                    mark = "  ←★수직이 아니다"
                print(f"          등각: ({a_i[0]:.0f},{a_i[1]:.0f})–"
                      f"({b_i[0]:.0f},{b_i[1]:.0f}) · 화면거리 {d_i:.1f}"
                      + (f" · 각 {ang:.1f}°{mark}" if ang is not None
                         else " · **길이 0**"))
                if real in noz:
                    stubs.append((tag, p.get("label"), a_i, b_i))

    # 그 스텁이 화면에서 무엇과 겹치는가
    print("\n  ── 등각에서 그 스텁과 겹치는 배관")
    segs = []
    for p in tbl.pipes:
        a, b = at_i.get(str(p.get("in"))), at_i.get(str(p.get("out")))
        if a and b and math.dist(a, b) > 1e-9:
            segs.append((str(p.get("label")), a, b))
    for tag, plabel, a_i, b_i in stubs:
        hits = [lab for lab, a, b in segs
                if lab != plabel and _cross(a_i, b_i, a, b)]
        print(f"    [{tag}] {plabel} — 겹침 {len(hits)}"
              + (f" {hits[:6]}" if hits else " (없음)"))


if __name__ == "__main__":
    raise SystemExit(main())
