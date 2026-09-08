# -*- coding: utf-8 -*-
"""[위상 ③] 설계 배치가 **도면 자리와 같은 자리**인가 — 잔차로 잰다.

앞선 둘로 이만큼 좁혀졌다:

    · 자료의 위상(갈림·끝·사슬·고리)은 평면과 설계가 **같다**.
    · 그런데 아이소 그림에는 없던 교차가 15건 생기고, 절점이 서로 겹친
      쌍이 195개다. 헤드 스텁 방향을 바꿔 봐도 15 → 9 로 줄 뿐이다.

그러면 남는 의심은 하나다 — **설계 절점의 «자리» 자체가 도면과 다르다**.
여기서는 그것을 직접 잰다:

    ① 설계 절점 ↔ board 절점 대응을 `edge_ref` 로 세운다.
    ② 화면이 밑그림에 쓰는 그 변환(`view.underlay` 의 k·tx·ty)으로 board 를
       설계 좌표로 옮긴다 — «별도 수학» 을 쓰지 않는다.
    ③ 옮긴 자리와 실제 설계 절점 자리의 **잔차**를 잰다.

잔차가 캔버스의 몇 %씩 난다면 그림이 뒤틀린 것이고, 그것이 곧 사람이 보는
«위상이 깨졌다» 다.

    python scripts/_probe_iso_fidelity.py [도면.dxf]
"""
from __future__ import annotations

import math
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def _cross2(p1, p2, p3, p4):
    def d(o, a, b):
        return ((a[0] - o[0]) * (b[1] - o[1])
                - (a[1] - o[1]) * (b[0] - o[0]))
    d1, d2 = d(p3, p4, p1), d(p3, p4, p2)
    d3, d4 = d(p1, p2, p3), d(p1, p2, p4)
    return ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0))


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEF
    if not dxf.is_file():
        print(f"표본 없음: {dxf}")
        return 0
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
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
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        groups = sorted((g.get("segs") or [] for g in st["body_groups"]),
                        key=len, reverse=True)
        seg = groups[0]
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": seg[0], "y": seg[1]})
        wait(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 30})
        c.post("/api/module-f/design/build", json={"sid": sid})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {j}")
            return 1

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        board = sess["edit"].board
        got = sess["design"]["got"]
        ref = got.get("edge_ref") or {}

        # 평면(iso=0) 좌표로 잰다 — 아이소 셰어는 선형이라 뒤틀림을 못 만든다.
        pv = c.get(f"/api/module-f/design/preview?sid={sid}&iso=0").get_json()
        view = pv.get("view") or {}
        under = view.get("underlay") or {}
        at = {str(n["label"]): (float(n["x"]), float(n["y"]))
              for n in (view.get("nodes") or [])}
        pipes = {str(p.get("label")): (str(p.get("a")), str(p.get("b")))
                 for p in (view.get("pipes") or [])}

        print(f"\n■ {dxf.name}")
        if not under:
            print("  밑그림 변환이 없다 — 이 도면으로는 못 잰다.")
            return 1
        k, tx, ty = float(under["k"]), float(under["tx"]), float(under["ty"])
        print(f"  변환 k={k:.6g} · tx={tx:.6g} · ty={ty:.6g}")

        def to_design(i):
            x, y = board.pts[i]
            return (k * float(x) + tx, k * float(y) + ty)

        # ── 대응 — 배관 하나가 board 간선 하나를 가리킨다. 두 끝을 «가까운 쪽»
        #    으로 짝지어 절점 대응을 모은다.
        pair = {}
        for pid, e in ref.items():
            ab = pipes.get(str(pid))
            if not ab:
                continue
            try:
                i, j2 = int(e[0]), int(e[1])
            except (TypeError, ValueError, IndexError):
                continue
            a, b = ab
            if a not in at or b not in at:
                continue
            pi, pj = to_design(i), to_design(j2)
            da = (math.dist(at[a], pi) + math.dist(at[b], pj))
            db = (math.dist(at[a], pj) + math.dist(at[b], pi))
            if da <= db:
                pair.setdefault(a, []).append(pi)
                pair.setdefault(b, []).append(pj)
            else:
                pair.setdefault(a, []).append(pj)
                pair.setdefault(b, []).append(pi)

        xs = [p[0] for p in at.values()]
        ys = [p[1] for p in at.values()]
        span = max(max(xs) - min(xs), max(ys) - min(ys)) or 1.0
        res = []
        for lab, cands in pair.items():
            # 여러 배관이 같은 절점을 가리키면 «가장 가까운» 후보로 본다.
            d = min(math.dist(at[lab], p) for p in cands)
            res.append((d, lab))
        res.sort()
        if not res:
            print("  대응을 못 세웠다.")
            return 1
        med = res[len(res) // 2][0]
        p90 = res[int(len(res) * 0.9)][0]
        print(f"  대응 절점 {len(res)} / {len(at)} · 캔버스 폭 {span:.0f}")
        print(f"  잔차: 중앙 {med:.2f} ({med / span * 100:.3f}%)"
              f" · 90% {p90:.2f} ({p90 / span * 100:.3f}%)"
              f" · 최대 {res[-1][0]:.2f} ({res[-1][0] / span * 100:.2f}%)")
        big = [(d, lab) for d, lab in res if d > span * 0.01]
        print(f"  캔버스의 1% 넘게 어긋난 절점 {len(big)}"
              f"{' — ' + str([lab for _d, lab in big[-8:]]) if big else ''}")

        # 겹친 절점 — 정확히 같은 자리에 몇이 서 있나
        exact = {}
        for lab, (x, y) in at.items():
            exact.setdefault((round(x, 3), round(y, 3)), []).append(lab)
        stacks = sorted((v for v in exact.values() if len(v) > 1),
                        key=len, reverse=True)
        print(f"  똑같은 자리에 선 절점 무리 {len(stacks)}"
              f" · 가장 큰 무리 {len(stacks[0]) if stacks else 0}"
              f" · 그런 절점 합 {sum(len(v) for v in stacks)}")
        for v in stacks[:5]:
            print(f"      {v}")

        # ── 표가 말하는 길이 ↔ 그림이 보이는 길이. 두 값이 다르면 그림이
        #    표를 배신한 것이다(길이 1.2m 인 배관이 점 하나로 그려진다).
        rows = (pv.get("tables") or {}).get("pipes") or []
        len_of = {str(r.get("label")): float(r.get("length") or 0.0)
                  for r in rows}
        bad = []
        for pid, (a, b) in pipes.items():
            if a not in at or b not in at:
                continue
            drawn = math.dist(at[a], at[b])
            ln = len_of.get(pid, 0.0)
            if drawn <= 0.5 and ln > 0.05:
                bad.append((pid, ln))
        print(f"  표는 길이가 있는데 그림에서는 점인 배관 {len(bad)}"
              f" · 그 길이 합 {sum(v for _p, v in bad):.2f} m")
        for pid, ln in bad[:6]:
            print(f"      {pid} — 표 {ln:.3f} m · 그림 0")
        # ── 표고가 어디서 사라졌나 — kfp 원본까지 되짚는다.
        kfp = got.get("kfp") or {}
        meta = kfp.get("nodes_meta_runtime") or {}
        zs = []
        for nid, m in meta.items():
            cc = (m or {}).get("coords") or ()
            zs.append(round(float(cc[2]), 3) if len(cc) > 2 else None)
        zset = sorted({z for z in zs if z is not None})
        print(f"  kfp 절점 {len(meta)} · coords 에 z 가 있는 것"
              f" {sum(1 for z in zs if z is not None)}"
              f" · z 종류 {len(zset)}: {zset[:8]}")
        pd = kfp.get("pipe_data") or {}
        vk = {}
        for pid, pr in list(pd.items())[:100000]:
            for key in ("kind", "type", "vertical", "role"):
                if key in (pr or {}):
                    vk[key] = vk.get(key, 0) + 1
        print(f"  kfp 배관 {len(pd)} · 배관 dict 에 있는 갈래 키 {vk}")
        # ── ★표는 표고를 갖고 있다. 미리보기 JSON 이 그것을 안 실을 뿐이다
        #    (nodes 는 label·x·y 만 보낸다) — 앞선 조사에서 «표고가 전부 0»
        #    으로 읽었던 것은 그 때문이다. 표에서 다시 읽는다.
        trows = (pv.get("tables") or {}).get("nodes") or []
        ev = {str(r.get("label")): float(r.get("elevation") or 0.0)
              for r in trows}
        es = sorted({round(v, 3) for v in ev.values()})
        print(f"  표의 표고 종류 {len(es)}: {es[:10]} · 폭"
              f" {max(es) - min(es) if es else 0:.3f} m")

        # ── 아이소 lift 를 «평면과 같은 자» 로 두면 교차가 어떻게 되나.
        #    지금 식: lift = 대각선·0.5·배율 / 표고폭  → 표고폭이 얼마든
        #    **화면 절반**으로 늘린다. 단위세대처럼 표고폭이 3 m 인 도면에서는
        #    0.3 m 스텁이 캔버스 몇 백 단위로 튀어, 없던 교차가 생긴다.
        norm = 1.0
        u = view.get("underlay") or {}
        if u:
            norm = float(u["k"])          # 캔버스 단위 / board mm
        iso_pv = c.get(f"/api/module-f/design/preview?sid={sid}"
                       f"&iso=1").get_json()
        stood = iso_pv.get("stood") or {}
        print(f"  지금 lift {float(stood.get('lift') or 0):.1f}"
              f" · 평면과 같은 자라면 {norm * 1000:.1f}"
              f" (배 {float(stood.get('lift') or 0) / max(1e-9, norm * 1000):.0f})")

        def cross_count(pos, segs_ab):
            n = 0
            for i in range(len(segs_ab)):
                a1, b1 = segs_ab[i]
                for j in range(i + 1, len(segs_ab)):
                    a2, b2 = segs_ab[j]
                    if {a1, b1} & {a2, b2}:
                        continue
                    if _cross2(pos[a1], pos[b1], pos[a2], pos[b2]):
                        n += 1
            return n

        segs_ab = [(a, b) for (a, b) in pipes.values()
                   if a in at and b in at]
        C30, S30 = math.cos(math.radians(30)), math.sin(math.radians(30))
        # 표의 «정규화 좌표» 는 iso=0 미리보기가 그대로 준다(at).
        for name, lift in (("지금 (화면 절반)", float(stood.get("lift") or 0)),
                           ("평면과 같은 자", norm * 1000),
                           ("그 3배", norm * 3000),
                           ("lift 0", 0.0)):
            pos = {}
            e_ref = float(stood.get("e_ref") or 0.0)
            for lab, (x, y) in at.items():
                e = ev.get(lab, 0.0)
                pos[lab] = ((x - y) * C30, (x + y) * S30 + (e - e_ref) * lift)
            print(f"      {name:14} lift {lift:9.1f} · 교차"
                  f" {cross_count(pos, segs_ab)}")

        n_head = sum(1 for n in (view.get("nodes") or []) if n.get("head"))
        print(f"  헤드 절점 {n_head}")

        # ── 산출물이 무엇이 바뀌나 — «그림» 만 바뀌고 «수» 는 그대로여야 한다.
        import copy as _copy
        from services.cad_import.design.emit import _head_attachments
        from services.cad_import.design.sdf_post import (
            bake_isometric, node_norm_params, normalize_node_coords)
        tbl = sess["design"]["tables"]
        outs = {}
        for tag, per_m in (("옛 식(화면 절반)", None), ("평면과 같은 자", True)):
            v2 = _copy.copy(tbl)
            v2.nodes = [dict(n) for n in tbl.nodes]
            cx, cy, sc = node_norm_params(v2, canvas_units=3000.0)
            normalize_node_coords(v2, canvas_units=3000.0)
            hn, hp, _loose = _head_attachments(v2)
            bake_isometric(v2, iso_z_scale=1.0, head_nodes=hn, head_parent=hp,
                           units_per_m=(sc * 1000.0 if per_m else None))
            outs[tag] = v2
        a, b = outs["옛 식(화면 절반)"], outs["평면과 같은 자"]
        moved = sum(1 for r1, r2 in zip(a.nodes, b.nodes)
                    if abs(r1["x"] - r2["x"]) > 1e-9
                    or abs(r1["y"] - r2["y"]) > 1e-9)
        ez = sum(1 for r1, r2 in zip(a.nodes, b.nodes)
                 if abs(float(r1.get("elevation") or 0)
                        - float(r2.get("elevation") or 0)) > 1e-9)
        print(f"  산출 비교: 자리가 바뀐 절점 {moved}/{len(a.nodes)}"
              f" · 표고가 바뀐 절점 {ez}"
              f" · 배관행 동일 {a.pipes == b.pipes}"
              f" · 노즐행 동일 {a.nozzles == b.nozzles}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
