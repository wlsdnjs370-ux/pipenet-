# -*- coding: utf-8 -*-
"""[BLOCKED §17 재검] board → 설계 표가 정말 «전역 변환» 이 아닌가.

§17 은 「전역 변환이 없다」고 결론지었다. 근거는 최소제곱 아핀 잔차
(중앙값 169.9 · 최대 2,003.8 = 도면 한 변의 9.3%)였고, 대응은 `edge_ref`
(설계 배관 ↔ board 간선)로 잡았다.

★그런데 §30 에서 같은 종류의 표(`node_ref`)가 **노드정리 전의 자리**를 가리켜
  30개 중 12개만 맞는다는 것이 드러났다. `edge_ref` 도 같은 함정을 밟는다:
  노드정리가 일직선 중간 노드를 지우면 설계 배관 하나가 board 간선 **여러 개**
  를 아우르는데, edge_ref 는 그중 «첫 간선» 만 기억한다. 그 짝을 그대로 쓰면
  설계 절점 X 를 사슬 한복판의 board 절점과 맺게 되고, 잔차는 사슬 길이만큼
  나온다 — **기하가 아니라 짝짓기가 만든 잔차** 다.

그래서 잔차를 다시 잰다. 이번에는 대응을 «믿을 수 있는 것» 으로만 고른다:
  · 헤드 — board 헤드 부착점 ↔ 설계 헤드 노드 (§30 에서 6~19mm 로 검증)
  · 접속점 — board 급수 노드 ↔ io_node=Input

    python scripts/_probe_f10e_affine2.py
"""
from __future__ import annotations

import math
import os
import statistics
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

KEY = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")

from routes.module_f.common import _boot                           # noqa: E402
_boot()
from services.cad_import.design.emit import display_tables         # noqa: E402
from services.cad_import.design.restrict import select_and_expand  # noqa: E402
from services.cad_import.design.tables import build_design_tables  # noqa: E402
from services.cad_import.design.worst import worst_k_heads         # noqa: E402
from services.cad_import.edit.session import EditSession           # noqa: E402


def fit_similarity(src, dst):
    """회전 없는 닮음(등방 배율 + 평행이동) 최소제곱.

    ★아핀(6자유도)이 아니라 닮음(3자유도)으로 맞춘다. board → kfp 는 평행이동
      +1/1000, kfp → 표시는 평행이동 + 등방 배율이므로, 두 좌표계가 정말
      이어져 있다면 **닮음으로 충분해야** 한다. 아핀으로 맞추면 자유도가 남아
      어긋남을 흡수해 버려 「이어져 있다」를 과대평가한다.
    """
    n = len(src)
    sx = sum(p[0] for p in src) / n
    sy = sum(p[1] for p in src) / n
    dx = sum(p[0] for p in dst) / n
    dy = sum(p[1] for p in dst) / n
    num = sum((s[0] - sx) * (d[0] - dx) + (s[1] - sy) * (d[1] - dy)
              for s, d in zip(src, dst))
    den = sum((s[0] - sx) ** 2 + (s[1] - sy) ** 2 for s in src)
    if den < 1e-12:
        return None
    k = num / den
    return k, dx - k * sx, dy - k * sy


def fit_affine(src, dst):
    """최소제곱 아핀(6자유도) — 정규방정식 3×3 을 직접 푼다."""
    n = len(src)
    sxx = sum(p[0] * p[0] for p in src)
    sxy = sum(p[0] * p[1] for p in src)
    syy = sum(p[1] * p[1] for p in src)
    sx = sum(p[0] for p in src)
    sy = sum(p[1] for p in src)
    m = [[sxx, sxy, sx], [sxy, syy, sy], [sx, sy, float(n)]]

    def solve(rhs):
        a = [row[:] + [r] for row, r in zip(m, rhs)]
        for i in range(3):
            p = max(range(i, 3), key=lambda r_: abs(a[r_][i]))
            if abs(a[p][i]) < 1e-12:
                return None
            a[i], a[p] = a[p], a[i]
            for r_ in range(3):
                if r_ == i:
                    continue
                f = a[r_][i] / a[i][i]
                for c_ in range(i, 4):
                    a[r_][c_] -= f * a[i][c_]
        return [a[i][3] / a[i][i] for i in range(3)]

    ax = solve([sum(s[0] * d[0] for s, d in zip(src, dst)),
                sum(s[1] * d[0] for s, d in zip(src, dst)),
                sum(d[0] for d in dst)])
    ay = solve([sum(s[0] * d[1] for s, d in zip(src, dst)),
                sum(s[1] * d[1] for s, d in zip(src, dst)),
                sum(d[1] for d in dst)])
    return None if ax is None or ay is None else (ax, ay)


def fit_affine3(src, dst):
    """(x,y,z) → (X,Y) 최소제곱 아핀. 4x4 정규방정식."""
    rows = [(p[0], p[1], p[2], 1.0) for p in src]
    m = [[sum(r[i] * r[j] for r in rows) for j in range(4)] for i in range(4)]

    def solve(rhs):
        a = [row[:] + [r] for row, r in zip(m, rhs)]
        for i in range(4):
            pv = max(range(i, 4), key=lambda t: abs(a[t][i]))
            if abs(a[pv][i]) < 1e-12:
                return None
            a[i], a[pv] = a[pv], a[i]
            for t in range(4):
                if t == i:
                    continue
                f = a[t][i] / a[i][i]
                for c in range(i, 5):
                    a[t][c] -= f * a[i][c]
        return [a[i][4] / a[i][i] for i in range(4)]

    ax = solve([sum(r[i] * d[0] for r, d in zip(rows, dst)) for i in range(4)])
    ay = solve([sum(r[i] * d[1] for r, d in zip(rows, dst)) for i in range(4)])
    return None if ax is None or ay is None else (ax, ay)


def main() -> int:
    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    b = es.board
    if not b.sources:
        print("접속점이 없는 저장본이다 — 잴 수 없다")
        return 3
    w = worst_k_heads(b.pts, b.edges, b._head_nodes(), b.sources, k=30)
    got = select_and_expand(es.convert_payload(), b, k=30)
    if not got.get("ok"):
        print("전개 실패:", got.get("error"))
        return 3
    tbl = build_design_tables(got["kfp"], w, got["edge_ref"], [],
                              board_pts=b.pts,
                              tree_loads=got.get("tree_loads"),
                              origin_mm=got.get("origin_mm"))
    view, _stood = display_tables(tbl, iso=False, canvas_units=3000.0)
    origin = got["origin_mm"]
    at = {str(n["label"]): n for n in view.nodes}
    kfp_nodes = got["kfp"]["nodes_meta_runtime"]

    def to_kfp(mx, my):
        return ((mx - origin[0]) / 1000.0 + 1.0,
                (my - origin[1]) / 1000.0 + 1.0)

    # 표 라벨 ↔ kfp 노드 — 표는 kfp 좌표를 mm 로 올려 실었다(§T3).
    lab_of = {}
    for n in tbl.nodes:
        lab_of[(round(n["x"] / 1000.0, 3), round(n["y"] / 1000.0, 3))] = \
            str(n["label"])

    src, dst, tags = [], [], []
    # ── ① 헤드 — §30 에서 검증된 대응
    hn = b._head_nodes()
    for hi in w["heads"]:
        reach = [n for n in hn[hi] if n < len(b.pts)]
        if not reach:
            continue
        for bn in reach:
            tx, ty = to_kfp(b.pts[bn][0], b.pts[bn][1])
            near = [(math.hypot(float((m.get("coords") or [0, 0])[0]) - tx,
                                float((m.get("coords") or [0, 0])[1]) - ty), nid)
                    for nid, m in kfp_nodes.items()
                    if str(m.get("type_id")) == "head"]
            near.sort()
            if not near or near[0][0] > 0.10:
                continue
            c = kfp_nodes[near[0][1]]["coords"]
            lab = lab_of.get((round(float(c[0]), 3), round(float(c[1]), 3)))
            if lab is None or lab not in at:
                continue
            src.append((float(b.pts[bn][0]), float(b.pts[bn][1])))
            dst.append((float(at[lab]["x"]), float(at[lab]["y"])))
            tags.append(f"헤드{hi}")
            break

    # ── ② 접속점
    root = next((str(n["label"]) for n in tbl.nodes
                 if str(n.get("io_node")) == "Input"), None)
    if root and root in at:
        bn = b.sources[0]
        src.append((float(b.pts[bn][0]), float(b.pts[bn][1])))
        dst.append((float(at[root]["x"]), float(at[root]["y"])))
        tags.append("접속점")

    print(f"\n■ 믿을 수 있는 대응만 — {len(src)}쌍 "
          f"(헤드 {sum(1 for t in tags if t.startswith('헤드'))} · "
          f"접속점 {sum(1 for t in tags if t == '접속점')})")
    if len(src) < 6:
        print("  ★대응이 너무 적다 — 판단 보류")
        return 3

    fit = fit_similarity(src, dst)
    if fit is None:
        print("  ★퇴화 — 맞출 수 없다")
        return 3
    k, tx0, ty0 = fit
    errs = sorted(math.hypot(k * s[0] + tx0 - d[0], k * s[1] + ty0 - d[1])
                  for s, d in zip(src, dst))
    span = max(max(d[0] for d in dst) - min(d[0] for d in dst),
               max(d[1] for d in dst) - min(d[1] for d in dst))
    print("\n■ 닮음(3자유도) 잔차 — 설계 표시 좌표 단위")
    print(f"    중앙값 {statistics.median(errs):,.2f} · "
          f"p90 {errs[int(len(errs) * .9)]:,.2f} · 최대 {errs[-1]:,.2f}")
    print(f"    표시 도면 한 변 {span:,.0f} · 최대 잔차 "
          f"{errs[-1] / max(1e-9, span) * 100:.3f}%")
    print(f"    배율 k = {k:.9f}  (board mm → 표시 단위)")

    ok = errs[-1] <= span * 0.005
    print(f"\n  {'전역 변환이 있다 — 밑그림을 어긋남 없이 깔 수 있다' if ok else '★전역 변환이 없다'}")

    # ── ③ 같은 빌드에서 «§17 의 짝짓기» 로도 재 본다.
    #
    # ★도면이 다르면 비교가 안 된다(§17 은 대명동, 여기는 다른 도면일 수 있다).
    #   같은 표·같은 좌표에 짝짓기만 갈아 끼워야 「기하냐 짝짓기냐」가 갈린다.
    print("\n■ 대조 — 같은 빌드, 짝짓기만 §17 방식(edge_ref)으로")
    from collections import Counter
    ends = {str(p.get("label")): (str(p.get("in")), str(p.get("out")))
            for p in tbl.pipes}
    pairs: dict = {}
    for pid, e in (got.get("edge_ref") or {}).items():
        lab = ends.get(str(pid))
        if not lab:
            continue
        try:
            i, j = int(e[0]), int(e[1])
        except (TypeError, ValueError, IndexError):
            continue
        if not (0 <= i < len(b.pts) and 0 <= j < len(b.pts)):
            continue
        if lab[0] not in at or lab[1] not in at:
            continue
        pairs.setdefault(lab[0], []).append(i)
        pairs.setdefault(lab[1], []).append(j)
    esrc, edst = [], []
    for lab, idxs in pairs.items():
        bi = Counter(idxs).most_common(1)[0][0]
        esrc.append((float(b.pts[bi][0]), float(b.pts[bi][1])))
        edst.append((float(at[lab]["x"]), float(at[lab]["y"])))
    print(f"    짝지어진 노드 {len(esrc)}개")
    if len(esrc) >= 6:
        f2 = fit_similarity(esrc, edst)
        e2 = sorted(math.hypot(f2[0] * s[0] + f2[1] - d[0],
                               f2[0] * s[1] + f2[2] - d[1])
                    for s, d in zip(esrc, edst))
        print(f"    닮음 잔차 중앙값 {statistics.median(e2):,.2f} · "
              f"최대 {e2[-1]:,.2f} · 한 변의 {e2[-1] / max(1e-9, span) * 100:.3f}%")
        print(f"\n  ★같은 좌표·같은 표인데 짝짓기만 바꿔 잔차가 "
              f"{e2[-1] / max(errs[-1], 1e-9):.0f}배 달라진다.")
        print("    → §17 의 «전역 변환 없음» 은 기하가 아니라 **짝짓기**가 만든 값이다.")

    # ── ④ §17 의 남은 두 근거도 확인한다.
    print("\n■ §17 근거 ② — K 를 키우면 «다른 공간» 인가")
    got2 = select_and_expand(es.convert_payload(), b, k=10)
    if got2.get("ok"):
        w2 = worst_k_heads(b.pts, b.edges, b._head_nodes(), b.sources, k=10)
        t2 = build_design_tables(got2["kfp"], w2, got2["edge_ref"], [],
                                 board_pts=b.pts,
                                 tree_loads=got2.get("tree_loads"),
                                 origin_mm=got2.get("origin_mm"))
        v2, _ = display_tables(t2, iso=False, canvas_units=3000.0)
        at2 = {str(n["label"]): n for n in v2.nodes}
        lab2 = {}
        for n in t2.nodes:
            lab2[(round(n["x"] / 1000.0, 3), round(n["y"] / 1000.0, 3))] = \
                str(n["label"])
        o2 = got2["origin_mm"]
        k2n = got2["kfp"]["nodes_meta_runtime"]
        s2, d2 = [], []
        for hi in w2["heads"]:
            for bn in [n for n in b._head_nodes()[hi] if n < len(b.pts)]:
                tx = (b.pts[bn][0] - o2[0]) / 1000.0 + 1.0
                ty = (b.pts[bn][1] - o2[1]) / 1000.0 + 1.0
                near = [(math.hypot(float((m.get("coords") or [0, 0])[0]) - tx,
                                    float((m.get("coords") or [0, 0])[1]) - ty),
                         nid)
                        for nid, m in k2n.items()
                        if str(m.get("type_id")) == "head"]
                near.sort()
                if not near or near[0][0] > 0.10:
                    continue
                c = k2n[near[0][1]]["coords"]
                lb = lab2.get((round(float(c[0]), 3), round(float(c[1]), 3)))
                if lb and lb in at2:
                    s2.append((float(b.pts[bn][0]), float(b.pts[bn][1])))
                    d2.append((float(at2[lb]["x"]), float(at2[lb]["y"])))
                break
        if len(s2) >= 6:
            f3 = fit_similarity(s2, d2)
            e3 = sorted(math.hypot(f3[0] * s[0] + f3[1] - d[0],
                                   f3[0] * s[1] + f3[2] - d[1])
                        for s, d in zip(s2, d2))
            print(f"    K=10 에서도 닮음 잔차 최대 {e3[-1]:,.2f} · 배율 {f3[0]:.9f}")
            print(f"    K=30 배율 {k:.9f} — 배율은 K 마다 다르다(정규화가 bbox "
                  "기준이므로 당연하다).")
            print("    ★그런데 밑그림은 «지금 그 표» 의 변환으로 깔면 된다 — "
                  "다른 K 표를 끌어올 이유가 없다.")

    print("\n■ §17 근거 ③ — 아이소 보기에서도 되는가")
    vi, stood = display_tables(tbl, iso=True, canvas_units=3000.0)
    ati = {str(n["label"]): n for n in vi.nodes}
    si, di, zs = [], [], []
    for (s, d, t) in zip(src, dst, tags):
        # 평면에서 쓴 라벨을 그대로 아이소 좌표로 바꿔 본다.
        lab = next((L for L, n in at.items()
                    if abs(float(n["x"]) - d[0]) < 1e-6
                    and abs(float(n["y"]) - d[1]) < 1e-6), None)
        if lab is None or lab not in ati:
            continue
        si.append(s)
        di.append((float(ati[lab]["x"]), float(ati[lab]["y"])))
        zs.append(float(ati[lab].get("elevation", 0.0) or 0.0))
    if len(si) >= 6:
        f4 = fit_similarity(si, di)
        e4 = sorted(math.hypot(f4[0] * s[0] + f4[1] - d[0],
                               f4[0] * s[1] + f4[2] - d[1])
                    for s, d in zip(si, di))
        spani = max(max(d[0] for d in di) - min(d[0] for d in di),
                    max(d[1] for d in di) - min(d[1] for d in di))
        print(f"    닮음(3자유도) 잔차 최대 {e4[-1]:,.2f} · "
              f"한 변의 {e4[-1] / max(1e-9, spani) * 100:.3f}%  ← 안 맞는다")
        # ★닮음으로 안 맞는 것이 곧 «못 깐다» 는 아니다. 아이소 투영은 평면을
        #   **기울여** 놓는다(비등방) — 그러면 닮음(회전+등배율)으로는 원리상
        #   못 맞추고, 아핀(6자유도)이라야 맞는다. 어느 쪽인지 갈라야 한다.
        f5 = fit_affine(si, di)
        if f5 is not None:
            a5, b5 = f5
            e5 = sorted(math.hypot(a5[0] * s[0] + a5[1] * s[1] + a5[2] - d[0],
                                   b5[0] * s[0] + b5[1] * s[1] + b5[2] - d[1])
                        for s, d in zip(si, di))
            print(f"    아핀(6자유도) 잔차 중앙값 {statistics.median(e5):,.2f} · "
                  f"최대 {e5[-1]:,.2f} · 한 변의 "
                  f"{e5[-1] / max(1e-9, spani) * 100:.3f}%")
            print(f"    (이 도면 헤드 표고 폭 {max(zs) - min(zs):.2f} m)")
            if e5[-1] <= spani * 0.005:
                print("    → 아이소에서도 **아핀 한 장**으로 깔 수 있다.")
            else:
                print("    → 이 짝으로는 기준(0.5%)을 못 넘긴다.")
            print("    ※ 다만 이 수치를 «아이소는 안 된다» 로 읽으면 안 된다 —")
            print("      짝이 전부 **헤드**인데, 아이소는 헤드를 일부러 화면")
            print("      수직으로 세운다(§G15 · head_stub). 그 세움은 절점마다")
            print("      더해지는 것이라 아핀 한 장에 담기지 않는다. 즉 여기")
            print("      잔차는 «바닥면 변환» 이 아니라 «헤드 세움» 을 재고 있다.")
            print("      아이소 밑그림의 가부는 세우지 않는 절점으로 다시 잰다 ↓")

    # ── ⑤ 아이소 바닥면 — board 대응이 필요 없다.
    #
    # board → 평면 보기는 위에서 닮음으로 증명됐다. 남은 다리는 «평면 → 아이소»
    # 인데, 그 둘은 **같은 표의 같은 라벨** 이라 대응을 짝지을 필요가 없다.
    # 세우는 헤드만 빼고 재면 바닥면 변환이 그대로 드러난다.
    print("\n■ §17 근거 ③ 다시 — 평면 → 아이소 (같은 라벨 · 헤드 제외)")
    head_labs = {str(r.get("in")) for r in (tbl.nozzles or ())
                 if r.get("in") is not None}
    ps, qs = [], []
    for lab, n in at.items():
        if lab in head_labs or lab not in ati:
            continue
        ps.append((float(n["x"]), float(n["y"])))
        qs.append((float(ati[lab]["x"]), float(ati[lab]["y"])))
    print(f"    세우지 않는 절점 {len(ps)}개 (헤드 {len(head_labs)}개 제외)")
    if len(ps) >= 6:
        fg = fit_affine(ps, qs)
        if fg is not None:
            ag, bg = fg
            eg = sorted(math.hypot(ag[0] * p[0] + ag[1] * p[1] + ag[2] - q[0],
                                   bg[0] * p[0] + bg[1] * p[1] + bg[2] - q[1])
                        for p, q in zip(ps, qs))
            spg = max(max(q[0] for q in qs) - min(q[0] for q in qs),
                      max(q[1] for q in qs) - min(q[1] for q in qs))
            print(f"    아핀 잔차 중앙값 {statistics.median(eg):,.2f} · "
                  f"최대 {eg[-1]:,.2f} · 한 변의 "
                  f"{eg[-1] / max(1e-9, spg) * 100:.3f}%")
            if eg[-1] <= spg * 0.005:
                print("    → 바닥면은 **아핀 한 장**이다. board → 평면(닮음) 을")
                print("      여기 이어 붙이면 board → 아이소도 한 장으로 선다.")
                print("      («세운 헤드» 만 그 규칙 밖이고, 그것은 표가 아는 값이다.)")
            else:
                print("    → 바닥면조차 아핀이 아니다.")
            # ★원인을 단정하기 전에 잰다. 아이소가 z 를 화면으로 섞는다면,
            #   «표고를 함께 넣은» 3차원 아핀은 맞아야 한다.
            zs2 = [float(ati[lab].get("elevation", 0.0) or 0.0)
                   for lab in at if lab not in head_labs and lab in ati]
            print(f"    세우지 않는 절점의 표고 폭 "
                  f"{max(zs2) - min(zs2):.2f} m (n={len(zs2)})")
            src3 = [(p[0], p[1], z) for p, z in zip(ps, zs2)]
            fz = fit_affine3(src3, qs)
            if fz is not None:
                e6 = sorted(math.hypot(
                    sum(a * v for a, v in zip(fz[0], (p[0], p[1], p[2], 1.0)))
                    - q[0],
                    sum(a * v for a, v in zip(fz[1], (p[0], p[1], p[2], 1.0)))
                    - q[1]) for p, q in zip(src3, qs))
                print(f"    표고를 넣은 3차원 아핀 잔차 최대 {e6[-1]:,.2f} · "
                      f"한 변의 {e6[-1] / max(1e-9, spg) * 100:.3f}%")
                if e6[-1] <= spg * 0.005:
                    print("    → 원인은 **표고**다. 평면 (x,y) 만으로는 아이소")
                    print("      자리가 안 정해지므로, «평평한» 밑그림은 아이소에")
                    print("      원리상 못 깐다. §17 의 이 부분은 옳다.")
                    # ★그래도 «못 한다» 로 끝내지 않는다. 아이소가 (x,y,z) 의
                    #   아핀이라는 것은 곧 **평면을 어느 표고에 놓고 같은 식으로
                    #   투영하면 된다** 는 뜻이다. 한 표고로 깔 때 얼마나
                    #   어긋나는지가 실용 판단의 전부다.
                    zc = math.hypot(fz[0][2], fz[1][2])
                    spread = max(zs2) - min(zs2)
                    print(f"\n    z 계수 {zc:.1f} (표시단위/m) · 표고 폭 "
                          f"{spread:.2f} m")
                    print(f"    → 한 표고로 깔면 최대 {zc * spread:,.1f} "
                          f"표시단위 = 한 변의 "
                          f"{zc * spread / max(1e-9, spg) * 100:.2f}% 어긋난다.")
                    print("      («어긋남» 이 아니라 실제 높이차다 — 바닥에 깔면")
                    print("       망이 그만큼 떠 보이는 것이 옳다.)")
                else:
                    print("    → 표고를 넣어도 안 맞는다 — 원인이 다른 데 있다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
