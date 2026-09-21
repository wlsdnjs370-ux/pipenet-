# -*- coding: utf-8 -*-
"""찍기 판 — 클릭 → 스펙. 화면 없음.

대화형 Qt 와 찍음 파일 재현이 같은 apply_click 을 탄다.
"""
import math
from collections import defaultdict

from services.cad_import.colors import cname
from services.cad_import.kinds import SLOT_UPDOWN, normalize_head_slot
from services.cad_import.pipeline import heads as flow_heads
from services.cad_import.pipeline import stage1 as s1

MODES = ("재료", "헤드")
CELL = 2000.0
PICK_PX = 10.0
# 겹쳐 그린 같은 원 — 본체 같은자리 지우기(§3-3)와 같은 눈금.
_DEDUP_TOL = float(s1.DEFAULT_KNOBS["rim_tol"])


def head_key(h):
    """헤드 픽 열쇠 — 같은 bundle×r 에 지문·칸이 다른 픽이 공존."""
    label = normalize_head_slot(h.get("label") or SLOT_UPDOWN)
    if "tri_side" in h:
        return (tuple(h["bundle"]), "tri", round(float(h["tri_side"]), 1),
                label, ())
    fp = h.get("mark_fp")
    if fp is None:
        fp = ["empty"]
    return (tuple(h["bundle"]), "r", round(float(h["r"]), 1),
            label, flow_heads.mark_fp_key(fp))


def head_full_key(h):
    """검증 대조용 — 정체 열쇠 + cluster_gap."""
    return head_key(h) + (h.get("cluster_gap"),)


def heads_from_spec(sp):
    """스펙 heads → 보드용. label·kind·mark_fp·mark_bundles 보존."""
    out = []
    for h in sp.get("heads") or []:
        h = dict(h)
        h["bundle"] = tuple(h["bundle"])
        if h.get("mark_fp") is None and "tri_side" not in h:
            h["mark_fp"] = ["empty"]
        if "tri_side" not in h:
            h["mark_bundles"] = [tuple(b) for b in
                                 (h.get("mark_bundles") or [])]
        h["label"] = normalize_head_slot(h.get("label"))
        if not str(h.get("kind") or "").strip():
            h["kind"] = "미지정"
        out.append(h)
    return out


class Board:
    """찍기 판 — 세계·격자·픽 목록."""

    def __init__(self, w, kn):
        self.w = w
        self.kn = kn
        self.mat = []
        self.heads = []
        self.clicks = []
        self.head_label = SLOT_UPDOWN
        self.mat_done = False
        self._sgrid = defaultdict(list)
        self.by_bundle = defaultdict(list)
        xs, ys = [], []
        for i, (ly, c, a, b) in enumerate(w.segs):
            self.by_bundle[(ly, c)].append((a, b))
            x0, x1 = sorted((a[0], b[0]))
            y0, y1 = sorted((a[1], b[1]))
            xs += [x0, x1]
            ys += [y0, y1]
            for gx in range(int(x0 // CELL), int(x1 // CELL) + 1):
                for gy in range(int(y0 // CELL), int(y1 // CELL) + 1):
                    self._sgrid[(gx, gy)].append(i)
        self._csmall = [t for t in w.circles if t[4] <= kn["small_r"]]
        self._cgrid = defaultdict(list)
        for i, (_ly, _c, x, y, _r) in enumerate(self._csmall):
            s1._grid_put(self._cgrid, CELL, x, y, i)
        span = (max(xs) - min(xs), max(ys) - min(ys)) if xs else (0.0, 0.0)
        self._maxring = int(max(span) / CELL) + 2
        self.fp_index = flow_heads.build_fp_index(w, kn["small_len"])
        self._head_match_cache = {}
        self._tri_seg_cache = {}

    def complete_materials(self):
        """재료 찍기 완료 → 헤드 모드 해금. 재료 0개면 거부."""
        if not self.mat:
            return False
        self.mat_done = True
        return True

    def unlock_materials(self):
        """재료 모드로만 연다 — 헤드는 아직 유지."""
        self.mat_done = False

    def clear_heads(self):
        """찍은 헤드·헤드 클릭 기록 해제. 반환: 해제한 픽 수."""
        n = len(self.heads)
        if not n:
            return 0
        self.heads = []
        self.clicks = [cl for cl in self.clicks
                       if (cl.get("모드") or cl.get("mode")) != "헤드"]
        return n

    def _ring(self, grid, x, y, ring):
        gx, gy = int(x // CELL), int(y // CELL)
        if ring == 0:
            yield from grid.get((gx, gy), ())
            return
        for dx in range(-ring, ring + 1):
            for dy in (-ring, ring):
                yield from grid.get((gx + dx, gy + dy), ())
        for dy in range(-ring + 1, ring):
            for dx in (-ring, ring):
                yield from grid.get((gx + dx, gy + dy), ())

    def nearest_seg(self, x, y, max_len=None):
        """클릭에서 가장 가까운 선분. 반환: (거리, (레이어,색), (a,b)) 또는 None."""
        best = None
        seen = set()
        for ring in range(self._maxring + 1):
            if best is not None and (ring - 1) * CELL > best[0]:
                break
            for i in self._ring(self._sgrid, x, y, ring):
                if i in seen:
                    continue
                seen.add(i)
                ly, c, a, b = self.w.segs[i]
                if max_len is not None and \
                        math.hypot(b[0] - a[0], b[1] - a[1]) > max_len:
                    continue
                d = s1.seg_geom(a, b, x, y)[0]
                if best is None or (d, ly, c) < (best[0], best[1][0],
                                                 best[1][1]):
                    best = (d, (ly, c), (a, b))
        return best

    def nearest_circle(self, x, y):
        """클릭에서 가장 가까운 작은 원 — 거리는 테두리까지."""
        best = None
        seen = set()
        small_r = self.kn["small_r"]
        for ring in range(self._maxring + 1):
            if best is not None and (ring - 1) * CELL - small_r > best[0]:
                break
            for i in self._ring(self._cgrid, x, y, ring):
                if i in seen:
                    continue
                seen.add(i)
                ly, c, cx, cy, r = self._csmall[i]
                d = abs(math.hypot(cx - x, cy - y) - r)
                if best is None or (d, ly, c, r) < (best[0], best[1][0],
                                                    best[1][1], best[1][2]):
                    best = (d, (ly, c, round(r, 1)), (cx, cy, r))
        return best

    def same_place_circles(self, circ):
        """같은 자리·같은 크기로 겹쳐 그려진 원 전부의 (레이어×색×r)."""
        cx, cy, r = circ
        r1 = round(r, 1)
        keys = set()
        for i in s1._grid_near(self._cgrid, CELL, cx, cy):
            ly, c, qx, qy, qr = self._csmall[i]
            if round(qr, 1) == r1 and \
                    math.hypot(qx - cx, qy - cy) <= _DEDUP_TOL:
                keys.add((ly, c, r1))
        return sorted(keys)

    def tri_head_segs(self, pick):
        """이 삼각형 픽이 잡는 삼각형들의 획."""
        key = head_full_key(pick)
        cached = self._tri_seg_cache.get(key)
        if cached is not None:
            return list(cached)
        spec1 = {"heads": [dict(pick, bundle=list(pick["bundle"]))]}
        cls = flow_heads.collect_head_clusters(self.w, spec1, self.kn)
        found, _marks, _info = flow_heads.split_head_circles(cls, self.kn)
        segs = tuple(s for h in found if "tri_side" in h
                     for s in h["tri_segs"])
        self._tri_seg_cache[key] = segs
        return list(segs)

    def _head_matches(self, pick, mat_set):
        """헤드 픽의 전체 원 목록. 같은 상태에서는 판정을 다시 돌리지 않는다."""
        marks = tuple(sorted(
            (tuple(b) for b in (pick.get("mark_bundles") or ())),
            key=repr))
        mat_key = None if mat_set is None else tuple(sorted(mat_set, key=repr))
        key = (head_key(pick), marks, mat_key)
        cached = self._head_match_cache.get(key)
        if cached is None:
            cached = tuple(flow_heads.circles_with_fp(
                self.w, pick["bundle"], pick["r"],
                pick.get("mark_fp") or ["empty"], self.kn["small_len"],
                index=self.fp_index, mat_set=mat_set,
                mark_bundles=marks if mat_set is not None else None))
            self._head_match_cache[key] = cached
        return cached

    def tri_at(self, x, y):
        """짧은 획이 닫힌 등변 삼각형의 변이면 후보."""
        got = self.nearest_seg(x, y, max_len=self.kn["small_len"])
        if got is None:
            return None
        d, (ly, c), seg = got
        side = round(math.hypot(seg[1][0] - seg[0][0],
                                seg[1][1] - seg[0][1]), 1)
        pick = {"bundle": (ly, c), "tri_side": side}
        for s3 in self.tri_head_segs(pick):
            if s3 == seg:
                return d, (ly, c, side), seg
        return None

    def apply_click(self, mode, x, y, cluster_gap=None, max_d=None,
                    label=None, mark_bundle=None):  # noqa: ARG002
        """찍기 한 번 — 성공하면 보고 dict, 못/안 찍으면 None."""
        if mode == "기타":
            return None
        if mode in ("상하향표식", "헤드문양"):
            return None
        if mode not in MODES:
            return None
        if mode == "헤드":
            if not self.mat_done:
                return None
            return self._click_head(x, y, cluster_gap, max_d,
                                    label=label or self.head_label)
        return self._click_line(x, y, max_d)

    def _click_line(self, x, y, max_d):
        got = self.nearest_seg(x, y)
        if got is None:
            return None
        d, (ly, c), _seg = got
        if max_d is not None and d > max_d:
            return None
        cleared = self.clear_heads() if self.heads else 0
        self.mat_done = False
        what = f"{ly}×{cname(c)}({c})"
        if (ly, c) in self.mat:
            self.mat.remove((ly, c))
            action = "취소"
        else:
            self.mat.append((ly, c))
            action = "추가"
        rec = {"모드": "재료", "x": round(float(x), 1),
               "y": round(float(y), 1),
               "픽": [ly, c], "동작": action, "찍힘": what}
        self.clicks.append(rec)
        return {"모드": "재료", "픽": what, "d": d, "동작": action,
                "헤드해제": cleared}

    def _commit_heads(self, x, y, picks, what, d, cluster_gap,
                      action_force=None):
        have = {head_key(h) for h in self.heads}
        new = [p for p in picks if head_key(p) not in have]
        if action_force == "추가" or (action_force is None and new):
            for p in new:
                if cluster_gap:
                    p["cluster_gap"] = float(cluster_gap)
                if not str(p.get("kind") or "").strip():
                    p["kind"] = "미지정"
                self.heads.append(p)
            action, applied = "추가", new
        else:
            keys = {head_key(p) for p in picks}
            applied = [h for h in self.heads if head_key(h) in keys]
            self.heads = [h for h in self.heads if head_key(h) not in keys]
            action = "취소"
        head_recs = []
        for p in applied:
            pr = dict(p, bundle=list(p["bundle"]))
            head_recs.append(pr)
        rec = {"모드": "헤드", "x": round(float(x), 1),
               "y": round(float(y), 1), "동작": action,
               "칸": (applied[0].get("label") if applied
                     else self.head_label),
               "헤드픽들": head_recs, "찍힘": what}
        if cluster_gap:
            rec["cluster_gap"] = float(cluster_gap)
        self.clicks.append(rec)
        return {"모드": "헤드", "픽": what, "d": d, "동작": action}

    def _click_head(self, x, y, cluster_gap, max_d, label=None):
        label = normalize_head_slot(label or self.head_label)
        cand_c = self.nearest_circle(x, y)
        got_s = self.nearest_seg(x, y, max_len=self.kn["small_len"])
        cand_t = None
        if got_s is not None and (cand_c is None or got_s[0] < cand_c[0]):
            cand_t = self.tri_at(x, y)
        if cand_c is None and cand_t is None:
            return None
        use_tri = cand_t is not None and \
            (cand_c is None or cand_t[0] < cand_c[0])
        d = cand_t[0] if use_tri else cand_c[0]
        if max_d is not None and d > max_d:
            return None
        if use_tri:
            _d, (ly, c, side), _seg = cand_t
            picks = [{"bundle": (ly, c), "tri_side": side, "label": label,
                      "mark_fp": ["empty"], "mark_bundles": []}]
            what = (f"{ly}×{cname(c)}({c}) 삼각형 변={side}"
                    f" [{label}]")
            return self._commit_heads(x, y, picks, what, d, cluster_gap)
        _d, (ly, c, _rr), (cx, cy, r1) = cand_c
        group = self.same_place_circles((cx, cy, r1))
        picks = []
        n_twin = 0
        bits = []
        sl = self.kn["small_len"]
        mat_set = set(self.mat)
        spatial = (self.fp_index or {}).get("spatial")
        for ly2, c2, r2 in group:
            fp, marks = flow_heads.pick_mark_fp(
                self.w, cx, cy, r2, sl, mat_set, spatial=spatial)
            pick = {"bundle": (ly2, c2), "r": float(r2),
                    "label": label, "mark_fp": fp,
                    "mark_bundles": [list(b) for b in marks]}
            twins = self._head_matches(pick, mat_set)
            n_twin += len(twins)
            picks.append(pick)
            bits.append(f"{ly2}×{cname(c2)}({c2}) r={r2}")
        mb = picks[0].get("mark_bundles") or []
        what = (" + ".join(bits) + f" [{label}] fp={picks[0]['mark_fp']}"
                f" · 문양묶음 {len(mb)} · 같은문양 {n_twin}원")
        return self._commit_heads(x, y, picks, what, d, cluster_gap)

    def undo(self):
        if not self.clicks:
            return None
        cl = self.clicks.pop()
        if cl["모드"] == "헤드":
            if cl.get("동작") == "문양대기":
                return cl
            picks = [dict(p, bundle=tuple(p["bundle"]))
                     for p in cl.get("헤드픽들") or []]
            keys = {head_key(p) for p in picks}
            if cl.get("동작") == "취소":
                self.heads.extend(picks)
            else:
                self.heads = [h for h in self.heads
                              if head_key(h) not in keys]
        elif cl["모드"] == "재료":
            pick = tuple(cl["픽"])
            if cl.get("동작") == "취소":
                self.mat.append(pick)
            else:
                self.mat.remove(pick)
        return cl

    def spec(self):
        """그릇 v2 spec — 찍은 그대로. 호 각은 world 실제값만."""
        sp = {"format": "v2",
              "material_picks": [list(t) for t in self.mat]}
        if self.heads:
            clean = []
            for h in self.heads:
                d = {"bundle": list(h["bundle"])}
                if "tri_side" in h:
                    d["tri_side"] = h["tri_side"]
                else:
                    d["r"] = h["r"]
                if h.get("label"):
                    d["label"] = normalize_head_slot(h["label"])
                d["kind"] = str(h.get("kind") or "").strip() or "미지정"
                if "tri_side" not in h:
                    d["mark_fp"] = h.get("mark_fp") or ["empty"]
                    d["mark_bundles"] = [
                        list(b) for b in (h.get("mark_bundles") or [])]
                if h.get("cluster_gap") is not None:
                    d["cluster_gap"] = h["cluster_gap"]
                clean.append(d)
            sp["heads"] = clean
        ho = []
        mat_layers = {ly for ly, _c in self.mat}
        small_r = float(self.kn["small_r"])
        angs = getattr(self.w, "arc_ang", ())
        for i, (ly, _c, cx, cy, r) in enumerate(self.w.arcs):
            if not 0 < r <= small_r or ly not in mat_layers:
                continue
            ang = angs[i] if i < len(angs) else None
            if not ang:
                continue
            ho.append({"cx": float(cx), "cy": float(cy), "r": float(r),
                       "sa": float(ang[0]), "sweep": float(ang[1])})
        sp["ho"] = ho
        return sp

    def highlight_geom(self):
        """화면 강조용 — 찍은 선·원·삼각형·마지막 클릭. Qt 없음."""
        pipe = []
        for ly, c in self.mat:
            pipe.extend(self.by_bundle.get((ly, c)) or [])
        hcirc = []
        seen = set()
        mat_set = set(self.mat) if self.mat_done else None
        for h in self.heads:
            if "tri_side" in h:
                continue
            for ly, c, cx, cy, r in self._head_matches(h, mat_set):
                key = (ly, c, round(cx, 1), round(cy, 1), round(r, 1))
                if key in seen:
                    continue
                seen.add(key)
                hcirc.append((cx, cy, r))
        tsegs = [s for h in self.heads if "tri_side" in h
                 for s in self.tri_head_segs(h)]
        last = None
        if self.clicks:
            cl = self.clicks[-1]
            last = (cl["x"], cl["y"])
        return {"pipe_bundles": list(self.mat), "pipe_segs": pipe,
                "head_circles": hcirc,
                "tri_segs": tsegs, "last_click": last}
