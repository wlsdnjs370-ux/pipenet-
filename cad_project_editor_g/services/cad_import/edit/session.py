# -*- coding: utf-8 -*-
"""손질 한 판 — 화면이 부르는 파사드. Qt 없음."""
import math

from services.cad_import.edit.board import (
    body_seg_groups, head_face_colors,
    wet_disk_keys, wet_set_from_disk_keys,
)
from services.cad_import.edit.io import load_edits, open_board, write_edits
from services.cad_import.pipeline.user_net import _pt_seg_d2, pick_head, pick_seg

MODE_JOIN = "이음"
MODE_DELETE = "삭제"
MODE_SOURCE = "급수시작위치"
MODE_VALVE = "알람밸브위치"


def _prefer_head(x, y, seg, head):
    """헤드가 배관보다 가깝거나 배관이 없으면 True. 시제품 이음 2클릭과 같음."""
    if head is None:
        return False
    if seg is None:
        return True
    seg_d = math.sqrt(_pt_seg_d2((x, y), *seg))
    head_d = max(0.0, math.hypot(x - head[0], y - head[1]) - head[2])
    return head_d <= seg_d


class EditSession:
    """자동망 한 장 + EditBoard. 모드·pending·고른 헤드는 여기, 판정은 Board."""

    def __init__(self, board, key=None, out_dir=None):
        self.board = board
        self.key = key if key is not None else board.key
        self.out_dir = out_dir
        self.mode = MODE_JOIN
        self.selected_head = None
        self._flowed = False
        self._wet_keys = None
        self._water = None
        self._frame = None
        self._hop_last = 0
        self._hop_step = 1
        self._wet_ranked = []
        self._head_hop = []

    @classmethod
    def open(cls, key, out_dir=None, load_saved=True, use_cache=True):
        board = open_board(key, use_cache=use_cache)
        if load_saved:
            load_edits(board, out_dir)
        return cls(board, key, out_dir)

    def set_mode(self, mode):
        self.mode = mode
        self.board.pending = None
        return True

    def click(self, x, y, max_d):
        """모드에 따라 보드를 두드린다. 좌표·반경만 받는다."""
        b = self.board
        if self.mode == MODE_SOURCE:
            node, added = b.toggle_source(x, y)
            if node is None:
                return None
            return {"동작": "급수", "node": node, "added": added}
        if self.mode == MODE_VALVE:
            node, added = b.toggle_valve(x, y, max_d)
            if node is None:
                return None
            return {"동작": "밸브", "node": node, "added": added}
        segs = b.segments()
        seg = pick_seg(segs, x, y, max_d)
        head = pick_head(b.disks, x, y, max_d)
        if self.mode == MODE_JOIN:
            return self._click_join(x, y, seg, head)
        if self.mode == MODE_DELETE:
            return self._click_delete(x, y, max_d, seg, head)
        return None

    def _click_join(self, x, y, seg, head):
        b = self.board
        if b.pending is None:
            if _prefer_head(x, y, seg, head):
                self.selected_head = head
                return {"동작": "헤드선택", "head": head}
            if seg is None:
                return None
            b.pending = seg
            return {"동작": "대기", "seg": seg}
        if _prefer_head(x, y, seg, head):
            self.selected_head = head
            made, blocked, covered, kind = b.join_head(b.pending, head)
            b.pending = None
            return {"동작": "이음", "kind": kind, "made": made,
                    "blocked": blocked, "covered": covered}
        if seg is not None:
            made, blocked, covered, kind = b.join(b.pending, seg)
            b.pending = None
            return {"동작": "이음", "kind": kind, "made": made,
                    "blocked": blocked, "covered": covered}
        return None

    def _click_delete(self, x, y, max_d, seg, head):
        """배관 삭제. 급수원/밸브 점이 더 가까우면 찍기만 지운다."""
        b = self.board
        src = b.pick_source(x, y, max_d)
        valv = b.pick_valve(x, y, max_d)
        seg_d = None if seg is None else math.sqrt(_pt_seg_d2((x, y), *seg))
        src_d = (None if src is None
                 else math.hypot(b.pts[src][0] - x, b.pts[src][1] - y))
        valv_d = (None if valv is None
                  else math.hypot(b.pts[valv][0] - x, b.pts[valv][1] - y))
        mark = []
        if src_d is not None:
            mark.append((src_d, "source", src))
        if valv_d is not None:
            mark.append((valv_d, "valve", valv))
        if mark:
            mark.sort()
            md, kind, node = mark[0]
            if seg_d is None or md <= seg_d:
                if kind == "source":
                    if not b.remove_source(node):
                        return None
                    return {"동작": "급수해제", "node": node}
                if not b.remove_valve(node):
                    return None
                return {"동작": "밸브해제", "node": node}
        if seg is not None:
            kind, n = b.delete(seg)
            return {"동작": "삭제", "kind": kind, "n": n}
        if head is not None:
            self.selected_head = head
            return {"동작": "헤드선택", "head": head}
        return None

    def set_kind(self, kind):
        """고른 헤드를 단추 종류로 덮는다. 고른 헤드가 없으면 None."""
        if self.selected_head is None:
            return None
        return self.board.set_head_kind(self.selected_head, kind)

    def flow(self):
        """급수가 있을 때만 water_state 를 그리고 20프레임 홉을 시작한다."""
        if not self.board.sources:
            return None
        state = self.board.water_state()
        self._flowed = True
        self._water = state
        self._wet_keys = wet_disk_keys(self.board.disks, state["wet_heads"])
        last = max(state["hop"].values(), default=0)
        self._hop_last = last
        self._hop_step = max(1, (last + 20 - 1) // 20)
        hop = state["hop"]
        wet_edges = state["wet_edges"]
        ranked = []
        for a, b in self.board.edges:
            if (min(a, b), max(a, b)) not in wet_edges:
                continue
            ranked.append((max(hop[a], hop[b]),
                           (self.board.pts[a], self.board.pts[b])))
        ranked.sort(key=lambda x: x[0])
        self._wet_ranked = ranked
        self._head_hop = []
        for nodes in self.board.hnodes:
            hs = [hop[n] for n in nodes if n in hop]
            self._head_hop.append(min(hs) if hs else None)
        self._frame = None if last == 0 else 0
        return state

    def flow_animating(self):
        return self._frame is not None

    def flow_tick(self):
        """홉 한 칸. 끝나면 False. 판정 없음."""
        if self._frame is None:
            return False
        self._frame = self._frame + self._hop_step
        if self._frame >= self._hop_last:
            self._frame = None
            return False
        return True

    def wet_kind_counts(self):
        """물흐름 후에만 젖은 헤드를 종류별로. 전에는 0."""
        counts = {"상향식": 0, "하향식": 0, "상하향식": 0}
        if not self._flowed:
            return counts
        wet_set = wet_set_from_disk_keys(self.board.disks, self._wet_keys)
        for i in wet_set or ():
            if i < 0 or i >= len(self.board.disk_kinds):
                continue
            k = self.board.disk_kinds[i]
            if k in counts:
                counts[k] += 1
        return counts

    def undo(self):
        ok = self.board.undo()
        if ok:
            self.selected_head = None
        return ok

    def commit(self, out_dir=None):
        dest = out_dir if out_dir is not None else self.out_dir
        return write_edits(self.board, dest)

    def convert_payload(self):
        """변환 화면이 받는 현재 망. 파일을 다시 읽지 않는다."""
        b = self.board
        data = b.payload()
        data["key"] = self.key
        data["pts"] = [list(p) for p in b.pts]
        data["edges"] = [list(e) for e in b.edges]
        data["hcov"] = [list(d) for d in b.disks]
        data["ups"] = [list(u) for u in b.ups]
        data["disk_kinds"] = list(b.disk_kinds)
        return data

    def display_geom(self, net=True):
        """화면용. 판정 없음. 연출 중에만 wet_pipes. net=False 는 망 재계산 생략."""
        b = self.board
        wet_set = None
        if self._flowed:
            if self._frame is not None:
                wet_set = {hi for hi, hh in enumerate(self._head_hop)
                           if hh is not None and hh <= self._frame}
            else:
                wet_set = wet_set_from_disk_keys(b.disks, self._wet_keys)
        colors = head_face_colors(b.disk_kinds, wet_set)
        npts = len(b.pts)
        sources = []
        for n in b.sources:
            if isinstance(n, int) and 0 <= n < npts:
                sources.append(tuple(b.pts[n]))
        valves = []
        for n in b.valves:
            if isinstance(n, int) and 0 <= n < npts:
                valves.append(tuple(b.pts[n]))
        return {
            "body_groups": (body_seg_groups(b.pts, b.edges, b.bodies())
                            if net else []),
            "heads": list(zip(b.disks, colors)),
            "multi_heads": [b.disks[di] for di in b.multi_arm_heads],
            "pending": b.pending,
            "selected_head": self.selected_head,
            "sources": sources,
            "valves": valves,
            "wet_pipes": self._wet_pipe_segs(),
        }

    def _wet_pipe_segs(self):
        """연출 중 홉 ≤ frame 인 젖은 간선. flow() 때 모아 둔 목록만."""
        if self._frame is None:
            return []
        return [s for h, s in self._wet_ranked if h <= self._frame]
