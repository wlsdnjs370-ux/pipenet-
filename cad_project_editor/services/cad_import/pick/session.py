# -*- coding: utf-8 -*-
"""찍기 한 판 — 화면이 부르는 파사드. Qt 없음."""
from services.cad_import.kinds import SLOT_UPDOWN, normalize_head_slot
from services.cad_import.pick.board import Board
from services.cad_import.pick.io import (
    open_dxf, resolve_dxf, write_backup, write_pick,
)

MODE_PIPE = "재료"
MODE_HEAD = "헤드"


class PickSession:
    """DXF 한 장 + Board. 모드·칸은 여기, 판정은 Board."""

    def __init__(self, world, key, knobs, board=None):
        self.world = world
        self.key = key
        self.knobs = knobs
        self.board = board or Board(world, knobs)
        self.mode = MODE_PIPE
        self.armed = False

    @classmethod
    def open(cls, key_or_path, knobs=None):
        path = resolve_dxf(key_or_path)
        w, key, kn = open_dxf(path, knobs)
        return cls(w, key, kn)

    def select_pipe(self):
        """배관선택 — 재료 모드. 헤드는 유지·잠금."""
        self.board.unlock_materials()
        self.mode = MODE_PIPE
        self.armed = True
        return True

    def complete_pipe(self):
        """선택완료 — 재료 0개면 False. 성공하면 헤드 해금."""
        if not self.board.complete_materials():
            return False
        self.mode = MODE_HEAD
        self.armed = True
        return True

    def set_slot(self, label):
        """헤드 칸. 재료 미완료면 False."""
        if not self.board.mat_done:
            return False
        self.board.head_label = normalize_head_slot(label)
        self.mode = MODE_HEAD
        self.armed = True
        return True

    def click(self, x, y, max_d=None):
        """armed 일 때만 찍기. 아니면 None."""
        if not self.armed:
            return None
        return self.board.apply_click(self.mode, x, y, max_d=max_d)

    def undo(self):
        return self.board.undo()

    def spec(self):
        return self.board.spec()

    def highlight_geom(self):
        return self.board.highlight_geom()

    def commit(self, out_dir=None):
        """다음 — 정식 저장. 재료 미완료면 None."""
        if not self.board.mat_done:
            return None
        return write_pick(self.key, self.world, self.board, out_dir=out_dir)

    def backup(self, out_dir=None):
        write_backup(self.key, self.board, out_dir=out_dir)

    @property
    def mat_done(self):
        return self.board.mat_done

    @property
    def head_label(self):
        return self.board.head_label or SLOT_UPDOWN
