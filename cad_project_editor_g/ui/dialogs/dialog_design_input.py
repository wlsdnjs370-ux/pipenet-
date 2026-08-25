# -*- coding: utf-8 -*-
"""[G7] 수리계산 입력 창 — 찍기 → 손질 → 변환 **다음** 네 번째.

판정은 하나도 여기 없다. `services/cad_import/design/` 이 다 하고 이 창은
입력을 받아 그 결과를 보여 줄 뿐이다(모듈 E 의 다른 창과 같은 계약).

★반드시 보여야 하는 것 — 「전개가 못 붙여 제외한 헤드 수」(BLOCKED B4).
최불리는 «배관에 붙는 헤드» 중에서만 고른다. 붙지 못한 헤드가 많다는 것은 그
도면의 배관이 끊겨 있다는 뜻이고, 조용히 빼면 **더 불리한 헤드가 있는데 못 본
채** 수리계산이 나간다. 그래서 이 수가 크면 화면이 경고한다.
"""
from __future__ import annotations

import os

from PySide6.QtCore import QThread, Qt
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
    QLabel, QMessageBox, QPushButton, QSpinBox, QComboBox, QVBoxLayout, QWidget,
)

_WARN_EXCLUDED_RATIO = 0.5      # 이 비율을 넘게 빠지면 경고색으로 알린다


class _DesignThread(QThread):
    """선정·제한 전개·테이블은 무거우므로 UI 스레드 밖에서 돈다.

    `_CadEditBuildThread` 와 같은 패턴이다 — 예외를 삼키지 않고 들고 나온다.
    """

    def __init__(self, job, parent=None):
        super().__init__(parent)
        self._job = job
        self.result_value = None
        self.error = None

    def run(self):
        try:
            self.result_value = self._job()
        except BaseException as exc:      # noqa: BLE001 — 창이 사유를 보여준다
            self.error = exc


class DesignInputDialog(QDialog):
    """기준개수 K 를 받아 최불리 배관망을 확정하고 .sdf + .slf 를 낸다."""

    def __init__(self, parent=None, *, session=None, payload=None,
                 selected_source=None, k=30):
        super().__init__(parent)
        self.setWindowTitle("수리계산 입력 (PIPENET SDF)")
        self._session = session
        self._payload = payload
        self._selected_source = selected_source
        self._result = None          # 마지막 계산 결과(창을 다시 열어도 유지)
        self._tables = None
        self._sheets = []
        self.resize(560, 620)
        self._build(k)
        self._load_sheets()

    # ── 화면 ────────────────────────────────────────────────────────────
    def _build(self, k):
        root = QVBoxLayout(self)

        box_in = QGroupBox("설계 범위")
        form = QFormLayout()
        self.spin_k = QSpinBox()
        self.spin_k.setRange(1, 200)
        self.spin_k.setValue(int(k))
        self.spin_k.setSuffix(" 개")
        form.addRow("기준개수 K (NFPC 103)", self.spin_k)
        self.cmb_sheet = QComboBox()
        self.cmb_sheet.addItem("전체", 0)
        self.lbl_sheet = QLabel("도면 장")
        form.addRow(self.lbl_sheet, self.cmb_sheet)
        box_in.setLayout(form)
        root.addWidget(box_in)

        row = QHBoxLayout()
        self.btn_run = QPushButton("최불리 배관망 확정")
        self.btn_run.clicked.connect(self._on_run)
        row.addWidget(self.btn_run)
        self.btn_save = QPushButton(".sdf + .slf 저장")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self._on_save)
        row.addWidget(self.btn_save)
        root.addLayout(row)

        self.lbl_warn = QLabel("")
        self.lbl_warn.setWordWrap(True)
        self.lbl_warn.setStyleSheet("QLabel { color: #b45309; font-weight: bold; }")
        self.lbl_warn.hide()
        root.addWidget(self.lbl_warn)

        box_out = QGroupBox("결과")
        out = QVBoxLayout()
        self.lbl_sum = QLabel("아직 계산하지 않았습니다.")
        self.lbl_sum.setWordWrap(True)
        self.lbl_sum.setTextFormat(Qt.PlainText)
        out.addWidget(self.lbl_sum)
        box_out.setLayout(out)
        root.addWidget(box_out, 1)

        close = QHBoxLayout()
        close.addStretch(1)
        btn_close = QPushButton("닫기")
        btn_close.clicked.connect(self.reject)
        close.addWidget(btn_close)
        root.addLayout(close)

    def _load_sheets(self):
        """도면이 여러 장일 때만 장 고르기를 보인다."""
        board = getattr(self._session, "board", None)
        if board is None:
            self.lbl_sheet.hide()
            self.cmb_sheet.hide()
            return
        try:
            from services.cad_import.design.worst import sheet_frames
            self._sheets = sheet_frames(board) or []
        except Exception:      # noqa: BLE001 — 장 나누기 실패가 창을 막지 않는다
            self._sheets = []
        if len(self._sheets) < 2:
            self.lbl_sheet.hide()
            self.cmb_sheet.hide()
            return
        for f in self._sheets:
            self.cmb_sheet.addItem(
                f"도면 {f.get('index')} — 헤드 {f.get('head_count')}개",
                int(f.get("index", 0)))

    # ── 계산 ────────────────────────────────────────────────────────────
    def _only_heads(self):
        idx = int(self.cmb_sheet.currentData() or 0)
        if not idx or not self._sheets:
            return None
        hit = next((f for f in self._sheets
                    if int(f.get("index", 0)) == idx), None)
        if hit is None:
            return None
        x0, y0, x1, y1 = [float(v) for v in hit["bbox"]]
        disks = getattr(self._session.board, "disks", []) or []
        return {i for i, d in enumerate(disks)
                if x0 <= float(d[0]) <= x1 and y0 <= float(d[1]) <= y1}

    def _on_run(self):
        board = getattr(self._session, "board", None)
        if board is None:
            self._blocked("손질 세션이 없습니다.")
            return
        # 변환 창과 같은 방식으로 막고 사유를 돌려준다 — 조용히 실패하지 않는다.
        if not getattr(board, "sources", None):
            self._blocked("급수 시작 위치를 먼저 찍어야 최불리를 고를 수 있습니다.")
            return
        undecided = sum(1 for k in getattr(board, "disk_kinds", []) or []
                        if k == "미지정")
        if undecided:
            self._blocked(f"헤드 종류가 미지정인 것이 {undecided}개 있습니다. "
                          "손질 창에서 종류를 정한 뒤 다시 시도하세요.")
            return

        k = int(self.spin_k.value())
        only = self._only_heads()
        payload = self._payload
        sel = self._selected_source

        def job():
            from services.cad_import.design.bore import extract_dia_text_points
            from services.cad_import.design.restrict import select_and_expand
            from services.cad_import.design.tables import build_design_tables

            got = select_and_expand(payload, board, k=k, only_heads=only,
                                    selected_source=sel)
            if not got.get("ok"):
                return {"ok": False, "error": got.get("error")}
            texts = self._dia_texts()
            tbl = build_design_tables(
                got["kfp"], got["worst"], got["edge_ref"], texts,
                board_pts=board.pts,
                excluded_heads=got.get("excluded_heads", 0),
                valve_nodes=None)
            return {"ok": True, "got": got, "tables": tbl}

        self.btn_run.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        th = _DesignThread(job, self)
        th.start()
        while th.isRunning():
            QApplication.processEvents()
            th.wait(30)
        while QApplication.overrideCursor() is not None:
            QApplication.restoreOverrideCursor()
        self.btn_run.setEnabled(True)

        if th.error is not None:
            self._blocked(f"{type(th.error).__name__}: {th.error}")
            return
        res = th.result_value or {}
        if not res.get("ok"):
            self._blocked(str(res.get("error") or "계산에 실패했습니다."))
            return
        self._result = res["got"]
        self._tables = res["tables"]
        self.btn_save.setEnabled(True)
        self._render(res["got"], res["tables"])

    def _dia_texts(self):
        """치수 텍스트 — handoff 캐시에서. 못 읽으면 빈 목록(관경은 별표1 폴백)."""
        try:
            import json
            from services.cad_import.pipeline import handoff, stage1 as s1
            from services.cad_import.design.bore import extract_dia_text_points
            key = getattr(self._session, "key", None)
            spec = os.path.join(handoff.pick_out_dir(), f"{key}_찍은스펙.json")
            with open(spec, encoding="utf-8") as f:
                src = json.load(f).get("source_dxf")
            w = handoff.load_world(key, src, s1.World)
            return extract_dia_text_points(w.texts) if w is not None else []
        except Exception as exc:      # noqa: BLE001
            print(f"[G7] 치수 텍스트를 읽지 못했습니다 — 관경은 별표1 로만: {exc}")
            return []

    # ── 표시 ────────────────────────────────────────────────────────────
    def _blocked(self, msg):
        self.lbl_warn.setText(f"막힘 — {msg}")
        self.lbl_warn.show()
        self.lbl_sum.setText("계산하지 못했습니다.")

    def _render(self, got, tbl):
        meta = dict(tbl.meta)
        w = got.get("worst") or {}
        total = int(got.get("total_heads") or 0)
        excluded = int(got.get("excluded_heads") or 0)
        cand = int(got.get("candidate_heads") or 0)

        lines = [
            f"설계면적 : 헤드 {len(w.get('heads') or [])}개",
            f"앵커(최원 유하거리) : {w.get('far_m')} m",
            f"설계면적 폭 : {w.get('span_m')} m",
            f"corridor 총연장 : {w.get('total_m')} m",
            f"주배관 담당 헤드 수 : {w.get('max_load')}",
            "",
            f"관경 근거 : 도면 텍스트 {meta.get('관경 근거 — 도면 텍스트')} · "
            f"별표1 보강 {meta.get('관경 근거 — 별표1 보강 (text<min)')} · "
            f"별표1 폴백 {meta.get('관경 근거 — 별표1 폴백 (text 없음)')}",
            f"부속 : {len(tbl.fittings)}건 · "
            f"판정 불가 {meta.get('부속 판정 불가')} · "
            f"등가길이 미해결 {meta.get('등가길이 미해결')}",
            f"표 : 노드 {len(tbl.nodes)} · 배관 {len(tbl.pipes)} · "
            f"노즐 {len(tbl.nozzles)} · 기기 {len(tbl.equipment)}",
            f"루프 잔여(표 꼬리) : {meta.get('루프 잔여 배관(표 꼬리)')}건",
        ]
        self.lbl_sum.setText("\n".join(lines))

        # ★제외 헤드 — 숨기지 않는다(BLOCKED B4).
        if excluded and total:
            ratio = excluded / total
            msg = (f"전개가 배관에 붙이지 못한 헤드 {excluded:,}개를 후보에서 "
                   f"제외했습니다 (후보 {cand:,} / 도면 {total:,}).")
            if ratio >= _WARN_EXCLUDED_RATIO:
                msg += ("\n그만큼 배관이 끊겨 있을 수 있습니다 — 더 불리한 헤드가 "
                        "빠졌을 수 있으니 손질 창에서 배관을 이어 주세요.")
            self.lbl_warn.setText(msg)
            self.lbl_warn.show()
        else:
            self.lbl_warn.hide()

    # ── 저장 ────────────────────────────────────────────────────────────
    def _on_save(self):
        if self._tables is None:
            return
        from services.cad_import.design.emit import AssetMissing, emit_design_sdf

        key = getattr(self._session, "key", None) or "design"
        path, _f = QFileDialog.getSaveFileName(
            self, "수리계산 입력 저장", f"{key}_수리계산입력.sdf",
            "PIPENET SDF (*.sdf);;All Files (*)")
        if not path:
            return
        try:
            out = emit_design_sdf(self._tables, path,
                                  project_title=f"{key} 수리계산 입력")
        except AssetMissing as exc:
            QMessageBox.critical(self, "수리계산 입력", str(exc))
            return
        except Exception as exc:      # noqa: BLE001
            QMessageBox.warning(self, "수리계산 입력", str(exc))
            return
        QMessageBox.information(
            self, "수리계산 입력",
            f"저장했습니다.\n\n{out}\n{out.with_suffix('.slf')}\n\n"
            "SDF 는 옆의 .slf 와 **한 쌍**입니다. 라이브러리 없이 열면 "
            "관경이 'Unset' 으로 뜹니다.")

    # 창을 닫았다 다시 열어도 직전 결과를 쓰기 위한 접근자
    @property
    def result(self):
        return {"worst": self._result, "tables": self._tables}
