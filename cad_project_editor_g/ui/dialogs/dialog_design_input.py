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

from PySide6.QtCore import QPointF, QThread, Qt
from PySide6.QtGui import QBrush, QColor, QPainter, QPen, QPolygonF
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog,
    QDoubleSpinBox, QFileDialog, QFormLayout, QGraphicsScene, QGraphicsView,
    QGroupBox, QHBoxLayout, QHeaderView, QLabel, QMessageBox, QPushButton,
    QScrollArea, QSpinBox, QSplitter, QTableWidget, QTableWidgetItem,
    QTabWidget, QVBoxLayout, QWidget,
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


class _PreviewView(QGraphicsView):
    """휠 줌 · 드래그 팬. 그리는 좌표는 저장에 쓰는 사본 그대로다(§G16)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScene(QGraphicsScene(self))
        self.setRenderHint(QPainter.Antialiasing, True)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setBackgroundBrush(QColor("#ffffff"))

    def wheelEvent(self, ev):
        f = 1.15 if ev.angleDelta().y() > 0 else 1 / 1.15
        self.scale(f, f)

    def fit(self):
        r = self.scene().itemsBoundingRect()
        if r.isEmpty():
            return
        self.fitInView(r.adjusted(-40, -40, 40, 40), Qt.KeepAspectRatio)


class DesignInputDialog(QDialog):
    """기준개수 K 를 받아 최불리 배관망을 확정하고 .sdf + .slf 를 낸다."""

    def __init__(self, parent=None, *, session=None, payload=None,
                 selected_source=None, k=30, convert_kwargs=None):
        super().__init__(parent)
        self.setWindowTitle("수리계산 입력 (PIPENET SDF)")
        self._session = session
        self._payload = payload
        self._selected_source = selected_source
        # ★변환 창에서 사람이 고른 헤드 접속관 길이. 이 값이 여기까지 와야
        #   `.kfp` 와 `.sdf` 가 같은 망이 된다 — 다르면 두 산출물이 다른 도면이다.
        self._convert_kwargs = convert_kwargs
        self._result = None          # 마지막 계산 결과(창을 다시 열어도 유지)
        self._tables = None
        self._sheets = []
        self._preview_points = {}     # 검사가 SDF Position 과 견주는 값(§G16)
        self._pipe_items = {}
        self.resize(1200, 780)
        self.setMinimumSize(1000, 640)
        self._build(k)
        self._load_sheets()

    # ── 화면 ────────────────────────────────────────────────────────────
    def _build(self, k):
        outer = QVBoxLayout(self)
        split = QSplitter(Qt.Horizontal, self)
        left_panel = QWidget()
        root = QVBoxLayout(left_panel)

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
        # [G10] 배관 규격 — 이름은 SLF 의 Item-name 과 같아야 PIPENET 이 내경을
        # 바인딩한다. 배관별 지정은 아직 미정이라(BLOCKED B8) 전체 기본값만 연다.
        from services.cad_import.design.sdf_post import SCHEDULE_NAMES
        self.cmb_sched = QComboBox()
        for n in SCHEDULE_NAMES:
            self.cmb_sched.addItem(n, n)
        form.addRow("배관 규격 (전체 기본값)", self.cmb_sched)
        box_in.setLayout(form)
        root.addWidget(box_in)

        # [G12] 표시 전용 조절 — 수리계산 결과는 바뀌지 않는다.
        box_view = QGroupBox("보기 (표시 전용 · 수리계산 결과는 바뀌지 않습니다)")
        vform = QFormLayout()
        self.chk_iso = QCheckBox("30° 등각으로 굽기")
        vform.addRow("아이소매트릭 보기", self.chk_iso)
        self.spin_zscale = QDoubleSpinBox()
        self.spin_zscale.setRange(0.5, 3.0)
        self.spin_zscale.setSingleStep(0.1)
        self.spin_zscale.setValue(1.0)
        vform.addRow("고도 펼침 배율", self.spin_zscale)
        self.spin_canvas = QSpinBox()
        self.spin_canvas.setRange(500, 20000)
        self.spin_canvas.setSingleStep(500)
        self.spin_canvas.setValue(3000)
        vform.addRow("캔버스 크기", self.spin_canvas)
        self.cmb_ref = QComboBox()
        self.cmb_ref.addItem("알람밸브 (없으면 표고 중앙)", "valve")
        self.cmb_ref.addItem("표고 중앙", "mid")
        vform.addRow("lift 영점", self.cmb_ref)
        # 헤드 스텁 — 등각에서 헤드를 화면 수직으로 세울 때 그 길이(§G15).
        # 캔버스 크기에 대한 비율이라 캔버스를 키워도 비례가 유지된다.
        self.spin_stub = QDoubleSpinBox()
        self.spin_stub.setRange(0.5, 10.0)
        self.spin_stub.setSingleStep(0.5)
        self.spin_stub.setValue(2.5)
        self.spin_stub.setSuffix(" %")
        vform.addRow("헤드 스텁 길이", self.spin_stub)
        box_view.setLayout(vform)
        root.addWidget(box_view)

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

        # ── 오른쪽: 미리보기. 저장하기 전에 형태와 표 값을 여기서 본다(§G16).
        scroll = QScrollArea()
        scroll.setWidget(left_panel)
        scroll.setWidgetResizable(True)
        scroll.setMinimumWidth(360)
        split.addWidget(scroll)

        tabs = QTabWidget()
        iso_page = QWidget()
        ilay = QVBoxLayout(iso_page)
        self.view_iso = _PreviewView(self)
        ilay.addWidget(self.view_iso, 1)
        bar = QHBoxLayout()
        btn_fit = QPushButton("화면에 맞추기")
        btn_fit.clicked.connect(lambda: self.view_iso.fit())
        bar.addWidget(btn_fit)
        bar.addWidget(QLabel("휠로 확대·축소 · 끌어서 이동"))
        bar.addStretch(1)
        ilay.addLayout(bar)
        tabs.addTab(iso_page, "아이소매트릭")

        tbl_page = QWidget()
        tlay = QVBoxLayout(tbl_page)
        self.cmb_table = QComboBox()
        for label, key in self._TABLES:
            self.cmb_table.addItem(label, key)
        self.cmb_table.currentIndexChanged.connect(self._on_table_switch)
        top = QHBoxLayout()
        top.addWidget(QLabel("표"))
        top.addWidget(self.cmb_table)
        top.addWidget(QLabel("저장될 값 그대로입니다."))
        top.addStretch(1)
        tlay.addLayout(top)
        self.tbl = QTableWidget()
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive)
        self.tbl.itemSelectionChanged.connect(self._on_table_row)
        tlay.addWidget(self.tbl, 1)
        tabs.addTab(tbl_page, "표")

        split.addWidget(tabs)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([400, 800])
        outer.addWidget(split)

        # 보기 설정을 바꾸면 **다시 그리기만** 한다 — 최불리 선정은 그대로다.
        for w in (self.chk_iso,):
            w.toggled.connect(self._redraw)
        for w in (self.spin_zscale, self.spin_canvas, self.spin_stub):
            w.valueChanged.connect(self._redraw)
        self.cmb_ref.currentIndexChanged.connect(self._redraw)

    def _on_table_switch(self):
        if self._tables is None:
            return
        from services.cad_import.design.emit import display_tables
        view, _ = display_tables(self._tables, **self._view_opts())
        self._fill_tables(view)

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
        sched = str(self.cmb_sched.currentData() or "")
        only = self._only_heads()
        payload = self._payload
        sel = self._selected_source
        ckw = self._convert_kwargs

        def job():
            from services.cad_import.design.anchor import valve_kfp_nodes
            from services.cad_import.design.bore import extract_dia_text_points
            from services.cad_import.design.restrict import select_and_expand
            from services.cad_import.design.tables import build_design_tables

            got = select_and_expand(payload, board, k=k, only_heads=only,
                                    selected_source=sel, convert_kwargs=ckw)
            if not got.get("ok"):
                return {"ok": False, "error": got.get("error")}
            texts = self._dia_texts()
            # [§29] 종전에는 여기가 `valve_nodes=None` 이었다 — 사람이 찍은
            #   알람밸브가 기기표에 한 번도 안 실린 두 자리 중 하나다.
            av_nodes, av_missed = valve_kfp_nodes(
                got["kfp"].get("nodes_meta_runtime") or {}, board.pts,
                list(getattr(board, "valves", None) or ()),
                got.get("origin_mm"))
            if av_missed:
                print(f"[설계] ★알람밸브 {len(av_missed)}곳을 전개 노드로 "
                      f"되짚지 못했습니다 — 그만큼 기기표에서 빠집니다.")
            tbl = build_design_tables(
                got["kfp"], got["worst"], got["edge_ref"], texts,
                board_pts=board.pts,
                excluded_heads=got.get("excluded_heads", 0),
                valve_nodes=av_nodes,
                default_schedule=sched,
                tree_loads=got.get("tree_loads"),
                origin_mm=got.get("origin_mm"))
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
        self._redraw()

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

    # ── 미리보기 ────────────────────────────────────────────────────────
    def _view_opts(self) -> dict:
        """보기 설정 — 미리보기와 저장이 **같은 값**을 쓴다(§G16)."""
        return {
            "iso": self.chk_iso.isChecked(),
            "iso_z_scale": float(self.spin_zscale.value()),
            "canvas_units": float(self.spin_canvas.value()),
            "iso_ref_label": (self._valve_label()
                              if self.cmb_ref.currentData() == "valve" else None),
            "head_stub_ratio": float(self.spin_stub.value()) / 100.0,
        }

    def _loads_by_pipe(self) -> dict:
        """배관별 담당 헤드 수 — 간선 굵기가 이걸 따른다. 없으면 빈 dict."""
        got = self._result or {}
        loads = ((got.get("worst") or {}).get("loads")) or {}
        ref = got.get("edge_ref") or {}
        out = {}
        for pid, edge in ref.items():
            try:
                i, j = int(edge[0]), int(edge[1])
            except (TypeError, ValueError, IndexError):
                continue
            out[str(pid)] = int(loads.get((min(i, j), max(i, j)), 0))
        return out

    def _redraw(self):
        """설정이 바뀌면 여기만 다시 돈다 — 최불리 선정은 다시 하지 않는다."""
        if self._tables is None:
            return
        from services.cad_import.design.emit import display_tables

        view, _stood = display_tables(self._tables, **self._view_opts())
        at = {str(n.get("label")): (float(n.get("x", 0)), float(n.get("y", 0)))
              for n in view.nodes}
        elev = {str(n.get("label")): float(n.get("elevation", 0) or 0)
                for n in view.nodes}
        # 검사가 이 값을 저장된 SDF 의 Position 과 견준다(§4).
        self._preview_points = at

        sc = self.view_iso.scene()
        sc.clear()
        self._pipe_items = {}
        loads = self._loads_by_pipe()
        max_load = max(loads.values(), default=0) or 1

        # ① 배관 — 굵기는 담당 헤드 수에 비례한다. 주배관이 한눈에 보인다.
        for row in view.pipes:
            a, b = at.get(str(row.get("in"))), at.get(str(row.get("out")))
            if a is None or b is None:
                continue
            n = loads.get(str(row.get("label")), 0)
            w = 1.0 + 5.0 * (n / max_load)
            pen = QPen(QColor("#334155"), w)
            pen.setCapStyle(Qt.RoundCap)
            # Qt 는 y 가 아래로 자란다 — 도면 위아래가 뒤집히지 않게 음수로 그린다.
            it = sc.addLine(a[0], -a[1], b[0], -b[1], pen)
            it.setToolTip(f"{row.get('label')} · {row.get('dia')}A · "
                          f"{row.get('length')} m · 담당 {n}")
            self._pipe_items[str(row.get("label"))] = (it, pen)

        # ② 헤드 — 상향 △ / 하향 ▽. 방향 규칙은 베이크와 **같다**(표고 차 부호).
        parent_of = {}
        for row in view.pipes:
            a, b = str(row.get("in")), str(row.get("out"))
            parent_of.setdefault(b, a)
            parent_of.setdefault(a, b)
        head_pen = QPen(QColor("#b91c1c"), 1.2)
        head_br = QBrush(QColor("#fecaca"))
        for row in view.nozzles:
            lab = str(row.get("in"))
            p = at.get(lab)
            if p is None:
                continue
            up = elev.get(lab, 0.0) - elev.get(parent_of.get(lab, ""), 0.0) >= 0
            s = 9.0
            y = -p[1]
            tri = (QPolygonF([QPointF(p[0], y - s), QPointF(p[0] - s, y + s * 0.6),
                              QPointF(p[0] + s, y + s * 0.6)])
                   if up else
                   QPolygonF([QPointF(p[0], y + s), QPointF(p[0] - s, y - s * 0.6),
                              QPointF(p[0] + s, y - s * 0.6)]))
            sc.addPolygon(tri, head_pen, head_br).setToolTip(
                f"헤드 {lab} · {'상향식' if up else '하향식'}")

        # ③ 급수원과 알람밸브 — 어디서 물이 들어오는지 먼저 보여야 한다.
        for n in view.nodes:
            if str(n.get("io_node")) != "Input":
                continue
            p = at.get(str(n.get("label")))
            if p:
                sc.addEllipse(p[0] - 11, -p[1] - 11, 22, 22,
                              QPen(QColor("#1d4ed8"), 2),
                              QBrush(QColor("#bfdbfe"))).setToolTip("급수원")
        av = self._valve_label()
        if av and at.get(str(av)):
            p = at[str(av)]
            sc.addRect(p[0] - 9, -p[1] - 9, 18, 18,
                       QPen(QColor("#15803d"), 2),
                       QBrush(QColor("#bbf7d0"))).setToolTip("알람밸브 A/V")

        self.view_iso.fit()
        self._fill_tables(view)

    _TABLES = (("노드", "nodes"), ("배관", "pipes"),
               ("노즐", "nozzles"), ("부속", "fittings"))
    _COL_NAMES = {"label": "이름", "in": "시작", "out": "끝", "type": "관종",
                  "dia": "호칭경(mm)", "length": "길이(m)", "elev": "표고차(m)",
                  "dia_src": "관경 근거", "elevation": "표고(m)",
                  "io_node": "입출력", "flow_lmin": "유량(L/min)",
                  "count": "개수", "pipe": "배관", "x": "x", "y": "y"}
    # 근거는 사람 말로 보여 준다 — 「nfpc_fallback」은 화면에 쓸 말이 아니다.
    _BORE_SRC = {"text": "도면 텍스트", "nfpc_min": "별표1 보강(텍스트<최소)",
                 "nfpc_fallback": "별표1 (텍스트 없음)"}

    def _fill_tables(self, view):
        """저장될 값 그대로 보여 준다. 배관표에는 관경 근거를 함께 둔다."""
        which = self.cmb_table.currentData() or "nodes"
        rows = list(getattr(view, which, []) or [])
        cols: list = []
        for r in rows:
            for k in r:
                if k not in cols:
                    cols.append(k)
        t = self.tbl
        t.clear()
        t.setColumnCount(len(cols))
        t.setRowCount(len(rows))
        t.setHorizontalHeaderLabels([self._COL_NAMES.get(c, str(c))
                                     for c in cols])
        for i, r in enumerate(rows):
            for j, c in enumerate(cols):
                v = r.get(c, "")
                if c == "dia_src":
                    v = self._BORE_SRC.get(str(v), v)
                t.setItem(i, j, QTableWidgetItem("" if v is None else str(v)))
        t.resizeColumnsToContents()

    def _on_table_row(self):
        """행을 고르면 그 배관을 도면에서 강조한다."""
        if (self.cmb_table.currentData() or "nodes") != "pipes":
            return
        rows = self.tbl.selectionModel().selectedRows() if self.tbl.selectionModel() else []
        picked = {self.tbl.item(ix.row(), 0).text() for ix in rows
                  if self.tbl.item(ix.row(), 0)}
        for lab, (it, pen) in self._pipe_items.items():
            if lab in picked:
                hp = QPen(QColor("#ea580c"), max(pen.widthF() + 2.0, 3.0))
                hp.setCapStyle(Qt.RoundCap)
                it.setPen(hp)
                it.setZValue(2)
            else:
                it.setPen(pen)
                it.setZValue(0)


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
            # ★미리보기와 **같은** 설정으로 낸다. 두 곳에 적으면 언젠가 어긋나고,
            #   그러면 화면에서 본 것과 파일이 달라진다(§G16).
            out = emit_design_sdf(self._tables, path,
                                  project_title=f"{key} 수리계산 입력",
                                  **self._view_opts())
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

    def _valve_label(self):
        """알람밸브 노드의 표 라벨. 손질에서 안 찍었으면 None(표고 중앙으로)."""
        for row in (self._tables.equipment if self._tables else ()):
            if str(row.get("desc")) == "A/V":
                return row.get("in")
        return None

    # 창을 닫았다 다시 열어도 직전 결과를 쓰기 위한 접근자
    @property
    def result(self):
        return {"worst": self._result, "tables": self._tables}
