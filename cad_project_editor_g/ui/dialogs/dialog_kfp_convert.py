# 파일명: ui/dialogs/dialog_kfp_convert.py
"""KFP 변환 폼 — 입력·표시만. 분류·Z·메인/가지 판정 없음."""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
if __package__ in (None, "") and str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette, QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QFormLayout, QGridLayout,
    QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton,
    QVBoxLayout, QWidget,
)

from services.cad_import.convert.engine import (
    BLOCKER_PLANAR_MISSING, convert_to_kfp, ensure_planar)
from services.cad_import.convert.planar import pick_convert_sources
from services.cad_import.convert.preflight import (
    BLOCKER_UNCONFIRMED_HEADS, preflight_kfp_convert)
from services.cad_import.kinds import CONFIRMED_KINDS, normalize_head_kind
from services.cad_import.dto import (
    BRANCH_DEFAULT_M, COMBO_1_DEFAULT_M, COMBO_2_DEFAULT_M,
    COMBO_3_DEFAULT_M, COMBO_UP_DEFAULT_M, HEAD_K_DEFAULT,
    PENDANT_2_DEFAULT_M, PENDANT_DEFAULT_M, UPRIGHT_DEFAULT_M,
    VALVE_1_DEFAULT_M, VALVE_2_DEFAULT_M, default_dto, dto_to_convert_kwargs)
from services.i18n_service import I18nService, _t
from ui.widgets.select_all_spinbox import SelectAllDoubleSpinBox

_DIAGRAM_DIR = Path(__file__).resolve().parents[2] / "docs" / "import"
_DIAGRAMS = {
    "branch": "_kfp_sample_가지.png",
    "upright": "_kfp_sample_상향식.png",
    "pendant": "_kfp_sample_하향식.png",
    "combo": "_kfp_sample_상하향식.png",
    "valve": "_kfp_sample_알람밸브.png",
}
_DIAGRAM_W = 200


def _ensure_i18n() -> None:
    if _t("cad.convert.title") == "cad.convert.title":
        I18nService.instance().load("ko")


def _head_k_sort_key(item: dict) -> float:
    try:
        return float(item.get("K_SI", item.get("K_val", 0.0)))
    except (TypeError, ValueError):
        return float("inf")


def _head_library_items(parent=None) -> list[dict]:
    editor = getattr(parent, "editor", None)
    if editor is not None and hasattr(editor, "get_head_items"):
        items = list(editor.get_head_items() or [])
        if items:
            return items
    # ★대비책은 **없을 수도 있는** 모듈에 기댄다(`misc_controller` 는 이 트리에
    #   없다). 편집기가 헤드를 하나도 안 돌려주면 여기서 ImportError 로 창이
    #   통째로 죽었다 — 목록이 비는 것과 창이 안 뜨는 것은 다른 일이다.
    try:
        from ui.controllers.misc_controller import MiscController
        lib = MiscController.load_library_json("nozzle_library.json") or {}
    except Exception as exc:      # noqa: BLE001
        print(f"[변환] 헤드 라이브러리를 읽지 못했습니다 — 목록을 비웁니다: {exc}")
        return []
    for cat in lib.get("categories", []):
        if cat.get("category_id") == "head":
            return list(cat.get("items", []))
    return []


def _mix_color(a: QColor, b: QColor, t: float) -> QColor:
    return QColor(
        int(a.red() + (b.red() - a.red()) * t),
        int(a.green() + (b.green() - a.green()) * t),
        int(a.blue() + (b.blue() - a.blue()) * t),
    )


def _fade_pair(widget):
    pal = widget.palette()
    on = pal.color(QPalette.Active, QPalette.WindowText)
    bg = pal.color(QPalette.Active, QPalette.Window)
    return on, _mix_color(on, bg, 0.78)


def _check_fade_style(widget) -> str:
    """체크 안 한 칸의 글자를 흐리게 하는 스타일. 색은 팔레트에서 계산한다.

    ★네모 표시기(indicator)는 여기서 칠하지 않는다. 예전에는 light.qss 규칙을
    그대로 베껴 적었는데, 그 바람에 ①다크에서도 밝은 테마 색(흰 바탕)이 나오고
    ②베낄 때 :disabled 한 줄을 빠뜨려 꺼진 칸이 켠 칸과 똑같이 보였다
    (밝은·어두운 테마 둘 다, 실측 픽셀 차이 0). 표시기는 테마에 맡긴다 —
    light.qss·dark.qss 가 :disabled 까지 갖추고 있다.
    """
    on, off = _fade_pair(widget)
    return (
        f"QCheckBox:!checked {{ color: {off.name()}; }}"
        f"QCheckBox:checked {{ color: {on.name()}; }}"
    )


def _circled_num(n: int) -> QLabel:
    lbl = QLabel(str(n))
    lbl.setFixedSize(22, 22)
    lbl.setAlignment(Qt.AlignCenter)
    on, off = _fade_pair(lbl)
    lbl.setStyleSheet(
        f"QLabel {{ border: 1.5px solid {on.name()}; border-radius: 11px; "
        f"background: palette(base); color: {on.name()}; font-weight: bold; }}"
        f"QLabel:disabled {{ border-color: {off.name()}; color: {off.name()}; }}"
    )
    return lbl


def _make_len_spin(value: float) -> SelectAllDoubleSpinBox:
    w = SelectAllDoubleSpinBox()
    w.setRange(0.0, 99.0)
    w.setDecimals(2)
    w.setSingleStep(0.1)
    w.setSuffix(" m")
    w.setValue(float(value))
    w.setButtonSymbols(SelectAllDoubleSpinBox.NoButtons)
    w.setMaximumWidth(90)
    return w


def _screen_dpr() -> float:
    app = QApplication.instance()
    if app is None:
        return 1.0
    screen = app.primaryScreen()
    return float(screen.devicePixelRatio()) if screen is not None else 1.0


def _diagram_label(kind: str) -> QLabel:
    lbl = QLabel()
    path = _DIAGRAM_DIR / _DIAGRAMS[kind]
    if path.is_file():
        pix = QPixmap(str(path))
        if not pix.isNull():
            dpr = _screen_dpr()
            target = max(1, int(_DIAGRAM_W * dpr))
            if pix.width() > target:
                pix = pix.scaledToWidth(target, Qt.SmoothTransformation)
            pix.setDevicePixelRatio(dpr)
            lbl.setPixmap(pix)
    lbl.setAlignment(Qt.AlignCenter)
    return lbl


def _parse_optional(text: str):
    s = (text or "").strip().replace(",", ".")
    if not s:
        return None
    return float(s)


def _block_message(result) -> str:
    for b in result.get("blockers") or []:
        if b.get("code") == BLOCKER_UNCONFIRMED_HEADS:
            return _t("cad.convert.unconfirmed")
    blockers = result.get("blockers") or []
    if blockers and blockers[0].get("message"):
        return blockers[0]["message"]
    return _t("cad.convert.unconfirmed")


def present_convert_kinds(payload):
    """손질 화면에 보이는 확정 헤드 종류 · 밸브 찍힘."""
    payload = payload or {}
    present = set()
    for kind in payload.get("disk_kinds") or ():
        k = normalize_head_kind(kind)
        if k in CONFIRMED_KINDS:
            present.add(k)
    for rec in payload.get("kind_overrides") or ():
        if not isinstance(rec, dict):
            continue
        k = normalize_head_kind(rec.get("kind"))
        if k in CONFIRMED_KINDS:
            present.add(k)
    for rec in payload.get("head_kinds") or ():
        if not isinstance(rec, dict):
            continue
        k = normalize_head_kind(rec.get("kind"))
        if k in CONFIRMED_KINDS:
            present.add(k)
    if payload.get("valve_picks"):
        present.add("밸브")
    return present


def try_convert(payload, dto, selected_source=None):
    """preflight 실패면 파일을 쓰지 않는다. 성공이면 메모리 kfp."""
    payload = dict(payload or {})
    if selected_source is not None:
        payload["selected_source"] = selected_source
    srcs = payload.get("sources") or ()
    if len(srcs) > 1:
        picked, err = pick_convert_sources(srcs, selected_source)
        if err:
            return {
                "ok": False, "path": None, "kfp": None,
                "blockers": [{"code": err[0], "message": err[1]}],
            }
        payload["sources"] = picked
    pf = preflight_kfp_convert(payload)
    empty = {
        "ok": False, "path": None, "kfp": None, "preflight": pf,
        "blockers": list(pf["blockers"]),
        "diagnostics": list(pf.get("diagnostics") or []),
        "stats": None,
    }
    if not pf["ok"]:
        return empty
    payload = ensure_planar(payload)
    if payload.get("kfp") is None and not payload.get("kfp_path"):
        empty["blockers"] = [{
            "code": BLOCKER_PLANAR_MISSING,
            "message": payload.get("_planar_error") or "평면 그래프 .kfp 가 없습니다.",
        }]
        return empty
    return convert_to_kfp(payload, None, **dto_to_convert_kwargs(dto))


# [G17] 산출물 선택은 이 실행 동안 기억한다 — 매번 같은 것을 다시 고르지 않게.
OUTPUT_CHOICE = {"kfp": True, "sdf": True}


class KfpConvertDialog(QDialog):
    def __init__(self, parent=None, payload=None, multi_heads=None):
        super().__init__(parent)
        _ensure_i18n()
        self.payload = payload
        self._multi_heads = list(multi_heads or ())
        self.result_kfp = None
        self.result = {"ok": False, "reason": "back", "path": None, "kfp": None}
        self.setWindowTitle(_t("cad.convert.title"))
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(12)

        root.addWidget(self._build_branch())

        grid = QGridLayout()
        grid.setSpacing(8)
        self.box_upright = self._kind_box(
            _t("cad.convert.group_upright"), "upright",
            [("upright_1_m", _t("cad.convert.seg_1"), UPRIGHT_DEFAULT_M)],
        )
        self.box_pendant = self._kind_box(
            _t("cad.convert.group_pendant"), "pendant",
            [("pendant_1_m", _t("cad.convert.seg_1"), PENDANT_DEFAULT_M),
             ("pendant_2_m", _t("cad.convert.seg_2"), PENDANT_2_DEFAULT_M)],
        )
        self.box_combo = self._kind_box(
            _t("cad.convert.group_combo"), "combo",
            [("combo_1_m", _t("cad.convert.seg_1"), COMBO_1_DEFAULT_M),
             ("combo_up_m", _t("cad.convert.seg_2"), COMBO_UP_DEFAULT_M),
             ("combo_2_m", _t("cad.convert.seg_3"), COMBO_2_DEFAULT_M),
             ("combo_3_m", _t("cad.convert.seg_4"), COMBO_3_DEFAULT_M)],
        )
        self.box_valve = self._kind_box(
            _t("cad.convert.group_valve"), "valve",
            [("valve_1_m", _t("cad.convert.seg_1"), VALVE_1_DEFAULT_M),
             ("valve_2_m", _t("cad.convert.seg_2"), VALVE_2_DEFAULT_M)],
        )
        grid.addWidget(self.box_upright, 0, 0)
        grid.addWidget(self.box_pendant, 0, 1)
        grid.addWidget(self.box_combo, 1, 0)
        grid.addWidget(self.box_valve, 1, 1)
        root.addLayout(grid)

        bottom = QHBoxLayout()
        bottom.addWidget(self._build_flex(), stretch=1)
        bottom.addWidget(self._build_sprinkler(), stretch=1)
        root.addLayout(bottom)

        self._source_row = self._build_source_row()
        root.addWidget(self._source_row)
        root.addWidget(self._build_outputs())
        self._fill_sources()
        self._apply_present_kinds()

        btns = QHBoxLayout()
        btns.addStretch(1)
        self.btn_back = QPushButton(_t("cad.convert.back"))
        self.btn_convert = QPushButton(_t("cad.convert.convert"))
        self.btn_back.clicked.connect(self._on_back)
        self.btn_convert.clicked.connect(self._on_convert)
        btns.addWidget(self.btn_back)
        btns.addWidget(self.btn_convert)
        root.addLayout(btns)

    def _build_outputs(self) -> QGroupBox:
        """[G17] 산출물 고르기 — `.kfp` 저장을 **강제로 통과시키지 않는다**.

        지금까지는 변환이 끝나면 무조건 `.kfp` 저장 대화상자와 완료 알림을 거친
        뒤에야 수리계산 입력 창이 떴다. SDF 만 필요한 사람에게는 불필요한 문이다.

        고른 값은 세션에 남는다 — 매번 같은 것을 다시 고르게 하지 않는다.
        """
        box = QGroupBox(_t("cad.convert.group_outputs"))
        row = QHBoxLayout(box)
        self.chk_out_kfp = QCheckBox(_t("cad.convert.out_kfp"))
        self.chk_out_sdf = QCheckBox(_t("cad.convert.out_sdf"))
        want = dict(OUTPUT_CHOICE)
        self.chk_out_kfp.setChecked(bool(want.get("kfp", True)))
        self.chk_out_sdf.setChecked(bool(want.get("sdf", True)))
        row.addWidget(self.chk_out_kfp)
        row.addWidget(self.chk_out_sdf)
        row.addStretch(1)
        return box

    def _outputs(self) -> dict:
        return {"kfp": bool(self.chk_out_kfp.isChecked()),
                "sdf": bool(self.chk_out_sdf.isChecked())}


    def _build_branch(self) -> QGroupBox:
        box = QGroupBox(_t("cad.convert.group_branch"))
        row = QHBoxLayout(box)
        form = QFormLayout()
        self.spin_branch = _make_len_spin(BRANCH_DEFAULT_M)
        form.addRow(_t("cad.convert.branch_rise"), self.spin_branch)
        row.addLayout(form)
        row.addStretch(1)
        row.addWidget(_diagram_label("branch"))
        return box

    def _kind_box(self, title, kind, fields) -> QGroupBox:
        box = QGroupBox(title)
        row = QHBoxLayout(box)
        row.addWidget(_diagram_label(kind))
        form = QFormLayout()
        for key, label, default in fields:
            spin = _make_len_spin(default)
            setattr(self, f"spin_{key}", spin)
            form.addRow(label, spin)
        row.addLayout(form)
        return box

    def _build_flex(self) -> QGroupBox:
        box = QGroupBox()
        vbox = QVBoxLayout(box)
        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        self.chk_flex = QCheckBox(_t("cad.convert.group_flex"))
        self.chk_flex.setChecked(False)
        self.chk_flex.setStyleSheet(_check_fade_style(self.chk_flex))
        header.addWidget(self.chk_flex)
        header.addSpacing(12)
        self._flex_hint = QWidget()
        hint = QHBoxLayout(self._flex_hint)
        hint.setContentsMargins(0, 0, 0, 0)
        hint.addWidget(QLabel(_t("cad.convert.flex_pendant")))
        self.lbl_flex_2 = _circled_num(2)
        hint.addWidget(self.lbl_flex_2)
        hint.addSpacing(8)
        hint.addWidget(QLabel(_t("cad.convert.flex_combo")))
        self.lbl_flex_4 = _circled_num(4)
        hint.addWidget(self.lbl_flex_4)
        hint.addStretch(1)
        header.addWidget(self._flex_hint)
        _, off = _fade_pair(self._flex_hint)
        self._flex_hint.setStyleSheet(
            f"QLabel:disabled {{ color: {off.name()}; }}")
        vbox.addLayout(header)
        self._flex_form = QWidget()
        form = QFormLayout(self._flex_form)
        form.setContentsMargins(0, 0, 0, 0)
        self.edit_flex_c = QLineEdit()
        self.edit_flex_c.setMaximumWidth(90)
        self.edit_flex_rough = QLineEdit()
        self.edit_flex_rough.setMaximumWidth(90)
        form.addRow(_t("cad.convert.flex_c"), self.edit_flex_c)
        form.addRow(_t("cad.convert.flex_rough"), self.edit_flex_rough)
        _, off_form = _fade_pair(self._flex_form)
        self._flex_form.setStyleSheet(
            f"QLabel:disabled {{ color: {off_form.name()}; }}")
        vbox.addWidget(self._flex_form)
        self.chk_flex.toggled.connect(self._flex_hint.setEnabled)
        self.chk_flex.toggled.connect(self._flex_form.setEnabled)
        self._flex_hint.setEnabled(False)
        self._flex_form.setEnabled(False)
        self.box_flex = box
        return box

    def _build_sprinkler(self) -> QGroupBox:
        box = QGroupBox(_t("cad.convert.group_sprinkler"))
        form = QFormLayout(box)
        self.combo_head = QComboBox()
        items = sorted(_head_library_items(self.parent()), key=_head_k_sort_key)
        for item in items:
            name = str(item.get("display_name", "") or "")
            if not name:
                continue
            try:
                k_si = float(item.get("K_SI", item.get("K_val", HEAD_K_DEFAULT)))
            except (TypeError, ValueError):
                k_si = HEAD_K_DEFAULT
            self.combo_head.addItem(name, k_si)
        if self.combo_head.count() == 0:
            self.combo_head.addItem(str(int(HEAD_K_DEFAULT)), HEAD_K_DEFAULT)
        idx = self.combo_head.findData(HEAD_K_DEFAULT)
        if idx < 0:
            for i in range(self.combo_head.count()):
                if abs(float(self.combo_head.itemData(i)) - HEAD_K_DEFAULT) < 1e-9:
                    idx = i
                    break
        if idx >= 0:
            self.combo_head.setCurrentIndex(idx)
        form.addRow(_t("cad.convert.head_k"), self.combo_head)
        self.chk_active = QCheckBox(_t("cad.convert.active"))
        self.chk_active.setChecked(False)
        self.chk_active.setStyleSheet(_check_fade_style(self.chk_active))
        form.addRow("", self.chk_active)
        return box

    def _build_source_row(self) -> QWidget:
        row = QWidget()
        lay = QFormLayout(row)
        lay.setContentsMargins(0, 0, 0, 0)
        self.combo_source = QComboBox()
        lay.addRow(_t("cad.convert.source"), self.combo_source)
        row.hide()
        return row

    def _fill_sources(self) -> None:
        srcs = [
            s for s in (self.payload or {}).get("sources") or ()
            if isinstance(s, dict) and s.get("xy") is not None
        ]
        if len(srcs) <= 1:
            self._source_row.hide()
            return
        self.combo_source.clear()
        for i, src in enumerate(srcs):
            tag = str(src.get("tag") or f"Z{i + 1}")
            self.combo_source.addItem(tag, tag)
        self._source_row.show()

    def _apply_present_kinds(self) -> None:
        present = present_convert_kinds(self.payload)
        self.box_upright.setEnabled("상향식" in present)
        self.box_pendant.setEnabled("하향식" in present)
        self.box_combo.setEnabled("상하향식" in present)
        self.box_valve.setEnabled("밸브" in present)
        self.box_flex.setEnabled("하향식" in present or "상하향식" in present)

    def read_dto(self) -> dict:
        dto = default_dto()
        dto["branch_rise_m"] = self.spin_branch.value()
        dto["upright_1_m"] = self.spin_upright_1_m.value()
        dto["pendant_1_m"] = self.spin_pendant_1_m.value()
        dto["pendant_2_m"] = self.spin_pendant_2_m.value()
        dto["combo_1_m"] = self.spin_combo_1_m.value()
        dto["combo_up_m"] = self.spin_combo_up_m.value()
        dto["combo_2_m"] = self.spin_combo_2_m.value()
        dto["combo_3_m"] = self.spin_combo_3_m.value()
        dto["valve_1_m"] = self.spin_valve_1_m.value()
        dto["valve_2_m"] = self.spin_valve_2_m.value()
        k = self.combo_head.currentData()
        dto["head_k"] = float(k) if k is not None else HEAD_K_DEFAULT
        name = str(self.combo_head.currentText() or "").strip()
        dto["head_spec_name"] = name or None
        dto["head_active"] = self.chk_active.isChecked()
        if self.chk_flex.isChecked():
            dto["flex_c"] = _parse_optional(self.edit_flex_c.text())
            dto["flex_roughness_mm"] = _parse_optional(
                self.edit_flex_rough.text())
        else:
            dto["flex_c"] = None
            dto["flex_roughness_mm"] = None
        return dto

    def _selected_source(self):
        if self._source_row.isHidden():
            return None
        return self.combo_source.currentData()

    def _on_back(self) -> None:
        self.result = {"ok": False, "reason": "back", "path": None, "kfp": None}
        self.reject()

    def _on_convert(self) -> None:
        # ★산출물 확인이 **맨 앞**이다. 무거운 변환을 다 돌린 뒤에 「만들 게
        #   없습니다」라고 말하는 것은 시간을 버리게 하는 일이다(§G17).
        want = self._outputs()
        if not (want["kfp"] or want["sdf"]):
            QMessageBox.warning(self, _t("cad.convert.title"),
                                _t("cad.convert.no_outputs"))
            return
        try:
            dto = self.read_dto()
        except ValueError:
            QMessageBox.warning(
                self, _t("cad.convert.title"), _t("cad.convert.err_number"))
            return
        if not self.payload:
            QMessageBox.warning(
                self, _t("cad.convert.title"), _t("cad.convert.err_no_graph"))
            return
        OUTPUT_CHOICE.update(want)

        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()
        try:
            r = try_convert(self.payload, dto, self._selected_source())
        finally:
            while QApplication.overrideCursor() is not None:
                QApplication.restoreOverrideCursor()
        if not r["ok"]:
            QMessageBox.warning(
                self, _t("cad.convert.title"), _block_message(r))
            if any(b.get("code") == BLOCKER_UNCONFIRMED_HEADS
                   for b in r.get("blockers") or []):
                self.result = {
                    "ok": False, "reason": "blocked", "path": None, "kfp": None}
                self.reject()
            return
        self.result_kfp = r.get("kfp")
        self.result = {
            "ok": True, "reason": "converted",
            "path": None, "kfp": self.result_kfp,
            "outputs": want,          # [G17] 흐름이 이 값만 보고 갈린다
        }
        if self._multi_heads:
            QMessageBox.warning(
                self, _t("cad.convert.title"),
                _t("cad.convert.warn_multi_arm"))
        st = r.get("stats") or {}
        if st.get("main_walk") is False and (self.payload or {}).get("ho"):
            QMessageBox.warning(
                self, _t("cad.convert.title"),
                _t("cad.convert.warn_main_walk"))
        self.accept()


if __name__ == "__main__":
    from ui.theme_manager import apply as apply_theme
    app = QApplication.instance() or QApplication(sys.argv)
    apply_theme(app, "system")
    KfpConvertDialog().exec()
