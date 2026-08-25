# -*- coding: utf-8 -*-
"""[G18] 미리보기 화면을 그림으로 남긴다 — 문서에 붙일 증빙.

PIPENET 은 여기서 띄울 수 없다. 대신 **저장될 좌표를 그대로 그린** 미리보기를
남긴다(미리보기 좌표 == SDF Position 은 `tests/test_design_dialog.py` 가 증명한다).

    QT_QPA_PLATFORM=offscreen python tests/_capture_preview.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_REPO = _ROOT.parent
for p in (str(_ROOT), str(_REPO)):
    if p not in sys.path:
        sys.path.insert(0, p)
os.chdir(_ROOT)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

KEY = "B1F 현장조사 소화설비 평면도"
IMG = _REPO / "docs" / "images"


def main() -> int:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QImage, QPainter
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    # ★offscreen 은 한글 글꼴을 못 찾아 글자가 전부 네모로 찍힌다. 증빙 그림에
    #   글자가 안 보이면 증빙이 아니다 — 있는 글꼴 중 하나를 명시한다.
    from PySide6.QtGui import QFont, QFontDatabase
    # offscreen 은 시스템 글꼴을 하나도 안 싣는다(families() 가 빈 목록).
    # 파일에서 직접 등록해야 한다.
    picked = None
    for f in ("C:/Windows/Fonts/malgun.ttf", "C:/Windows/Fonts/gulim.ttc",
              "/usr/share/fonts/truetype/nanum/NanumGothic.ttf"):
        if not os.path.isfile(f):
            continue
        fid = QFontDatabase.addApplicationFont(f)
        fams = QFontDatabase.applicationFontFamilies(fid) if fid >= 0 else []
        if fams:
            picked = fams[0]
            app.setFont(QFont(picked, 9))
            break
    print(f"글꼴: {picked or '못 찾음 — 글자가 네모로 찍힐 수 있습니다'}")
    from services.cad_import.edit.session import EditSession
    from ui.dialogs.dialog_design_input import DesignInputDialog

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None
    dlg = DesignInputDialog(None, session=es, payload=payload,
                            selected_source=sel, k=30)
    dlg.resize(1400, 860)
    dlg.show()
    dlg.chk_iso.setChecked(True)
    dlg._on_run()
    IMG.mkdir(parents=True, exist_ok=True)

    # ① 아이소매트릭 — 장면만 크게 뽑는다(창 테두리는 증빙에 도움이 안 된다).
    sc = dlg.view_iso.scene()
    r = sc.itemsBoundingRect().adjusted(-40, -40, 40, 40)
    h = int(1400 * r.height() / max(r.width(), 1)) or 700
    img = QImage(1400, h, QImage.Format_ARGB32)
    img.fill(Qt.white)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing, True)
    sc.render(p, target=img.rect(), source=r)
    p.end()
    out1 = IMG / "module_g_preview_iso.png"
    img.save(str(out1))

    # ② 표 — 관종·호칭경·관경 근거가 행마다 채워졌는지 보이는 장면.
    dlg.cmb_table.setCurrentIndex(1)
    dlg._on_table_switch()
    QApplication.processEvents()
    out2 = IMG / "module_g_preview_table.png"
    dlg.tbl.grab().save(str(out2))

    for o in (out1, out2):
        print(f"{o.relative_to(_REPO)} · {o.stat().st_size:,} bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
