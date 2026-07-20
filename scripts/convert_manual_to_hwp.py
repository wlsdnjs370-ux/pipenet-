# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

import win32com.client
from docx import Document


def count_hangul(text: str) -> int:
    return sum(1 for ch in text if 0xAC00 <= ord(ch) <= 0xD7A3)


def read_docx_text(path: Path) -> str:
    doc = Document(str(path))
    lines = [(p.text or "").strip() for p in doc.paragraphs]
    return "\r\n".join(x for x in lines if x)


def pick_best_source(desktop: Path) -> tuple[Path, str]:
    candidates = sorted(desktop.glob("PIPENET*.docx"))
    if not candidates:
        raise FileNotFoundError("Desktop에서 PIPENET*.docx 파일을 찾지 못했습니다.")

    scored: list[tuple[int, Path, str]] = []
    for p in candidates:
        text = read_docx_text(p)
        scored.append((count_hangul(text), p, text))

    scored.sort(key=lambda x: x[0], reverse=True)
    best_count, best_path, best_text = scored[0]
    if best_count <= 0:
        raise ValueError("한글 텍스트가 포함된 DOCX를 찾지 못했습니다.")
    return best_path, best_text


def save_hwp(text: str, output_path: Path) -> None:
    if output_path.exists():
        output_path.unlink()

    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
    try:
        hwp.XHwpWindows.Item(0).Visible = False
    except Exception:
        pass

    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = str(output_path)
    hwp.HParameterSet.HFileOpenSave.Format = "HWP"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    ok = hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    if not ok:
        hwp.Quit()
        raise RuntimeError("HWP 저장에 실패했습니다.")
    hwp.Quit()


def main() -> None:
    desktop = Path.home() / "Desktop"
    src_path, text = pick_best_source(desktop)
    out_path = desktop / "PIPENET_수리계산_검증프로그램_설명서.hwp"
    save_hwp(text, out_path)
    print(f"source={src_path}")
    print(f"hangul_chars={count_hangul(text)}")
    print(f"output={out_path}")


if __name__ == "__main__":
    main()
