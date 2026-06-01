from __future__ import annotations

import os
import shutil
import subprocess
import time
from pathlib import Path

import win32clipboard
import win32con
import win32com.client
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
DESKTOP = Path(os.environ.get("USERPROFILE", r"C:\Users\admin")) / "Desktop"
OUT_DIR = ROOT / "docs" / "generated_reports" / "template_v5"
OUT_DIR.mkdir(parents=True, exist_ok=True)

TEMPLATE_KR = DESKTOP / "PIPENET 수리계산 검증 프로그램 중 2. 설계 최적화 가이드 로직.hwp"
TEMPLATE_ASCII = ROOT / "templates_hwp" / "template_optimization_guide.hwp"
OUT_ASCII = DESKTOP / "PIPENET_2_optimization_guide_template_tables_v5.hwp"
OUT_KR = DESKTOP / "PIPENET 수리계산 검증 프로그램 중 2. 설계 최적화 가이드 로직_양식유지_표삽도_완성본.hwp"

FONT_REG = r"C:\Windows\Fonts\HANBatang.ttf"
FONT_BOLD = r"C:\Windows\Fonts\HANBatangB.ttf"
FONT_NAME_HWP = "함초롬바탕"


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BOLD if bold and Path(FONT_BOLD).exists() else FONT_REG, size)


def wrap(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont, width: int) -> list[str]:
    out: list[str] = []
    for raw in str(text).split("\n"):
        line = ""
        for token in raw.split(" "):
            cand = token if not line else f"{line} {token}"
            if draw.textbbox((0, 0), cand, font=fnt)[2] <= width:
                line = cand
            else:
                if line:
                    out.append(line)
                line = token
        out.append(line)
    return out


def make_table_image(title: str, headers: list[str], rows: list[list[str]], filename: str) -> Path:
    col_count = len(headers)
    width = 1800
    margin = 42
    title_h = 78
    header_h = 60
    row_min_h = 72
    col_w = (width - margin * 2) // col_count
    tmp = Image.new("RGB", (width, 200), "white")
    dtmp = ImageDraw.Draw(tmp)
    body_font = font(30)
    heights = []
    for row in rows:
        max_lines = 1
        for value in row:
            max_lines = max(max_lines, len(wrap(dtmp, value, body_font, col_w - 22)))
        heights.append(max(row_min_h, max_lines * 39 + 24))
    height = margin + title_h + header_h + sum(heights) + margin
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    draw.rectangle((0, 0, width - 1, height - 1), outline="#d1d5db", width=2)
    draw.text((margin, margin - 6), title, font=font(38, True), fill="#111827")

    y = margin + title_h
    x = margin
    for h in headers:
        draw.rectangle((x, y, x + col_w, y + header_h), fill="#f3f4f6", outline="#111827", width=2)
        lines = wrap(draw, h, font(28, True), col_w - 20)
        ty = y + (header_h - len(lines) * 34) // 2
        for line in lines:
            tw = draw.textbbox((0, 0), line, font=font(28, True))[2]
            draw.text((x + (col_w - tw) / 2, ty), line, font=font(28, True), fill="#111827")
            ty += 34
        x += col_w
    y += header_h

    for idx, row in enumerate(rows):
        row_h = heights[idx]
        x = margin
        fill = "#ffffff" if idx % 2 == 0 else "#fafafa"
        for value in row:
            draw.rectangle((x, y, x + col_w, y + row_h), fill=fill, outline="#9ca3af", width=1)
            lines = wrap(draw, value, body_font, col_w - 22)
            ty = y + 14
            for line in lines:
                draw.text((x + 11, ty), line, font=body_font, fill="#111827")
                ty += 39
            x += col_w
        y += row_h

    path = OUT_DIR / filename
    img.save(path, quality=96)
    return path


def make_flow_image(title: str, boxes: list[tuple[str, str]], filename: str, palette: list[str]) -> Path:
    width, height = 1800, 620
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle((0, 0, width - 1, height - 1), outline="#d1d5db", width=2)
    draw.text((50, 35), title, font=font(42, True), fill="#111827")
    y = 170
    box_w = 300
    gap = 52
    x = 55
    for i, (head, body) in enumerate(boxes):
        fill = palette[i % len(palette)]
        draw.rounded_rectangle((x, y, x + box_w, y + 285), radius=18, fill=fill, outline="#111827", width=2)
        draw.text((x + 22, y + 24), head, font=font(32, True), fill="#111827")
        ty = y + 82
        for line in wrap(draw, body, font(26), box_w - 44):
            draw.text((x + 22, ty), line, font=font(26), fill="#111827")
            ty += 34
        if i < len(boxes) - 1:
            sx, sy = x + box_w, y + 142
            ex, ey = x + box_w + gap - 12, y + 142
            draw.line((sx, sy, ex, ey), fill="#4b5563", width=5)
            draw.polygon([(ex, ey), (ex - 18, ey - 12), (ex - 18, ey + 12)], fill="#4b5563")
        x += box_w + gap
    path = OUT_DIR / filename
    img.save(path, quality=96)
    return path


def build_images() -> list[tuple[str, Path]]:
    images: list[tuple[str, Path]] = []
    images.append(
        (
            "전체 데이터 흐름도",
            make_flow_image(
                "설계 최적화 가이드 전체 데이터 흐름",
                [
                    ("입력", "결과서 DOCX/PDF\nSDF 배관망\n선택: meta/policy"),
                    ("파싱", "배관, 헤드, 압력,\n유속, 손실, 토폴로지"),
                    ("필수 PASS", "법적/기술 기준\n먼저 통과"),
                    ("공학/경제", "안정성 후보와\n비용절감 후보 분리"),
                    ("출력", "표, 카드, 맵,\n최종 리포트"),
                ],
                "flow_overall.png",
                ["#f8fafc", "#eef2ff", "#ecfdf5", "#fff7ed", "#fafafa"],
            ),
        )
    )
    images.append(
        (
            "공통 전제 기준표",
            make_table_image(
                "필수 PASS 전제 기준",
                ["No", "전제 기준", "판정 의미"],
                [
                    ["1", "헤드 최소 방수압 0.1 MPa 이상", "말단 방수압 기준을 만족해야 최적화 검토 가능"],
                    ["2", "헤드 최소 방수량 기준 만족", "노즐별 요구 유량 이상 확보"],
                    ["3", "가지배관 유속 6 m/s 이하", "Topology 기준 branch 배관에 적용"],
                    ["4", "그 밖의 배관 유속 10 m/s 이하", "교차배관, 주배관, 50A 초과 배관에 적용"],
                    ["5", "재질별 C-Factor 기준 만족", "KSD 계열 C=120, CPVC 계열 C=150"],
                    ["6", "특수설비/밸브 등가길이 검증", "FX, AV/PV, PRV 등 입력값 검토"],
                    ["7", "Hazen-Williams 재계산 검증", "결과서 Frict. Loss를 수식으로 재현"],
                ],
                "table_prereq.png",
            ),
        )
    )
    images.append(
        (
            "공통 계산 지표표",
            make_table_image(
                "공통 계산 지표와 수식",
                ["지표", "수식", "해석"],
                [
                    ["m당 마찰손실", "friction_loss / length", "길이 영향을 제거한 손실 집중도"],
                    ["마찰손실 변화율", "(현재 m당 손실 - 직전 m당 손실) / 직전 m당 손실", "직전 배관 대비 급증 여부"],
                    ["유속 여유율", "1 - actual_velocity / velocity_limit", "유속 기준까지 남은 여유"],
                    ["압력 여유", "actual_pressure - required_pressure", "압력 안정성 또는 축소 여지"],
                    ["등가길이 비율", "(fitting_eq + special_eq) / total_length", "피팅/특수설비 집중도"],
                    ["HW 재계산", "6.174e4 * Q^1.85 * L / (C^1.85 * D^4.87) / 0.1", "결과서 손실 재현 검산"],
                ],
                "table_formula.png",
            ),
        )
    )
    images.append(
        (
            "공학 최적화 흐름도",
            make_flow_image(
                "공학적 마찰손실 최적화 판정 흐름",
                [
                    ("지표 계산", "m당 손실\n변화율\n유속/압력 여유"),
                    ("후보 세분화", "마찰손실\n변화율\n피팅집중\n압력여유"),
                    ("Spike Map", "m당 손실 > 1.0\n전용 빨간 표시"),
                    ("원인 분류", "긴 배관인지\n짧지만 손실 집중인지"),
                    ("조치", "구경 상향\n피팅 축소\n경로 단순화"),
                ],
                "flow_engineering.png",
                ["#eff6ff", "#fefce8", "#fee2e2", "#f8fafc", "#ecfdf5"],
            ),
        )
    )
    images.append(
        (
            "공학 최적화 기준표",
            make_table_image(
                "공학적 마찰손실 최적화 후보 기준",
                ["후보 사유", "판정 조건", "권장 조치"],
                [
                    ["마찰손실 절대값 과다", "m당 마찰손실이 공학 기준 초과", "관경 상향, 피팅 축소, 경로 단순화"],
                    ["마찰손실 변화율 급증", "직전 배관 대비 변화율이 큼", "이전/현재 구간의 구경, 피팅, 특수설비 비교"],
                    ["유속 여유 부족", "유속 여유율이 낮아 기준에 근접", "관경 상향 또는 분기/경로 재구성"],
                    ["피팅/특수설비 집중", "등가길이 비율이 과다", "엘보/Tee 축소, 특수설비 위치 조정"],
                    ["압력 여유 부족", "말단 압력이 요구 기준에 근접", "손실 저감 또는 관경 상향"],
                ],
                "table_engineering.png",
            ),
        )
    )
    images.append(
        (
            "경제성 최적화 흐름도",
            make_flow_image(
                "시공사 경제성 확보 판정 흐름",
                [
                    ("후보 탐색", "저유속\n압력여유 과다\n대구경 밸브"),
                    ("축소 가정", "관경 한 단계\n축소 시뮬레이션"),
                    ("재계산", "HW 손실\n유속\n말단압력"),
                    ("통과조건", "유속 기준\n압력 기준\n유량 기준"),
                    ("제안", "관경 축소\n밸브 최적화\nCPVC 회피"),
                ],
                "flow_economy.png",
                ["#f0fdf4", "#ecfdf5", "#f8fafc", "#fff7ed", "#fafafa"],
            ),
        )
    )
    images.append(
        (
            "경제성 최적화 기준표",
            make_table_image(
                "시공사 경제성 확보 후보 기준",
                ["경제성 후보", "판정 조건", "검토 방향"],
                [
                    ["저유속 과설계", "velocity < 2.0 m/s and nominal_bore > 25A", "한 단계 관경 축소 시뮬레이션"],
                    ["압력 여유 과다", "말단압력이 최소 기준보다 과도하게 높음", "0.11~0.12 MPa 수준 목표 검토"],
                    ["대구경 밸브 비용", "valve_connected_bore > 100A", "100A 이하 조정 가능성 검토"],
                    ["CPVC 대구경", "CPVC nominal_bore > 65A", "50A/65A 복수 라인 분산 검토"],
                    ["하류 헤드 수 대비 과대 관경", "downstream_nozzle_count 대비 nominal_bore가 큼", "40A→32A, 32A→25A 등 검토"],
                ],
                "table_economy.png",
            ),
        )
    )
    images.append(
        (
            "맵 표시 기준표",
            make_table_image(
                "시각화 맵 표시 기준",
                ["맵", "표시 대상", "색상/기능"],
                [
                    ["Friction Loss Spike Map", "m당 마찰손실 > 1.0 kg/cm²/m", "빨간 배관, 드래그/호버 시 Pipe 번호와 손실 정보 표시"],
                    ["Economy Optimization Candidate Map", "저유속, 압력여유, 대구경 밸브, CPVC 대구경 등", "경제성 후보 배관 표시, 줌/드래그 지원"],
                    ["결과 데이터 테이블", "공학 후보와 경제 후보 전체", "후보 사유를 마찰손실, 변화율, 유속여유, 피팅집중, 압력여유 등으로 분리"],
                ],
                "table_maps.png",
            ),
        )
    )
    images.append(
        (
            "충돌 절충 판단표",
            make_table_image(
                "공학-경제 충돌 및 절충 판단",
                ["충돌 유형", "공학 해석", "경제 해석", "절충안"],
                [
                    ["손실 큰 배관", "구경 상향 또는 피팅 축소 필요", "구경 상향은 비용 증가", "관경 변경 전 피팅/경로부터 점검"],
                    ["저유속 대구경", "압력 안정성에는 유리", "과설계 가능성", "축소 시뮬레이션 후 기준 재확인"],
                    ["CPVC 대구경", "C=150으로 손실 저감 가능", "대구경 CPVC 비용 상승", "복수 소구경 라인 또는 재질 대안 비교"],
                    ["대구경 밸브", "손실과 유지관리 측면에서 안정", "밸브 단가 상승", "100A 이하 조정 가능성 검토"],
                ],
                "table_conflict.png",
            ),
        )
    )
    return images


def clipboard_text(text: str) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def set_hwp_font(hwp) -> None:
    try:
        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
        cs = hwp.HParameterSet.HCharShape
        cs.FaceNameHangul = FONT_NAME_HWP
        cs.FaceNameLatin = "HCR Batang"
        cs.FaceNameHanja = FONT_NAME_HWP
        cs.FaceNameJapanese = FONT_NAME_HWP
        cs.FaceNameOther = FONT_NAME_HWP
        cs.Height = 1000
        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
    except Exception:
        pass


def paste_text(hwp, text: str) -> None:
    set_hwp_font(hwp)
    clipboard_text(text)
    hwp.HAction.Run("Paste")
    time.sleep(0.05)


def insert_picture(hwp, path: Path) -> None:
    try:
        hwp.InsertPicture(str(path), True, 1, False, False, 0, 150, 0)
    except Exception:
        # Fallback through InsertPicture action for versions where direct API differs.
        hwp.HAction.GetDefault("InsertPicture", hwp.HParameterSet.HInsertPicture.HSet)
        hwp.HParameterSet.HInsertPicture.filename = str(path)
        hwp.HParameterSet.HInsertPicture.Embedded = True
        hwp.HAction.Execute("InsertPicture", hwp.HParameterSet.HInsertPicture.HSet)
    hwp.HAction.Run("BreakPara")


def build_report_text() -> str:
    return (
        "\n\n"
        "2. 설계 최적화 가이드 로직 보고서\n"
        "\n"
        "본 문서는 PIPENET 수리계산 검증 프로그램의 2. 설계 최적화 가이드가 어떤 입력값을 읽고, 어떤 수식과 판정 기준을 거쳐 공학적 최적화 후보와 시공사 경제성 후보를 제시하는지 정리한 문서이다.\n"
        "\n"
        "설계 최적화는 법적 기준을 무시하고 관경을 줄이는 작업이 아니다. 헤드 방수압, 방수량, 배관 유속, 재질별 C-Factor, 특수설비 등가길이, Hazen-Williams 재계산 검증이 먼저 PASS 되어야 한다. 그 이후에 공학적 안정성 개선 후보와 경제성 개선 후보를 별도로 제시한다.\n"
        "\n"
    )


def build_hwp(images: list[tuple[str, Path]]) -> None:
    if not TEMPLATE_KR.exists():
        raise FileNotFoundError(TEMPLATE_KR)
    shutil.copy2(TEMPLATE_KR, TEMPLATE_ASCII)
    shutil.copy2(TEMPLATE_ASCII, OUT_ASCII)

    hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        if not hwp.Open(str(OUT_ASCII)):
            raise RuntimeError(f"Cannot open template copy: {OUT_ASCII}")
        hwp.HAction.Run("MoveDocEnd")
        set_hwp_font(hwp)
        paste_text(hwp, build_report_text())
        for caption, img_path in images:
            paste_text(hwp, f"\n{caption}\n")
            insert_picture(hwp, img_path)
        paste_text(
            hwp,
            "\n적용상 주의사항\n"
            "1. Friction Loss Spike Map은 m당 마찰손실 > 1.0 kg/cm²/m 조건만 빨간색으로 표시한다.\n"
            "2. 결과 데이터 테이블의 공학 최적화 후보는 마찰손실, 변화율, 유속여유, 피팅집중, 압력여유를 모두 포함하므로 Spike Map보다 후보 수가 많을 수 있다.\n"
            "3. Economy Optimization Candidate Map은 비용 절감 후보를 표시하며, 최종 축소 여부는 Hazen-Williams 재계산, 유속 기준, 말단압력, 헤드 유량 기준을 다시 통과해야 한다.\n"
            "4. project_meta.json과 design_policy.json이 없으면 세대 내부/세대 유입/정책성 관경 기준은 REVIEW로 남는다.\n",
        )
        hwp.HAction.Run("SelectAll")
        set_hwp_font(hwp)
        hwp.HAction.Run("MoveDocEnd")
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(OUT_ASCII)
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        if not hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet):
            raise RuntimeError("HWP save failed")
    finally:
        hwp.Quit()
    shutil.copy2(OUT_ASCII, OUT_KR)


def validate() -> None:
    print(f"OUT_ASCII={OUT_ASCII} exists={OUT_ASCII.exists()} size={OUT_ASCII.stat().st_size if OUT_ASCII.exists() else 0}")
    print(f"OUT_KR={OUT_KR} exists={OUT_KR.exists()} size={OUT_KR.stat().st_size if OUT_KR.exists() else 0}")
    if OUT_ASCII.exists():
        text = subprocess.check_output(["hwp5txt.exe", str(OUT_ASCII)], text=True, encoding="utf-8", errors="replace", timeout=30)
        markers = ["Analysis Report", "문서 목적", "목차", "설계 최적화", "Friction Loss Spike Map"]
        print("MARKERS=" + ", ".join(f"{m}:{m in text}" for m in markers))
        print("QUESTION_MARK_COUNT=", text.count("?"))
    hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        print("COM_OPEN=", bool(hwp.Open(str(OUT_ASCII))))
    finally:
        hwp.Quit()


if __name__ == "__main__":
    imgs = build_images()
    build_hwp(imgs)
    validate()
