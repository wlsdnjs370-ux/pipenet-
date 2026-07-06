from __future__ import annotations

import os
import shutil
from pathlib import Path

import fitz


ROOT = Path(__file__).resolve().parents[1]
DESKTOP = Path(os.environ.get("USERPROFILE", r"C:\Users\admin")) / "Desktop"
OUT_ASCII = DESKTOP / "PIPENET_hydraulic_calculation_guide_v3_screenshots.pdf"
OUT_KR = DESKTOP / "PIPENET 수리계산 검증 프로그램 가이드_v3_화면캡처포함.pdf"
SCREENSHOT_DIR = ROOT / "docs" / "generated_reports" / "server_screenshots"

PAGE_W = 841.92
PAGE_H = 594.96
SIDEBAR_X = 0
SIDEBAR_W = 222
CONTENT_X = SIDEBAR_X + SIDEBAR_W + 34
CONTENT_R = PAGE_W - 34
MARGIN = CONTENT_X

FONT = r"C:\Windows\Fonts\HANBatang.ttf"
FONT_BOLD = r"C:\Windows\Fonts\HANBatangB.ttf"

COLORS = {
    "navy": (0.07, 0.11, 0.18),
    "blue": (0.08, 0.31, 0.72),
    "sky": (0.88, 0.94, 1.0),
    "green": (0.05, 0.48, 0.30),
    "green_bg": (0.90, 0.98, 0.94),
    "orange": (0.75, 0.30, 0.06),
    "orange_bg": (1.0, 0.95, 0.88),
    "red": (0.75, 0.10, 0.10),
    "red_bg": (1.0, 0.92, 0.92),
    "yellow_bg": (1.0, 0.98, 0.84),
    "gray": (0.95, 0.96, 0.98),
    "line": (0.78, 0.81, 0.86),
    "text": (0.08, 0.10, 0.14),
    "muted": (0.34, 0.38, 0.45),
    "white": (1, 1, 1),
}


def rect(page, x0, y0, x1, y1, fill, stroke=None, width=0.8, radius=False):
    shape = page.new_shape()
    r = fitz.Rect(x0, y0, x1, y1)
    if radius:
        shape.draw_rect(r)
    else:
        shape.draw_rect(r)
    shape.finish(color=stroke or fill, fill=fill, width=width)
    shape.commit()


def line(page, x0, y0, x1, y1, color=COLORS["line"], width=1):
    shape = page.new_shape()
    shape.draw_line(fitz.Point(x0, y0), fitz.Point(x1, y1))
    shape.finish(color=color, width=width)
    shape.commit()


def text(page, box, value, size=10, color=COLORS["text"], bold=False, align=0, leading=None):
    fontname = "hanbold" if bold else "han"
    page.insert_textbox(
        fitz.Rect(*box),
        value,
        fontsize=size,
        fontname=fontname,
        color=color,
        align=align,
        lineheight=leading or 1.18,
    )


def bullet_list(page, x, y, items, width=520, size=9.3, gap=20, color=COLORS["text"]):
    yy = y
    for item in items:
        rect(page, x, yy + 4, x + 5, yy + 9, COLORS["blue"])
        text(page, (x + 14, yy, x + width, yy + gap + 12), item, size=size, color=color)
        yy += gap
    return yy


def card(page, x, y, w, h, title, body, accent=COLORS["blue"], fill=COLORS["white"]):
    rect(page, x, y, x + w, y + h, fill, COLORS["line"], 0.8)
    rect(page, x, y, x + 7, y + h, accent)
    text(page, (x + 16, y + 12, x + w - 12, y + 34), title, size=11.5, bold=True, color=accent)
    text(page, (x + 16, y + 38, x + w - 12, y + h - 10), body, size=8.8, color=COLORS["text"])


def table(page, x, y, widths, row_h, headers, rows, title=None):
    if title:
        text(page, (x, y - 23, x + sum(widths), y - 4), title, size=11, bold=True, color=COLORS["text"])
    total_w = sum(widths)
    rect(page, x, y, x + total_w, y + row_h, COLORS["navy"])
    cx = x
    for i, h in enumerate(headers):
        text(page, (cx + 5, y + 7, cx + widths[i] - 5, y + row_h - 3), h, size=8.3, color=COLORS["white"], bold=True, align=1)
        line(page, cx, y, cx, y + row_h * (len(rows) + 1), COLORS["line"], 0.5)
        cx += widths[i]
    line(page, x + total_w, y, x + total_w, y + row_h * (len(rows) + 1), COLORS["line"], 0.5)
    yy = y + row_h
    for r_i, row in enumerate(rows):
        rect(page, x, yy, x + total_w, yy + row_h, COLORS["white"] if r_i % 2 == 0 else COLORS["gray"], COLORS["line"], 0.4)
        cx = x
        for c_i, cell in enumerate(row):
            text(page, (cx + 5, yy + 6, cx + widths[c_i] - 5, yy + row_h - 3), str(cell), size=7.8, color=COLORS["text"])
            cx += widths[c_i]
        yy += row_h
    return yy


def arrow(page, x0, y0, x1, y1, color=COLORS["blue"]):
    line(page, x0, y0, x1, y1, color, 1.5)
    shape = page.new_shape()
    shape.draw_polyline([fitz.Point(x1, y1), fitz.Point(x1 - 8, y1 - 5), fitz.Point(x1 - 8, y1 + 5), fitz.Point(x1, y1)])
    shape.finish(color=color, fill=color)
    shape.commit()


def flow_boxes(page, x, y, labels, box_w=96, box_h=54, gap=22):
    cx = x
    for i, (head, body) in enumerate(labels):
        rect(page, cx, y, cx + box_w, y + box_h, COLORS["sky"], COLORS["blue"], 0.8)
        text(page, (cx + 7, y + 8, cx + box_w - 7, y + 23), head, size=8.8, bold=True, color=COLORS["blue"], align=1)
        text(page, (cx + 7, y + 25, cx + box_w - 7, y + box_h - 4), body, size=7.2, color=COLORS["text"], align=1)
        if i < len(labels) - 1:
            arrow(page, cx + box_w + 3, y + box_h / 2, cx + box_w + gap - 3, y + box_h / 2)
        cx += box_w + gap


def image_box(page, x, y, w, h, image_name: str, caption: str):
    path = SCREENSHOT_DIR / image_name
    rect(page, x, y, x + w, y + h, COLORS["white"], COLORS["line"], 0.8)
    if path.exists():
        page.insert_image(fitz.Rect(x + 5, y + 5, x + w - 5, y + h - 25), filename=str(path), keep_proportion=True)
    else:
        text(page, (x + 12, y + 20, x + w - 12, y + h - 30), f"캡처 이미지 없음: {image_name}", size=9, color=COLORS["red"])
    text(page, (x + 8, y + h - 22, x + w - 8, y + h - 6), caption, size=7.3, color=COLORS["muted"], align=1)


def sidebar(page, section_no, title, keywords, note_title, note_lines):
    sx1 = SIDEBAR_X + SIDEBAR_W
    rect(page, SIDEBAR_X, 0, sx1, PAGE_H, COLORS["navy"])
    text(page, (SIDEBAR_X + 24, 34, sx1 - 22, 62), "PIPENET GUIDE", size=15, bold=True, color=COLORS["white"])
    text(page, (SIDEBAR_X + 24, 68, sx1 - 22, 100), f"{section_no:02d}", size=32, bold=True, color=(0.52, 0.68, 1.0))
    text(page, (SIDEBAR_X + 24, 116, sx1 - 22, 162), title, size=13, bold=True, color=COLORS["white"])
    line(page, SIDEBAR_X + 24, 178, sx1 - 24, 178, (0.25, 0.30, 0.40), 1)
    text(page, (SIDEBAR_X + 24, 200, sx1 - 24, 224), "KEY POINTS", size=10.5, bold=True, color=(0.65, 0.78, 1.0))
    y = 234
    for kw in keywords:
        rect(page, SIDEBAR_X + 24, y + 2, SIDEBAR_X + 29, y + 7, (0.65, 0.78, 1.0))
        text(page, (SIDEBAR_X + 38, y - 2, sx1 - 24, y + 24), kw, size=8.4, color=COLORS["white"])
        y += 24
    rect(page, SIDEBAR_X + 24, 420, sx1 - 24, 548, (0.12, 0.17, 0.27), (0.25, 0.30, 0.40), 0.8)
    text(page, (SIDEBAR_X + 38, 438, sx1 - 38, 460), note_title, size=10.5, bold=True, color=(0.65, 0.78, 1.0))
    text(page, (SIDEBAR_X + 38, 468, sx1 - 38, 536), "\n".join(note_lines), size=7.9, color=COLORS["white"])


def header(page, section_no, title, subtitle):
    text(page, (MARGIN, 28, CONTENT_R, 58), f"{section_no:02d}. {title}", size=20, bold=True, color=COLORS["text"])
    text(page, (MARGIN, 63, CONTENT_R, 88), subtitle, size=9.4, color=COLORS["muted"])
    line(page, MARGIN, 94, CONTENT_R, 94, COLORS["line"], 1)


def footer(page, page_no):
    text(page, (MARGIN, PAGE_H - 28, MARGIN + 180, PAGE_H - 12), "HBF&C R&D Center", size=7.5, color=COLORS["muted"])
    text(page, (CONTENT_R - 55, PAGE_H - 28, CONTENT_R, PAGE_H - 12), f"{page_no:02d}", size=8, color=COLORS["muted"], align=2)


def add_page(doc, no, title, subtitle, sidebar_keywords, sidebar_note_title, sidebar_note_lines):
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    page.insert_font(fontname="han", fontfile=FONT)
    page.insert_font(fontname="hanbold", fontfile=FONT_BOLD)
    rect(page, 0, 0, PAGE_W, PAGE_H, COLORS["white"])
    header(page, no, title, subtitle)
    sidebar(page, no, title, sidebar_keywords, sidebar_note_title, sidebar_note_lines)
    footer(page, no)
    return page


def build_pdf():
    doc = fitz.open()

    # Page 1
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    page.insert_font(fontname="han", fontfile=FONT)
    page.insert_font(fontname="hanbold", fontfile=FONT_BOLD)
    rect(page, 0, 0, PAGE_W, PAGE_H, COLORS["white"])
    rect(page, SIDEBAR_X, 0, SIDEBAR_X + SIDEBAR_W, PAGE_H, COLORS["navy"])
    text(page, (MARGIN, 42, CONTENT_R, 78), "PIPENET 수리계산 검증 프로그램", size=23, bold=True)
    text(page, (MARGIN, 86, CONTENT_R, 126), "검증 로직 · 결과 해석 · 설계 최적화 가이드", size=15, color=COLORS["blue"], bold=True)
    text(page, (MARGIN, 145, CONTENT_R, 218), "결과서 DOCX/PDF와 SDF 배관망을 기반으로 수리계산 결과의 법적 기준 충족 여부, Hazen-Williams 재계산, 배관 토폴로지 유속 판정, 공학/경제 최적화 후보를 확인하는 실무용 안내서입니다.", size=10.5, color=COLORS["text"])
    card(page, MARGIN, 255, 170, 105, "검증", "헤드, 배관 유속, C-Factor, 특수설비, HW 마찰손실 재계산", COLORS["blue"], COLORS["sky"])
    card(page, MARGIN + 190, 255, 170, 105, "해석", "결과 데이터 테이블, 상세 카드, 색상 범례, 판정 사유 확인", COLORS["green"], COLORS["green_bg"])
    card(page, MARGIN + 380, 255, 170, 105, "최적화", "공학적 안정성 후보와 시공사 경제성 후보를 분리 제시", COLORS["orange"], COLORS["orange_bg"])
    flow_boxes(page, MARGIN, 420, [("입력", "DOCX/PDF\nSDF"), ("파싱", "표/그래프\n구조화"), ("검증", "PASS/FAIL\nREVIEW"), ("최적화", "공학/경제\n후보"), ("출력", "테이블\n맵")], 92, 58, 16)
    sx1 = SIDEBAR_X + SIDEBAR_W
    text(page, (SIDEBAR_X + 28, 46, sx1 - 24, 80), "HBF&C", size=24, bold=True, color=COLORS["white"])
    text(page, (SIDEBAR_X + 28, 86, sx1 - 24, 128), "PIPENET\nHYDRAULIC\nGUIDE", size=22, bold=True, color=(0.65, 0.78, 1.0))
    rect(page, SIDEBAR_X + 28, 350, sx1 - 28, 505, (0.12, 0.17, 0.27), (0.25, 0.30, 0.40), 0.8)
    text(page, (SIDEBAR_X + 42, 370, sx1 - 42, 395), "DOCUMENT SCOPE", size=10.5, bold=True, color=(0.65, 0.78, 1.0))
    text(page, (SIDEBAR_X + 42, 404, sx1 - 42, 492), "수리계산 결과 검증 기준\n데이터 테이블 해석\nSDF 배관망 시각화\n설계 최적화 리포트\nCAD 대조 모듈 개요", size=9, color=COLORS["white"])
    footer(page, 1)

    # Page 2
    page = add_page(doc, 2, "입력 데이터와 사용 흐름", "결과서와 SDF 파일이 서버 내부에서 어떻게 구조화되는지 설명합니다.", ["결과서 DOCX/PDF", "SDF 배관망", "선택 입력: CAD, project_meta, design_policy", "결과 테이블과 맵 출력"], "DATA FLOW", ["입력값은 원자료를 보존한 상태로", "검증용 구조 데이터로 변환됩니다.", "판정은 서버 로직 기준으로 수행됩니다."])
    table(page, MARGIN, 130, [92, 165, 255], 40, ["입력", "주요 추출 항목", "사용 목적"], [
        ["결과서", "PIPE CONFIGURATION, FLOW IN PIPES, NOZZLE", "압력, 유량, 유속, 마찰손실, 구경, C-Factor 검증"],
        ["SDF", "Node, Pipe, Nozzle, Waypoint", "배관망 토폴로지, 하류 헤드 수, 네트워크 맵 구성"],
        ["CAD/DXF", "선분, 레이어, 객체 타입", "도면-아이소매트릭 대조 모듈의 선택 입력"],
        ["정책 파일", "project_meta, design_policy", "세대 내부/유입 배관, 헤드 수별 관경 정책 보완"],
    ], "입력 데이터 구조")
    flow_boxes(page, MARGIN, 345, [("업로드", "결과서\nSDF"), ("파서", "표 추출\n노드 추출"), ("룰 엔진", "법적 기준\n공식 검산"), ("UI", "테이블\n맵"), ("보고서", "PDF/Excel\n리포트")], 92, 60, 16)
    bullet_list(page, MARGIN, 440, ["결과서가 없어도 SDF만으로 모든 검증을 수행할 수는 없습니다.", "SDF는 토폴로지와 시각화에 강하고, 결과서는 계산 결과 검증에 필수입니다.", "CAD는 도면 대조용 선택 입력이며, 없으면 REVIEW로 처리합니다."], 530, 9.3, 25)

    # Page 3
    page = add_page(doc, 3, "공통 검증 기준", "수리계산 결과가 기본 기준을 만족하는지 먼저 확인합니다.", ["Hazen-Williams 선언", "마찰손실 재계산", "허용오차 비교", "PASS/FAIL 메시지"], "COMMON RULE", ["공통 검증은 최적화보다 우선합니다.", "공식 재현 실패 시 결과서 신뢰성 확인이 필요합니다."])
    card(page, MARGIN, 130, 250, 105, "COMMON.HW_DECLARED", "DESIGN INFORMATION에 'Using the Hazen-Williams Equation' 선언이 있는지 확인합니다. 문구가 있으면 PASS, 없으면 FAIL입니다.", COLORS["blue"], COLORS["sky"])
    card(page, MARGIN + 270, 130, 250, 105, "COMMON.HW_RECALC", "Pipe별 total_length와 Act. Bore, Flowrate, C-Factor로 Frict. Loss를 재계산해 결과서 값과 비교합니다.", COLORS["green"], COLORS["green_bg"])
    table(page, MARGIN, 285, [160, 185, 160], 38, ["항목", "수식/기준", "해석"], [
        ["총 등가길이", "length + fitting_eq + special_eq", "배관 본길이와 부속/특수설비 등가길이 합산"],
        ["HW 손실", "6.174e4·Q^1.85·L/(C^1.85·D^4.87)/0.1", "kg/cm² 기준 결과서 Frict. Loss와 비교"],
        ["허용오차", "max(0.005, reported × 0.005)", "절대오차 0.005 또는 상대 0.5% 중 큰 값"],
    ], "Hazen-Williams 재계산 기준")

    # Page 4
    page = add_page(doc, 4, "법적 기준 충족 검증", "헤드와 배관 유속의 핵심 판정 기준을 정리합니다.", ["헤드 압력/유량", "Topology 기반 유속", "가지배관 6 m/s", "그 밖의 배관 10 m/s"], "LEGAL CHECK", ["구경만으로 가지배관을 판단하지 않습니다.", "하류 헤드 수와 교차분기를 함께 봅니다."])
    card(page, MARGIN, 128, 250, 120, "헤드 검증", "표준헤드는 최소 방수압 0.1 MPa 이상, 요구 유량 이상이어야 합니다. 결과서 노즐 표에서 압력과 유량을 추출해 판정합니다.", COLORS["blue"], COLORS["sky"])
    card(page, MARGIN + 270, 128, 250, 120, "배관 유속 검증", "PIPE CONFIGURATION과 NOZZLE CONFIGURATION으로 그래프를 만들고, 하류 교차분기 여부로 branch/other를 판정합니다.", COLORS["orange"], COLORS["orange_bg"])
    flow_boxes(page, MARGIN, 300, [("Graph", "Pipe\ninput/output"), ("Nozzle", "하류 헤드\n계산"), ("Split", "교차분기\n탐지"), ("Role", "branch\nother"), ("Limit", "6 또는\n10 m/s")], 92, 60, 16)
    table(page, MARGIN, 405, [115, 155, 245], 34, ["분류", "기준", "설명"], [
        ["branch", "6.0 m/s 이하", "50A 이하이면서 하류 subtree에 multi-head cross split이 없는 배관"],
        ["other", "10.0 m/s 이하", "50A 초과 또는 하류 subtree에 교차분기가 존재하는 배관"],
        ["review", "판정 보류", "그래프 누락, 사이클, 노드 참조 오류 등 topology ambiguity 발생"],
    ], "Topology 기반 유속 기준")

    # Page 5
    page = add_page(doc, 5, "배관 항목 검증", "배관 재질, C-Factor, 고압 구간, 정책성 기준을 구조화합니다.", ["PIPE.001~PIPE.006", "KSD3562 고압 구간", "C-Factor", "REVIEW 처리"], "PIPE RULES", ["CAD나 메타데이터가 없으면 억지 FAIL이 아니라 REVIEW입니다.", "결과서는 계산 결과의 권위 출력입니다."])
    table(page, MARGIN, 125, [76, 170, 250], 35, ["Rule", "검증 항목", "판정 방식"], [
        ["PIPE.001", "도면과 Schematic 일치", "CAD 입력이 없으면 REVIEW, 있으면 mismatch 목록 기반 판정"],
        ["PIPE.002", "1.2 MPa 이상 KSD3562", "max(inlet,outlet) ≥ 12 kg/cm² 구간은 KSD3562 요구"],
        ["PIPE.003", "재질별 C-Factor", "KSD 계열 C=120, CPVC 계열 C=150"],
        ["PIPE.004", "단위세대 내부 CPVC", "project_meta 또는 CAD zone 입력 필요"],
        ["PIPE.005", "단위세대 유입 65A 이하", "unit_inlet_pipe_labels 필요"],
        ["PIPE.006", "헤드 수별 관경 정책", "design_policy.json 정책표 기반 판정"],
    ], "4. 배관 대제목 규칙")
    card(page, MARGIN, 405, 250, 85, "PASS 예시", "고압 구간이 없고 C-Factor가 재질 기준과 일치하면 배관 기본 검증은 PASS입니다.", COLORS["green"], COLORS["green_bg"])
    card(page, MARGIN + 270, 405, 250, 85, "REVIEW 예시", "CAD, project_meta, design_policy가 없으면 도면 일치·세대 구분·정책성 관경은 REVIEW입니다.", COLORS["orange"], COLORS["orange_bg"])

    # Page 6
    page = add_page(doc, 6, "결과 데이터 테이블 해석", "색상과 상세 카드가 어떤 의미인지 설명합니다.", ["색상 범례", "클릭 사유 카드", "HW 검산 상세", "Topology 유속 열"], "RESULT TABLE", ["필터링된 행을 클릭하면", "수식·기준·실제값·결론이 카드로 표시됩니다."])
    table(page, MARGIN, 125, [85, 120, 305], 35, ["색상", "의미", "대표 조건"], [
        ["빨강", "기준 위반", "유속 초과, HW 재계산 실패, 법적 기준 미달"],
        ["파랑", "공학 최적화 후보", "마찰손실, 변화율, 유속여유, 피팅집중, 압력여유 관련 후보"],
        ["초록", "경제성 검토 후보", "저유속 과설계, 압력여유 과다, 대구경 밸브, CPVC 대구경"],
        ["흰색", "일반/REVIEW", "REVIEW는 별도 정보 버튼에서 규칙 상태 확인"],
    ], "색상 의미")
    table(page, MARGIN, 310, [142, 180, 185], 34, ["열 그룹", "주요 필드", "클릭 시 표시"], [
        ["기본 배관 정보", "Pipe, 구경, 길이, 유량, 유속", "원자료 출처와 단위"],
        ["HW 검산 상세", "총등가길이, 재계산 손실, 오차", "공식과 허용오차"],
        ["Topology 유속", "배관 역할, 하류 헤드 수, 교차분기", "branch/other 판정 이유"],
        ["공학/경제 후보", "후보 사유, 추천 조치", "원인과 개선 방향"],
    ], "테이블 구성")

    # Page 7
    page = add_page(doc, 7, "설계 최적화 가이드", "공학적 안정성과 경제성 절감을 분리해서 제시합니다.", ["공학 최적화", "경제성 최적화", "충돌 구간", "절충 설계"], "OPTIMIZATION", ["최적화는 기준 PASS 이후 단계입니다.", "하나의 정답이 아니라 후보와 근거를 제공합니다."])
    card(page, MARGIN, 125, 250, 118, "공학적 마찰손실 최적화", "안정성, 압력 여유, 손실 분산, 유속 안정성을 우선합니다. m당 손실, 변화율, 유속 여유, 피팅 집중, 압력 여유를 종합합니다.", COLORS["blue"], COLORS["sky"])
    card(page, MARGIN + 270, 125, 250, 118, "시공사 경제성 확보", "법적 기준을 유지하면서 과설계를 줄입니다. 저유속, 압력 여유 과다, 대구경 밸브, CPVC 대구경, 과대 관경을 찾습니다.", COLORS["green"], COLORS["green_bg"])
    table(page, MARGIN, 295, [132, 190, 190], 38, ["관점", "대표 지표", "권장 조치"], [
        ["공학", "m당 손실, 변화율, 유속여유 부족", "구경 상향, 피팅 축소, 경로 단순화"],
        ["경제", "저유속, 압력여유 과다, 대구경 비용", "관경 축소 시뮬레이션, 밸브 구경 최적화"],
        ["절충", "공학 후보와 경제 후보가 충돌", "위험구간은 공학 우선, 여유 구간만 경제안 검토"],
    ], "공학/경제 관점 비교")

    # Page 8
    page = add_page(doc, 8, "SDF 아이소매트릭 배관망", "검증 결과를 네트워크 맵과 연결해 해석합니다.", ["줌/드래그", "적합/부적합 표시", "Spike Map", "Economy Map"], "NETWORK VIEW", ["배관망은 판정 결과를 공간적으로 이해하기 위한 보조 화면입니다.", "SDF가 없으면 맵 기능은 제한됩니다."])
    flow_boxes(page, MARGIN, 128, [("SDF", "Node\nPipe"), ("Parser", "좌표\n연결"), ("Graph", "노드\n엣지"), ("Layer", "적합\n부적합"), ("Map", "줌\n드래그")], 92, 60, 16)
    table(page, MARGIN, 245, [155, 160, 190], 38, ["맵", "표시 대상", "인터랙션"], [
        ["Friction Loss Spike Map", "m당 마찰손실 > 1.0 kg/cm²/m", "빨간 배관 드래그 시 Pipe 번호, 손실, 변화율 표시"],
        ["Economy Optimization Candidate Map", "저유속/압력여유/대구경 후보", "줌/드래그로 후보 위치 확인"],
        ["검증 결과 레이어", "선택한 PASS/FAIL 분류", "해당 노드와 엣지만 색상 표시"],
    ], "네트워크 시각화 기능")
    card(page, MARGIN, 420, 520, 70, "해석 주의", "네트워크 맵은 결과를 이해하기 위한 시각화 도구입니다. 최종 설계 변경은 결과서 재계산, 유속 기준, 말단압력, 헤드 유량 기준을 다시 확인해야 합니다.", COLORS["orange"], COLORS["orange_bg"])

    # Page 9
    page = add_page(doc, 9, "CAD-아이소매트릭 대조 모듈", "DXF 도면과 SDF 배관망을 비교하기 위한 보조 모듈입니다.", ["DXF 레이어 필터", "선분 그래프화", "영역 선택", "구성요소 대조"], "CAD COMPARE", ["현재 대조 모듈은 보조 검토 도구입니다.", "정확한 Pipe 단위 대조에는 도면 정제와 좌표 보정이 필요합니다."])
    table(page, MARGIN, 125, [135, 180, 190], 36, ["단계", "처리 내용", "목적"], [
        ["1. DXF 정제", "배관 후보 레이어 분리, 불필요 객체 삭제", "건축 도면의 복잡도를 낮춤"],
        ["2. 선분 추출", "LINE, LWPOLYLINE, ARC endpoint 추출", "도면 선분을 네트워크 후보로 변환"],
        ["3. 점 병합", "endpoint tolerance로 인접점 병합", "노드, 말단, 분기점 생성"],
        ["4. 그래프 생성", "노드-엣지 연결 구조화", "SDF 그래프와 비교 가능한 형태 구성"],
        ["5. 대조 분석", "헤드 수, 밸브, 부속류, 길이, 연결성 비교", "명시적 mismatch 목록 생성"],
    ], "DXF 기반 대조 흐름")
    flow_boxes(page, MARGIN, 390, [("DXF", "레이어\n정제"), ("Segment", "선분\n추출"), ("Node", "점 병합\n분기점"), ("Graph", "연결망\n생성"), ("Compare", "SDF와\n대조")], 92, 60, 16)

    # Page 10
    page = add_page(doc, 10, "사용 절차와 FAQ", "실무자가 검증 결과를 해석할 때 확인할 항목입니다.", ["업로드 순서", "REVIEW 처리", "다운로드", "주의사항"], "FAQ", ["검증 결과는 설계자 판단을 대체하지 않습니다.", "수정 후에는 반드시 재검증해야 합니다."])
    table(page, MARGIN, 125, [72, 190, 245], 36, ["순서", "사용자 작업", "확인 사항"], [
        ["1", "결과서 DOCX/PDF 업로드", "결과서 표 파싱 가능 여부 확인"],
        ["2", "SDF 파일 선택 업로드", "토폴로지, 배관망 맵, 하류 헤드 수 확인"],
        ["3", "검증 실행", "PASS/FAIL/WARNING/REVIEW 메시지 확인"],
        ["4", "결과 데이터 테이블 클릭", "수식, 기준값, 실제값, 판정 사유 확인"],
        ["5", "공학/경제 최적화 검토", "후보 조치 후 재계산 필요"],
        ["6", "Excel/PDF 결과 저장", "검토 기록과 설계 수정 근거로 활용"],
    ], "기본 사용 절차")
    table(page, MARGIN, 380, [160, 345], 34, ["질문", "답변"], [
        ["REVIEW는 오류인가?", "아닙니다. CAD, project_meta, design_policy 등 추가 입력이 없어 판정 보류된 상태입니다."],
        ["경제성 후보는 바로 축소해도 되나?", "아닙니다. 축소 후 유속, 말단압력, 헤드 유량, HW 검산을 다시 통과해야 합니다."],
        ["Spike Map의 빨간색 의미는?", "m당 마찰손실 > 1.0 kg/cm²/m인 급증 전용 조건입니다."],
    ], "자주 묻는 질문")

    # Page 11
    page = add_page(doc, 11, "서버 웹 인터페이스 캡처", "실제/재현 서버 화면을 삽입하여 사용자가 기능 위치를 빠르게 이해하도록 합니다.", ["업로드 화면", "검증 결과", "결과 테이블", "통계 화면"], "SCREENSHOT", ["표 설명만으로 부족한 부분은", "서버 화면 캡처를 함께 보면서", "기능 위치를 확인합니다."])
    image_box(page, MARGIN, 120, 245, 172, "server_actual_main.png", "파일 업로드 및 검증 실행 첫 화면")
    image_box(page, MARGIN + 265, 120, 245, 172, "server_validation.png", "검증 결과와 SDF 배관망 화면")
    image_box(page, MARGIN, 330, 245, 172, "server_table.png", "결과 데이터 테이블과 색상 범례")
    image_box(page, MARGIN + 265, 330, 245, 172, "server_stats.png", "검진 통계와 그래프 화면")

    # Page 12
    page = add_page(doc, 12, "설계 최적화 화면 예시", "공학 최적화와 경제성 최적화 후보가 화면에서 어떻게 분리되는지 보여줍니다.", ["Friction Loss Spike Map", "Economy Candidate Map", "공학 후보", "경제성 후보"], "OPTIMIZATION UI", ["공학 맵은 손실 급증 구간 중심,", "경제성 맵은 비용 절감 후보 중심으로", "서로 다른 목적을 갖습니다."])
    image_box(page, MARGIN, 125, 510, 315, "server_optimization.png", "공학적 마찰손실 최적화 조치와 시공사 경제성 확보 방안 화면")
    card(page, MARGIN, 470, 245, 62, "공학 맵 해석", "빨간 구간은 m당 마찰손실 급증 전용 조건을 만족한 배관입니다.", COLORS["blue"], COLORS["sky"])
    card(page, MARGIN + 265, 470, 245, 62, "경제 맵 해석", "초록 후보는 기준 만족 범위 내에서 관경/밸브/재질 최적화를 검토할 구간입니다.", COLORS["green"], COLORS["green_bg"])

    doc.save(OUT_ASCII, garbage=4, deflate=True)
    doc.close()
    shutil.copy2(OUT_ASCII, OUT_KR)


if __name__ == "__main__":
    build_pdf()
    doc = fitz.open(OUT_ASCII)
    print("pages", doc.page_count)
    print("text markers", all(s in "\n".join(p.get_text() for p in doc) for s in ["PIPENET", "Hazen-Williams", "설계 최적화", "CAD-아이소매트릭"]))
    print(OUT_ASCII, OUT_ASCII.exists(), OUT_ASCII.stat().st_size)
    print(OUT_KR, OUT_KR.exists(), OUT_KR.stat().st_size)
