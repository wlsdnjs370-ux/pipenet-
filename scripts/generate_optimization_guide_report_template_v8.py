from __future__ import annotations

import os
import shutil
import subprocess
import time
from pathlib import Path

import win32clipboard
import win32con
import win32com.client.dynamic


ROOT = Path(__file__).resolve().parents[1]
DESKTOP = Path(os.environ.get("USERPROFILE", r"C:\Users\admin")) / "Desktop"
TEMPLATE_ASCII = ROOT / "templates_hwp" / "template_optimization_guide_from_user_v8.hwp"
OUT_ASCII = DESKTOP / "PIPENET_2_optimization_guide_on_template_v8.hwp"
OUT_KR = DESKTOP / "PIPENET 수리계산 검증 프로그램 중 2. 설계 최적화 가이드 로직_원본양식위_완성본_v8.hwp"


FONT_HWP = "함초롬바탕"
FONT_LATIN = "HCR Batang"


def set_clipboard_text(text: str) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def line(char: str = "─", n: int = 74) -> str:
    return char * n


def table(title: str, headers: list[str], rows: list[list[str]], widths: list[int]) -> str:
    def fmt_row(vals: list[str]) -> str:
        parts = []
        for val, width in zip(vals, widths):
            s = str(val).replace("\n", " ")
            if len(s) > width:
                s = s[: width - 1] + "…"
            parts.append(s.ljust(width))
        return "│ " + " │ ".join(parts) + " │"

    total = sum(widths) + len(widths) * 3 + 1
    out = ["", f"[표] {title}", "┌" + "─" * (total - 2) + "┐", fmt_row(headers), "├" + "─" * (total - 2) + "┤"]
    for row in rows:
        out.append(fmt_row(row))
    out.append("└" + "─" * (total - 2) + "┘")
    return "\r\n".join(out) + "\r\n"


def flow(title: str, steps: list[tuple[str, str]]) -> str:
    out = ["", f"[삽도] {title}", "┌" + "─" * 72 + "┐"]
    for idx, (head, body) in enumerate(steps, 1):
        out.append(f"│ {idx}. {head}".ljust(74) + "│")
        for b in body.split("\n"):
            out.append(f"│    - {b}".ljust(74) + "│")
        if idx != len(steps):
            out.append("│" + " " * 34 + "↓" + " " * 37 + "│")
    out.append("└" + "─" * 72 + "┘")
    return "\r\n".join(out) + "\r\n"


def build_report() -> str:
    parts: list[str] = []
    parts.append(
        "\r\n\r\n"
        + line("=")
        + "\r\n"
        + "2. 설계 최적화 가이드 로직 보고서\r\n"
        + line("=")
        + "\r\n"
        + "본 문서는 PIPENET 수리계산 검증 프로그램의 2. 설계 최적화 가이드가 어떤 입력값을 읽고, 어떤 수식과 판정 기준을 거쳐 공학적 최적화 후보와 시공사 경제성 후보를 제시하는지 정리한 문서이다.\r\n"
        + "설계 최적화는 법적 기준을 무시하고 관경을 줄이는 작업이 아니다. 공학안과 경제안 모두 필수 검증 기준을 통과한 설계를 전제로 하며, 이후 안정성 또는 비용 절감 관점의 개선 후보를 분리해 제시한다.\r\n"
    )
    parts.append(
        table(
            "목차",
            ["No", "항목명"],
            [
                ["1", "문서 목적 및 기본 전제"],
                ["2", "입력 데이터와 파서 구조"],
                ["3", "공통 계산 지표와 수식"],
                ["4", "공학적 마찰손실 최적화 로직"],
                ["5", "Friction Loss Spike Map"],
                ["6", "시공사 경제성 확보 로직"],
                ["7", "Economy Optimization Candidate Map"],
                ["8", "결과 데이터 테이블 표시 구조"],
                ["9", "공학-경제 충돌 및 절충 판단"],
                ["10", "관경 축소 시뮬레이션 흐름"],
                ["11", "한계 및 추가 입력"],
            ],
            [6, 54],
        )
    )
    parts.append(
        flow(
            "설계 최적화 가이드 전체 데이터 흐름",
            [
                ("입력", "결과서 DOCX/PDF, SDF 배관망, 선택 입력(project_meta, design_policy)"),
                ("파싱", "배관, 헤드, 압력, 유속, 마찰손실, C-Factor, 토폴로지 구조화"),
                ("필수 기준 PASS 확인", "법적/기술 기준을 먼저 검증하고, 미달 시 최적화보다 수정이 우선"),
                ("공학/경제 후보 분리", "안정성 개선 후보와 비용 절감 후보를 별도 로직으로 산출"),
                ("출력", "결과 데이터 테이블, 공학/경제 맵, 상세 카드, 최종 리포트 표시"),
            ],
        )
    )
    parts.append(
        "\r\n1. 문서 목적 및 기본 전제\r\n"
        + line()
        + "\r\n"
        + "공학안과 경제안은 모두 필수 기준 PASS가 전제이다. 기준 미달 배관은 최적화 후보가 아니라 수정 대상이다.\r\n"
    )
    parts.append(
        table(
            "필수 PASS 전제 기준",
            ["No", "전제 기준", "판정 의미"],
            [
                ["1", "헤드 최소 방수압 0.1 MPa 이상", "말단 방수압 기준 만족 후 최적화 가능"],
                ["2", "헤드 최소 방수량 기준 만족", "노즐별 요구 유량 이상 확보"],
                ["3", "가지배관 유속 6 m/s 이하", "Topology 기준 branch 배관에 적용"],
                ["4", "그 밖의 배관 유속 10 m/s 이하", "교차배관/주배관/50A 초과 배관에 적용"],
                ["5", "재질별 C-Factor 기준 만족", "KSD 계열 C=120, CPVC 계열 C=150"],
                ["6", "특수설비/밸브 등가길이 검증", "FX, AV/PV, PRV 등 입력값 검토"],
                ["7", "Hazen-Williams 재계산 검증", "결과서 Frict. Loss를 수식으로 재현"],
            ],
            [4, 30, 32],
        )
    )
    parts.append(
        "\r\n2. 입력 데이터와 파서 구조\r\n"
        + line()
        + "\r\n"
    )
    parts.append(
        table(
            "입력 데이터와 사용 목적",
            ["입력", "주요 추출 필드", "역할"],
            [
                ["결과서 DOCX/PDF", "PIPE CONFIGURATION, FLOW IN PIPES", "압력, 유량, 유속, 손실, 구경, C값의 권위 출력"],
                ["SDF 파일", "Node, Pipe input/output, Nozzle input", "연결 관계, 하류 헤드 수, 네트워크 맵 표시"],
                ["project_meta.json", "unit_internal, unit_inlet", "공간 의미 정보 보완"],
                ["design_policy.json", "head_count_min_nominal", "프로젝트별 관경 정책 적용"],
            ],
            [18, 28, 28],
        )
    )
    parts.append(
        "\r\n3. 공통 계산 지표와 수식\r\n"
        + line()
        + "\r\n"
    )
    parts.append(
        table(
            "공통 계산 지표와 수식",
            ["지표", "수식", "해석"],
            [
                ["m당 마찰손실", "friction_loss / length", "길이 영향을 제거한 손실 집중도"],
                ["마찰손실 변화율", "(현재 m당 손실 - 직전 m당 손실) / 직전 m당 손실", "직전 배관 대비 급증 여부"],
                ["유속 여유율", "1 - actual_velocity / velocity_limit", "유속 기준까지 남은 여유"],
                ["압력 여유", "actual_pressure - required_pressure", "압력 안정성 또는 축소 여지"],
                ["등가길이 비율", "(fitting_eq + special_eq) / total_length", "피팅/특수설비 집중도"],
                ["HW 재계산", "6.174e4*Q^1.85*L/(C^1.85*D^4.87)/0.1", "결과서 손실 재현 검산"],
            ],
            [16, 38, 26],
        )
    )
    parts.append(
        "\r\n4. 공학적 마찰손실 최적화 로직\r\n"
        + line()
        + "\r\n"
        + "공학적 최적화는 안정성, 압력 여유, 마찰손실 분산, 유속 안정성을 우선한다. 비용 절감보다 특정 배관에 손실이 집중되는 현상과 말단 압력 불안정을 줄이는 것이 목적이다.\r\n"
    )
    parts.append(
        flow(
            "공학적 마찰손실 최적화 판정 흐름",
            [
                ("지표 계산", "m당 손실, 변화율, 유속 여유율, 압력 여유, 등가길이 비율 산정"),
                ("후보 세분화", "마찰손실, 변화율, 유속여유, 피팅집중, 압력여유 사유로 분리"),
                ("Spike Map 표시", "m당 마찰손실 > 1.0 kg/cm²/m 조건만 빨간색 표시"),
                ("원인 분류", "긴 배관인지, 짧지만 손실이 과도한지, 구경/유속/피팅 때문인지 구분"),
                ("조치 제안", "구경 상향, 피팅 축소, 배관 경로 단순화, 특수설비 위치 조정"),
            ],
        )
    )
    parts.append(
        table(
            "공학적 마찰손실 최적화 후보 기준",
            ["후보 사유", "판정 조건", "권장 조치"],
            [
                ["마찰손실 절대값 과다", "m당 마찰손실이 공학 기준 초과", "관경 상향, 피팅 축소, 경로 단순화"],
                ["마찰손실 변화율 급증", "직전 배관 대비 변화율이 큼", "이전/현재 구간의 구경, 피팅, 특수설비 비교"],
                ["유속 여유 부족", "유속 여유율이 낮아 기준에 근접", "관경 상향 또는 분기/경로 재구성"],
                ["피팅/특수설비 집중", "등가길이 비율이 과다", "엘보/Tee 축소, 특수설비 위치 조정"],
                ["압력 여유 부족", "말단 압력이 요구 기준에 근접", "손실 저감 또는 관경 상향"],
            ],
            [24, 34, 34],
        )
    )
    parts.append(
        "\r\n5. Friction Loss Spike Map\r\n"
        + line()
        + "\r\n"
    )
    parts.append(
        table(
            "Friction Loss Spike Map 표시 기준",
            ["구분", "적용 기준", "인터랙션"],
            [
                ["표시 목적", "마찰손실 급증 전용 조건만 네트워크에 표시", "빨간 배관 드래그/호버 시 Pipe 번호와 손실 정보 표시"],
                ["전용 조건", "m당 마찰손실 > 1.0 kg/cm²/m", "결과 데이터 테이블의 전체 공학 후보와 구분"],
                ["해석", "짧은 배관인데 손실이 크면 피팅/특수설비 집중 가능성", "원인과 조치 방향 동시 표시"],
            ],
            [16, 40, 40],
        )
    )
    parts.append(
        "\r\n6. 시공사 경제성 확보 로직\r\n"
        + line()
        + "\r\n"
        + "경제성 최적화는 가능한 싸게 만드는 것이 아니라, 법적/기술 기준을 유지하는 범위에서 과설계를 줄이는 방향이다. 관경 축소 후에도 유속, 말단압력, 헤드 유량, HW 검산 조건을 재확인해야 한다.\r\n"
    )
    parts.append(
        flow(
            "시공사 경제성 확보 판정 흐름",
            [
                ("후보 탐색", "저유속, 압력여유 과다, 대구경 밸브, CPVC 대구경, 과대 관경 탐색"),
                ("축소 가정", "관경 한 단계 축소를 가정하고 재계산 준비"),
                ("HW 재계산", "축소 후 마찰손실, 유속, 말단압력, 헤드 유량 검토"),
                ("통과 조건", "가지 ≤ 6 m/s, 그 밖 ≤ 10 m/s, 방수압 ≥ 0.1 MPa"),
                ("경제 제안", "관경 축소, 밸브 100A 이하 조정, CPVC 대구경 회피 제안"),
            ],
        )
    )
    parts.append(
        table(
            "시공사 경제성 확보 후보 기준",
            ["경제성 후보", "판정 조건", "권장 조치"],
            [
                ["저유속 과설계", "velocity < 2.0 m/s and nominal_bore > 25A", "한 단계 관경 축소 시뮬레이션"],
                ["압력 여유 과다", "말단압력이 최소 기준보다 과도하게 높음", "0.11~0.12 MPa 수준 목표 검토"],
                ["대구경 밸브 비용", "valve_connected_bore > 100A", "100A 이하 조정 가능성 검토"],
                ["CPVC 대구경", "CPVC nominal_bore > 65A", "50A/65A 복수 라인 분산 검토"],
                ["하류 헤드 수 대비 과대 관경", "downstream_nozzle_count 대비 nominal_bore가 큼", "40A→32A, 32A→25A 등 검토"],
            ],
            [24, 42, 34],
        )
    )
    parts.append(
        "\r\n7. Economy Optimization Candidate Map\r\n"
        + line()
        + "\r\n"
    )
    parts.append(
        table(
            "Economy Optimization Candidate Map 표시 기준",
            ["구분", "표시 대상", "주의사항"],
            [
                ["목적", "경제성 후보 배관을 네트워크상에서 확인", "빨간 Spike Map과 목적이 다르므로 혼동 금지"],
                ["대상", "저유속, 압력여유 과다, 대구경 밸브, CPVC 대구경", "최종 축소 여부는 재계산 통과 후 결정"],
                ["조작", "Friction Loss Spike Map과 동일하게 줌/드래그 가능", "시각화는 후보 탐색 도구이며 자동 설계 변경은 아님"],
            ],
            [16, 40, 40],
        )
    )
    parts.append(
        "\r\n8. 결과 데이터 테이블 표시 구조\r\n"
        + line()
        + "\r\n"
    )
    parts.append(
        table(
            "결과 데이터 테이블 열 그룹",
            ["열 그룹", "주요 필드", "클릭 시 설명"],
            [
                ["기본 배관 정보", "Pipe, 구경, 실내경, 길이, 유량, 유속", "원자료 출처와 단위 설명"],
                ["HW 검산 상세", "C-Factor, 총등가길이, 결과서 손실, 재계산 손실, 오차", "수식, 허용오차, PASS/FAIL 근거"],
                ["Topology 유속", "배관 역할, 하류 헤드 수, 하류 교차분기, 유속 기준", "branch/other 판정 이유"],
                ["공학 후보", "마찰손실, 변화율, 유속여유, 피팅집중, 압력여유", "각 사유별 계산값과 조치"],
                ["경제 후보", "저유속, 압력여유 과다, 대구경 밸브, CPVC 대구경", "축소 시뮬레이션 필요 조건"],
            ],
            [20, 48, 34],
        )
    )
    parts.append(
        "\r\n9. 공학-경제 충돌 및 절충 판단\r\n"
        + line()
        + "\r\n"
    )
    parts.append(
        table(
            "공학-경제 충돌 및 절충 판단",
            ["충돌 유형", "공학 해석", "경제 해석", "절충안"],
            [
                ["손실 큰 배관", "구경 상향 또는 피팅 축소 필요", "구경 상향은 비용 증가", "관경 변경 전 피팅/경로부터 점검"],
                ["저유속 대구경", "압력 안정성에는 유리", "과설계 가능성", "축소 시뮬레이션 후 기준 재확인"],
                ["CPVC 대구경", "C=150으로 손실 저감 가능", "대구경 CPVC 비용 상승", "복수 소구경 라인 또는 재질 대안 비교"],
                ["대구경 밸브", "손실과 유지관리 측면에서 안정", "밸브 단가 상승", "100A 이하 조정 가능성 검토"],
            ],
            [18, 30, 28, 34],
        )
    )
    parts.append(
        flow(
            "공학 관점과 경제 관점의 충돌/절충 구조",
            [
                ("공학 관점", "안정성, 압력 여유, 손실 균일성 증가. 대체로 비용 증가 가능"),
                ("경제 관점", "기준 만족 범위 내 관경, 밸브, 재질, 시공비 축소. 여유 감소 가능"),
                ("충돌 구간", "공학은 상향을 요구하지만 경제는 축소를 요구하는 구간"),
                ("절충안", "관경 변경 전 피팅/경로 점검, 축소 시뮬레이션 후 기준 재확인"),
            ],
        )
    )
    parts.append(
        "\r\n10. 관경 축소 시뮬레이션 흐름\r\n"
        + line()
        + "\r\n"
    )
    parts.append(
        table(
            "관경 축소 시뮬레이션 절차",
            ["단계", "처리 내용", "통과 조건"],
            [
                ["1", "현재 설계 압력 여유와 유속 여유 확인", "필수 기준 PASS 상태"],
                ["2", "관경 한 단계 축소 가정", "125A→100A, 100A→80A, 80A→65A 등"],
                ["3", "Hazen-Williams로 마찰손실 재계산", "총등가길이와 C-Factor 반영"],
                ["4", "유속 기준 재검토", "가지 ≤ 6 m/s, 그 밖 ≤ 10 m/s"],
                ["5", "말단압력 및 헤드 유량 재검토", "방수압 ≥ 0.1 MPa, 유량 기준 만족"],
                ["6", "통과 시 경제성 후보로 제시", "자동 확정이 아니라 설계자 검토 대상으로 출력"],
            ],
            [8, 46, 38],
        )
    )
    parts.append(
        "\r\n11. 한계 및 추가 입력\r\n"
        + line()
        + "\r\n"
    )
    parts.append(
        table(
            "한계 및 추가 입력",
            ["항목", "현재 한계", "보완 입력"],
            [
                ["세대 내부 CPVC", "결과서만으로 공간 구분 불가", "project_meta.json 또는 CAD zone 분류"],
                ["세대 유입 65A", "어떤 배관이 세대 유입인지 결과서만으로 확정 불가", "unit_inlet_pipe_labels"],
                ["헤드 수별 관경 정책", "법정 기준이 아니라 회사/프로젝트 정책", "design_policy.json"],
                ["도면-스케매틱 일치", "CAD 레이어/축척/좌표계 정리가 필요", "정제 DXF, transform metadata"],
                ["경제성 금액 산출", "현재는 후보 로직 중심", "자재 단가표, 밸브 단가표, 시공 단가표"],
            ],
            [22, 42, 34],
        )
    )
    return "".join(parts)


def set_font(hwp) -> None:
    try:
        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
        cs = hwp.HParameterSet.HCharShape
        cs.FaceNameHangul = FONT_HWP
        cs.FaceNameLatin = FONT_LATIN
        cs.FaceNameHanja = FONT_HWP
        cs.FaceNameJapanese = FONT_HWP
        cs.FaceNameOther = FONT_HWP
        cs.Height = 1000
        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
    except Exception:
        pass


def paste_unicode(hwp, text: str) -> None:
    set_clipboard_text(text)
    hwp.HAction.Run("Paste")
    time.sleep(0.2)


def build_hwp() -> None:
    if not TEMPLATE_ASCII.exists():
        raise FileNotFoundError(f"Template copy not found: {TEMPLATE_ASCII}")
    shutil.copy2(TEMPLATE_ASCII, OUT_ASCII)
    hwp = win32com.client.dynamic.Dispatch("HWPFrame.HwpObject")
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        if not hwp.Open(str(OUT_ASCII), "HWP", "forceopen:true"):
            raise RuntimeError("HWP open failed")
        hwp.HAction.Run("MoveDocEnd")
        set_font(hwp)
        paste_unicode(hwp, build_report())
        hwp.HAction.Run("SelectAll")
        set_font(hwp)
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
    text = subprocess.check_output(["hwp5txt.exe", str(OUT_ASCII)], text=True, encoding="utf-8", errors="replace", timeout=30)
    markers = ["Analysis Report", "문서 목적", "목차", "공학적 마찰손실", "시공사 경제성", "관경 축소"]
    print("MARKERS=" + ", ".join(f"{m}:{m in text}" for m in markers))
    print("QUESTION_MARK_COUNT=", text.count("?"))
    print("TEXT_LENGTH=", len(text))
    hwp = win32com.client.dynamic.Dispatch("HWPFrame.HwpObject")
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        print("COM_OPEN=", bool(hwp.Open(str(OUT_ASCII), "HWP", "forceopen:true")))
    finally:
        hwp.Quit()


if __name__ == "__main__":
    build_hwp()
    validate()
