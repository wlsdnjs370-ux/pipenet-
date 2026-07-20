# -*- coding: utf-8 -*-
"""다이소 양주 Layer 2 — 도면→추출 vs 답안 실측 격차.

실행:
    python calibration/daiso_yangju_layer2.py

흐름
====
1) 대상 도면 DWG 를 작업폴더로 복사 (참고용 도서 폴더 오염 방지)
2) ODA File Converter 로 DWG→DXF
3) run_remote30_extraction 으로 우리 추출망 SDF 생성 → parse_sdf → CommonNetwork
4) 답안(K-160 지상4층 REMOTE) 과 프로파일 비교 + 보수적 채점
5) 파탄 지점 진단(정렬 실패 = 형상 격차 신호)

주의: MF-101 은 '지하1층~지상4층' 다층 합본 시트라 추출이 거칠 수 있다 — 그 거칠음
자체가 첫 실측 신호다.
"""
from __future__ import annotations

import os
import shutil
import sys
import traceback
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from kfp_sdf_converter import parse_sdf  # noqa: E402
import score_network as sn  # noqa: E402

REF = _ROOT / "수리계산 참고용 도서" / "1. 저수조_아성다이소 양주허브센터"
DRAWING = REF / "도면" / "MF-101~ 지하1층~지상4층 소화설비 평면도.dwg"
ANSWER = REF / "수리계산 원본" / "물류동" / "1-1. 지상4층 K-160 가장 먼 구간_REMOTE.sdf"
WORK = _ROOT / "calibration" / "_work"


def _locate_oda_exe():
    env = os.environ.get("ODA_FILE_CONVERTER_EXE")
    if env and Path(env).is_file():
        return env
    for base in (Path(r"C:/Program Files/ODA"), Path(r"C:/Program Files (x86)/ODA")):
        if base.is_dir():
            hits = sorted(base.glob("*/ODAFileConverter.exe"), reverse=True)
            if hits:
                return str(hits[0])
    return None


def dwg_to_dxf(dwg_path: Path, out_dir: Path) -> Path:
    import ezdxf
    from ezdxf.addons import odafc
    exe = _locate_oda_exe()
    if exe:
        try:
            ezdxf.options.set("odafc-addon", "win_exec_path", exe)
        except Exception:
            pass
    if not odafc.is_installed():
        raise RuntimeError("ODA File Converter 미설치 — DWG→DXF 불가")
    out_dir.mkdir(parents=True, exist_ok=True)
    local_dwg = out_dir / dwg_path.name
    shutil.copy2(dwg_path, local_dwg)
    dxf_path = local_dwg.with_suffix(".dxf")
    odafc.convert(str(local_dwg), str(dxf_path), replace=True)
    if not dxf_path.exists():
        raise RuntimeError("DXF 변환 결과 없음")
    return dxf_path


def main():
    print("=" * 100)
    print("LAYER 2 — 도면→추출 vs 답안 실측 격차 (다이소 양주 / 지상4층 K-160)")
    print("=" * 100)
    print(f"도면 : {DRAWING.name}")
    print(f"답안 : {ANSWER.name}\n")

    # 1~2) DWG→DXF
    print("[1/4] DWG → DXF 변환 중...")
    try:
        dxf = dwg_to_dxf(DRAWING, WORK)
        print(f"      OK: {dxf.name} ({dxf.stat().st_size/1e6:.1f} MB)\n")
    except Exception as exc:  # noqa: BLE001
        print(f"      변환 실패: {exc}")
        traceback.print_exc()
        return 1

    # 3) 추출
    print("[2/4] run_remote30_extraction 실행 중 (N=30)...")
    try:
        from sprinkler_remote30_extractor import run_remote30_extraction
        result = run_remote30_extraction(
            str(dxf), out_dir=WORK / "extract_out",
            overrides={"remote_head_count": 30},
        )
    except Exception as exc:  # noqa: BLE001
        print(f"      추출 실패: {exc}")
        traceback.print_exc()
        return 1
    print("      counts:", result.get("counts"))
    if result.get("warnings"):
        print("      warnings:")
        for w in result["warnings"][:10]:
            print("        -", w)
    sdf_path = result.get("sdf_path")
    if not sdf_path:
        print("      !! SDF 미생성 (selected heads 0 또는 emit_sdf off) — 추출이 망을 못 만듦.")
        print("      → 이것이 첫 실측 신호: 다층 합본 시트에서 단일 설계구역 격리 실패 가능성.")
        return 2
    print(f"      추출 SDF: {Path(sdf_path).name}\n")

    # 4) 프로파일 + 채점
    print("[3/4] 프로파일 비교")
    extracted = parse_sdf(sdf_path)
    answer = parse_sdf(str(ANSWER))
    pe, pa = sn.profile(extracted), sn.profile(answer)
    print("  추출:", pe.as_row())
    print("  답안:", pa.as_row())
    print("  추출 kinds:", pe.kind_counts, "구경:", pe.diameter_counts)
    print("  답안 kinds:", pa.kind_counts, "구경:", pa.diameter_counts, "\n")

    print("[4/4] 보수적 채점 (추출 vs 답안)")
    r = sn.score(extracted, answer)
    print(r.report())
    print()
    print("-" * 100)
    if r.overall_pass:
        print("결과: PASS — 첫 시도부터 정합(예상 밖). 정렬 신뢰성 재확인 필요.")
    else:
        print("결과: FAIL (예상) — 아래가 격차 진단:")
        print("  · head_count 차이 = 구역 격리/헤드선정 격차")
        print("  · components 차이 = 망 연결 복원 격차")
        print("  · total_length/구간 = 길이/축척 격차")
        print("  · matched_segments 0 근처 = 좌표 정렬 실패 → 강건한 등록(registration) 필요")
    return 0


if __name__ == "__main__":
    sys.exit(main())
