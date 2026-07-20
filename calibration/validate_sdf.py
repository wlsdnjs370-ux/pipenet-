# -*- coding: utf-8 -*-
"""추출 결과(SDF/KFP) 1건을 prior 밴드에 비춰 '배관망스러운가' 점검하는 CLI.

답안-키 1:1 채점이 불가한 코퍼스(도면≫답안·좌표계 불공유)의 차선책. 정상 답안
(REMOTE)에서 학습한 분포 밴드에, 우리 추출망이 드는지 자동 점검한다. 정답을 외우는
게 아니라 '배관망스러움'을 본다 — 정답 없는 새 도면에도 적용 가능.

사용:
    python calibration/validate_sdf.py <추출.sdf | 추출.kfp>
    python calibration/validate_sdf.py            # 인자 없으면 답안 자기검증

밴드는 ``수리계산 참고용 도서`` 전 패키지의 **스프링클러 REMOTE 설계구역**에서
학습한다. 옥내소화전은 위상이 전혀 달라(소화전 라이저가 분리돼 comp>1, n/head≈7)
스프링클러 추출기 밴드를 오염시키므로 제외한다.
"""
from __future__ import annotations

import glob
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from kfp_sdf_converter import parse_sdf, parse_kfp  # noqa: E402
import score_network as sn  # noqa: E402

ANSWER_GLOBS = [
    str(_ROOT / "수리계산 참고용 도서" / "**" / "*.sdf"),
]
# 옥내소화전은 스프링클러와 위상이 달라 밴드 학습에서 제외.
EXCLUDE_KEYWORDS = ("옥내소화전",)


def _remote_answer_files():
    files = []
    for pat in ANSWER_GLOBS:
        files += glob.glob(pat, recursive=True)
    out = []
    for f in files:
        if "REMOTE" not in f.upper():
            continue
        if any(kw in f for kw in EXCLUDE_KEYWORDS):
            continue
        out.append(f)
    return sorted(out)


def learn_default_bands() -> sn.BandSet:
    profs = [sn.profile(parse_sdf(f)) for f in _remote_answer_files()]
    return sn.learn_bands(profs)


def _load_net(path: str):
    p = Path(path)
    if p.suffix.lower() == ".kfp":
        return parse_kfp(p)
    return parse_sdf(str(p))


def main(argv):
    bands = learn_default_bands()
    if len(argv) < 2:
        print("인자 없음 — 답안 자기검증으로 대체합니다.\n")
        print(bands.report(), "\n")
        for f in _remote_answer_files():
            rep = sn.validate(parse_sdf(f), bands)
            print(f"  {rep.verdict:<4} {os.path.basename(f)}")
        return 0

    target = argv[1]
    if not Path(target).is_file():
        print(f"파일 없음: {target}")
        return 1

    print(bands.report(), "\n")
    print(f"검증 대상: {target}\n")
    net = _load_net(target)
    rep = sn.validate(net, bands)
    print(rep.report())
    print()
    if rep.has_fail:
        print("판정: FAIL — 추출 실패(구조적 결함). 구역 격리/연결 복원 재점검 필요.")
    elif rep.has_warn:
        print("판정: WARN — 정상 분포 밖. 위 항목이 실제 설계 변형인지 추출 오류인지 확인.")
    else:
        print("판정: OK — 정상 설계구역 범위. (단, 범위 통과가 정확성 보장은 아님.)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
