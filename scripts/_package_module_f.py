# -*- coding: utf-8 -*-
"""모듈 F 소스 묶음 — 바탕화면에 .zip 하나로.

**경계는 추측이 아니라 실제 import 에서 뽑았다.** `routes/module_f/*.py` 가
부르는 저장소 모듈을 훑어 나온 것만 담는다(모듈 A · core 6종 · G 엔진).

담지 «않는» 것과 그 이유는 README.md 에 그대로 적어 넣는다 — 무엇이 빠졌는지
모르는 묶음은 받는 사람이 되살릴 수 없다.

실행: python scripts/_package_module_f.py [출력폴더]
"""
from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
DESKTOP = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.home() / "Desktop"

# ── 담을 것 ────────────────────────────────────────────────────────────
# (glob 또는 파일 경로, 묶음 안 설명)
SPEC: list[tuple[str, str]] = [
    ("routes/module_f/*.py",            "모듈 F 라우트 — 본체"),
    ("templates/module_f.html",         "화면 — 틀"),
    ("static/module_f.js",              "화면 — 동작"),
    ("static/module_f.css",             "화면 — 모양"),
    ("static/upload_stream.js",         "화면 — 업로드(압축·진행률) 공용 헬퍼"),
    ("대조 서버.py",                     "호스트 — 모듈 F 등록 · _save_upload"),
    ("serve.py",                        "호스트 — waitress 진입점"),
    ("remote30_prototype.py",           "모듈 A — 정찰·헤드검출·레이어분류·도면장"),
    ("kfp_sdf_converter.py",            "KFP ↔ SDF 변환"),
    ("core/fitting_rules.py",           "부속 규칙"),
    ("core/nftc_rules.py",              "NFTC 103 기준개수 표"),
    ("core/has_converter.py",           "HAS(하스) 입출력"),
    ("core/upload_names.py",            "업로드 파일명 정화"),
    ("core/remote30_full_network.py",   "제5국면 S700 원시함수(라이저·결합)"),
    ("core/remote30_graph.py",          "그래프 유틸"),
    ("tests/test_module_f_*.py",        "시험 — 모듈 F"),
    ("tests/test_g_graph_index_incremental.py",
                                        "시험 — G 그래프 색인 동등성"),
    ("scripts/_verify_module_f*.py",    "검증 — 전 경로·브라우저·이름"),
    ("scripts/_compare_kfp_canonical.py",
                                        "검증 — .kfp 정준 비교(산출 동등성)"),
    ("scripts/_package_module_f.py",    "이 묶음을 만든 스크립트"),
]

# G 엔진은 «소스만». docs/ 는 작업폴더(캐시·도면 산출물)라 112MB 다.
ENGINE_ROOT = "cad_project_editor_g"
ENGINE_SKIP_DIRS = {"__pycache__", "docs", "tests", ".git", "_out"}

EXCLUDED_NOTE = """\
■ 일부러 담지 않은 것

  cad_project_editor_g/docs/     112 MB. 소스가 아니라 **작업폴더**다 —
                                 찍은스펙·표시캐시·유저손질·변환 산출물이
                                 쌓이는 곳으로, 도면을 열면 다시 생긴다.
  cad_project_editor_g/tests/    2.4 MB. G 자체 수용시험 + _out 산출물.
                                 모듈 F 시험은 tests/ 에 담겨 있다.
  data/ · samples/ · 도면(.dxf)  실도면은 수백 MB 다. 검증 스크립트가
                                 가리키는 경로는 README 아래에 적어 둔다.
  __pycache__ · .venv · .git     빌드·환경 산물.

■ 이 묶음만으로 «돌지» 는 않는다

  모듈 F 는 호스트(대조 서버.py)와 G 엔진 위에서 도는 웹 모듈이고, 실제로
  열려면 실도면과 파이썬 환경(flask·waitress·ezdxf·shapely 등)이 필요하다.
  이 묶음은 **읽고 이어받기 위한 소스 일습**이지 배포본이 아니다.
"""


def _git(*args: str) -> str:
    try:
        return subprocess.run(["git", *args], cwd=ROOT, capture_output=True,
                              text=True, encoding="utf-8", timeout=60
                              ).stdout.strip()
    except Exception:  # noqa: BLE001 — git 이 없어도 묶음은 만든다
        return ""


def collect() -> list[tuple[Path, str]]:
    """(실제 경로, 묶음 안 상대경로) 목록. 중복은 한 번만."""
    out: list[tuple[Path, str]] = []
    seen: set[str] = set()

    def add(p: Path) -> None:
        rel = p.relative_to(ROOT).as_posix()
        if rel not in seen and p.is_file():
            seen.add(rel)
            out.append((p, rel))

    for pattern, _desc in SPEC:
        if any(ch in pattern for ch in "*?["):
            for p in sorted(ROOT.glob(pattern)):
                add(p)
        else:
            p = ROOT / pattern
            if p.is_file():
                add(p)
            else:
                print(f"  !! 없음(건너뜀): {pattern}")

    eng = ROOT / ENGINE_ROOT
    for dirpath, dirnames, filenames in os.walk(eng):
        dirnames[:] = [d for d in dirnames if d not in ENGINE_SKIP_DIRS]
        for fn in sorted(filenames):
            if fn.endswith(".py"):
                add(Path(dirpath) / fn)
    return out


def readme(files: list[tuple[Path, str]], head: str, subject: str) -> str:
    total = sum(p.stat().st_size for p, _ in files)
    by_group: dict[str, list[str]] = {}
    for _p, rel in files:
        top = rel.split("/")[0] if "/" in rel else "(루트)"
        if rel.startswith(ENGINE_ROOT):
            top = ENGINE_ROOT
        by_group.setdefault(top, []).append(rel)

    lines = [
        "# 모듈 F — 소스 일습",
        "",
        f"뽑은 날짜 : {datetime.now():%Y-%m-%d %H:%M}",
        f"커밋      : {head}",
        f"            {subject}",
        f"파일      : {len(files)}개 · {total / 1024 / 1024:.1f} MB",
        "",
        "모듈 F 는 «CAD 도면 → 수리계산 입력» 웹 워크벤치다. 모듈 G(데스크톱",
        "CAD 편집기)의 찍기·손질·변환 엔진을 소스 수정 없이 HTTP 로 열고,",
        "캔버스만 새로 그린 것이다. 화면 경로는 `/module-f`.",
        "",
        "## 무엇이 들어 있나",
        "",
    ]
    for pattern, desc in SPEC:
        lines.append(f"  {pattern:<44}  {desc}")
    lines += [
        f"  {ENGINE_ROOT + '/**/*.py':<44}  모듈 F 가 무는 엔진(G) — 소스만",
        "",
        "## 폴더별 파일 수",
        "",
    ]
    for top in sorted(by_group):
        lines.append(f"  {top:<28} {len(by_group[top]):>4}개")
    lines += [
        "",
        "## 최근 작업 (2026-09-01~02)",
        "",
        "  · 업로드 지연 — 도면을 «잡이 끝나기 전에» 그린다(world_ready).",
        "    실측 B1F 110.6MB: 42초 중 33초가 «이미 서버에 있는 도면» 을 안",
        "    그린 채 흘렀다. 업로드에 gzip+진행률(13.3 → 1.5 MB · 8.9배).",
        "  · 전체망 변환 942.7초 → 15.0초. 병목은 알고리즘이 아니라 그래프",
        "    색인의 전수 재구축이었다(변경 1건마다). 증분 유지 + 겹침 검사",
        "    색인(frozen_geometry)으로 해소. 산출물은 정준 비교로 동일 확인.",
        "  · 재검토로 «조용히 틀릴 자리» 4건 차단(잠금·색인 자동 차단·그래프",
        "    바꿔치기·무효화 함정). 시험 1,585건 통과.",
        "",
        "## 시험·검증 돌리는 법 (저장소 루트에서)",
        "",
        "  python -m pytest tests/test_module_f_*.py -q",
        "  python -m pytest tests/test_g_graph_index_incremental.py -q",
        "  python scripts/_verify_module_f.py           # 전 경로(실도면 필요)",
        "  python scripts/_verify_module_f_names.py     # 자유이름 정적 검사",
        "  MF_BASE=http://127.0.0.1:5065 LOGIN_PASSWORD=… \\",
        "    python scripts/_verify_module_f_upload_browser.py",
        "",
        "  ★`_verify_module_f.py` 는 실도면을 가리킨다(파일 머리 `DXF =`).",
        "    없으면 2~4단을 건너뛰고 그 사실을 화면에 적는다.",
        "",
        EXCLUDED_NOTE,
    ]
    return "\n".join(lines) + "\n"


def main() -> None:
    head = _git("rev-parse", "--short", "HEAD") or "(git 없음)"
    subject = _git("log", "-1", "--pretty=%s") or ""
    files = collect()
    stamp = f"{datetime.now():%Y%m%d}"
    base = f"모듈F_소스_{stamp}_{head}"
    DESKTOP.mkdir(parents=True, exist_ok=True)
    out = DESKTOP / f"{base}.zip"

    print(f"대상 {len(files)}개 파일 → {out}")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        z.writestr(f"{base}/README.md", readme(files, head, subject))
        for p, rel in files:
            z.write(p, f"{base}/{rel}")

    size = out.stat().st_size
    print(f"완료 · {out}")
    print(f"     {len(files) + 1}개 항목 · {size / 1024 / 1024:.1f} MB")


if __name__ == "__main__":
    main()
