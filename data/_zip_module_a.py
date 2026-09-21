# -*- coding: utf-8 -*-
"""모듈 A(Remote 30) 소스만 골라 바탕화면에 zip 으로 낸다. 데이터·도면·venv 제외."""
from __future__ import annotations

import io
import os
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DESK = Path(os.path.expanduser("~")) / "Desktop"
OUT = DESK / "모듈A_소스코드_20260824.zip"

ENGINE = [
    "remote30_prototype.py",
    "sprinkler_remote30_extractor.py",
    "kfp_sdf_converter.py",
]
CORE = [
    "core/remote30_constants.py",
    "core/remote30_full_network.py",
    "core/remote30_graph.py",
    "core/remote30_ml.py",
    "core/fitting_rules.py",
    "core/has_converter.py",
]
ROUTES = [
    "routes/r30_prototype.py",
    "routes/r30_overall.py",
    "routes/r30_combined.py",
    "routes/r30_inspect.py",
    "routes/r30_machineroom.py",
    "routes/r30_system.py",
    "routes/r30_gnn.py",
    "routes/r30_has.py",
]
TEMPLATES = [
    "templates/remote30_prototype.html",
    "templates/remote30_overall.html",
    "templates/remote30_workbench.html",
    "templates/remote30_workbench_gnn.html",
    "templates/sprinkler_pipeline.html",
    "templates/_sprinkler_pipeline_panel.html",
]
SCRIPTS = [
    "scripts/generate_sprinkler_yolo_dataset.py",
    "scripts/train_sprinkler_yolo.py",
    "scripts/_diag_head_detect_generalize.py",
    "scripts/_verify_head_generalize.py",
    "scripts/_verify_graph_corpus.py",
]

GROUPS = [("엔진", ENGINE), ("core", CORE), ("라우트", ROUTES),
          ("템플릿", TEMPLATES), ("도구·검증", SCRIPTS)]

rows, missing, total = [], [], 0
for label, files in GROUPS:
    for rel in files:
        p = ROOT / rel
        if not p.is_file():
            missing.append(rel)
            continue
        n = p.stat().st_size
        lines = sum(1 for _ in io.open(p, encoding="utf-8", errors="replace"))
        rows.append((label, rel, lines, n))
        total += n

man = ["# 모듈 A (Remote 30) 소스 묶음", "",
       "생성: 2026-08-24 · 저장소 `JupyterProject` @ main", "",
       "## 담긴 것", ""]
cur = None
for label, rel, lines, n in rows:
    if label != cur:
        man.append(f"### {label}")
        cur = label
    man.append(f"- `{rel}` — {lines:,} 줄 · {n:,} bytes")
man += ["", f"합계 {len(rows)}개 파일 · {total:,} bytes", "",
        "## 일부러 뺀 것", "",
        "- `routes/r30_design.py` · `templates/design_workbench.html`",
        "  — `/design-workbench` 는 설계 워크벤치(모듈 C·D)라 모듈 A 가 아니다.",
        "- `pipenet_converter/` — 11번 모듈(별도). 엔진이 `models`/`sdf_writer` 만",
        "  가져다 쓴다.",
        "- `대조 서버.py` — 앱 호스트. 위 라우트들을 `register(app, ...)` 로 등록만 한다.",
        "- 도면(.dxf/.dwg)·산출물·`data/`·`.venv` — 소스가 아니다.", ""]
if missing:
    man += ["## 목록에 있었으나 없는 파일", ""] + [f"- `{m}`" for m in missing] + [""]

OUT.parent.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    for _label, rel, _l, _n in rows:
        z.write(ROOT / rel, rel)
    z.writestr("MANIFEST.md", "\n".join(man))

print(f"파일 {len(rows)}개 · 원본 {total:,} bytes")
if missing:
    print("빠짐:", ", ".join(missing))
print(f"→ {OUT}  ({OUT.stat().st_size:,} bytes)")
