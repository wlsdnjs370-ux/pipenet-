"""Remote 30 프로토타입 소스코드 zip 패키징."""

from __future__ import annotations

import datetime as _dt
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
OUT_DIR = BASE / "data" / "prototype_source"
OUT_DIR.mkdir(parents=True, exist_ok=True)

stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M")
ZIP_PATH = OUT_DIR / f"remote30_prototype_source_{stamp}.zip"


# ─────────────────────────────────────────────────────────────────
# 1. Flask 라우트 발췌 (대조 서버.py L2664-2837)
# ─────────────────────────────────────────────────────────────────
SERVER_SRC = BASE / "대조 서버.py"
server_lines = SERVER_SRC.read_text(encoding="utf-8").splitlines()
# 2664-2837 (1-indexed)
routes_block = "\n".join(server_lines[2663:2837])

FLASK_EXCERPT = f'''"""대조 서버.py — Remote 30 프로토타입 Flask 라우트 발췌.

원본 위치: 대조 서버.py L2664-L2837
의존 import (대조 서버.py 상단에서 가져오는 것들):

    import json, secrets
    from pathlib import Path
    from flask import Flask, jsonify, make_response, render_template, request, Response, send_file
    from werkzeug.utils import secure_filename

    BASE_DIR = Path(__file__).resolve().parent
    app = Flask(__name__)

`_save_upload(field, exts, required)` 헬퍼는 대조 서버.py 내부에서 정의된 공용
유틸리티로, multipart/form-data 의 dxf_file 필드를 받아 data/uploads/ 에
저장하고 Path 객체를 반환합니다.
"""

import json
import secrets
from pathlib import Path

from flask import Flask, jsonify, make_response, render_template, request, Response, send_file
from werkzeug.utils import secure_filename

# 아래 라우트들이 의존하는 전역:
#   app          : Flask 인스턴스
#   BASE_DIR     : 프로젝트 루트
#   _save_upload : 업로드 헬퍼 (대조 서버.py 내부 유틸)


{routes_block}
'''

# ─────────────────────────────────────────────────────────────────
# 2. index.html 메뉴 9번째 카드 발췌
# ─────────────────────────────────────────────────────────────────
INDEX_EXCERPT = """<!--
templates/index.html 발췌 — Remote 30 프로토타입 메뉴 카드 + iframe 패널
원본은 1500+ 라인의 SPA 진입점.
-->

<!-- ① 모듈 선택 버튼 (대략 L61) -->
<button type="button" class="home-module-btn" data-panel="remote30-prototype-panel">
  <strong>Remote 30 프로토타입</strong>
  <span>DXF 한 장 업로드 → 배관망 추출 → 가장 불리한 30 헤드 색출 →
        PIPENET 입력 XLSX/CSV → SDF 출력까지의 4-stage 전과정을
        캔버스에 실시간 애니메이션으로 보여줍니다.</span>
</button>

<!-- ② iframe 패널 (대략 L268) -->
<section id="remote30-prototype-panel" class="card hidden menu-panel module-fullscreen-panel embedded-iframe-panel">
  <iframe
    id="remote30-prototype-frame"
    class="module-iframe module-iframe-full"
    src="{{ url_for('remote30_prototype') }}"
    title="Remote 30 프로토타입"
  ></iframe>
</section>
"""

# ─────────────────────────────────────────────────────────────────
# 3. README
# ─────────────────────────────────────────────────────────────────
README = """# Remote 30 프로토타입 — 소스코드 패키지

DXF 한 장을 업로드하면 PIPENET SDF 까지 자동 생성하는 6-stage
파이프라인 모듈의 소스코드 모음입니다.

## 파일 목록

| 파일 | 역할 | 라인 수 |
|------|------|---------|
| `remote30_prototype.py` | 파이프라인 코어 (Stage 0~5 전체 로직) | ~1700 |
| `templates/remote30_prototype.html` | 프론트엔드 (캔버스 + SSE 클라이언트) | ~950 |
| `server_routes_excerpt.py` | Flask 엔드포인트 5종 (대조 서버.py L2664-2837 발췌) | ~175 |
| `templates/index_excerpt.html` | 메뉴 진입점 발췌 | ~25 |
| `scripts/test_prototype_2phase.py` | 2-phase end-to-end 검증 스크립트 | ~85 |

## 파이프라인 6단계

```
Stage 0: DXF 파싱 (ezdxf + INSERT 가상 분해)
Stage 1: PIPENET-only 필터링 (배관/노즐 레이어만)
Stage 2: 헤드 후보 바운딩박스 인식 (5-Rule + 250 mm 클러스터)
   ─── ⏸ 사용자 편집 (head add/delete + zone) ───
Stage 3: 전체 배관망 그래프 시각화 (snap 200mm + bridge 200/500/1000/2000mm)
Stage 4: 알람밸브 → 헤드 Dijkstra → 최불리 30 헤드 + 경로
Stage 5: 4-table 빌드 (Pipes/Nozzles/Sources/Pumps) + XLSX/CSV/SDF emit
```

## API 엔드포인트 (5)

```
GET  /remote30-prototype                       # HTML 진입
POST /api/remote30/prototype/run               # DXF 업로드 → job_id
GET  /api/remote30/prototype/stream/<id>       # Stage 0~2 SSE
POST /api/remote30/prototype/finalize/<id>     # 편집 데이터 수신
GET  /api/remote30/prototype/finalize_stream/<id>  # Stage 3~5 SSE
GET  /api/remote30/prototype/result/<id>/<filename>  # 결과 다운로드
```

## 외부 의존성

- ezdxf >= 1.0
- numpy, scipy (Dijkstra 외)
- pipenet_converter (로컬 패키지 — Pydantic 모델 + SDF writer)
  - `pipenet_converter.models`
  - `pipenet_converter.sdf_writer.write_sdf`
- Flask, openpyxl (XLSX 출력)

## 실행 방법

서버를 띄운 뒤 (`python "대조 서버.py"` → http://127.0.0.1:5050) 메뉴 9번
"Remote 30 프로토타입" 카드를 누르면 됩니다.

테스트 스크립트는 서버를 켠 상태에서:

```bash
python scripts/test_prototype_2phase.py
```

## 결과물 (Stage 5 산출)

```
data/prototype_runs/<job_id>/
├─ pipenet_input.xlsx
├─ pipenet_input_*.csv (4종)
└─ output.sdf  (PIPENET 3-1형, Graphics 블록 포함)
```
"""


# ─────────────────────────────────────────────────────────────────
# zip 생성
# ─────────────────────────────────────────────────────────────────
files_to_zip: list[tuple[Path | str, str, str | None]] = [
    # (source_or_marker, archive_name, optional_text_override)
    (BASE / "remote30_prototype.py", "remote30_prototype.py", None),
    (BASE / "templates" / "remote30_prototype.html",
     "templates/remote30_prototype.html", None),
    (BASE / "scripts" / "test_prototype_2phase.py",
     "scripts/test_prototype_2phase.py", None),
    ("__TEXT__", "server_routes_excerpt.py", FLASK_EXCERPT),
    ("__TEXT__", "templates/index_excerpt.html", INDEX_EXCERPT),
    ("__TEXT__", "README.md", README),
]

with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
    for src, arcname, override in files_to_zip:
        if override is not None:
            zf.writestr(arcname, override)
        else:
            assert isinstance(src, Path)
            zf.write(src, arcname=arcname)
        print(f"  + {arcname}")

print(f"\n[done] {ZIP_PATH}  ({ZIP_PATH.stat().st_size / 1024:.1f} KB)")
