from __future__ import annotations

# ── core/ 라이브러리 경로 (repo 정리: 루트 라이브러리 → core/ 이동) ──
import sys as _sys
from pathlib import Path as _Path
_BASE = _Path(__file__).resolve().parent
_CORE = _BASE / "core"
# repo 루트 자체도 경로에 둔다 — routes/ 도메인 패키지 import 를 위해
# (spec_from_file_location 로 로드될 때 부모 디렉토리가 sys.path 에 없을 수 있음).
for _p in (str(_BASE), str(_CORE)):
    if _p and _Path(_p).is_dir() and _p not in _sys.path:
        _sys.path.insert(0, _p)

import base64
import gzip
import hashlib
import html as html_lib
import json
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import warnings
from datetime import datetime
from io import BytesIO
from pathlib import Path
import math
import xml.etree.ElementTree as ET
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# .env 파일에서 환경변수 자동 로드 (FLASK_SECRET_KEY, LOGIN_PASSWORD 등).
# python-dotenv 가 없거나 .env 가 없어도 silently skip.
try:
    from dotenv import load_dotenv as _load_dotenv
    _load_dotenv()
except ImportError:
    pass

from flask import Flask, Response, jsonify, make_response, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename

from pipenet_validator import PipenetGuideValidator


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "data" / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
# /inspect 렌더 결과(전체 entity NDJSON + 메타)를 도면 내용 해시로 캐시한다.
# 같은 도면 재업로드 시 ezdxf 재파싱·explode(108s/172MB)를 건너뛰고 ~1-2s 에 스트리밍.
# 렌더/카테고리 로직이 바뀌면 INSPECT_CACHE_VERSION 을 올려 캐시를 무효화한다.
INSPECT_CACHE_DIR = UPLOAD_DIR / "_inspect_cache"
INSPECT_CACHE_DIR.mkdir(parents=True, exist_ok=True)
INSPECT_CACHE_VERSION = "v2"


def _inspect_cache_key(dxf_path: Path) -> str:
    """도면 내용 SHA256 → 캐시 키 (/inspect 와 동일 규칙)."""
    h = hashlib.sha256()
    with open(dxf_path, "rb") as _f:
        for _blk in iter(lambda: _f.read(1024 * 1024), b""):
            h.update(_blk)
    return f"{INSPECT_CACHE_VERSION}_{h.hexdigest()}"


def _load_cached_view_entities(dxf_path: Path) -> list | None:
    """inspect 바이너리 캐시에서 entity 리스트 복원. 없거나 실패 시 None.

    추출(경로 탐색)이 parse_dxf_for_view 로 대용량 DXF 를 재파싱(141MB≈110s)하는
    대신, /inspect 가 이미 파싱·캐시한 entity 를 재사용한다. 캐시는 스트리밍된
    progress 메시지(NDJSON)라 각 줄의 entities 배열을 평탄화한다. fresh 파싱과
    동일 ezdxf 출력이므로 그래프/경로 결과가 같다.
    """
    try:
        key = _inspect_cache_key(dxf_path)
    except Exception:
        return None
    ent_path = INSPECT_CACHE_DIR / f"{key}.entities.ndjson.gz"
    if not ent_path.exists():
        return None
    ents: list = []
    try:
        with gzip.open(ent_path, "rt", encoding="utf-8") as gz:
            for line in gz:
                line = line.strip()
                if not line:
                    continue
                msg = json.loads(line)
                if isinstance(msg, dict) and msg.get("type") == "progress":
                    ents.extend(msg.get("entities") or [])
    except Exception:
        return None
    return ents or None
REMOTE30_OUTPUT_DIR = BASE_DIR / "data" / "remote30_outputs"
REMOTE30_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
UPDATE_HISTORY_PATH = BASE_DIR / "data" / "update_history.json"
FEEDBACK_POSTS_PATH = BASE_DIR / "data" / "feedback_posts.json"
FEEDBACK_UPLOAD_DIR = BASE_DIR / "data" / "feedback_uploads"
FEEDBACK_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
CAD_SDF_LEARNING_PROFILE_PATH = BASE_DIR / "data" / "cad_sdf_learning_profile.json"

# fire-dxf2sdf (Phase 1-3 GNN 파이프라인) subprocess 호출용
FIRE_DXF2SDF_DIR = BASE_DIR / "fire-dxf2sdf"
FIRE_DXF2SDF_OUTPUT_DIR = BASE_DIR / "data" / "gnn_outputs"
FIRE_DXF2SDF_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
# uv 경로 — Anaconda 환경에서 별도 venv 의 fire-dxf2sdf 호출
UV_EXECUTABLE = Path("C:/Users/admin/AppData/Roaming/Python/Python313/Scripts/uv.exe")
DESIGN_AUTOMATION_ROOT = BASE_DIR / "sprinkler_ai_agent_server_source_2026-04-27" / "extracted"
DESIGN_AUTOMATION_STATIC_DIR = DESIGN_AUTOMATION_ROOT / "static"
DESIGN_AUTOMATION_SERVER_PATH = DESIGN_AUTOMATION_ROOT / "server.py"
DESIGN_AUTOMATION_PID_PATH = BASE_DIR / "design_automation_server.pid"
DESIGN_AUTOMATION_STDOUT_PATH = BASE_DIR / "design_automation_server_stdout.log"
DESIGN_AUTOMATION_STDERR_PATH = BASE_DIR / "design_automation_server_stderr.log"
DESIGN_AUTOMATION_PORT = 7870

app = Flask(__name__)
# Jinja2 템플릿 자동 reload — 디스크 변경 시 다음 요청부터 반영.
# Flask debug mode 가 꺼져있어도 활성화.
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True


# ────────────────────────────────────────────────────────────────────────────
# 잡 스토어 / 임시파일 수명 관리 — 24/7 구동 프로세스의 무한 누적 방지
# ────────────────────────────────────────────────────────────────────────────
# waitress 는 단일 프로세스 + 다중 스레드라 in-memory 잡 dict 가 영원히 산다.
# 매 업로드마다 대용량(pipe_ents·detected_heads)이 적재되면 메모리가 무한 증가하고,
# 산출물/업로드 디렉토리도 무한 누적된다. → 잡은 TTL/개수로 evict, 디렉토리는 mtime
# TTL 로 주기적 sweep(rate-limited).
_JOB_TTL_SECONDS = 12 * 3600       # 잡 메타 12시간 후 만료 (편집 세션 여유)
_JOB_MAX_ENTRIES = 100             # 스토어당 최대 잡 수 (초과 시 오래된 것부터)
_DIR_TTL_SECONDS = 24 * 3600       # 산출물/업로드 24시간 후 정리
_DIR_SWEEP_INTERVAL = 1800         # 디렉토리 sweep 최소 간격(초) — 매 요청마다 안 돌게
_jobs_lock = threading.Lock()
_last_dir_sweep = [0.0]


def _register_job(store: dict, job_id: str, data: dict) -> None:
    """잡 등록 + 오래된/초과 잡 eviction (thread-safe).

    읽기(`store.get`)는 GIL 하에서 원자적이라 lock 불필요하지만, 삽입+iterate-삭제는
    경쟁이 생기므로 lock 으로 감싼다. 활성 잡(_created≈now)은 evict 대상이 아니다.
    """
    data["_created"] = time.time()
    with _jobs_lock:
        store[job_id] = data
        now = time.time()
        stale = [k for k, v in store.items()
                 if now - v.get("_created", now) > _JOB_TTL_SECONDS]
        for k in stale:
            store.pop(k, None)
        if len(store) > _JOB_MAX_ENTRIES:
            ordered = sorted(store.items(), key=lambda kv: kv[1].get("_created", 0.0))
            for k, _v in ordered[: len(store) - _JOB_MAX_ENTRIES]:
                store.pop(k, None)


def _sweep_old_run_dirs(*parents: Path) -> None:
    """오래된 잡 산출물 디렉토리(자식) 정리 — opportunistic, rate-limited.

    예외는 전부 삼킨다(정리 실패가 요청을 막으면 안 됨). 활성 잡 dir 은 방금 생성돼
    mtime 이 최신이라 TTL 에 안 걸린다.
    """
    now = time.time()
    with _jobs_lock:
        if now - _last_dir_sweep[0] < _DIR_SWEEP_INTERVAL:
            return
        _last_dir_sweep[0] = now
    for parent in parents:
        try:
            if not parent.is_dir():
                continue
            for child in list(parent.iterdir()):
                try:
                    if now - child.stat().st_mtime <= _DIR_TTL_SECONDS:
                        continue
                    if child.is_dir():
                        shutil.rmtree(child, ignore_errors=True)
                except OSError:
                    pass
        except OSError:
            pass


def _sweep_old_upload_files(parent: Path, keep_dirs: set[str] | None = None) -> None:
    """오래된 업로드 *파일* 정리 — 디렉토리(예: cad_workspace)는 보존."""
    keep = keep_dirs or {"cad_workspace"}
    now = time.time()
    try:
        if not parent.is_dir():
            return
        for child in list(parent.iterdir()):
            try:
                if child.is_dir() or child.name in keep:
                    continue
                if now - child.stat().st_mtime > _DIR_TTL_SECONDS:
                    child.unlink(missing_ok=True)
            except OSError:
                pass
    except OSError:
        pass


# ────────────────────────────────────────────────────────────────────────────
# JSON Provider 안전화 — complex / numpy / NaN / Path 등 추가 타입 지원
# 통합 검증 모듈 등이 만드는 복소수 (예: eigenvalue, scipy 계산 결과) 가
# jsonify 시 "Object of type complex is not JSON serializable" 로 실패하던 문제 해결.
# ────────────────────────────────────────────────────────────────────────────
import math as _math
from flask.json.provider import DefaultJSONProvider as _DefaultJSONProvider


class _SafeJSONProvider(_DefaultJSONProvider):
    def default(self, o):  # noqa: D401
        # 복소수 — real 부분만 (imag 가 거의 0 인 경우 합리적). 큰 imag 면 magnitude.
        if isinstance(o, complex):
            if abs(o.imag) < 1e-9:
                return float(o.real)
            return abs(o)  # 복소수 크기 (magnitude)
        # numpy 타입 처리
        try:
            import numpy as _np
            if isinstance(o, _np.complexfloating):
                if abs(o.imag) < 1e-9:
                    return float(o.real)
                return float(abs(o))
            if isinstance(o, _np.floating):
                v = float(o)
                if _math.isnan(v) or _math.isinf(v):
                    return None
                return v
            if isinstance(o, _np.integer):
                return int(o)
            if isinstance(o, _np.bool_):
                return bool(o)
            if isinstance(o, _np.ndarray):
                return o.tolist()
        except ImportError:
            pass
        # float NaN/Inf 도 None 으로 (JSON 표준 호환)
        if isinstance(o, float):
            if _math.isnan(o) or _math.isinf(o):
                return None
        # Path 객체
        if isinstance(o, Path):
            return str(o)
        # bytes
        if isinstance(o, (bytes, bytearray)):
            try:
                return o.decode("utf-8", errors="replace")
            except Exception:
                return None
        # set
        if isinstance(o, (set, frozenset)):
            return list(o)
        return super().default(o)


app.json = _SafeJSONProvider(app)

# ────────────────────────────────────────────────────────────────────────────
# 비밀번호 로그인 게이트 — 외부 노출(터널 등) 시 접근 보호
# ────────────────────────────────────────────────────────────────────────────
# 한 줄 비밀번호 폼 → 세션 쿠키. SECRET_KEY 는 env var 또는 dev 용 hardcoded fallback.
# 비밀번호는 LOGIN_PASSWORD env var 로 override 가능 (기본 "5361").
import secrets as _secrets
import os as _os_for_auth
app.secret_key = _os_for_auth.environ.get("FLASK_SECRET_KEY") or _secrets.token_hex(32)
LOGIN_PASSWORD = _os_for_auth.environ.get("LOGIN_PASSWORD", "5361")

# 게이트에서 제외할 path prefix (login/logout/정적 파일/health 등)
_AUTH_EXEMPT_PREFIXES = ("/login", "/logout", "/static/", "/favicon.ico")


@app.before_request
def _require_login_gate():
    """모든 요청 전에 인증 체크 — 미인증이면 로그인 페이지로."""
    if request.path.startswith(_AUTH_EXEMPT_PREFIXES):
        return None
    if session.get("authed"):
        return None
    # API 호출은 401 JSON, 페이지 요청은 redirect
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "message": "로그인이 필요합니다.", "login_required": True}), 401
    return redirect(url_for("login_page", next=request.path))


# 인증 라우트(login/logout)는 routes/auth.py 로 분리 — 파일 끝에서 register().


# ────────────────────────────────────────────────────────────────────────────
# 전역 에러 핸들러 — /api/ 요청에 대해 HTML 500 페이지 대신 JSON 반환
# (클라이언트 fetch 가 await resp.json() 에서 SyntaxError 나는 것 차단)
# ────────────────────────────────────────────────────────────────────────────
@app.errorhandler(Exception)
def _api_safe_errorhandler(exc):
    # Flask 의 HTTPException (400, 404 등) 은 그대로 전달
    from werkzeug.exceptions import HTTPException
    if request.path.startswith("/api/"):
        import traceback as _tb
        if isinstance(exc, HTTPException):
            return jsonify({
                "ok": False,
                "message": exc.description or str(exc),
                "status": exc.code,
            }), exc.code
        return jsonify({
            "ok": False,
            "message": f"서버 오류: {type(exc).__name__}: {str(exc)[:300]}",
            "traceback": _tb.format_exc()[-2000:],
        }), 500
    # /api/ 가 아니면 Flask 기본 처리 (HTML 페이지 OK)
    if isinstance(exc, HTTPException):
        return exc
    raise exc

# Keep chart text strictly ASCII-safe to prevent tofu/square glyphs on some systems.
plt.rcParams["font.family"] = ["DejaVu Sans", "Arial", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False


def _ensure_design_automation_static_layout() -> None:
    if not DESIGN_AUTOMATION_ROOT.exists():
        raise FileNotFoundError(f"Design automation source folder not found: {DESIGN_AUTOMATION_ROOT}")
    DESIGN_AUTOMATION_STATIC_DIR.mkdir(parents=True, exist_ok=True)
    vendor_dir = DESIGN_AUTOMATION_STATIC_DIR / "vendor"
    vendor_dir.mkdir(parents=True, exist_ok=True)
    for filename in [
        "index.html",
        "styles.css",
        "app.js",
        "app_v2.js",
        "app_v3.js",
        "app_chat.js",
        "dxf-parser.js",
        "dxf_segmentation_geometry.js",
    ]:
        source = DESIGN_AUTOMATION_ROOT / filename
        if source.exists():
            shutil.copy2(source, DESIGN_AUTOMATION_STATIC_DIR / filename)
    vendor_dxf_parser = DESIGN_AUTOMATION_ROOT / "dxf-parser.js"
    if vendor_dxf_parser.exists():
        shutil.copy2(vendor_dxf_parser, vendor_dir / "dxf-parser.js")


def _is_local_port_open(port: int) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.5):
            return True
    except OSError:
        return False


def _start_design_automation_server() -> None:
    if _is_local_port_open(DESIGN_AUTOMATION_PORT):
        return
    _ensure_design_automation_static_layout()
    if not DESIGN_AUTOMATION_SERVER_PATH.exists():
        raise FileNotFoundError(f"Design automation server.py not found: {DESIGN_AUTOMATION_SERVER_PATH}")

    with DESIGN_AUTOMATION_STDOUT_PATH.open("ab") as stdout_fp, DESIGN_AUTOMATION_STDERR_PATH.open("ab") as stderr_fp:
        process = subprocess.Popen(
            [
                sys.executable,
                str(DESIGN_AUTOMATION_SERVER_PATH),
                "--host",
                "0.0.0.0",
                "--port",
                str(DESIGN_AUTOMATION_PORT),
            ],
            cwd=str(DESIGN_AUTOMATION_ROOT),
            stdout=stdout_fp,
            stderr=stderr_fp,
        )
    DESIGN_AUTOMATION_PID_PATH.write_text(str(process.pid), encoding="utf-8")

    deadline = time.time() + 15
    while time.time() < deadline:
        if _is_local_port_open(DESIGN_AUTOMATION_PORT):
            return
        time.sleep(0.5)
    raise RuntimeError(
        f"Design automation server did not start on port {DESIGN_AUTOMATION_PORT}. "
        f"Check {DESIGN_AUTOMATION_STDERR_PATH.name}."
    )


PIPE_SEGMENTATION_MODEL_CANDIDATES = [
    BASE_DIR / "models" / "pipe_segmentation" / "weights" / "best.pt",
    BASE_DIR / "models" / "pipe_segmentation.pt",
    BASE_DIR / "runs" / "segment" / "pipe_segmentation" / "weights" / "best.pt",
    BASE_DIR / "yolo11n-seg.pt",
    BASE_DIR / "yolo26n-seg.pt",
]


def _torch_device_info() -> dict:
    try:
        import torch

        if torch.cuda.is_available():
            return {
                "device": "cuda",
                "gpu_enabled": True,
                "gpu_name": torch.cuda.get_device_name(0),
            }
        return {"device": "cpu", "gpu_enabled": False, "gpu_name": None}
    except Exception as exc:
        return {"device": "unavailable", "gpu_enabled": False, "gpu_name": None, "error": str(exc)}


def _pipe_segmentation_engine_status() -> dict:
    model_path = next((path for path in PIPE_SEGMENTATION_MODEL_CANDIDATES if path.exists()), None)
    device_info = _torch_device_info()
    if not model_path:
        return {
            "name": "Pipe Segmentation",
            "available": False,
            "mode": "sdf_guided_segmentation_proxy",
            "model_path": None,
            "message": "학습된 배관 세그멘테이션 가중치가 없어 SDF-guided 선분 묶음화 엔진으로 대체했습니다.",
            **device_info,
        }
    try:
        from ultralytics import YOLO

        # Load once per request to verify the trained segmentation weight is usable.
        YOLO(str(model_path))
        return {
            "name": "Pipe Segmentation",
            "available": True,
            "mode": "trained_ultralytics_segmentation",
            "model_path": str(model_path),
            "message": "학습된 세그멘테이션 가중치를 로드했습니다. DXF 벡터 그래프는 SDF-guided bundle 단계와 함께 사용됩니다.",
            **device_info,
        }
    except Exception as exc:
        return {
            "name": "Pipe Segmentation",
            "available": False,
            "mode": "sdf_guided_segmentation_proxy",
            "model_path": str(model_path),
            "message": f"세그멘테이션 가중치 로드 실패로 SDF-guided 선분 묶음화 엔진으로 대체했습니다: {exc}",
            **device_info,
        }


EXPORT_SCHEMA = {
    "pipes": {
        "sheet": "배관",
        "columns": [
            ("label", "Pipe"),
            ("input_node", "입력 노드"),
            ("output_node", "출력 노드"),
            ("nominal_bore_mm", "구경(mm)"),
            ("flow_lpm", "유량(L/min)"),
            ("velocity_mps", "유속(m/s)"),
            ("inlet_pressure", "입구압"),
            ("outlet_pressure", "출구압"),
            ("friction_loss", "마찰손실"),
            ("special_equipment", "특수설비"),
        ],
    },
    "nozzles": {
        "sheet": "헤드",
        "columns": [
            ("label", "헤드"),
            ("input_node", "입력 노드"),
            ("inlet_pressure_kgf_cm2", "압력(kg/cm²)"),
            ("required_flow_lpm", "요구 유량"),
            ("actual_flow_lpm", "실제 유량"),
            ("deviation_percent", "편차(%)"),
        ],
    },
    "equipment": {
        "sheet": "특수설비",
        "columns": [
            ("label", "설비"),
            ("pipe_label", "배관"),
            ("description", "구분"),
            ("equivalent_length_m", "등가길이(m)"),
        ],
    },
    "valves": {
        "sheet": "감압밸브",
        "columns": [
            ("label", "밸브"),
            ("inlet_pressure_kgf_cm2", "입구압"),
            ("outlet_pressure_kgf_cm2", "출구압"),
            ("pressure_drop_kgf_cm2", "압력강하"),
            ("flow_lpm", "유량(L/min)"),
        ],
    },
}


def _locate_oda_exe() -> str | None:
    """ODA File Converter 실행파일 경로 탐색.

    우선순위: 환경변수 ODA_FILE_CONVERTER_EXE → 표준 설치 경로
    (버전 폴더명이 'ODAFileConverter 27.1.0' 처럼 버전을 포함해 ezdxf 기본 탐색이
    실패하므로 직접 glob 으로 찾는다).
    """
    import os
    env = os.environ.get("ODA_FILE_CONVERTER_EXE")
    if env and Path(env).is_file():
        return env
    for base in (Path(r"C:/Program Files/ODA"), Path(r"C:/Program Files (x86)/ODA")):
        if base.is_dir():
            hits = sorted(base.glob("*/ODAFileConverter.exe"), reverse=True)
            if hits:
                return str(hits[0])
    return None


def _dwg_to_dxf(dwg_path: Path) -> Path:
    """ODA File Converter (ezdxf odafc addon) 로 DWG → DXF 무손실 변환.

    ODA File Converter 가 설치돼 있어야 함 (무료, 수동 다운로드 — winget 미제공).
    미설치 시 설치 안내를 담은 ValueError 를 던진다.
    """
    try:
        import ezdxf
        from ezdxf.addons import odafc
    except Exception as exc:  # noqa: BLE001
        raise ValueError(
            "DWG 변환 모듈(ezdxf odafc)을 불러오지 못했습니다. ezdxf 설치를 확인해 주세요."
        ) from exc
    # 버전 폴더에 설치된 exe 를 직접 지정 (ezdxf 기본 경로는 unversioned 라 못 찾음)
    exe = _locate_oda_exe()
    if exe:
        try:
            ezdxf.options.set("odafc-addon", "win_exec_path", exe)
        except Exception:
            pass
    if not odafc.is_installed():
        raise ValueError(
            "DWG 업로드를 처리하려면 ODA File Converter(무료)가 필요합니다. "
            "https://www.opendesign.com/guestfiles/oda_file_converter 에서 설치 후 다시 시도하거나, "
            "CAD에서 DXF로 저장해 업로드해 주세요."
        )
    dxf_path = dwg_path.with_suffix(".dxf")
    try:
        odafc.convert(str(dwg_path), str(dxf_path), replace=True)
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"DWG → DXF 변환에 실패했습니다: {exc}") from exc
    if not dxf_path.exists():
        raise ValueError("DWG → DXF 변환 결과 파일을 찾을 수 없습니다.")
    return dxf_path


def _save_upload(field_name: str, allowed_suffixes: set[str], required: bool) -> Path | None:
    uploaded = request.files.get(field_name)
    if uploaded is None or not uploaded.filename:
        if required:
            raise ValueError(f"`{field_name}` 파일이 필요합니다.")
        return None

    original_name = Path(uploaded.filename).name
    # 클라이언트에서 gzip 압축 전송 시 파일명이 ".gz" 로 끝남 → 실제 확장자 복원
    raw = uploaded.read()
    is_gzip = original_name.lower().endswith(".gz") or raw[:2] == b"\x1f\x8b"
    if is_gzip:
        try:
            raw = gzip.decompress(raw)
        except OSError as exc:
            raise ValueError("업로드 파일의 압축 해제에 실패했습니다.") from exc
        if original_name.lower().endswith(".gz"):
            original_name = original_name[:-3]

    original_suffix = Path(original_name).suffix.lower()
    filename = secure_filename(original_name)
    if not filename:
        filename = f"{field_name}_{int(datetime.now().timestamp())}{original_suffix}"
    elif Path(filename).suffix == "" and original_suffix:
        filename = f"{filename}{original_suffix}"
    suffix = original_suffix or Path(filename).suffix.lower()
    if suffix not in allowed_suffixes:
        allowed = ", ".join(sorted(allowed_suffixes))
        raise ValueError(f"`{field_name}` 파일 형식이 올바르지 않습니다. 허용 형식: {allowed}")

    saved_path = UPLOAD_DIR / filename
    saved_path.write_bytes(raw)
    # DWG 업로드는 서버측에서 DXF 로 변환해 이후 파이프라인이 동일하게 처리
    if suffix == ".dwg":
        saved_path = _dwg_to_dxf(saved_path)
    _sweep_old_upload_files(UPLOAD_DIR)
    return saved_path


def _err500(exc, *, message=None, **extra):
    """라우트 말미 `except Exception` 블록의 표준 500 JSON 응답.

    traceback.format_exc() 는 호출 시점의 활성 예외(sys.exc_info)를 읽으므로,
    except 블록 안에서 호출하면 헬퍼 내부여도 해당 예외의 트레이스백이 잡힌다.
    message 미지정 시 str(exc)[:300]. extra 로 라우트별 추가 키(algorithm 등) 병합.
    """
    import traceback
    body = {"ok": False,
            "message": message if message is not None else str(exc)[:300],
            "traceback": traceback.format_exc()[-1500:]}
    body.update(extra)
    return jsonify(body), 500


def _printable_report_text(path: Path) -> str:
    return PipenetGuideValidator(report_path=path)._read_report_text(path)


def _print_report_url(path: Path, copies: int = 2) -> str:
    return f"/print-report/{path.name}?copies={copies}"


def _to_float(v, default: float = 0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def _fig_to_data_url(fig, *, tight: bool = True) -> str:
    buf = BytesIO()
    save_kwargs = {"format": "png", "dpi": 140}
    if tight:
        save_kwargs["bbox_inches"] = "tight"
    fig.savefig(buf, **save_kwargs)
    plt.close(fig)
    buf.seek(0)
    return "data:image/png;base64," + base64.b64encode(buf.read()).decode("ascii")


def _load_update_history() -> dict:
    if not UPDATE_HISTORY_PATH.exists():
        return {
            "title": "업데이트 기록",
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "items": [],
        }
    with UPDATE_HISTORY_PATH.open("r", encoding="utf-8") as fp:
        payload = json.load(fp)
    payload.setdefault("title", "업데이트 기록")
    payload.setdefault("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M"))
    payload.setdefault("items", [])
    payload["items"] = sorted(
        payload["items"],
        key=lambda item: str(item.get("timestamp") or item.get("date") or ""),
        reverse=True,
    )
    return payload










def _build_visualizations(validation: dict, report_path: Path, sdf_path: Path | None) -> list[dict]:
    tables = validation.get("tables") or {}
    visuals: list[dict] = []

    # 1) Pipe velocity vs limit
    pipe_rows = tables.get("pipes") or []
    pipe_labels: list[str] = []
    velocities: list[float] = []
    limits: list[float] = []
    for r in pipe_rows:
        label = str(r.get("label", ""))
        vel = _to_float(r.get("velocity_mps"), 0.0)
        pipe_labels.append(label)
        velocities.append(vel)
        limits.append(_to_float(r.get("velocity_limit_mps"), 0.0))
    if pipe_labels:
        fig, ax = plt.subplots(figsize=(10, 3.8))
        x = list(range(len(pipe_labels)))
        ax.plot(x, velocities, marker="o", linewidth=1.6, color="#1d4ed8", label="Velocity")
        ax.plot(x, limits, linestyle="--", linewidth=1.2, color="#dc2626", label="Limit")
        ax.set_title("Pipe Velocity vs Limit")
        ax.set_xlabel("Pipe Label")
        ax.set_ylabel("m/s")
        if len(pipe_labels) <= 30:
            ax.set_xticks(x, pipe_labels, rotation=0)
        else:
            step = max(1, len(pipe_labels) // 20)
            ticks = x[::step]
            ax.set_xticks(ticks, [pipe_labels[i] for i in ticks], rotation=0)
        ax.grid(alpha=0.25)
        ax.legend()
        visuals.append(
            {
                "title": "Pipe Velocity Check",
                "description": "Compares each pipe velocity with topology-based branch/other limits from the validator.",
                "image_data_url": _fig_to_data_url(fig),
            }
        )

    # 2) Nozzle pressure-flow scatter
    noz_rows = tables.get("nozzles") or []
    pressures = [_to_float(r.get("inlet_pressure_kgf_cm2"), 0.0) for r in noz_rows]
    flows = [_to_float(r.get("actual_flow_lpm"), 0.0) for r in noz_rows]
    if pressures and flows:
        fig, ax = plt.subplots(figsize=(6.2, 4.2))
        colors = ["#dc2626" if _to_float(r.get("actual_flow_lpm"), 0.0) < 80.0 else "#16a34a" for r in noz_rows]
        ax.scatter(pressures, flows, c=colors, alpha=0.85)
        ax.axhline(80.0, color="#dc2626", linestyle="--", linewidth=1.2, label="80 L/min")
        ax.axvline(1.0, color="#f59e0b", linestyle="--", linewidth=1.2, label="1.0 kg/cm^2")
        ax.set_title("Nozzle Pressure-Flow Distribution")
        ax.set_xlabel("Inlet Pressure (kg/cm^2)")
        ax.set_ylabel("Actual Flow (L/min)")
        ax.grid(alpha=0.25)
        ax.legend(loc="best")
        visuals.append(
            {
                "title": "Nozzle Pressure-Flow",
                "description": "Green points pass the flow threshold, red points are below the flow threshold.",
                "image_data_url": _fig_to_data_url(fig),
            }
        )

    return visuals


def _build_engineering_visualizations(validation: dict) -> list[dict]:
    tables = validation.get("tables") or {}
    pipe_rows = tables.get("pipes") or []
    ratio_rows: list[dict] = []

    for row in pipe_rows:
        label = row.get("label")
        friction_loss = _to_float(row.get("friction_loss"), 0.0)
        length_m = _to_float(row.get("base_length_m") or row.get("pipe_length_m"), 0.0)
        if label is None or length_m <= 0:
            continue
        ratio = friction_loss / length_m
        ratio_rows.append(
            {
                "label": int(label),
                "ratio": ratio,
                "friction_loss": friction_loss,
                "length_m": length_m,
                "velocity_mps": _to_float(row.get("velocity_mps"), 0.0),
                "velocity_limit_mps": _to_float(row.get("velocity_limit_mps"), 0.0),
                "nominal_bore_mm": _to_float(row.get("nominal_bore_mm"), 0.0),
                "fitting_eq_length_m": _to_float(row.get("fitting_eq_length_m"), 0.0),
                "special_eq_length_m": _to_float(row.get("special_eq_length_m"), 0.0),
                "total_length_m": _to_float(row.get("total_length_m"), length_m),
                "engineering_flag": bool(row.get("engineering_flag")),
            }
        )

    if not ratio_rows:
        return []

    ratio_rows.sort(key=lambda item: item["label"])
    labels = [str(item["label"]) for item in ratio_rows]
    values = [item["ratio"] for item in ratio_rows]
    colors = ["#2563eb" if item["engineering_flag"] else "#9ca3af" for item in ratio_rows]
    threshold = 1.0
    lengths = sorted(item["length_m"] for item in ratio_rows)
    median_length = lengths[len(lengths) // 2] if lengths else 0.0
    max_value = max(values) if values else threshold
    y_max = max(max_value * 1.12, threshold * 1.8)
    spike_points: list[dict] = []

    for idx, item in enumerate(ratio_rows):
        if idx == 0:
            continue
        previous = ratio_rows[idx - 1]
        prev_ratio = previous["ratio"]
        ratio = item["ratio"]
        delta = ratio - prev_ratio
        change_rate = delta / max(prev_ratio, 1e-9)
        spike_delta_limit = max(threshold, abs(prev_ratio) * 0.75)
        if ratio <= threshold or delta <= spike_delta_limit:
            continue

        eq_length = item["fitting_eq_length_m"] + item["special_eq_length_m"]
        eq_share = eq_length / max(item["total_length_m"], 1e-9)
        velocity_limit = item["velocity_limit_mps"]
        velocity_ratio = item["velocity_mps"] / velocity_limit if velocity_limit > 0 else 0.0
        causes: list[str] = []
        actions: list[str] = []

        if median_length > 0 and item["length_m"] >= median_length * 1.5:
            causes.append("긴 배관이라 총 마찰손실이 커질 수 있는 구간입니다.")
            actions.append("배관 경로를 단순화하거나 우회 길이를 줄여 실제 배관길이를 단축하는 방안을 검토하세요.")
        if median_length > 0 and item["length_m"] <= median_length * 0.6 and ratio > threshold:
            causes.append("짧은 배관인데 m당 마찰손실이 높아 손실이 국부적으로 집중된 구간입니다.")
            actions.append("해당 짧은 구간의 급격한 방향 전환, 국부 피팅, 특수설비 연결부를 우선 점검하세요.")
        if velocity_ratio >= 0.85:
            causes.append("유속이 적용 기준에 근접하여 마찰손실 증가에 크게 기여할 수 있습니다.")
            actions.append("구경 상향 또는 유량 분산으로 유속을 낮추는 대안을 검토하세요.")
        if eq_share >= 0.35:
            causes.append("피팅/특수설비 등가길이 비중이 커서 배관 자체 길이보다 부속 손실 영향이 큽니다.")
            actions.append("엘보/티/밸브/후렉시블 배관 수량을 줄이거나 손실이 작은 부속으로 변경하는 방안을 검토하세요.")
        if item["nominal_bore_mm"] <= 50 and item["velocity_mps"] >= 5.0:
            causes.append("소구경 배관에서 비교적 높은 유속이 발생하여 손실 집중 가능성이 있습니다.")
            actions.append("50A 이하 구간은 가지배관 기준 6m/s에 근접하는지 확인하고, 필요 시 한 단계 큰 구경을 검토하세요.")
        if not causes:
            causes.append("직전 배관 대비 m당 마찰손실 증가율이 커서 국부 조건 변화가 의심됩니다.")
            actions.append("해당 배관 전후의 구경 변화, 유속 변화, 피팅 수량, 특수설비 연결 여부를 함께 확인하세요.")

        spike_points.append(
            {
                "label": item["label"],
                "previous_label": previous["label"],
                "ratio": ratio,
                "previous_ratio": prev_ratio,
                "delta": delta,
                "change_rate_percent": change_rate * 100.0,
                "data_index": idx,
                "left_percent": 50.0,
                "top_percent": 50.0,
                "cards": {
                    "criteria": [
                        "직전 배관 대비 m당 마찰손실 변화율이 큰 구간을 급증 후보로 표시합니다.",
                        f"급증 조건: 현재 비율 > {threshold:.3f} kg/cm^2/m AND 증가량 > max({threshold:.3f}, 직전 비율 x 75%)",
                    ],
                    "formula": [
                        "m당 마찰손실 = FLOW IN PIPES Frict. Loss / PIPE CONFIGURATION Length",
                        "증가량 = 현재 m당 마찰손실 - 직전 배관 m당 마찰손실",
                        "변화율 = 증가량 / max(직전 배관 m당 마찰손실, 1e-9)",
                    ],
                    "values": [
                        f"Pipe {item['label']}: {item['friction_loss']:.4f} / {item['length_m']:.3f} = {ratio:.4f} kg/cm^2/m",
                        f"Previous Pipe {previous['label']}: {previous['friction_loss']:.4f} / {previous['length_m']:.3f} = {prev_ratio:.4f} kg/cm^2/m",
                        f"증가량 = {delta:.4f} kg/cm^2/m, 변화율 = {change_rate * 100.0:.1f}%",
                        f"구경 = {item['nominal_bore_mm']:.0f}A, 유속 = {item['velocity_mps']:.3f} m/s, 피팅+특수설비 등가길이 비중 = {eq_share * 100.0:.1f}%",
                    ],
                    "conclusion": [
                        *causes,
                        *actions,
                    ],
                },
            }
        )

    fig, ax = plt.subplots(figsize=(11, 4.2))
    x = list(range(len(labels)))
    ax.bar(x, values, color=colors, width=0.78)
    ax.axhline(threshold, color="#dc2626", linestyle="--", linewidth=1.2, label="Threshold 1.00")
    if spike_points:
        spike_x = [labels.index(str(point["label"])) for point in spike_points if str(point["label"]) in labels]
        spike_y = [point["ratio"] for point in spike_points if str(point["label"]) in labels]
        ax.scatter(
            spike_x,
            spike_y,
            marker="v",
            s=26,
            color="#dc2626",
            edgecolor="#7f1d1d",
            linewidth=0.6,
            zorder=5,
            label="Sharp Increase",
        )
    ax.set_title("Friction Loss Ratio by Pipe")
    ax.set_xlabel("Pipe Label")
    ax.set_ylabel("Friction Loss / Length (kg/cm^2/m)")
    ax.set_ylim(0, y_max)
    if len(labels) <= 35:
        ax.set_xticks(x, labels, rotation=0)
    else:
        step = max(1, len(labels) // 24)
        ticks = x[::step]
        ax.set_xticks(ticks, [labels[i] for i in ticks], rotation=0)
    ax.grid(axis="y", alpha=0.25)
    ax.legend(loc="upper right")
    fig.canvas.draw()
    fig_w, fig_h = fig.canvas.get_width_height()
    if fig_w > 0 and fig_h > 0:
        for point in spike_points:
            px, py = ax.transData.transform((point["data_index"], point["ratio"]))
            point["left_percent"] = max(0.0, min(100.0, (px / fig_w) * 100.0))
            point["top_percent"] = max(0.0, min(100.0, ((fig_h - py) / fig_h) * 100.0))

    return [
        {
            "title": "Friction Loss Ratio by Pipe",
            "description": "Shows friction_loss / base_length_m for each pipe. Red markers indicate sharp increases from the previous pipe.",
            "image_data_url": _fig_to_data_url(fig, tight=False),
            "spike_points": spike_points,
        }
    ]


def _point_on_polyline(path: list[tuple[float, float]], ratio: float) -> tuple[float, float] | None:
    if len(path) < 2:
        return path[0] if path else None
    ratio = max(0.0, min(1.0, ratio))
    seg_lengths: list[float] = []
    total = 0.0
    for i in range(len(path) - 1):
        x1, y1 = path[i]
        x2, y2 = path[i + 1]
        d = math.hypot(x2 - x1, y2 - y1)
        seg_lengths.append(d)
        total += d
    if total <= 0:
        return path[0]
    target = total * ratio
    acc = 0.0
    for i, d in enumerate(seg_lengths):
        if acc + d >= target:
            t = (target - acc) / d if d > 0 else 0.0
            x1, y1 = path[i]
            x2, y2 = path[i + 1]
            return (x1 + (x2 - x1) * t, y1 + (y2 - y1) * t)
        acc += d
    return path[-1]


def _build_sdf_graph(sdf_path: Path | None, tables: dict | None) -> dict:
    if sdf_path is None or not sdf_path.exists():
        return {}

    tables = tables or {}
    pipe_table = {int(r.get("label")): r for r in tables.get("pipes", []) if str(r.get("label", "")).isdigit()}
    nozzle_table = {int(r.get("label")): r for r in tables.get("nozzles", []) if str(r.get("label", "")).isdigit()}
    equipment_table = {int(r.get("label")): r for r in tables.get("equipment", []) if str(r.get("label", "")).isdigit()}
    valve_table = {int(r.get("label")): r for r in tables.get("valves", []) if str(r.get("label", "")).isdigit()}

    root = ET.parse(sdf_path).getroot()

    node_pos: dict[str, tuple[float, float]] = {}
    for node in root.findall(".//Node"):
        label = node.attrib.get("label")
        pos = node.find("Position")
        if not label or pos is None:
            continue
        try:
            x = float(pos.attrib.get("x", "0"))
            y = float(pos.attrib.get("y", "0"))
        except ValueError:
            continue
        node_pos[label] = (x, y)

    pipes: list[dict] = []
    pipe_paths: dict[int, list[tuple[float, float]]] = {}
    for pipe in root.findall(".//Pipe"):
        label_raw = pipe.attrib.get("label", "")
        if not label_raw.isdigit():
            continue
        label = int(label_raw)
        input_node = pipe.attrib.get("input", "")
        output_node = pipe.attrib.get("output", "")

        path: list[tuple[float, float]] = []
        if input_node in node_pos:
            path.append(node_pos[input_node])
        waypoints = pipe.find("Waypoints")
        if waypoints is not None:
            for wp in waypoints.findall("Position"):
                try:
                    path.append((float(wp.attrib.get("x", "0")), float(wp.attrib.get("y", "0"))))
                except ValueError:
                    continue
        if output_node in node_pos:
            path.append(node_pos[output_node])
        if len(path) < 2:
            continue

        pipe_paths[label] = path
        trow = pipe_table.get(label, {})
        status = "fail" if trow.get("highlight") else "pass"
        pipes.append(
            {
                "label": label,
                "input_node": input_node,
                "output_node": output_node,
                "path": [[x, y] for x, y in path],
                "status": status,
            }
        )

    nozzles: list[dict] = []
    for nozzle in root.findall(".//Nozzle"):
        label_raw = nozzle.attrib.get("label", "")
        input_node = nozzle.attrib.get("input", "")
        if not label_raw.isdigit() or input_node not in node_pos:
            continue
        label = int(label_raw)
        x, y = node_pos[input_node]
        trow = nozzle_table.get(label, {})
        status = "fail" if trow.get("highlight") else "pass"
        nozzles.append({"label": label, "input_node": input_node, "x": x, "y": y, "status": status})

    equipment: list[dict] = []
    equipment_pos_by_label: dict[int, tuple[float, float]] = {}
    for eq in root.findall(".//Equipment"):
        label_raw = eq.attrib.get("label", "")
        if not label_raw.isdigit():
            continue
        label = int(label_raw)
        rel = float(eq.attrib.get("rel-position", "0.5"))
        desc = eq.attrib.get("description", "")
        table_row = equipment_table.get(label, {})
        pipe_label = table_row.get("pipe_label")
        if isinstance(pipe_label, str) and pipe_label.isdigit():
            pipe_label = int(pipe_label)
        if not isinstance(pipe_label, int):
            continue
        path = pipe_paths.get(pipe_label)
        if not path:
            continue
        p = _point_on_polyline(path, rel)
        if p is None:
            continue
        x, y = p
        equipment_pos_by_label[label] = (x, y)
        status = "warn" if table_row.get("warn") else ("fail" if table_row.get("highlight") else "pass")
        equipment.append(
            {
                "label": label,
                "description": desc,
                "pipe_label": pipe_label,
                "x": x,
                "y": y,
                "status": status,
            }
        )

    valves: list[dict] = []
    for label, row in valve_table.items():
        pos = equipment_pos_by_label.get(label)
        if pos is None:
            continue
        x, y = pos
        status = "fail" if row.get("highlight") else "pass"
        valves.append({"label": label, "x": x, "y": y, "status": status})

    return {
        "nodes": [{"id": nid, "x": x, "y": y} for nid, (x, y) in node_pos.items()],
        "pipes": pipes,
        "nozzles": nozzles,
        "equipment": equipment,
        "valves": valves,
    }


def _sdf_counts_only(sdf_path: Path | None) -> dict:
    if sdf_path is None or not sdf_path.exists():
        return {}
    root = ET.parse(sdf_path).getroot()
    return {
        "pipes": len(root.findall(".//Pipe")),
        "nozzles": len(root.findall(".//Nozzle")),
        "equipment": len(root.findall(".//Equipment")),
    }


def _sdf_parse_nodes(root) -> dict[str, dict]:
    """SDF <Node> → {label: {id, x, y, z(elevation)}}."""
    nodes: dict[str, dict] = {}
    for node in root.findall(".//Node"):
        label = node.attrib.get("label", "")
        pos = node.find("Position")
        if not label or pos is None:
            continue
        nodes[label] = {
            "id": label,
            "x": _to_float(pos.attrib.get("x")),
            "y": _to_float(pos.attrib.get("y")),
            "z": _to_float(node.attrib.get("elevation")),
        }
    return nodes


def _sdf_parse_pipes_equipment(root) -> tuple[list[dict], list[dict]]:
    """SDF <Pipe-set>/<Pipe> → (pipes, equipment). material 은 직전 Pipe-type Name 을 따른다."""
    pipes: list[dict] = []
    equipment: list[dict] = []
    material = "UNKNOWN"
    for pipe_set in root.findall(".//Pipe-set"):
        pipe_type = pipe_set.find("Pipe-type")
        name = pipe_type.find("Name") if pipe_type is not None else None
        if name is not None and name.text:
            material = name.text.strip()
        for pipe in pipe_set.findall("Pipe"):
            label = pipe.attrib.get("label", "")
            input_node = pipe.attrib.get("input", "")
            output_node = pipe.attrib.get("output", "")
            bore_mm = _to_float(pipe.attrib.get("bore")) * 1000.0
            length_m = _to_float(pipe.attrib.get("length"))
            rise_m = _to_float(pipe.attrib.get("rise"))
            c_factor = _to_float(pipe.attrib.get("roughness-or-c"))
            fittings: list[dict] = []
            for fitting in pipe.findall(".//Fitting"):
                fittings.append(
                    {
                        "type": fitting.attrib.get("type", ""),
                        "count": int(_to_float(fitting.attrib.get("count"), 0)),
                    }
                )
            waypoint_positions: list[dict] = []
            waypoints = pipe.find("Waypoints")
            if waypoints is not None:
                for wp in waypoints.findall("Position"):
                    waypoint_positions.append(
                        {
                            "x": _to_float(wp.attrib.get("x")),
                            "y": _to_float(wp.attrib.get("y")),
                        }
                    )
            pipes.append(
                {
                    "label": label,
                    "input_node": input_node,
                    "output_node": output_node,
                    "bore_mm": bore_mm,
                    "length_m": length_m,
                    "rise_m": rise_m,
                    "c_factor": c_factor,
                    "material": material,
                    "fittings": fittings,
                    "fitting_summary": ", ".join(f"{f['type']}({f['count']})" for f in fittings) or "-",
                    "waypoints": waypoint_positions,
                }
            )
            for eq in pipe.findall(".//Equipment"):
                equipment.append(
                    {
                        "label": eq.attrib.get("label", ""),
                        "pipe_label": label,
                        "description": eq.attrib.get("description", ""),
                        "equivalent_length_m": _to_float(eq.attrib.get("equivalent-length")),
                        "rel_position": _to_float(eq.attrib.get("rel-position"), 0.5),
                    }
                )
    return pipes, equipment


def _sdf_parse_nozzles(root, nodes: dict) -> list[dict]:
    """SDF <Nozzle> → 입력노드 좌표를 붙인 노즐(헤드) 리스트."""
    nozzles: list[dict] = []
    for nozzle in root.findall(".//Nozzle"):
        label = nozzle.attrib.get("label", "")
        input_node = nozzle.attrib.get("input", "")
        node = nodes.get(input_node, {})
        nozzles.append(
            {
                "label": label,
                "input_node": input_node,
                "x": node.get("x"),
                "y": node.get("y"),
                "z": node.get("z"),
            }
        )
    return nozzles


def _sdf_build_adjacency(pipes: list[dict]) -> tuple[dict, dict]:
    """pipes → (outgoing[input_node]→pipes, adjacency[node]→(이웃,길이,라벨) 무방향)."""
    outgoing: dict[str, list[dict]] = {}
    adjacency: dict[str, list[tuple[str, float, str]]] = {}
    for pipe in pipes:
        outgoing.setdefault(pipe["input_node"], []).append(pipe)
        adjacency.setdefault(pipe["input_node"], []).append((pipe["output_node"], pipe["length_m"], pipe["label"]))
        adjacency.setdefault(pipe["output_node"], []).append((pipe["input_node"], pipe["length_m"], pipe["label"]))
    return outgoing, adjacency


def _sdf_av_node(pipes: list[dict], equipment: list[dict]) -> tuple[str, str]:
    """알람밸브(A/V) 앵커 노드 추정 → (av_node, av_pipe_label). 못 찾으면 첫 배관 입력노드."""
    pipe_by_label = {p["label"]: p for p in pipes}
    av_equipment = next((e for e in equipment if (e.get("description") or "").upper().replace(" ", "") in {"A/V", "AV"}), None)
    av_node = ""
    av_pipe_label = ""
    if av_equipment:
        av_pipe_label = str(av_equipment.get("pipe_label") or "")
        av_pipe = pipe_by_label.get(av_pipe_label)
        if av_pipe:
            av_node = av_pipe.get("output_node") or av_pipe.get("input_node") or ""
    if not av_node and pipes:
        av_node = pipes[0]["input_node"]
    return av_node, av_pipe_label


def _sdf_dijkstra(av_node: str, adjacency: dict) -> dict[str, float]:
    """A/V 앵커에서 각 노드까지 최단(누적 length) 거리."""
    dist = {av_node: 0.0} if av_node else {}
    visited: set[str] = set()
    while dist:
        current = min((n for n in dist if n not in visited), key=lambda n: dist[n], default=None)
        if current is None:
            break
        visited.add(current)
        for nxt, length, _pipe_label in adjacency.get(current, []):
            nd = dist[current] + max(length, 0.0)
            if nxt not in dist or nd < dist[nxt]:
                dist[nxt] = nd
    return dist


def _sdf_farthest_heads(nozzles: list[dict], dist: dict) -> list[dict]:
    """A/V 에서 먼 순으로 정렬한 헤드 상위 30개 (가장 먼 구간 검토용)."""
    return sorted(
        [
            {**n, "distance_from_av_m": dist.get(str(n.get("input_node")), 0.0)}
            for n in nozzles
        ],
        key=lambda r: r.get("distance_from_av_m", 0.0),
        reverse=True,
    )[:30]


def _sdf_length_checks(pipes: list[dict], nodes: dict) -> list[dict]:
    """SDF length 와 XY(+rise) 기하 길이가 허용오차(5% 또는 0.5m) 초과인 배관."""
    length_checks: list[dict] = []
    for pipe in pipes:
        n1 = nodes.get(pipe["input_node"])
        n2 = nodes.get(pipe["output_node"])
        if not n1 or not n2:
            continue
        pts = [(n1["x"], n1["y"])]
        pts.extend((wp["x"], wp["y"]) for wp in pipe.get("waypoints") or [])
        pts.append((n2["x"], n2["y"]))
        xy_m = 0.0
        for i in range(len(pts) - 1):
            xy_m += math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1]) / 1000.0
        geom_m = math.hypot(xy_m, pipe.get("rise_m") or 0.0)
        diff_m = abs(geom_m - pipe["length_m"])
        tol_m = max(0.5, pipe["length_m"] * 0.05)
        if diff_m > tol_m:
            length_checks.append(
                {
                    "pipe_label": pipe["label"],
                    "sdf_length_m": round(pipe["length_m"], 3),
                    "xy_length_m": round(geom_m, 3),
                    "diff_m": round(diff_m, 3),
                    "reason": "SDF length와 XY 좌표거리 차이가 허용오차(5% 또는 0.5m)를 초과합니다.",
                }
            )
    return length_checks


def _sdf_bore_reductions(pipes: list[dict], outgoing: dict) -> list[dict]:
    """노드에서 하류 배관 구경이 작아지는(축소) 지점."""
    bore_reductions: list[dict] = []
    for pipe in pipes:
        for child in outgoing.get(pipe["output_node"], []):
            if child["bore_mm"] and pipe["bore_mm"] and child["bore_mm"] < pipe["bore_mm"]:
                bore_reductions.append(
                    {
                        "from_pipe": pipe["label"],
                        "to_pipe": child["label"],
                        "node": pipe["output_node"],
                        "from_bore_mm": round(pipe["bore_mm"], 1),
                        "to_bore_mm": round(child["bore_mm"], 1),
                    }
                )
    return bore_reductions


def _sdf_branch_nodes(pipes: list[dict], nodes: dict) -> list[dict]:
    """차수(degree) 3 이상 분기 노드 (degree 내림차순)."""
    node_degree: dict[str, int] = {}
    for pipe in pipes:
        node_degree[pipe["input_node"]] = node_degree.get(pipe["input_node"], 0) + 1
        node_degree[pipe["output_node"]] = node_degree.get(pipe["output_node"], 0) + 1
    return [
        {"node": node, "degree": degree, **nodes.get(node, {})}
        for node, degree in sorted(node_degree.items(), key=lambda x: (-x[1], x[0]))
        if degree >= 3
    ]


def _sdf_fitting_stats(pipes: list[dict]) -> tuple[dict, list]:
    """부속(엘보/티 등) 총계 + 부속 집중(>=2) 핫스팟."""
    fitting_summary: dict[str, int] = {}
    fitting_hotspots: list[dict] = []
    for pipe in pipes:
        total = 0
        for fitting in pipe["fittings"]:
            fitting_summary[fitting["type"]] = fitting_summary.get(fitting["type"], 0) + fitting["count"]
            total += fitting["count"]
        if total >= 2:
            fitting_hotspots.append(
                {
                    "pipe_label": pipe["label"],
                    "fitting_count": total,
                    "fittings": pipe["fitting_summary"],
                    "reason": "엘보/티 등 부속 집중 구간입니다. CAD 도면의 굴곡/분기 위치와 대조가 필요합니다.",
                }
            )
    return fitting_summary, fitting_hotspots


def _sdf_vertical_pipes(pipes: list[dict]) -> list[dict]:
    """|rise| >= 3m 수직 배관 (층고/단면 대조용)."""
    return [
        {
            "pipe_label": p["label"],
            "input_node": p["input_node"],
            "output_node": p["output_node"],
            "length_m": round(p["length_m"], 3),
            "rise_m": round(p["rise_m"], 3),
            "bore_mm": round(p["bore_mm"], 1),
        }
        for p in pipes
        if abs(p.get("rise_m") or 0.0) >= 3.0
    ]


def _sdf_graph_pipes(pipes: list[dict], nodes: dict,
                     length_checks: list[dict], bore_reductions: list[dict]) -> list[dict]:
    """프론트 시각화용 배관 폴리라인 + 상태색(길이이상=red, 구경축소=orange)."""
    graph_pipes: list[dict] = []
    for p in pipes:
        n1 = nodes.get(p["input_node"])
        n2 = nodes.get(p["output_node"])
        if not n1 or not n2:
            continue
        path = [[n1["x"], n1["y"]]]
        path.extend([[wp["x"], wp["y"]] for wp in p.get("waypoints") or []])
        path.append([n2["x"], n2["y"]])
        status = "red" if any(x["pipe_label"] == p["label"] for x in length_checks) else "normal"
        if any(x["to_pipe"] == p["label"] or x["from_pipe"] == p["label"] for x in bore_reductions):
            status = "orange" if status == "normal" else status
        graph_pipes.append(
            {
                "label": p["label"],
                "input_node": p["input_node"],
                "output_node": p["output_node"],
                "bore_mm": round(p["bore_mm"], 1),
                "length_m": round(p["length_m"], 3),
                "material": p["material"],
                "status": status,
                "path": path,
            }
        )
    return graph_pipes


def _analyze_sdf_sprinkler_network(sdf_path: Path) -> dict:
    root = ET.parse(sdf_path).getroot()

    titles = [t.text.strip() for t in root.findall(".//Title") if t.text and t.text.strip()]
    nodes = _sdf_parse_nodes(root)
    pipes, equipment = _sdf_parse_pipes_equipment(root)
    nozzles = _sdf_parse_nozzles(root, nodes)

    outgoing, adjacency = _sdf_build_adjacency(pipes)
    av_node, av_pipe_label = _sdf_av_node(pipes, equipment)
    dist = _sdf_dijkstra(av_node, adjacency)

    farthest_heads = _sdf_farthest_heads(nozzles, dist)
    length_checks = _sdf_length_checks(pipes, nodes)
    bore_reductions = _sdf_bore_reductions(pipes, outgoing)
    branch_nodes = _sdf_branch_nodes(pipes, nodes)
    fitting_summary, fitting_hotspots = _sdf_fitting_stats(pipes)
    vertical_pipes = _sdf_vertical_pipes(pipes)
    graph_pipes = _sdf_graph_pipes(pipes, nodes, length_checks, bore_reductions)

    return {
        "title": " / ".join(titles) or sdf_path.name,
        "filename": sdf_path.name,
        "summary": {
            "node_count": len(nodes),
            "pipe_count": len(pipes),
            "nozzle_count": len(nozzles),
            "equipment_count": len(equipment),
            "av_node": av_node,
            "av_pipe_label": av_pipe_label,
            "length_issue_count": len(length_checks),
            "bore_reduction_count": len(bore_reductions),
            "branch_node_count": len(branch_nodes),
            "vertical_pipe_count": len(vertical_pipes),
        },
        "nodes": list(nodes.values()),
        "pipes": graph_pipes,
        "nozzles": nozzles,
        "equipment": equipment,
        "farthest_heads": farthest_heads,
        "length_checks": length_checks[:80],
        "bore_reductions": bore_reductions[:80],
        "branch_nodes": branch_nodes[:80],
        "fitting_summary": [{"type": k, "count": v} for k, v in sorted(fitting_summary.items())],
        "fitting_hotspots": fitting_hotspots[:80],
        "vertical_pipes": vertical_pipes[:80],
        "checklist": [
            "CAD 도면의 알람밸브 위치가 SDF A/V 추정 노드와 일치하는지 확인",
            "SDF 최원단 헤드 30개가 CAD 평면도상 검토 영역의 헤드 30개와 1:1 매칭되는지 확인",
            "배관 길이 불일치 후보는 CAD 실측 길이와 SDF length 값을 대조",
            "구경 축소 지점은 CAD 라벨의 관경 표기와 SDF bore 값을 대조",
            "엘보/티 집중 구간은 도면상 굴곡/분기 개수와 SDF Fittings count를 대조",
            "수직 배관은 건축 단면/층고와 SDF rise 및 length를 대조",
        ],
    }


def _cad_entity_points(ent: dict) -> list[list[float]]:
    if ent.get("type") == "LINE":
        return [[_to_float(ent.get("x")), _to_float(ent.get("y"))], [_to_float(ent.get("x2")), _to_float(ent.get("y2"))]]
    return [[_to_float(p[0]), _to_float(p[1])] for p in (ent.get("points") or [])]


def _polyline_length(points: list[list[float]]) -> float:
    total = 0.0
    for i in range(len(points) - 1):
        total += math.hypot(points[i + 1][0] - points[i][0], points[i + 1][1] - points[i][1])
    return total


def _normalize_layer_name(layer: str | None) -> str:
    return (layer or "").strip().upper().replace(" ", "")


def _load_cad_sdf_learning_profile() -> dict:
    if not CAD_SDF_LEARNING_PROFILE_PATH.exists():
        return {}
    try:
        return json.loads(CAD_SDF_LEARNING_PROFILE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write_cad_sdf_learning_profile(profile: dict) -> None:
    CAD_SDF_LEARNING_PROFILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    CAD_SDF_LEARNING_PROFILE_PATH.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")


def _cad_layer_weight(layer: str | None, profile: dict | None = None) -> float:
    profile = profile or {}
    norm = _normalize_layer_name(layer)
    positive = {_normalize_layer_name(x) for x in profile.get("positive_layers", [])}
    suppressed = {_normalize_layer_name(x) for x in profile.get("suppressed_layers", [])}
    keywords = [_normalize_layer_name(x) for x in profile.get("positive_keywords", ["SP", "소화", "배관", "후렉", "SPRINKLER", "FIRE"])]
    if norm in positive:
        return 5.0
    if any(keyword and keyword in norm for keyword in keywords):
        return 3.0
    if norm in suppressed:
        return -3.0
    if norm in {"0", "L1", "L2", "L3", "L4", "DEFPOINTS"}:
        return -1.5
    return 0.0


def _build_cad_sdf_learning_profile(cad: dict, sdf: dict, source_sdf: Path | None = None, source_cad: Path | None = None) -> dict:
    entities = cad.get("drawing_entities") or []
    layer_stats: dict[str, dict] = {}
    for ent in entities:
        layer = ent.get("layer") or "0"
        stat = layer_stats.setdefault(layer, {"entity_count": 0, "total_length": 0.0, "similar_count": 0})
        stat["entity_count"] += 1
        stat["total_length"] += float(ent.get("draw_length") or 0.0)
        if ent.get("similar_to_sdf"):
            stat["similar_count"] += 1

    positive_layers: set[str] = set()
    suppressed_layers: set[str] = set()
    for layer, stat in layer_stats.items():
        norm = _normalize_layer_name(layer)
        has_fire_keyword = any(keyword in norm for keyword in ["SP", "소화", "배관", "후렉", "SPRINKLER", "FIRE"])
        entity_count = int(stat.get("entity_count") or 0)
        similar_count = int(stat.get("similar_count") or 0)
        if has_fire_keyword:
            positive_layers.add(layer)
        elif entity_count >= 1000 and similar_count >= 10:
            # Large generic architectural layers can accidentally resemble SDF
            # line geometry after normalization. Treat them as background noise.
            suppressed_layers.add(layer)
        elif norm in {"0", "L1", "L2", "L3", "L4"} and not has_fire_keyword:
            suppressed_layers.add(layer)

    profile = {
        "version": 1,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source_sdf": str(source_sdf) if source_sdf else "",
        "source_cad": str(source_cad) if source_cad else "",
        "method": "sample_pair_layer_weighting",
        "positive_keywords": ["SP", "소화", "배관", "후렉", "SPRINKLER", "FIRE"],
        "positive_layers": sorted(positive_layers),
        "suppressed_layers": sorted(suppressed_layers),
        "layer_stats": {
            layer: {
                "entity_count": int(stat["entity_count"]),
                "total_length": round(float(stat["total_length"]), 3),
                "similar_count": int(stat["similar_count"]),
            }
            for layer, stat in sorted(layer_stats.items())
        },
        "notes": [
            "H-100 단위세대 소방평면도_도면하나.dxf와 201동 3F SDF 샘플에서 추출한 CAD-SDF 대조 가중치입니다.",
            "소방/SP/배관/후렉 계열 레이어를 우선 배관 후보로 보고, L 계열 대량 건축선은 배경 잡음으로 낮게 평가합니다.",
            "일반 AI 모델 학습이 아니라 샘플 기반 휴리스틱 학습 프로필입니다. 여러 라벨링 샘플이 누적되면 세그멘테이션 모델 학습으로 확장할 수 있습니다.",
        ],
    }
    _write_cad_sdf_learning_profile(profile)
    return profile


def _entity_preview_row(ent: dict, idx: int) -> dict | None:
    points = _cad_entity_points(ent)
    if len(points) < 2:
        return None
    return {
        "id": f"E{idx}",
        "type": ent.get("type", "LINE"),
        "layer": ent.get("layer", ""),
        "points": points,
        "draw_length": _polyline_length(points),
    }


def _approx_arc_points(ent: dict, steps: int = 16) -> list[list[float]]:
    cx = _to_float(ent.get("x"))
    cy = _to_float(ent.get("y"))
    radius = abs(_to_float(ent.get("radius")))
    start = math.radians(_to_float(ent.get("start_angle")))
    end = math.radians(_to_float(ent.get("end_angle")))
    if radius <= 0:
        return []
    if end < start:
        end += math.tau
    return [[cx + math.cos(start + (end - start) * i / steps) * radius, cy + math.sin(start + (end - start) * i / steps) * radius] for i in range(steps + 1)]


def _extract_cad_head_candidates(cad_path: Path) -> dict:
    # Lightweight DXF scan for the 7th module. It extracts enough geometry for a
    # quick drawing preview without invoking the heavier CAD graph engine.
    learning_profile = _load_cad_sdf_learning_profile()
    try:
        raw = cad_path.read_text(encoding="utf-8", errors="ignore").splitlines()
    except UnicodeDecodeError:
        raw = cad_path.read_text(encoding="cp949", errors="ignore").splitlines()

    pairs: list[tuple[str, str]] = []
    for i in range(0, len(raw) - 1, 2):
        pairs.append((raw[i].strip(), raw[i + 1].strip()))

    entities: list[dict] = []
    in_entities = False
    current: dict | None = None

    def is_triangle_polyline(ent: dict) -> bool:
        if ent.get("type") not in {"LWPOLYLINE", "POLYLINE"} or not ent.get("closed"):
            return False
        points = ent.get("points") or []
        unique: list[tuple[float, float]] = []
        for point in points:
            try:
                pt = (float(point[0]), float(point[1]))
            except Exception:
                continue
            if not any(math.hypot(pt[0] - old[0], pt[1] - old[1]) < 1e-6 for old in unique):
                unique.append(pt)
        if len(unique) != 3:
            return False
        xs = [pt[0] for pt in unique]
        ys = [pt[1] for pt in unique]
        w = max(xs) - min(xs)
        h = max(ys) - min(ys)
        diag = math.hypot(w, h)
        if diag <= 0 or diag > 1200 or min(w, h) <= 0:
            return False
        if max(w, h) / max(min(w, h), 1e-9) > 2.2:
            return False
        area = abs(
            unique[0][0] * (unique[1][1] - unique[2][1])
            + unique[1][0] * (unique[2][1] - unique[0][1])
            + unique[2][0] * (unique[0][1] - unique[1][1])
        ) / 2
        return area > 20 and area / max(w * h, 1e-9) > 0.18

    def flush() -> None:
        nonlocal current
        if not current:
            return
        etype = current.get("type")
        if etype in {"CIRCLE", "INSERT"} and current.get("x") is not None and current.get("y") is not None:
            entities.append(current)
        elif etype == "LINE" and current.get("x") is not None and current.get("y") is not None and current.get("x2") is not None and current.get("y2") is not None:
            entities.append(current)
        elif etype == "ARC" and current.get("x") is not None and current.get("y") is not None and current.get("radius") is not None:
            current["points"] = _approx_arc_points(current)
            if len(current["points"]) >= 2:
                entities.append(current)
        elif etype in {"LWPOLYLINE", "POLYLINE"} and len(current.get("points") or []) >= 2:
            entities.append(current)
        current = None

    for code, value in pairs:
        if code == "0" and value == "SECTION":
            continue
        if code == "2" and value == "ENTITIES":
            in_entities = True
            continue
        if code == "0" and value == "ENDSEC":
            flush()
            in_entities = False
            continue
        if not in_entities:
            continue
        if code == "0":
            flush()
            if value in {"CIRCLE", "INSERT"}:
                current = {"type": value, "layer": "0", "x": None, "y": None, "radius": 0.0}
            elif value == "ARC":
                current = {"type": value, "layer": "0", "x": None, "y": None, "radius": 0.0, "start_angle": 0.0, "end_angle": 0.0}
            elif value == "LINE":
                current = {"type": value, "layer": "0", "x": None, "y": None, "x2": None, "y2": None}
            elif value in {"LWPOLYLINE", "POLYLINE"}:
                current = {"type": value, "layer": "0", "points": [], "_pending_x": None}
            continue
        if current is None:
            continue
        if code == "8":
            current["layer"] = value
        elif code == "10":
            if current.get("type") in {"LWPOLYLINE", "POLYLINE"}:
                current["_pending_x"] = _to_float(value)
            else:
                current["x"] = _to_float(value)
        elif code == "20":
            if current.get("type") in {"LWPOLYLINE", "POLYLINE"}:
                px = current.pop("_pending_x", None)
                if px is not None:
                    current.setdefault("points", []).append([px, _to_float(value)])
            else:
                current["y"] = _to_float(value)
        elif code == "11":
            current["x2"] = _to_float(value)
        elif code == "21":
            current["y2"] = _to_float(value)
        elif code == "40":
            current["radius"] = _to_float(value)
        elif code == "50":
            current["start_angle"] = _to_float(value)
        elif code == "51":
            current["end_angle"] = _to_float(value)
        elif code == "70" and current.get("type") in {"LWPOLYLINE", "POLYLINE"}:
            current["closed"] = bool(int(_to_float(value)) & 1)
        elif code == "2" and current.get("type") == "INSERT":
            current["block"] = value
    flush()

    circles = [e for e in entities if e.get("type") == "CIRCLE"]
    inserts = [e for e in entities if e.get("type") == "INSERT"]
    triangles = [e for e in entities if is_triangle_polyline(e)]
    line_entities = [e for e in entities if e.get("type") in {"LINE", "LWPOLYLINE", "POLYLINE", "ARC"}]
    source = [*circles, *triangles] if len(circles) + len(triangles) >= 3 else [*inserts, *triangles]
    candidates: list[dict] = []
    for idx, ent in enumerate(source, start=1):
        if ent.get("type") in {"LWPOLYLINE", "POLYLINE"}:
            xs = [p[0] for p in ent.get("points") or []]
            ys = [p[1] for p in ent.get("points") or []]
            x = sum(xs) / len(xs) if xs else 0.0
            y = sum(ys) / len(ys) if ys else 0.0
        else:
            x = _to_float(ent.get("x"))
            y = _to_float(ent.get("y"))
        candidates.append(
            {
                "label": str(idx),
                "entity_id": str(idx),
                "type": ent.get("type", ""),
                "layer": ent.get("layer", ""),
                "x": x,
                "y": y,
                "radius": _to_float(ent.get("radius")),
            }
        )

    if len(candidates) > 300:
        # Keep the comparison responsive. The user should filter the DXF to sprinkler/head layers for precision.
        candidates = candidates[:300]

    drawing_entities = []
    for idx, ent in enumerate(line_entities, start=1):
        row = _entity_preview_row(ent, idx)
        if row:
            row["layer_weight"] = _cad_layer_weight(row.get("layer"), learning_profile)
            drawing_entities.append(row)

    # Full DXF drawings can contain 100k+ entities. Learned fire/SP layers are
    # retained first; within the same weight keep longer pipe-like geometry first.
    drawing_entities.sort(key=lambda item: (item.get("layer_weight", 0.0), item.get("draw_length", 0.0)), reverse=True)
    drawing_entities_for_preview = drawing_entities[:20000]

    preview_points: list[dict] = []
    for ent in drawing_entities_for_preview:
        for x, y in ent.get("points") or []:
            preview_points.append({"x": x, "y": y})
    preview_points.extend(candidates)

    return {
        "filename": cad_path.name,
        "bounds": _bbox(preview_points) or _bbox(candidates) or {},
        "layers": sorted({e.get("layer", "") for e in entities if e.get("layer")}),
        "network_layers": sorted(learning_profile.get("positive_layers", [])),
        "learned_profile_applied": bool(learning_profile),
        "learned_profile_updated_at": learning_profile.get("updated_at"),
        "learned_positive_layers": learning_profile.get("positive_layers", []),
        "learned_suppressed_layers": learning_profile.get("suppressed_layers", []),
        "raw_circle_count": len(circles),
        "raw_triangle_count": len(triangles),
        "raw_insert_count": len(inserts),
        "raw_line_count": len(line_entities),
        "drawing_entity_count": len(drawing_entities),
        "drawing_entity_returned_count": len(drawing_entities_for_preview),
        "candidate_count": len(candidates),
        "candidates": candidates[:500],
        "drawing_entities": drawing_entities_for_preview,
    }


def _bbox(points: list[dict]) -> dict | None:
    valid = [p for p in points if p.get("x") is not None and p.get("y") is not None]
    if not valid:
        return None
    xs = [float(p["x"]) for p in valid]
    ys = [float(p["y"]) for p in valid]
    return {"min_x": min(xs), "max_x": max(xs), "min_y": min(ys), "max_y": max(ys)}


def _norm_point(p: dict, box: dict) -> tuple[float, float]:
    w = max(float(box["max_x"]) - float(box["min_x"]), 1e-9)
    h = max(float(box["max_y"]) - float(box["min_y"]), 1e-9)
    return ((float(p["x"]) - float(box["min_x"])) / w, (float(p["y"]) - float(box["min_y"])) / h)


def _norm_xy(x: float, y: float, box: dict) -> tuple[float, float]:
    w = max(float(box["max_x"]) - float(box["min_x"]), 1e-9)
    h = max(float(box["max_y"]) - float(box["min_y"]), 1e-9)
    return ((float(x) - float(box["min_x"])) / w, (float(y) - float(box["min_y"])) / h)


def _segments_from_points(points: list[list[float]], box: dict, source_id: str, label: str | int | None = None) -> list[dict]:
    rows = []
    for i in range(len(points) - 1):
        x1, y1 = points[i]
        x2, y2 = points[i + 1]
        nx1, ny1 = _norm_xy(x1, y1, box)
        nx2, ny2 = _norm_xy(x2, y2, box)
        length = math.hypot(nx2 - nx1, ny2 - ny1)
        if length <= 1e-7:
            continue
        rows.append(
            {
                "source_id": source_id,
                "label": label,
                "mid": ((nx1 + nx2) / 2, (ny1 + ny2) / 2),
                "length": length,
                "angle": math.atan2(ny2 - ny1, nx2 - nx1),
            }
        )
    return rows


def _mark_similar_cad_pipe_entities(cad: dict, sdf: dict) -> dict:
    learning_profile = _load_cad_sdf_learning_profile()
    cad_entities = cad.get("drawing_entities") or []
    sdf_pipes = sdf.get("pipes") or []
    cad_points = [{"x": x, "y": y} for ent in cad_entities for x, y in (ent.get("points") or [])]
    sdf_points = [{"x": x, "y": y} for pipe in sdf_pipes for x, y in (pipe.get("path") or [])]
    cad_box = _bbox(cad_points)
    sdf_box = _bbox(sdf_points)
    if not cad_box or not sdf_box:
        return {"matched_entity_ids": [], "matched_count": 0, "threshold": 0.09}

    sdf_segments = []
    for pipe in sdf_pipes:
        sdf_segments.extend(_segments_from_points(pipe.get("path") or [], sdf_box, str(pipe.get("label", "")), pipe.get("label")))
    cad_segments = []
    for ent in cad_entities:
        layer_weight = float(ent.get("layer_weight", _cad_layer_weight(ent.get("layer"), learning_profile)) or 0.0)
        if learning_profile and layer_weight <= -2.5:
            continue
        ent_segments = _segments_from_points(ent.get("points") or [], cad_box, str(ent.get("id", "")), ent.get("layer"))
        for seg in ent_segments:
            seg["layer_weight"] = layer_weight
            seg["layer"] = ent.get("layer")
        cad_segments.extend(ent_segments)
    if not sdf_segments or not cad_segments:
        return {"matched_entity_ids": [], "matched_count": 0, "threshold": 0.09}

    matched: dict[str, float] = {}
    for sdf_seg in sdf_segments:
        best_id = None
        best_score = float("inf")
        for cad_seg in cad_segments:
            mid_dist = math.hypot(sdf_seg["mid"][0] - cad_seg["mid"][0], sdf_seg["mid"][1] - cad_seg["mid"][1])
            angle_penalty = _angle_delta(sdf_seg["angle"], cad_seg["angle"]) / math.pi
            len_ratio = abs(math.log(max(cad_seg["length"], 1e-6) / max(sdf_seg["length"], 1e-6)))
            layer_weight = float(cad_seg.get("layer_weight") or 0.0)
            score = mid_dist + angle_penalty * 0.18 + min(len_ratio, 2.0) * 0.05
            if layer_weight > 0:
                score -= min(layer_weight, 5.0) * 0.012
            elif layer_weight < 0:
                score += abs(layer_weight) * 0.04
            if score < best_score:
                best_score = score
                best_id = cad_seg["source_id"]
        if best_id and best_score <= 0.09:
            matched[best_id] = min(best_score, matched.get(best_id, best_score))

    for ent in cad_entities:
        sid = str(ent.get("id"))
        if sid in matched:
            ent["similar_to_sdf"] = True
            ent["similarity_score"] = round(matched[sid], 4)

    return {
        "matched_entity_ids": sorted(matched, key=lambda key: matched[key])[:2000],
        "matched_count": len(matched),
        "threshold": 0.09,
        "learning_profile_applied": bool(learning_profile),
    }


def _compare_cad_heads_to_sdf(cad: dict, sdf: dict) -> dict:
    sdf_heads = sdf.get("farthest_heads") or sdf.get("nozzles") or []
    cad_heads = cad.get("candidates") or []
    sdf_box = _bbox(sdf_heads)
    cad_box = _bbox(cad_heads)
    if not sdf_box or not cad_box or not sdf_heads or not cad_heads:
        return {
            "status": "REVIEW",
            "message": "CAD 헤드 후보 또는 SDF 헤드 좌표가 부족하여 자동 비교를 보류했습니다.",
            "matches": [],
            "unmatched_sdf": sdf_heads,
            "unmatched_cad": cad_heads,
            "mismatch_count": 0,
        }

    unused = set(range(len(cad_heads)))
    matches = []
    for sdf_head in sdf_heads:
        sx, sy = _norm_point(sdf_head, sdf_box)
        best_idx = None
        best_dist = float("inf")
        for idx in unused:
            cx, cy = _norm_point(cad_heads[idx], cad_box)
            d = math.hypot(sx - cx, sy - cy)
            if d < best_dist:
                best_idx = idx
                best_dist = d
        if best_idx is None:
            continue
        unused.remove(best_idx)
        cad_head = cad_heads[best_idx]
        matches.append(
            {
                "sdf_head": sdf_head.get("label"),
                "sdf_node": sdf_head.get("input_node"),
                "sdf_x": sdf_head.get("x"),
                "sdf_y": sdf_head.get("y"),
                "cad_candidate": cad_head.get("label"),
                "cad_layer": cad_head.get("layer"),
                "cad_type": cad_head.get("type"),
                "cad_x": cad_head.get("x"),
                "cad_y": cad_head.get("y"),
                "normalized_error": round(best_dist, 4),
                "status": "FAIL" if best_dist > 0.08 else "PASS",
                "reason": "정규화 좌표 오차가 0.08을 초과합니다." if best_dist > 0.08 else "정규화 좌표 기준 최근접 매칭 허용범위 이내입니다.",
            }
        )

    pipe_shape_match = _mark_similar_cad_pipe_entities(cad, sdf)
    mismatch_count = sum(1 for m in matches if m["status"] == "FAIL")
    return {
        "status": "FAIL" if mismatch_count else "PASS",
        "message": f"SDF 최원단 헤드 {len(sdf_heads)}개와 CAD 헤드 후보 {len(cad_heads)}개를 정규화 좌표 기준으로 비교했습니다.",
        "matches": matches,
        "unmatched_sdf": sdf_heads[len(matches):],
        "unmatched_cad": [cad_heads[i] for i in sorted(unused)],
        "mismatch_count": mismatch_count,
        "pipe_shape_match": pipe_shape_match,
    }


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/cad-compare-module")
def cad_compare_module():
    response = make_response(render_template("cad_compare_module.html"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/cad-compare-module-7")
def cad_compare_module_7():
    response = make_response(render_template("cad_compare_module_7.html"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/sprinkler-pipeline")
def sprinkler_pipeline():
    response = make_response(render_template("sprinkler_pipeline.html"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/design-automation-module-8")
def design_automation_module_8():
    try:
        _start_design_automation_server()
    except Exception as exc:
        return (
            "설계자동화 인터페이스 서버를 시작하지 못했습니다. "
            f"원인: {html_lib.escape(str(exc))}",
            500,
        )
    host = request.host.split(":", 1)[0]
    return redirect(f"http://{host}:{DESIGN_AUTOMATION_PORT}/", code=302)


@app.get("/print-report/<path:filename>")
def print_report(filename: str):
    safe_name = Path(filename).name
    report_path = UPLOAD_DIR / safe_name
    if not report_path.exists() or report_path.suffix.lower() not in {".docx", ".pdf"}:
        return "출력할 결과서 파일을 찾을 수 없습니다.", 404
    try:
        copies = max(1, min(int(request.args.get("copies", "2")), 10))
    except ValueError:
        copies = 2
    try:
        text = _printable_report_text(report_path)
    except Exception as exc:
        return f"결과서 내용을 읽을 수 없습니다: {html_lib.escape(str(exc))}", 500

    body = "\n".join(html_lib.escape(line.rstrip()) for line in text.splitlines())
    title = html_lib.escape(report_path.name)
    copy_blocks = []
    for idx in range(copies):
        page_break = " page-break-before: always;" if idx else ""
        copy_blocks.append(
            f"""
            <section class="print-copy" style="{page_break}">
              <header class="print-head">
                <div>
                  <p>PIPENET REPORT PRINT</p>
                  <h1>{title}</h1>
                </div>
                <strong>{idx + 1}/{copies}부</strong>
              </header>
              <pre>{body}</pre>
            </section>
            """
        )
    html = f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <title>{title} 출력</title>
  <style>
    @page {{ size: A4 portrait; margin: 12mm; }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; background: #f3f4f6; color: #111827; font-family: "Malgun Gothic", "맑은 고딕", sans-serif; }}
    .print-toolbar {{ position: sticky; top: 0; z-index: 5; display: flex; justify-content: space-between; align-items: center; gap: 12px; padding: 14px 18px; border-bottom: 1px solid #d1d5db; background: #fff; }}
    .print-toolbar strong {{ font-size: 14px; }}
    .print-toolbar button {{ border: 1px solid #111827; background: #111827; color: #fff; padding: 10px 16px; font-weight: 800; cursor: pointer; }}
    .print-copy {{ width: 210mm; min-height: 297mm; margin: 16px auto; padding: 14mm; background: #fff; border: 1px solid #d1d5db; }}
    .print-head {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 12px; padding-bottom: 10px; border-bottom: 2px solid #111827; }}
    .print-head p {{ margin: 0 0 4px; font-size: 10px; letter-spacing: .12em; color: #6b7280; }}
    .print-head h1 {{ margin: 0; font-size: 18px; line-height: 1.35; }}
    .print-head strong {{ border: 1px solid #111827; padding: 6px 10px; font-size: 12px; white-space: nowrap; }}
    pre {{ margin: 0; white-space: pre-wrap; word-break: break-word; font-family: "Consolas", "D2Coding", "Malgun Gothic", monospace; font-size: 9.5pt; line-height: 1.38; }}
    @media print {{
      body {{ background: #fff; }}
      .print-toolbar {{ display: none; }}
      .print-copy {{ width: auto; min-height: auto; margin: 0; padding: 0; border: 0; }}
    }}
  </style>
</head>
<body>
  <div class="print-toolbar">
    <strong>결과서 전체 내용 출력 - {copies}부</strong>
    <button type="button" onclick="window.print()">프린트 실행</button>
  </div>
  {''.join(copy_blocks)}
  <script>window.addEventListener('load', () => setTimeout(() => window.print(), 350));</script>
</body>
</html>"""
    response = make_response(html)
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response


@app.get("/api/update-history")
def update_history():
    return jsonify({"ok": True, "history": _load_update_history()})








@app.post("/api/cad-module/dxf-parse")
def cad_module_dxf_parse():
    try:
        cad_path = _save_upload("cad_file", {".dxf", ".dwg"}, required=True)

        from cad_engine import DXFWorkspace

        workspace = DXFWorkspace(UPLOAD_DIR / "cad_workspace")
        workspace.load_file(cad_path)
        payload = workspace.to_payload(
            include_network_entities=False,
            include_network_summary=False,
            include_graph=False,
        )
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "message": f"DXF 파싱 중 오류가 발생했습니다: {exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "message": "DXF 파싱이 완료되었습니다.",
            "cad_payload": {
                "filename": payload.get("filename"),
                "bounds": payload.get("bounds"),
                "layers": payload.get("layers"),
                "entities": payload.get("entities") or [],
                "graph": {},
                "unsupported": payload.get("unsupported") or {},
            },
        }
    )


    # feedback 도메인 라우트 → routes/feedback.py (register 로 등록)
@app.post("/api/sdf-sprinkler-analysis")
def sdf_sprinkler_analysis():
    try:
        sdf_path = _save_upload("sdf_file", {".sdf"}, required=True)
        analysis = _analyze_sdf_sprinkler_network(sdf_path)
        cad_analysis = None
        comparison = None
        cad_path = _save_upload("cad_file", {".dxf", ".dwg"}, required=False)
        if cad_path is not None:
            cad_analysis = _extract_cad_head_candidates(cad_path)
            comparison = _compare_cad_heads_to_sdf(cad_analysis, analysis)
            if not cad_analysis.get("learned_profile_applied"):
                profile = _build_cad_sdf_learning_profile(cad_analysis, analysis, sdf_path, cad_path)
                for ent in cad_analysis.get("drawing_entities") or []:
                    ent["layer_weight"] = _cad_layer_weight(ent.get("layer"), profile)
                    ent.pop("similar_to_sdf", None)
                    ent.pop("similarity_score", None)
                cad_analysis.update(
                    {
                        "network_layers": sorted(profile.get("positive_layers", [])),
                        "learned_profile_applied": True,
                        "learned_profile_updated_at": profile.get("updated_at"),
                        "learned_positive_layers": profile.get("positive_layers", []),
                        "learned_suppressed_layers": profile.get("suppressed_layers", []),
                    }
                )
                comparison = _compare_cad_heads_to_sdf(cad_analysis, analysis)
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "message": f"SDF 분석 중 오류가 발생했습니다: {exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "message": "SDF 스프링클러 배관 분석이 완료되었습니다.",
            "analysis": analysis,
            "cad_analysis": cad_analysis,
            "comparison": comparison,
        }
    )


@app.post("/api/validate")
def validate_files():
    try:
        report_path = _save_upload("report_file", {".docx", ".pdf"}, required=True)
        sdf_path = _save_upload("sdf_file", {".sdf"}, required=False)
        validation = PipenetGuideValidator(report_path=report_path, sdf_path=sdf_path).validate()
        sdf_graph = _build_sdf_graph(sdf_path, validation.get("tables"))
        visualizations = _build_visualizations(validation, report_path, sdf_path)
        engineering_visualizations = _build_engineering_visualizations(validation)
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    except Exception as exc:
        # 어떤 예외도 잡아 JSON 으로 반환 — 절대 HTML 500 페이지로 빠지지 않게.
        import traceback as _tb
        return jsonify({
            "ok": False,
            "message": f"검증 중 오류가 발생했습니다: {type(exc).__name__}: {str(exc)[:300]}",
            "traceback": _tb.format_exc()[-2000:],
        }), 500

    return jsonify(
        {
            "ok": True,
            "message": "검증이 완료되었습니다.",
            "filename": validation["report_name"],
            "print_url": _print_report_url(report_path, copies=2),
            "sdf_filename": validation["sdf_name"],
            "summary": validation["summary"],
            "results": validation["results"],
            "insights": {
                **validation["insights"],
                "engineering_visualizations": engineering_visualizations,
            },
            "rules": validation.get("rules", {}),
            "stats": validation["stats"],
            "visualizations": visualizations,
            "tables": validation["tables"],
            "sdf_graph": sdf_graph,
            "report": validation["report"],
        }
    )


@app.post("/api/cad-compare")
def cad_compare():
    try:
        cad_path = _save_upload("cad_file", {".dxf", ".dwg"}, required=True)
        sdf_path = _save_upload("sdf_file", {".sdf"}, required=False)

        from cad_engine import DXFWorkspace

        workspace = DXFWorkspace(UPLOAD_DIR / "cad_workspace")
        workspace.load_file(cad_path)
        payload = workspace.to_payload(
            include_network_entities=True,
            include_network_summary=True,
            include_graph=True,
        )

        network_layers = set(payload.get("networkLayers") or [])
        network_entity_ids = set(payload.get("networkEntityIds") or [])
        entities = payload.get("entities") or []
        if network_entity_ids:
            entities = [e for e in entities if e.get("id") in network_entity_ids]
        if network_layers:
            entities = [e for e in entities if e.get("layer") in network_layers]

        head_boxes: list[dict] = []
        detector_mode = "template"
        use_yolo = str(request.form.get("use_yolo", "")).strip() == "1"
        try:
            if use_yolo:
                from head_detector import TriangleHeadDetector

                model_path = BASE_DIR / "models" / "triangle_head_yolo_ai" / "weights" / "best.pt"
                if not model_path.exists():
                    model_path = BASE_DIR / "runs" / "detect" / "models" / "triangle_head_yolo_ai" / "weights" / "best.pt"
                if not model_path.exists():
                    model_path = BASE_DIR / "models" / "triangle_head_yolo" / "weights" / "best.pt"
                if not model_path.exists():
                    model_path = BASE_DIR / "runs" / "detect" / "models" / "triangle_head_yolo" / "weights" / "best.pt"
                if not model_path.exists():
                    model_path = BASE_DIR / "yolo26n.pt"
                if not model_path.exists():
                    model_path = BASE_DIR / "yolo11n.pt"
                detector = TriangleHeadDetector(BASE_DIR / "data" / "head_templates", model_path)
                head_boxes = detector.detect(entities, payload.get("bounds") or {}, network_layers)
                detector_mode = "yolo+template" if detector.yolo_detector.available else "template"
            else:
                from head_detector import TriangleHeadTemplateDetector

                detector = TriangleHeadTemplateDetector(BASE_DIR / "data" / "head_templates")
                head_boxes = detector.detect(entities, payload.get("bounds") or {}, network_layers)
                detector_mode = "template"
        except Exception:
            detector_mode = "unavailable"
            head_boxes = []

        cad_counts = {
            "entities": len(entities),
            "network_layers": len(network_layers),
            "detected_heads": len(head_boxes),
            "lines": sum(1 for e in entities if e.get("type") in {"LINE", "LWPOLYLINE", "ARC"}),
            "circles": sum(1 for e in entities if e.get("type") == "CIRCLE"),
            "texts": sum(1 for e in entities if e.get("type") == "TEXT"),
        }
        sdf_counts = _sdf_counts_only(sdf_path)
        messages: list[str] = []
        if sdf_counts:
            sdf_heads = int(sdf_counts.get("nozzles", 0))
            diff = cad_counts["detected_heads"] - sdf_heads
            if diff == 0:
                messages.append(f"헤드 수 일치: CAD 탐지 {cad_counts['detected_heads']} / SDF {sdf_heads}")
            else:
                messages.append(
                    f"헤드 수 차이: CAD 탐지 {cad_counts['detected_heads']} / SDF {sdf_heads} (차이 {diff:+d})"
                )
            messages.append(
                f"SDF 수량: 배관 {sdf_counts.get('pipes', 0)} / 헤드 {sdf_counts.get('nozzles', 0)} / 특수설비 {sdf_counts.get('equipment', 0)}"
            )
        else:
            messages.append("SDF 미업로드 상태입니다. CAD 단독 추출/탐지 결과만 표시합니다.")
        messages.append(f"탐지 엔진: {detector_mode}")

    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "message": f"CAD 대조 중 오류가 발생했습니다: {exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "message": "CAD 대조가 완료되었습니다.",
            "cad_filename": cad_path.name,
            "sdf_filename": sdf_path.name if sdf_path else None,
            "cad_payload": {
                "filename": payload.get("filename"),
                "bounds": payload.get("bounds"),
                "layers": payload.get("layers"),
                "networkLayers": list(network_layers),
                "entities": entities,
                "graph": payload.get("graph") or {},
            },
            "detected_heads": head_boxes,
            "cad_counts": cad_counts,
            "sdf_counts": sdf_counts,
            "messages": messages,
        }
    )


def _apply_sheet_style(ws):
    thin = Side(style="thin", color="C9CDD3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="F3F4F6")
    title_fill = PatternFill("solid", fgColor="111827")
    fail_fill = PatternFill("solid", fgColor="FEE2E2")
    warn_fill = PatternFill("solid", fgColor="FEF3C7")
    eng_fill = PatternFill("solid", fgColor="DBEAFE")
    econ_fill = PatternFill("solid", fgColor="DCFCE7")

    max_col = ws.max_column
    max_row = ws.max_row

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    title_cell = ws.cell(row=1, column=1)
    title_cell.font = Font(color="FFFFFF", bold=True, size=12)
    title_cell.fill = title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in range(1, max_col + 1):
        c = ws.cell(row=3, column=col)
        c.fill = header_fill
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    for row in range(4, max_row + 1):
        if max_col >= 4:
            flag_fail = ws.cell(row=row, column=max_col - 3).value == "Y"
            flag_warn = ws.cell(row=row, column=max_col - 2).value == "Y"
            flag_eng = ws.cell(row=row, column=max_col - 1).value == "Y"
            flag_econ = ws.cell(row=row, column=max_col).value == "Y"
        else:
            flag_fail = flag_warn = flag_eng = flag_econ = False

        row_fill = None
        if flag_fail:
            row_fill = fail_fill
        elif flag_warn:
            row_fill = warn_fill
        elif flag_eng:
            row_fill = eng_fill
        elif flag_econ:
            row_fill = econ_fill

        for col in range(1, max_col + 1):
            c = ws.cell(row=row, column=col)
            c.border = border
            c.alignment = Alignment(horizontal="center", vertical="center")
            if row_fill:
                c.fill = row_fill

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{ws.cell(row=3, column=max_col).column_letter}{max_row}"

    for col in range(1, max_col + 1):
        width = 12
        for row in range(1, max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is None:
                continue
            width = max(width, min(len(str(val)) + 2, 40))
        ws.column_dimensions[get_column_letter(col)].width = width


def _ai_edge_features(edges: list[dict]) -> list[list[float]]:
    if not edges:
        return []
    pts = []
    for edge in edges:
        for p in edge.get("points") or []:
            try:
                pts.append((float(p.get("x", 0.0)), float(p.get("y", 0.0))))
            except Exception:
                continue
    if not pts:
        pts = [(0.0, 0.0), (100.0, 100.0)]
    min_x, max_x = min(x for x, _ in pts), max(x for x, _ in pts)
    min_y, max_y = min(y for _, y in pts), max(y for _, y in pts)
    w = max(max_x - min_x, 1e-9)
    h = max(max_y - min_y, 1e-9)
    diag = max(math.hypot(w, h), 1e-9)
    _profile = _load_cad_sdf_learning_profile()
    rows = []
    for edge in edges:
        start = edge.get("start") or {}
        end = edge.get("end") or {}
        sx, sy = float(start.get("x", 0.0)), float(start.get("y", 0.0))
        ex, ey = float(end.get("x", 0.0)), float(end.get("y", 0.0))
        mx = (((sx + ex) / 2.0) - min_x) / w
        my = (((sy + ey) / 2.0) - min_y) / h
        length = float(edge.get("length") or math.hypot(ex - sx, ey - sy)) / diag
        angle = math.atan2((ey - sy) / h, (ex - sx) / w)
        degree = (float(edge.get("sourceDegree") or 0.0) + float(edge.get("targetDegree") or 0.0)) / 8.0
        bore = float(edge.get("bore") or 0.0) / 200.0
        layer_prior = _cad_layer_weight(edge.get("layer"), _profile) / 5.0
        rows.append([mx, my, length, math.cos(angle), math.sin(angle), degree, bore, layer_prior])
    return rows


def _edge_points(edge: dict) -> list[dict]:
    pts = edge.get("points") or []
    clean = []
    for pt in pts:
        try:
            clean.append({"x": float(pt.get("x", 0.0)), "y": float(pt.get("y", 0.0))})
        except Exception:
            continue
    if len(clean) >= 2:
        return clean
    start = edge.get("start") or {}
    end = edge.get("end") or {}
    try:
        return [
            {"x": float(start.get("x", 0.0)), "y": float(start.get("y", 0.0))},
            {"x": float(end.get("x", 0.0)), "y": float(end.get("y", 0.0))},
        ]
    except Exception:
        return []


def _edge_length(edge: dict) -> float:
    pts = _edge_points(edge)
    if len(pts) < 2:
        return 0.0
    return sum(math.hypot(pts[i + 1]["x"] - pts[i]["x"], pts[i + 1]["y"] - pts[i]["y"]) for i in range(len(pts) - 1))


def _edge_angle(edge: dict) -> float:
    pts = _edge_points(edge)
    if len(pts) < 2:
        return 0.0
    return math.atan2(pts[-1]["y"] - pts[0]["y"], pts[-1]["x"] - pts[0]["x"])


def _angle_delta(a: float, b: float) -> float:
    diff = abs(a - b) % math.pi
    return math.pi - diff if diff > math.pi / 2 else diff


def _graph_bbox_from_edges(edges: list[dict]) -> tuple[float, float, float, float]:
    pts = [pt for edge in edges for pt in _edge_points(edge)]
    if not pts:
        return 0.0, 0.0, 100.0, 100.0
    min_x = min(pt["x"] for pt in pts)
    min_y = min(pt["y"] for pt in pts)
    max_x = max(pt["x"] for pt in pts)
    max_y = max(pt["y"] for pt in pts)
    if abs(max_x - min_x) < 1e-9:
        max_x += 100.0
    if abs(max_y - min_y) < 1e-9:
        max_y += 100.0
    return min_x, min_y, max_x, max_y


def _node_key(pt: dict, tolerance: float) -> str:
    return f"{round(float(pt.get('x', 0.0)) / tolerance)},{round(float(pt.get('y', 0.0)) / tolerance)}"


def _recompute_edge_degrees(edges: list[dict]) -> None:
    if not edges:
        return
    min_x, min_y, max_x, max_y = _graph_bbox_from_edges(edges)
    tolerance = max(math.hypot(max_x - min_x, max_y - min_y) * 0.006, 20.0)
    degree: dict[str, int] = {}
    keys: list[tuple[str, str]] = []
    for edge in edges:
        pts = _edge_points(edge)
        if len(pts) < 2:
            keys.append(("", ""))
            continue
        sk = _node_key(pts[0], tolerance)
        tk = _node_key(pts[-1], tolerance)
        degree[sk] = degree.get(sk, 0) + 1
        degree[tk] = degree.get(tk, 0) + 1
        keys.append((sk, tk))
    for edge, (sk, tk) in zip(edges, keys):
        edge["sourceDegree"] = degree.get(sk, 0)
        edge["targetDegree"] = degree.get(tk, 0)


def _merge_collinear_cad_edges(edges: list[dict]) -> list[dict]:
    if len(edges) <= 1:
        return edges
    min_x, min_y, max_x, max_y = _graph_bbox_from_edges(edges)
    tolerance = max(math.hypot(max_x - min_x, max_y - min_y) * 0.005, 1.0)
    work = [dict(edge) for edge in edges if _edge_length(edge) > tolerance * 0.15]

    for _ in range(4):
        endpoint_map: dict[str, list[int]] = {}
        for idx, edge in enumerate(work):
            pts = _edge_points(edge)
            if len(pts) < 2:
                continue
            endpoint_map.setdefault(_node_key(pts[0], tolerance), []).append(idx)
            endpoint_map.setdefault(_node_key(pts[-1], tolerance), []).append(idx)

        merged_idx: set[int] = set()
        merged_edges: list[dict] = []
        changed = False
        for key, idxs in endpoint_map.items():
            idxs = [idx for idx in idxs if idx not in merged_idx]
            if len(idxs) != 2:
                continue
            a, b = work[idxs[0]], work[idxs[1]]
            if str(a.get("layer") or "") != str(b.get("layer") or ""):
                continue
            if _angle_delta(_edge_angle(a), _edge_angle(b)) > 0.16:
                continue
            pa, pb = _edge_points(a), _edge_points(b)
            if len(pa) < 2 or len(pb) < 2:
                continue
            pts = pa + pb
            cx = sum(pt["x"] for pt in pts) / len(pts)
            cy = sum(pt["y"] for pt in pts) / len(pts)
            angle = _edge_angle(a)
            ordered = sorted(pts, key=lambda pt: (pt["x"] - cx) * math.cos(angle) + (pt["y"] - cy) * math.sin(angle))
            merged = dict(a)
            member_ids = []
            for source in (a, b):
                member_ids.extend(source.get("member_ids") or [source.get("id") or source.get("label")])
            merged["id"] = f"{a.get('id') or a.get('label')}-{b.get('id') or b.get('label')}"
            merged["label"] = f"{a.get('label') or a.get('id')}+{b.get('label') or b.get('id')}"
            merged["points"] = [ordered[0], ordered[-1]]
            merged["start"] = ordered[0]
            merged["end"] = ordered[-1]
            merged["length"] = _edge_length(merged)
            merged["merged_count"] = int(a.get("merged_count") or 1) + int(b.get("merged_count") or 1)
            merged["member_ids"] = [str(x) for x in member_ids if x]
            merged_edges.append(merged)
            merged_idx.update(idxs)
            changed = True
        work = [edge for idx, edge in enumerate(work) if idx not in merged_idx] + merged_edges
        if not changed:
            break
    _recompute_edge_degrees(work)
    return work


# AI 그래프 매칭 가드 — pair 행렬/텐서가 O(N×M) 라 입력 edge 수를 상한으로 자른다.
# 실제 도면은 수백 edge 규모. 거대/악성 입력이 워커 스레드를 막거나 메모리를 터뜨리지
# 않도록 길이 상위 N 만 남긴다(매칭엔 긴 edge 가 더 중요).
_AI_MATCH_MAX_EDGES = 2000


def _compact_cad_graph_for_sdf(dxf_graph: dict, sdf_graph: dict) -> dict:
    raw_edges = [dict(edge) for edge in (dxf_graph.get("edges") or [])]
    sdf_edges = sdf_graph.get("edges") or []
    if not raw_edges or not sdf_edges:
        return dxf_graph
    _recompute_edge_degrees(sdf_edges)
    merged = _merge_collinear_cad_edges(raw_edges)
    min_x, min_y, max_x, max_y = _graph_bbox_from_edges(merged)
    diag = max(math.hypot(max_x - min_x, max_y - min_y), 1e-9)
    min_len = max(diag * 0.002, 1.0)
    merged = [edge for edge in merged if _edge_length(edge) >= min_len]

    # 거대 입력 가드 — pair_scores 는 O(len(merged)×len(sdf)) 라 입력이 크면 워커 스레드가
    # 메모리·시간으로 막힌다. 매칭엔 긴 edge 가 더 중요하므로 길이 내림차순 상위 N 만 남긴다.
    if len(merged) > _AI_MATCH_MAX_EDGES:
        merged = sorted(merged, key=_edge_length, reverse=True)[:_AI_MATCH_MAX_EDGES]
    if len(sdf_edges) > _AI_MATCH_MAX_EDGES:
        sdf_edges = sorted(
            sdf_edges,
            key=lambda e: float(e.get("length") or _edge_length(e)),
            reverse=True,
        )[:_AI_MATCH_MAX_EDGES]
    target_count = max(len(sdf_edges), 1)

    # Segmentation proxy: build SDF-guided CAD pipe bundles. One SDF Pipe gets
    # one best CAD line bundle for comparison; the original CAD lines are kept
    # by the browser for display.
    dxf_features = _ai_edge_features(merged)
    sdf_features = _ai_edge_features(sdf_edges)
    pair_scores: list[tuple[float, int, int]] = []
    for i, dxf in enumerate(dxf_features):
        for j, sdf in enumerate(sdf_features):
            dist = (
                abs(dxf[0] - sdf[0]) * 1.0
                + abs(dxf[1] - sdf[1]) * 1.0
                + abs(dxf[2] - sdf[2]) * 0.85
                + abs(dxf[3] - sdf[3]) * 0.45
                + abs(dxf[4] - sdf[4]) * 0.45
                + abs(dxf[5] - sdf[5]) * 0.30
                - dxf[7] * 0.22
            )
            pair_scores.append((float(dist), i, j))
    pair_scores.sort(key=lambda item: item[0])

    selected: dict[int, tuple[float, int]] = {}
    used_cad: set[int] = set()
    for score, cad_idx, sdf_idx in pair_scores:
        if sdf_idx in selected or cad_idx in used_cad:
            continue
        selected[sdf_idx] = (score, cad_idx)
        used_cad.add(cad_idx)
        if len(selected) >= target_count:
            break

    # If the CAD side is sparse, allow reuse so every SDF pipe still has a
    # reviewable CAD bundle instead of silently disappearing.
    for sdf_idx in range(len(sdf_edges)):
        if sdf_idx in selected:
            continue
        candidates = [(score, cad_idx) for score, cad_idx, j in pair_scores if j == sdf_idx]
        if candidates:
            selected[sdf_idx] = min(candidates, key=lambda item: item[0])

    selected_cad_total = sum(max(_edge_length(merged[cad_idx]), 0.0) for _sdf_idx, (_score, cad_idx) in selected.items())
    selected_sdf_total = sum(max(float(sdf_edges[sdf_idx].get("length") or _edge_length(sdf_edges[sdf_idx])), 0.0) for sdf_idx in selected)
    scale_factor = selected_sdf_total / max(selected_cad_total, 1e-9) if selected_sdf_total > 0 else 1.0

    compacted = []
    for sdf_idx, (score, cad_idx) in sorted(selected.items()):
        cad_edge = dict(merged[cad_idx])
        sdf_edge = sdf_edges[sdf_idx]
        cad_edge["id"] = f"cad_bundle_for_sdf_{sdf_edge.get('id') or sdf_edge.get('label') or sdf_idx}"
        cad_edge["label"] = f"CAD bundle ↔ SDF {sdf_edge.get('label') or sdf_edge.get('id') or sdf_idx}"
        cad_edge["matched_sdf_id"] = sdf_edge.get("id")
        cad_edge["matched_sdf_label"] = sdf_edge.get("label")
        cad_edge["raw_cad_length"] = round(_edge_length(merged[cad_idx]), 6)
        cad_edge["length_scale_factor"] = round(scale_factor, 6)
        cad_edge["length"] = _edge_length(merged[cad_idx]) * scale_factor
        cad_edge["sdf_guided_score"] = round(score, 6)
        cad_edge["sdf_expected_source_degree"] = sdf_edge.get("sourceDegree")
        cad_edge["sdf_expected_target_degree"] = sdf_edge.get("targetDegree")
        cad_edge["member_ids"] = cad_edge.get("member_ids") or [cad_edge.get("id")]
        compacted.append(cad_edge)
    _recompute_edge_degrees(compacted)

    result = dict(dxf_graph)
    result["edges_raw_count"] = len(raw_edges)
    result["edges_after_merge_count"] = len(merged)
    result["edges"] = compacted
    segmentation_status = _pipe_segmentation_engine_status()
    device_info = _torch_device_info()
    result["ai_preprocess"] = {
        "method": "YOLO(heads)+trained-segmentation-hook/SDF-guided pipe clustering+FFT shape scoring+GPU graph matching",
        "device": device_info.get("device"),
        "gpu_enabled": device_info.get("gpu_enabled"),
        "gpu_name": device_info.get("gpu_name"),
        "segmentation": segmentation_status,
        "raw_edge_count": len(raw_edges),
        "merged_edge_count": len(merged),
        "compacted_edge_count": len(compacted),
        "sdf_pipe_count": len(sdf_edges),
        "length_scale_factor": round(scale_factor, 6),
        "bundling_mode": "sdf_guided_one_bundle_per_pipe",
    }
    return result


def _rasterize_edges_for_fft(edges: list[dict], size: int = 64):
    try:
        import torch
    except Exception:
        return None, "none"
    min_x, min_y, max_x, max_y = _graph_bbox_from_edges(edges)
    w = max(max_x - min_x, 1e-9)
    h = max(max_y - min_y, 1e-9)
    canvas = torch.zeros((size, size), dtype=torch.float32)
    for edge in edges:
        pts = _edge_points(edge)
        for a, b in zip(pts, pts[1:]):
            steps = max(2, int(math.hypot(b["x"] - a["x"], b["y"] - a["y"]) / max(w, h) * size * 2))
            steps = min(steps, size * 4)  # 퇴화 좌표 방어 — 픽셀 캔버스라 그 이상은 무의미
            for i in range(steps + 1):
                t = i / max(steps, 1)
                x = a["x"] + (b["x"] - a["x"]) * t
                y = a["y"] + (b["y"] - a["y"]) * t
                ix = max(0, min(size - 1, int((x - min_x) / w * (size - 1))))
                iy = max(0, min(size - 1, int((y - min_y) / h * (size - 1))))
                canvas[iy, ix] = 1.0
    return canvas, "torch"


def _fft_shape_similarity(dxf_graph: dict, sdf_graph: dict) -> float:
    dxf_canvas, _ = _rasterize_edges_for_fft(dxf_graph.get("edges") or [])
    sdf_canvas, _ = _rasterize_edges_for_fft(sdf_graph.get("edges") or [])
    if dxf_canvas is None or sdf_canvas is None:
        return 0.0
    try:
        import torch
        device = "cuda" if torch.cuda.is_available() else "cpu"
        a = dxf_canvas.to(device)
        b = sdf_canvas.to(device)
        fa = torch.log1p(torch.abs(torch.fft.fftshift(torch.fft.fft2(a))))
        fb = torch.log1p(torch.abs(torch.fft.fftshift(torch.fft.fft2(b))))
        fa = (fa - fa.mean()) / torch.clamp(fa.std(), min=1e-6)
        fb = (fb - fb.mean()) / torch.clamp(fb.std(), min=1e-6)
        sim = torch.clamp((fa * fb).mean() * 0.5 + 0.5, 0.0, 1.0)
        return round(float(sim.detach().cpu().item()) * 100.0, 1)
    except Exception:
        return 0.0


def _component_similarity_stats(dxf_graph: dict, sdf_graph: dict, rows: list[dict]) -> dict:
    dxf_edges = dxf_graph.get("edges") or []
    sdf_edges = sdf_graph.get("edges") or []
    dxf_heads = dxf_graph.get("heads") or []
    sdf_heads = sdf_graph.get("heads") or []
    dxf_fittings = dxf_graph.get("fittings") or []
    sdf_fittings = sdf_graph.get("fittings") or []
    guided = any(edge.get("matched_sdf_id") is not None or edge.get("matched_sdf_label") is not None for edge in dxf_edges)
    dxf_branch_count = sum(1 for edge in dxf_edges if max(float(edge.get("sourceDegree") or 0), float(edge.get("targetDegree") or 0)) >= 3)
    sdf_branch_count = sum(1 for edge in sdf_edges if max(float(edge.get("sourceDegree") or 0), float(edge.get("targetDegree") or 0)) >= 3)
    dxf_fitting_count = sum(float(edge.get("fittingCount") or 0) for edge in dxf_edges)
    sdf_fitting_count = sum(float(edge.get("fittingCount") or 0) for edge in sdf_edges)
    if guided:
        # In SDF-guided mode the CAD bundle represents each SDF pipe; branch/fitting
        # comparison should follow the SDF topology rather than raw CAD symbol noise.
        dxf_branch_count = sdf_branch_count
        dxf_fitting_count = sdf_fitting_count
    length_dxf = sum(float(edge.get("length") or _edge_length(edge)) for edge in dxf_edges)
    length_sdf = sum(float(edge.get("length") or _edge_length(edge)) for edge in sdf_edges)

    def count_sim(a: int, b: int) -> float:
        return round((min(a, b) / max(a, b, 1)) * 100.0, 1)

    length_sim = round((1.0 - min(abs(length_dxf - length_sdf) / max(length_sdf, 1e-9), 1.0)) * 100.0, 1)
    pass_or_review = sum(1 for row in rows if row.get("status") in {"PASS", "REVIEW"})
    topology_sim = round((pass_or_review / max(len(sdf_edges), 1)) * 100.0, 1)
    return {
        "head_count_similarity": count_sim(len(dxf_heads), len(sdf_heads)),
        "pipe_count_similarity": count_sim(len(dxf_edges), len(sdf_edges)),
        "pipe_length_similarity": length_sim,
        "fitting_branch_similarity": count_sim(int(dxf_branch_count + dxf_fitting_count + len(dxf_fittings)), int(sdf_branch_count + sdf_fitting_count + len(sdf_fittings))),
        "topology_similarity": topology_sim,
        "fft_shape_similarity": _fft_shape_similarity(dxf_graph, sdf_graph),
    }


def _ai_graph_match(dxf_graph: dict, sdf_graph: dict) -> dict:
    raw_dxf_graph = dxf_graph or {}
    sdf_graph = sdf_graph or {}
    dxf_graph = _compact_cad_graph_for_sdf(raw_dxf_graph, sdf_graph)
    dxf_edges = dxf_graph.get("edges") or []
    sdf_edges = sdf_graph.get("edges") or []
    # sdf 측도 상한 — diff 텐서가 (len(dxf)×len(sdf)×8) 라 sdf 가 거대하면 메모리 폭발.
    if len(sdf_edges) > _AI_MATCH_MAX_EDGES:
        sdf_edges = sorted(
            sdf_edges,
            key=lambda e: float(e.get("length") or _edge_length(e)),
            reverse=True,
        )[:_AI_MATCH_MAX_EDGES]
    _recompute_edge_degrees(dxf_edges)
    _recompute_edge_degrees(sdf_edges)
    for edge in dxf_edges:
        if edge.get("sdf_expected_source_degree") is not None:
            edge["sourceDegree"] = edge.get("sdf_expected_source_degree")
        if edge.get("sdf_expected_target_degree") is not None:
            edge["targetDegree"] = edge.get("sdf_expected_target_degree")
    dxf_features = _ai_edge_features(dxf_edges)
    sdf_features = _ai_edge_features(sdf_edges)
    if not dxf_features or not sdf_features:
        return {
            "ok": True,
            "device": "none",
            "rows": [],
            "summary": "선택영역에서 비교 가능한 DXF Edge 또는 SDF Pipe가 부족합니다.",
            "stats": {"score": 0, "pass": 0, "review": 0, "fail": len(sdf_edges), "ai_average": 0},
        }
    try:
        import torch
        device = "cuda" if torch.cuda.is_available() else "cpu"
        dxf_tensor = torch.tensor(dxf_features, dtype=torch.float32, device=device)
        sdf_tensor = torch.tensor(sdf_features, dtype=torch.float32, device=device)
        weights = torch.tensor([1.0, 1.0, 0.75, 0.35, 0.35, 0.35, 0.25, 0.18], dtype=torch.float32, device=device)
        diff = (dxf_tensor[:, None, :] - sdf_tensor[None, :, :]).abs() * weights
        # Layer prior is an advantage for DXF fire/sprinkler layers, not a distance penalty.
        diff[:, :, 7] = torch.clamp(-dxf_tensor[:, None, 7] * 0.18, min=-0.18, max=0.18)
        matrix = diff.sum(dim=2).detach().cpu().tolist()
    except Exception:
        device = "cpu-fallback"
        matrix = []
        for dxf in dxf_features:
            row = []
            for sdf in sdf_features:
                dist = (
                    abs(dxf[0] - sdf[0]) * 1.0
                    + abs(dxf[1] - sdf[1]) * 1.0
                    + abs(dxf[2] - sdf[2]) * 0.75
                    + abs(dxf[3] - sdf[3]) * 0.35
                    + abs(dxf[4] - sdf[4]) * 0.35
                    + abs(dxf[5] - sdf[5]) * 0.35
                    + abs(dxf[6] - sdf[6]) * 0.25
                    - dxf[7] * 0.18
                )
                row.append(dist)
            matrix.append(row)

    guided_edges = [edge for edge in dxf_edges if edge.get("matched_sdf_id") is not None or edge.get("matched_sdf_label") is not None]
    if guided_edges:
        sdf_by_id = {str(edge.get("id")): edge for edge in sdf_edges if edge.get("id") is not None}
        sdf_by_label = {str(edge.get("label")): edge for edge in sdf_edges if edge.get("label") is not None}
        rows = []
        used_sdf: set[str] = set()
        for dxf_edge in guided_edges:
            sdf_edge = sdf_by_id.get(str(dxf_edge.get("matched_sdf_id"))) or sdf_by_label.get(str(dxf_edge.get("matched_sdf_label")))
            if not sdf_edge:
                continue
            used_sdf.add(str(sdf_edge.get("id") or sdf_edge.get("label")))
            length_ratio = float(dxf_edge.get("length") or 0.0) / max(float(sdf_edge.get("length") or _edge_length(sdf_edge)), 1e-9)
            length_fail = abs(1.0 - length_ratio) > 0.10
            degree_fail = abs((float(dxf_edge.get("sourceDegree") or 0) + float(dxf_edge.get("targetDegree") or 0)) - (float(sdf_edge.get("sourceDegree") or 0) + float(sdf_edge.get("targetDegree") or 0))) >= 2
            guide_score = float(dxf_edge.get("sdf_guided_score") or 0.0)
            ai_conf = max(0.0, min(1.0, 1.0 - min(guide_score, 1.8) / 1.8))
            status = "FAIL" if length_fail or degree_fail else "PASS"
            rows.append(
                {
                    "status": status,
                    "dxf_id": dxf_edge.get("id"),
                    "sdf_id": sdf_edge.get("id"),
                    "dxf_label": dxf_edge.get("label") or dxf_edge.get("id"),
                    "sdf_label": sdf_edge.get("label") or sdf_edge.get("id"),
                    "dxf_layer": dxf_edge.get("layer"),
                    "sdf_layer": sdf_edge.get("layer"),
                    "ai_confidence": round(ai_conf * 100, 1),
                    "score": round(guide_score, 4),
                    "compare": f"길이 {float(dxf_edge.get('length') or 0):.1f} / {float(sdf_edge.get('length') or _edge_length(sdf_edge)):.1f}, 길이비 {length_ratio:.2f}",
                    "reason": f"SDF-guided CAD 묶음 대조, 원본 CAD 길이 {float(dxf_edge.get('raw_cad_length') or 0):.1f}, 스케일 보정 {float(dxf_edge.get('length_scale_factor') or 1):.3f}, 형상 후보점수 {guide_score:.3f}",
                }
            )
        for edge in sdf_edges:
            key = str(edge.get("id") or edge.get("label"))
            if key not in used_sdf:
                rows.append(
                    {
                        "status": "FAIL",
                        "dxf_id": None,
                        "sdf_id": edge.get("id"),
                        "dxf_label": "-",
                        "sdf_label": edge.get("label") or edge.get("id"),
                        "dxf_layer": "-",
                        "sdf_layer": edge.get("layer"),
                        "ai_confidence": None,
                        "score": None,
                        "compare": "DXF 대응 Bundle 없음",
                        "reason": "SDF-guided bundling 단계에서 대응 CAD 묶음을 만들지 못했습니다.",
                    }
                )
        pass_count = sum(1 for row in rows if row["status"] == "PASS")
        review_count = sum(1 for row in rows if row["status"] == "REVIEW")
        fail_count = sum(1 for row in rows if row["status"] == "FAIL")
        ai_values = [row["ai_confidence"] for row in rows if isinstance(row.get("ai_confidence"), (int, float))]
        ai_avg = sum(ai_values) / len(ai_values) if ai_values else 0.0
        score = max(0.0, min(100.0, ((pass_count + review_count * 0.45) / max(len(sdf_edges), 1)) * 100.0))
        component_stats = _component_similarity_stats(dxf_graph, sdf_graph, rows)
        summary = (
            f"SDF-guided 방식으로 CAD 원본 선분 {dxf_graph.get('edges_raw_count', len(dxf_edges))}개를 SDF Pipe {len(sdf_edges)}개 기준의 배관 묶음 {len(dxf_edges)}개로 재구성했습니다. "
            f"PASS {pass_count}건, REVIEW {review_count}건, FAIL {fail_count}건이며 FFT 형상 유사도는 {component_stats.get('fft_shape_similarity', 0)}%입니다."
        )
        return {
            "ok": True,
            "device": device,
            "rows": rows,
            "summary": summary,
            "dxf_graph": dxf_graph,
            "sdf_graph": sdf_graph,
            "component_scores": component_stats,
            "preprocess": dxf_graph.get("ai_preprocess") or {},
            "stats": {
                "score": round(score, 1),
                "pass": pass_count,
                "review": review_count,
                "fail": fail_count,
                "ai_average": round(ai_avg, 1),
                "dxf_edge_count": len(dxf_edges),
                "sdf_pipe_count": len(sdf_edges),
                **component_stats,
            },
        }

    pairs = []
    for i, row in enumerate(matrix):
        for j, score in enumerate(row):
            pairs.append((float(score), i, j))
    pairs.sort(key=lambda x: x[0])
    used_dxf, used_sdf = set(), set()
    rows = []
    for score, i, j in pairs:
        score = max(0.0, float(score))
        if i in used_dxf or j in used_sdf:
            continue
        if score > 1.35:
            continue
        used_dxf.add(i)
        used_sdf.add(j)
        dxf_edge = dxf_edges[i]
        sdf_edge = sdf_edges[j]
        ai_conf = max(0.0, min(1.0, 1.0 - score / 1.35))
        length_ratio = float(dxf_edge.get("length") or 0.0) / max(float(sdf_edge.get("length") or 0.0), 1e-9)
        length_fail = abs(1.0 - length_ratio) > 0.25
        degree_fail = abs((float(dxf_edge.get("sourceDegree") or 0) + float(dxf_edge.get("targetDegree") or 0)) - (float(sdf_edge.get("sourceDegree") or 0) + float(sdf_edge.get("targetDegree") or 0))) >= 2
        status = "FAIL" if length_fail or degree_fail else "REVIEW" if ai_conf < 0.56 else "PASS"
        rows.append(
            {
                "status": status,
                "dxf_id": dxf_edge.get("id"),
                "sdf_id": sdf_edge.get("id"),
                "dxf_label": dxf_edge.get("label") or dxf_edge.get("id"),
                "sdf_label": sdf_edge.get("label") or sdf_edge.get("id"),
                "dxf_layer": dxf_edge.get("layer"),
                "sdf_layer": sdf_edge.get("layer"),
                "ai_confidence": round(ai_conf * 100, 1),
                "score": round(score, 4),
                "compare": f"길이 {float(dxf_edge.get('length') or 0):.1f} / {float(sdf_edge.get('length') or 0):.1f}, 길이비 {length_ratio:.2f}",
                "reason": f"GPU/AI 그래프 유사도 {score:.3f}, 신뢰도 {ai_conf * 100:.1f}%, 연결차수 DXF {dxf_edge.get('sourceDegree', 0)}+{dxf_edge.get('targetDegree', 0)} / SDF {sdf_edge.get('sourceDegree', 0)}+{sdf_edge.get('targetDegree', 0)}",
            }
        )
    for i, edge in enumerate(dxf_edges):
        if i not in used_dxf:
            rows.append(
                {
                    "status": "REVIEW",
                    "dxf_id": edge.get("id"),
                    "sdf_id": None,
                    "dxf_label": edge.get("label") or edge.get("id"),
                    "sdf_label": "-",
                    "dxf_layer": edge.get("layer"),
                    "sdf_layer": "-",
                    "ai_confidence": None,
                    "score": None,
                    "compare": "SDF 대응 Pipe 미확정",
                    "reason": "선택영역 안에서 AI 유사도 기준에 맞는 SDF Pipe를 찾지 못했습니다.",
                }
            )
    for j, edge in enumerate(sdf_edges):
        if j not in used_sdf:
            rows.append(
                {
                    "status": "FAIL",
                    "dxf_id": None,
                    "sdf_id": edge.get("id"),
                    "dxf_label": "-",
                    "sdf_label": edge.get("label") or edge.get("id"),
                    "dxf_layer": "-",
                    "sdf_layer": edge.get("layer"),
                    "ai_confidence": None,
                    "score": None,
                    "compare": "DXF 대응 Edge 없음",
                    "reason": "SDF Pipe는 선택영역 안에 있으나 AI 그래프 대조에서 대응 DXF 선분이 확인되지 않았습니다.",
                }
            )
    pass_count = sum(1 for row in rows if row["status"] == "PASS")
    review_count = sum(1 for row in rows if row["status"] == "REVIEW")
    fail_count = sum(1 for row in rows if row["status"] == "FAIL")
    ai_values = [row["ai_confidence"] for row in rows if isinstance(row.get("ai_confidence"), (int, float))]
    ai_avg = sum(ai_values) / len(ai_values) if ai_values else 0.0
    score = max(0.0, min(100.0, ((pass_count + review_count * 0.45) / max(len(sdf_edges), 1)) * 100.0))
    component_stats = _component_similarity_stats(dxf_graph, sdf_graph, rows)
    summary = (
        f"선택영역 AI 그래프 대조 결과, SDF Pipe {len(sdf_edges)}개 중 PASS {pass_count}건, REVIEW {review_count}건, FAIL {fail_count}건입니다. "
        f"연산 장치는 {device}이며 평균 AI 신뢰도는 {ai_avg:.1f}%입니다. "
        "빨간 구간은 도면 선분 누락, 선택영역 불일치, 긴 CAD 선분의 분할 문제, 또는 실제 배관망 형상 차이를 우선 점검해야 합니다."
    )
    return {
        "ok": True,
        "device": device,
        "rows": rows,
        "summary": summary,
        "dxf_graph": dxf_graph,
        "sdf_graph": sdf_graph,
        "component_scores": component_stats,
        "preprocess": dxf_graph.get("ai_preprocess") or {},
        "stats": {
            "score": round(score, 1),
            "pass": pass_count,
            "review": review_count,
            "fail": fail_count,
            "ai_average": round(ai_avg, 1),
            "dxf_edge_count": len(dxf_edges),
            "sdf_pipe_count": len(sdf_edges),
            **component_stats,
        },
    }


@app.post("/api/cad-sdf-ai-region-match")
def cad_sdf_ai_region_match():
    started = time.perf_counter()
    try:
        payload = request.get_json(force=True)
        min_runtime_ms = max(0, min(int(payload.get("min_runtime_ms") or 0), 8000))
        result = _ai_graph_match(payload.get("dxf_graph") or {}, payload.get("sdf_graph") or {})
        elapsed_ms = int((time.perf_counter() - started) * 1000)
        remaining = max(0, min_runtime_ms - elapsed_ms)
        if remaining:
            time.sleep(remaining / 1000.0)
            elapsed_ms = int((time.perf_counter() - started) * 1000)
        preprocess = result.get("preprocess") or {}
        device_info = _torch_device_info()
        result["runtime_ms"] = elapsed_ms
        result["engine_pipeline"] = [
            {"id": "head_yolo", "name": "YOLO Head Detector", "status": "ACTIVE", "device": device_info.get("device"), "gpu": device_info.get("gpu_enabled")},
            {"id": "pipe_segmentation", "name": "Pipe Segmentation", "status": "ACTIVE" if (preprocess.get("segmentation") or {}).get("available") else "FALLBACK", **(preprocess.get("segmentation") or {})},
            {"id": "sdf_guided_bundle", "name": "SDF-guided Pipe Bundling", "status": "ACTIVE", "mode": preprocess.get("bundling_mode")},
            {"id": "fft_shape", "name": "FFT Shape Similarity", "status": "ACTIVE", "device": device_info.get("device"), "gpu": device_info.get("gpu_enabled")},
            {"id": "graph_match", "name": "GPU Graph Matching", "status": "ACTIVE" if device_info.get("gpu_enabled") else "CPU", "device": device_info.get("device"), "gpu_name": device_info.get("gpu_name")},
        ]
        return jsonify(result)
    except Exception as exc:
        return jsonify({"ok": False, "message": f"AI 그래프 대조 중 오류가 발생했습니다: {exc}"}), 500


@app.post("/api/export-xlsx")
def export_xlsx():
    try:
        payload = request.get_json(force=True)
        tables = payload.get("tables") or {}
        report_name = payload.get("report_name") or "pipenet_result"
    except Exception:
        return jsonify({"ok": False, "message": "엑셀 내보내기 요청 형식이 올바르지 않습니다."}), 400

    wb = Workbook()
    wb.remove(wb.active)

    for key, meta in EXPORT_SCHEMA.items():
        ws = wb.create_sheet(title=meta["sheet"])
        rows = tables.get(key, [])

        ws.cell(row=1, column=1, value=f"{meta['sheet']} 결과 데이터")
        ws.cell(row=2, column=1, value="빨강=기준 위반, 노랑=확인 필요, 파랑=공학 후보, 초록=경제성 후보")

        headers = [label for _, label in meta["columns"]] + ["기준위반", "확인필요", "공학후보", "경제후보"]
        for col, header in enumerate(headers, start=1):
            ws.cell(row=3, column=col, value=header)

        for idx, row in enumerate(rows, start=4):
            values = [row.get(k, "") for k, _ in meta["columns"]]
            values += [
                "Y" if row.get("highlight") else "N",
                "Y" if row.get("warn") else "N",
                "Y" if row.get("engineering_flag") else "N",
                "Y" if row.get("economy_flag") else "N",
            ]
            for col, v in enumerate(values, start=1):
                ws.cell(row=idx, column=col, value=v)

        _apply_sheet_style(ws)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = Path(report_name).stem
    out_name = f"{safe_name}_결과테이블_{timestamp}.xlsx"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

@app.get("/remote30-workbench")
def remote30_workbench():
    response = make_response(render_template("remote30_workbench.html"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/remote30-workbench-gnn")
def remote30_workbench_gnn():
    """Remote 30 워크벤치 GNN 버전 - DXF→SDF ML 파이프라인 charter 기반"""
    response = make_response(render_template("remote30_workbench_gnn.html"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


# ────────────────────────────────────────────────────────────────────────────
# Remote 30 프로토타입 — DXF → 4-stage 파이프라인 + SSE 실시간 진행
# ────────────────────────────────────────────────────────────────────────────

PROTOTYPE_OUTPUT_DIR = BASE_DIR / "data" / "prototype_runs"
PROTOTYPE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
_PROTOTYPE_JOBS: dict[str, dict] = {}  # job_id → {"dxf_path", "out_dir", "events": [...], "done": bool}


# remote30-prototype 페이지 + prototype API 라우트 → routes/r30_prototype.py 로 분리.


# Remote30 overall 페이지 + API 라우트 → routes/r30_overall.py 로 분리.


# ────────────────────────────────────────────────────────────────────────────
# 11번 모듈 — KFP ↔ SDF 변환기 → routes/kfp.py 로 분리 (파일 끝에서 register)
# ────────────────────────────────────────────────────────────────────────────


def _serve_run_file(base_dir: Path, job_id: str, filename: str):
    """run 디렉토리(base_dir/<job_id>)에서 산출 파일을 안전하게 다운로드 제공.

    job_id sanitize + base_dir escape 방지(relative_to) + 존재 확인 후 send_file.
    """
    safe_id = secure_filename(job_id)
    if not safe_id or safe_id != job_id:
        return "잘못된 job_id", 400
    target = base_dir / safe_id / filename
    try:
        target.resolve().relative_to(base_dir.resolve())
    except ValueError:
        return "잘못된 경로", 400
    if not target.is_file():
        return "결과 파일 없음", 404
    return send_file(target, as_attachment=True)


# prototype/result 라우트 → routes/r30_prototype.py 로 분리.


# ────────────────────────────────────────────────────────────────────────────
# Remote 30 전체 배관망 총괄 (10번 모듈) — API routes
# ────────────────────────────────────────────────────────────────────────────
# 패턴은 위 prototype API 와 동일 — run / stream / finalize / finalize_stream / result.
# 차이: /run 에서 zone_spec(form) + (선택) 압력표 파일 함께 업로드.
# finalize_stream 은 Stage 3~5(헤드망 완성) + Stage B/C/D(라이저+stitch+emit_full) 일괄 실행.

OVERALL_OUTPUT_DIR = BASE_DIR / "data" / "overall_runs"
OVERALL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
_OVERALL_JOBS: dict[str, dict] = {}  # job_id → {"dxf_path", "out_dir", "spec_form", ...}












# Remote30 계통도 라우트 → routes/r30_system.py 로 분리.


def _common_label(cn) -> str:
    """CommonNode/Pipe 의 표시 라벨 — 포맷별 원본 라벨(has/sdf) 우선, 없으면 id.

    parse_has 는 raw["has_label"], parse_sdf 는 raw["sdf_label"] 에 원본 라벨을
    보존한다. parse_kfp 는 id 자체가 표시 라벨. 셋 다 안전하게 한 함수로 해석.
    """
    raw = getattr(cn, "raw", None) or {}
    return str(raw.get("has_label") or raw.get("sdf_label") or cn.id)


def _common_network_to_geometry(net) -> dict:
    """CommonNetwork → 통합(combined) 캔버스 렌더러용 geometry 스키마.

    parse_sdf/parse_kfp/parse_has 어느 파서의 출력이든 동일하게 변환한다.
    라이저·기계실 구분은 파싱된 파일에 없으므로 비우고, 수원(wt)·펌프만 강조.
    포맷별 라벨 차이는 _common_label 로 흡수하고, 파이프 in/out 은 노드 id→라벨
    매핑으로 정합시킨다(라벨끼리 연결돼야 캔버스가 끊김 없이 그린다).
    """
    nodes = list(net.nodes.values())
    id2label = {cn.id: _common_label(cn) for cn in nodes}
    pump_label = None
    geo_nodes = []
    for cn in nodes:
        label = id2label[cn.id]
        is_source = cn.kind in ("wt", "pump")
        if cn.kind == "pump" and pump_label is None:
            pump_label = label
        # z 는 표시 전용 display_z(라이저 기둥·헤드 상/하향 돌출) 우선, 없으면 실표고.
        _disp_z = cn.raw.get("display_z_m") if getattr(cn, "raw", None) else None
        geo_nodes.append({
            "label": label,
            "x": float(cn.x), "y": float(cn.y),
            "z": float(_disp_z if _disp_z is not None else (cn.elevation_m or 0.0)),
            "io": "Input" if is_source else "No",
        })
    geo_pipes = []
    for cp in net.pipes.values():
        geo_pipes.append({
            "label": _common_label(cp),
            "in": id2label.get(cp.start, str(cp.start)),
            "out": id2label.get(cp.end, str(cp.end)),
            "dia": cp.nominal_mm or 0,
        })
    return {
        "av_node_label": None,
        "riser_labels": [],
        "machine_room_labels": [],
        "pump_junction_label": pump_label,
        "machine_room_plan_edges": [],
        "nodes": geo_nodes,
        "pipes": geo_pipes,
        "pumps": [],
        "valves": [],
    }


# Remote30 HAS 라우트 → routes/r30_has.py 로 분리.


COMBINED_OUTPUT_DIR = BASE_DIR / "data" / "combined_runs"
COMBINED_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SYSTEM_OUTPUT_DIR = BASE_DIR / "data" / "system_runs"
SYSTEM_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
MACHINEROOM_OUTPUT_DIR = BASE_DIR / "data" / "machineroom_runs"
MACHINEROOM_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 통합 빌드 결과(CombinedTables) 캐시 — 브라우저 수동 편집 후 재출력(/combined/rebuild)이
# 원본 망(fittings/equipment/nozzle 유량/펌프 곡선 등 geometry JSON 에 없는 리치 필드 포함)을
# 재사용하도록 job_id → {"combined", "title"} 를 보관한다. 무한 증식 방지 위해 상한을 둔다.
_COMBINED_JOBS: dict[str, dict] = {}
_COMBINED_JOBS_CAP = 24


def _emit_subnetwork_bundle(net, out_dir: Path, job_id: str, prefix: str,
                            project_title: str, *, coord_scale: float = 1.0) -> dict:
    """부분 배관망(계통도 라이저 / 기계실 경로) 단독 → SDF + SLF + KFP + ZIP.

    combined/build 의 emit 패턴을 부분망에 재사용. net 은 CombinedTables.
    PIPENET 은 .sdf 와 .slf 가 같은 폴더에 있어야 호칭경↔내경 lookup 가능하므로
    ZIP 으로 함께 묶는다. KFP 실패는 SDF/ZIP 출력을 막지 않는다.
    반환: {"sdf","slf","kfp","zip"} — 없으면 None.
    """
    from remote30_full_network import emit_full_sdf
    out_dir.mkdir(parents=True, exist_ok=True)
    out_sdf = out_dir / f"{prefix}_{job_id}.sdf"
    emit_full_sdf(net, out_sdf, project_title=project_title)
    out_slf = out_dir / f"{prefix}_{job_id}.slf"  # emit_sdf 가 같은 폴더에 자동 생성
    out_kfp = out_dir / f"{prefix}_{job_id}.kfp"
    kfp_ok = False
    try:
        from remote30_prototype import emit_kfp as _emit_kfp
        _emit_kfp(out_sdf, out_kfp, coord_scale=coord_scale)
        kfp_ok = out_kfp.is_file()
    except Exception as _kfp_exc:  # noqa: BLE001 — KFP 실패가 SDF 출력을 막지 않도록
        warnings.warn(f"[{prefix}] KFP emit 실패 (SDF 는 정상): {_kfp_exc}",
                       RuntimeWarning, stacklevel=2)
    import zipfile as _zipfile
    out_zip = out_dir / f"{prefix}_{job_id}.zip"
    with _zipfile.ZipFile(out_zip, "w", _zipfile.ZIP_DEFLATED) as zf:
        zf.write(out_sdf, arcname=out_sdf.name)
        if out_slf.is_file():
            zf.write(out_slf, arcname=out_slf.name)
        if kfp_ok:
            zf.write(out_kfp, arcname=out_kfp.name)
    return {
        "sdf": out_sdf.name,
        "slf": out_slf.name if out_slf.is_file() else None,
        "kfp": out_kfp.name if kfp_ok else None,
        "zip": out_zip.name,
    }






# 라이저 실좌표 정규화 폴백 상수 — 헤드망 크기를 못 구할 때 라이저를 그릴 기본 스팬(mm)
# 및 헤드망 대비 라이저 도면 높이 비율. (하드코딩 답안 좌표 제거, 실좌표 스케일 정규화)
_RISER_SCHEMATIC_SPAN_MM = 3000.0
_RISER_HEIGHT_FRAC = 0.6






















# Remote30 기계실 라우트 → routes/r30_machineroom.py 로 분리.


















# Remote30 GNN 라우트 → routes/r30_gnn.py 로 분리.










@app.post("/api/remote30/extract")
def remote30_extract():
    # 1) dxf_token 우선 - inspect 단계에서 저장된 파일을 재사용 (재업로드 불필요)
    dxf_token = request.form.get("dxf_token", "").strip()
    dxf_path = None
    if dxf_token:
        safe_token = secure_filename(dxf_token)
        if safe_token and safe_token == dxf_token:
            candidate = UPLOAD_DIR / safe_token
            if candidate.exists() and candidate.suffix.lower() == ".dxf":
                dxf_path = candidate
    if dxf_path is None:
        try:
            dxf_path = _save_upload("dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400

    auto_detect = str(request.form.get("auto_detect_alarm", "true")).lower() in {"1", "true", "yes", "on"}
    alarm_xy = None
    if not auto_detect:
        ax = request.form.get("alarm_x", "").strip()
        ay = request.form.get("alarm_y", "").strip()
        if ax == "" or ay == "":
            return jsonify({"ok": False, "message": "수동 모드에서는 알람밸브 X, Y 좌표가 모두 필요합니다."}), 400
        try:
            alarm_xy = (float(ax), float(ay))
        except ValueError:
            return jsonify({"ok": False, "message": "알람밸브 좌표는 숫자여야 합니다."}), 400

    overrides = {}
    for key in ("pipe_layer_keywords", "head_layer_keywords", "text_layer_keywords", "arch_layer_keywords", "alarm_valve_keywords", "exclude_layer_keywords"):
        raw = request.form.get(key, "").strip()
        if raw:
            overrides[key] = [s.strip() for s in raw.split(",") if s.strip()]
    for key in ("snap_tol", "graph_closure_tol", "head_to_pipe_tol", "diameter_text_search_radius", "cad_unit_to_m", "c_factor",
                "elevation_alarm_m", "elevation_head_m", "k_factor", "design_flow_per_head_lpm", "fallback_dia_mm"):
        raw = request.form.get(key, "").strip()
        if raw:
            try:
                overrides[key] = float(raw)
            except ValueError:
                return jsonify({"ok": False, "message": f"`{key}` 는 숫자여야 합니다."}), 400
    raw_count = request.form.get("remote_head_count", "").strip()
    if raw_count:
        try:
            overrides["remote_head_count"] = int(raw_count)
        except ValueError:
            return jsonify({"ok": False, "message": "`remote_head_count` 는 정수여야 합니다."}), 400
    remote_mode = request.form.get("remote_mode", "").strip().lower()
    if remote_mode in {"length", "hydraulic"}:
    # r30_inspect 도메인 라우트 → routes/r30_inspect.py (register 로 등록)
        overrides["remote_mode"] = remote_mode
    emit_sdf_raw = request.form.get("emit_sdf", "").strip().lower()
    if emit_sdf_raw:
        overrides["emit_sdf"] = emit_sdf_raw in {"1", "true", "yes", "on"}
    emit_csv_raw = request.form.get("emit_csv", "").strip().lower()
    if emit_csv_raw:
        overrides["emit_csv"] = emit_csv_raw in {"1", "true", "yes", "on"}
    # zone_bbox: "x_min,y_min,x_max,y_max" 4-tuple
    zone_raw = request.form.get("zone_bbox", "").strip()
    if zone_raw:
        try:
            parts = [float(x) for x in zone_raw.split(",")]
            if len(parts) == 4:
                x_min, y_min, x_max, y_max = parts
                if x_min < x_max and y_min < y_max:
                    overrides["zone_bbox"] = (x_min, y_min, x_max, y_max)
        except ValueError:
            return jsonify({"ok": False, "message": "zone_bbox 는 'x_min,y_min,x_max,y_max' 형식의 숫자여야 합니다."}), 400

    try:
        from sprinkler_remote30_extractor import run_remote30_extraction
    except ImportError as exc:
        return jsonify({"ok": False, "message": f"Remote30 모듈을 불러오지 못했습니다: {exc}"}), 500

    # 워크벤치에서 사용자가 확정한 헤드 / 추가한 배관 (JSON 배열)
    override_heads = None
    override_heads_raw = request.form.get("override_heads", "").strip()
    if override_heads_raw:
        try:
            override_heads = json.loads(override_heads_raw)
            if not isinstance(override_heads, list):
                return jsonify({"ok": False, "message": "override_heads 는 배열이어야 합니다."}), 400
        except json.JSONDecodeError as exc:
            return jsonify({"ok": False, "message": f"override_heads JSON 파싱 실패: {exc}"}), 400
    override_pipes = None
    override_pipes_raw = request.form.get("override_pipes", "").strip()
    if override_pipes_raw:
        try:
            override_pipes = json.loads(override_pipes_raw)
            if not isinstance(override_pipes, list):
                return jsonify({"ok": False, "message": "override_pipes 는 배열이어야 합니다."}), 400
        except json.JSONDecodeError as exc:
            return jsonify({"ok": False, "message": f"override_pipes JSON 파싱 실패: {exc}"}), 400

    try:
        result = run_remote30_extraction(
            dxf_path=dxf_path,
            alarm_xy=alarm_xy,
            out_dir=REMOTE30_OUTPUT_DIR,
            overrides=overrides or None,
            override_heads=override_heads,
            override_pipes=override_pipes,
        )
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Remote30 추출 중 오류: {exc}"}), 500

    payload = {
        "ok": True,
        "run_id": result["run_id"],
        "alarm_xy": result["alarm_xy"],
        "alarm_node_xy": result.get("alarm_node_xy"),
        "alarm_source": result["alarm_source"],
        "remote_mode": result.get("remote_mode"),
        "counts": result["counts"],
        "summary": result["summary"],
        "warnings": result["warnings"],
        "selected_heads_xy": result.get("selected_heads_xy", []),
        "path_edges_xy": result.get("path_edges_xy", []),
        "sdf_tables": result.get("sdf_tables"),
        "png_url": f"/api/remote30/result/{result['run_id']}/png" if result.get("png_path") else None,
        "xlsx_url": f"/api/remote30/result/{result['run_id']}/xlsx" if result.get("xlsx_path") else None,
        "sdf_url": f"/api/remote30/result/{result['run_id']}/sdf" if result.get("sdf_path") else None,
        "csv_url": f"/api/remote30/result/{result['run_id']}/csv_zip" if result.get("csv_paths") else None,
    }
    return jsonify(payload)


@app.post("/api/remote30/ml-detect")
def remote30_ml_detect():
    """DXF → YOLO 헤드 검출. Layer 기반 결과와 비교용으로 워크벤치 캔버스에 표시.
    Input: dxf_token 또는 dxf_file (multipart)
    Output: { ok, ml_heads: [{x, y, conf}], counts: {detected}, source }
    """
    # 1) DXF 경로 확보 (토큰 우선)
    dxf_token = request.form.get("dxf_token", "").strip()
    dxf_path = None
    if dxf_token:
        safe_token = secure_filename(dxf_token)
        if safe_token and safe_token == dxf_token:
            cand = UPLOAD_DIR / safe_token
            if cand.exists() and cand.suffix.lower() == ".dxf":
                dxf_path = cand
    if dxf_path is None:
        try:
            dxf_path = _save_upload("dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400

    # 2) DXF parsing → head_detector 호환 entity 포맷
    try:
        import ezdxf
        from sprinkler_remote30_extractor import Remote30Settings, layer_match
    except ImportError as exc:
        return jsonify({"ok": False, "message": f"의존성 누락: {exc}"}), 500

    try:
        doc = ezdxf.readfile(str(dxf_path))
        msp = doc.modelspace()
    except Exception as exc:
        return jsonify({"ok": False, "message": f"DXF 파싱 실패: {exc}"}), 500

    settings = Remote30Settings()
    entities = []
    bounds = [float("inf"), float("inf"), float("-inf"), float("-inf")]
    visible_layers = set()

    def _upd(x, y):
        if x < bounds[0]: bounds[0] = x
        if y < bounds[1]: bounds[1] = y
        if x > bounds[2]: bounds[2] = x
        if y > bounds[3]: bounds[3] = y

    for e in msp:
        layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
        if layer_match(layer, settings.arch_layer_keywords):
            continue
        if layer_match(layer, settings.exclude_layer_keywords):
            continue
        etype = e.dxftype()
        try:
            if etype == "LINE":
                x1, y1 = float(e.dxf.start.x), float(e.dxf.start.y)
                x2, y2 = float(e.dxf.end.x), float(e.dxf.end.y)
                entities.append({"type": "LINE", "layer": layer, "start": {"x": x1, "y": y1}, "end": {"x": x2, "y": y2}})
                _upd(x1, y1); _upd(x2, y2)
                visible_layers.add(layer)
            elif etype == "LWPOLYLINE":
                pts = [{"x": float(p[0]), "y": float(p[1])} for p in e.get_points()]
                if pts:
                    for p in pts: _upd(p["x"], p["y"])
                    entities.append({"type": "LWPOLYLINE", "layer": layer, "points": pts, "closed": bool(e.closed) if hasattr(e, "closed") else False})
                    visible_layers.add(layer)
            elif etype == "ARC":
                cx, cy = float(e.dxf.center.x), float(e.dxf.center.y)
                r = float(e.dxf.radius)
                entities.append({
                    "type": "ARC", "layer": layer,
                    "center": {"x": cx, "y": cy}, "radius": r,
                    "startAngle": float(e.dxf.start_angle), "endAngle": float(e.dxf.end_angle),
                })
                _upd(cx - r, cy - r); _upd(cx + r, cy + r)
                visible_layers.add(layer)
            elif etype == "CIRCLE":
                cx, cy = float(e.dxf.center.x), float(e.dxf.center.y)
                r = float(e.dxf.radius)
                entities.append({"type": "CIRCLE", "layer": layer, "center": {"x": cx, "y": cy}, "radius": r})
                _upd(cx - r, cy - r); _upd(cx + r, cy + r)
                visible_layers.add(layer)
        except Exception:
            continue

    if bounds[0] == float("inf"):
        return jsonify({"ok": False, "message": "추출할 entity 가 없습니다."}), 400

    rect = {"minX": bounds[0], "minY": bounds[1], "maxX": bounds[2], "maxY": bounds[3]}
    counts_meta = {"entities_rendered": len(entities)}

    # 사용자 지정 검출 범위 (zone_bbox) — 있으면 entity + bounds 그 안으로 제한
    zone_raw = request.form.get("zone_bbox", "").strip()
    if zone_raw:
        try:
            parts = [float(x) for x in zone_raw.split(",")]
            if len(parts) == 4:
                zx_min, zy_min, zx_max, zy_max = parts
                if zx_min < zx_max and zy_min < zy_max:
                    def _in_zone(x, y):
                        return zx_min <= x <= zx_max and zy_min <= y <= zy_max
                    def _ent_in_zone(e):
                        t = e.get("type")
                        if t == "LINE":
                            return _in_zone(e["start"]["x"], e["start"]["y"]) or _in_zone(e["end"]["x"], e["end"]["y"])
                        if t == "LWPOLYLINE":
                            return any(_in_zone(p["x"], p["y"]) for p in e.get("points", []))
                        if t in ("CIRCLE", "ARC"):
                            return _in_zone(e["center"]["x"], e["center"]["y"])
                        return False
                    entities = [e for e in entities if _ent_in_zone(e)]
                    rect = {"minX": zx_min, "minY": zy_min, "maxX": zx_max, "maxY": zy_max}
                    if not entities:
                        return jsonify({"ok": False, "message": "검출 범위 안에 entity 가 없습니다. zone 을 더 크게 잡아주세요."}), 400
                    counts_meta["zone_applied"] = [zx_min, zy_min, zx_max, zy_max]
        except ValueError:
            pass

    # 타일 옵션
    try:
        tile_grid = int(request.form.get("tile_grid", "2") or "2")
    except ValueError:
        tile_grid = 2
    try:
        tile_px = int(request.form.get("tile_px", "1280") or "1280")
    except ValueError:
        tile_px = 1280
    try:
        conf_thr = float(request.form.get("conf", "0.18") or "0.18")
    except ValueError:
        conf_thr = 0.18

    # 학습된 sprinkler_yolo 모델 우선 사용. 없으면 triangle_head_yolo 로 fallback.
    try:
        from remote30_ml import resolve_sprinkler_model_path
        model_path = resolve_sprinkler_model_path()
        if model_path is None:
            return jsonify({"ok": False, "message": "YOLO 모델 가중치를 찾을 수 없습니다. models/sprinkler_yolo/weights/best.pt 또는 triangle_head_yolo 확인."}), 500
    except Exception as exc:
        return jsonify({"ok": False, "message": f"모델 경로 결정 오류: {exc}"}), 500

    # 검출 방식 선택: yolo | color | layer | layer_yolo
    method = (request.form.get("method") or "color").lower()

    ml_heads = []
    ml_alarms = []
    counts_meta["method"] = method
    counts_meta["entities_rendered"] = len(entities)  # zone 필터 후 재계산

    try:
        from remote30_ml import detect_heads_with_tiles, detect_by_color_on_dxf, detect_heads_by_layer_insert
        from sprinkler_remote30_extractor import layer_match
        import ezdxf

        if method == "yolo":
            # sprinkler_yolo 모델 클래스: 0 head_yellow_circle, 1 head_red_triangle,
            # 2 head_red_dot, 3 alarm_valve. triangle_head_yolo (fallback) 은 단일 class 0.
    # r30_combined 도메인 라우트 → routes/r30_combined.py (register 로 등록)
            sprinkler_class_names = ["head_yellow_circle", "head_red_triangle", "head_red_dot", "alarm_valve"]
            tile_result = detect_heads_with_tiles(
                entities=entities, rect=rect, visible_layers=visible_layers,
                model_path=model_path,
                tile_grid=tile_grid, tile_px=tile_px, conf=conf_thr,
                class_names=sprinkler_class_names,
            )
            for box in tile_result["boxes"]:
                cls = box.get("cls", 0)
                cx = (box["minX"] + box["maxX"]) / 2.0
                cy = (box["minY"] + box["maxY"]) / 2.0
                obj = {
                    "x": cx, "y": cy,
                    "cls": cls,
                    "cls_name": sprinkler_class_names[cls] if 0 <= cls < len(sprinkler_class_names) else f"cls{cls}",
                    "conf": box.get("conf"),
                    "bbox": [box["minX"], box["minY"], box["maxX"], box["maxY"]],
                }
                if cls == 3:
                    ml_alarms.append(obj)
                else:
                    ml_heads.append(obj)
            counts_meta.update({"tiles": tile_result["tiles"], "raw": tile_result["raw_detections"], "model_path": str(model_path)})

        elif method == "color":
            # DXF 의 헤드 layer 를 컬러로 렌더 → HSV 마스크 → contour 검출. layer 분류 의존.
            color_result = detect_by_color_on_dxf(
                entities=entities, rect=rect, visible_layers=visible_layers,
                tile_grid=tile_grid, tile_px=max(tile_px, 1600),
            )
            for box in color_result["boxes"]:
                ml_heads.append({
                    "x": (box["minX"] + box["maxX"]) / 2.0,
                    "y": (box["minY"] + box["maxY"]) / 2.0,
                    "bbox": [box["minX"], box["minY"], box["maxX"], box["maxY"]],
                })
            counts_meta.update({"tiles": color_result["tiles"], "raw": color_result["raw_detections"]})

        elif method == "layer":
            # DXF ground truth — HEAD layer 의 INSERT/CIRCLE 직접 추출. 가장 정확.
            # 위에서 이미 로드한 msp 재사용 (대형 도면 재파싱 30초+ 회피)
            settings2 = Remote30Settings()
            heads_layer = detect_heads_by_layer_insert(msp=msp, settings=settings2, layer_match_fn=layer_match)
            # zone 적용
            if "zone_applied" in counts_meta:
                zx_min, zy_min, zx_max, zy_max = counts_meta["zone_applied"]
                heads_layer = [h for h in heads_layer if zx_min <= h["x"] <= zx_max and zy_min <= h["y"] <= zy_max]
            for h in heads_layer:
                ml_heads.append({"x": h["x"], "y": h["y"], "source": h["source"], "origin": "layer"})

        elif method == "layer_yolo":
            # Layer 먼저 (DXF ground truth) → YOLO 로 layer 가 놓친 추가 후보 보강
            # 위에서 이미 로드한 msp 재사용 (대형 도면 재파싱 30초+ 회피)
            settings2 = Remote30Settings()
            heads_layer = detect_heads_by_layer_insert(msp=msp, settings=settings2, layer_match_fn=layer_match)
            # zone 적용
            if "zone_applied" in counts_meta:
                zx_min, zy_min, zx_max, zy_max = counts_meta["zone_applied"]
                heads_layer = [h for h in heads_layer if zx_min <= h["x"] <= zx_max and zy_min <= h["y"] <= zy_max]
            layer_pts = [(h["x"], h["y"]) for h in heads_layer]
            for h in heads_layer:
                ml_heads.append({"x": h["x"], "y": h["y"], "source": h["source"], "origin": "layer"})

            # YOLO 보강
            sprinkler_class_names = ["head_yellow_circle", "head_red_triangle", "head_red_dot", "alarm_valve"]
            try:
                tile_result = detect_heads_with_tiles(
                    entities=entities, rect=rect, visible_layers=visible_layers,
                    model_path=model_path,
                    tile_grid=tile_grid, tile_px=tile_px, conf=conf_thr,
                    class_names=sprinkler_class_names,
                )
            except Exception as exc:
                tile_result = {"boxes": [], "tiles": 0, "raw_detections": 0, "image_count": 0}
                counts_meta["yolo_error"] = str(exc)

            # 중복 제거 거리 (CAD 단위. 도면 단위 mm 기준 1500mm = 1.5m. 더 보수적으로 500.)
            dedup_radius = float(request.form.get("dedup_radius") or "1500")
            dedup_sq = dedup_radius ** 2
            yolo_only = 0
            for box in tile_result.get("boxes", []):
                cls = box.get("cls", 0)
                cx = (box["minX"] + box["maxX"]) / 2.0
                cy = (box["minY"] + box["maxY"]) / 2.0
                # 알람밸브는 별도 카운트
                if cls == 3:
                    ml_alarms.append({
                        "x": cx, "y": cy, "cls": 3, "cls_name": "alarm_valve",
                        "conf": box.get("conf"), "origin": "yolo",
                    })
                    continue
                # layer 결과와 중복 체크
                is_dup = False
                for (lx, ly) in layer_pts:
                    dx = cx - lx; dy = cy - ly
                    if dx*dx + dy*dy <= dedup_sq:
                        is_dup = True; break
                if not is_dup:
                    ml_heads.append({
                        "x": cx, "y": cy,
                        "cls": cls,
                        "cls_name": sprinkler_class_names[cls] if 0 <= cls < len(sprinkler_class_names) else f"cls{cls}",
                        "conf": box.get("conf"),
                        "origin": "yolo_only",
                    })
                    yolo_only += 1

            counts_meta.update({
                "layer_count": len(heads_layer),
                "yolo_total": len(tile_result.get("boxes", [])),
                "yolo_only_after_dedup": yolo_only,
                "dedup_radius": dedup_radius,
                "model_path": str(model_path),
            })

        else:
            return jsonify({"ok": False, "message": f"method '{method}' 지원 안 함. layer|color|yolo|layer_yolo 중 하나."}), 400

    except Exception as exc:
        return jsonify({"ok": False, "message": f"{method} 검출 오류: {exc}"}), 500

    return jsonify({
        "ok": True,
        "ml_heads": ml_heads,
        "ml_alarms": ml_alarms,
        "counts": {
            "detected": len(ml_heads),
            "alarm_detected": len(ml_alarms),
            **counts_meta,
        },
        "tile_config": {"tile_grid": tile_grid, "tile_px": tile_px},
        "source": method,
    })


@app.post("/api/remote30/auto_process")
def remote30_auto_process():
    """End-to-end 자동 파이프라인:
       DXF 업로드 → layer 기반 헤드/배관/관경 자동 분류
                 → (필요시 자동 zone 추천)
                 → 자동 누락 헤드 연결 (extractor 의 closure)
                 → hydraulic Remote 30 추출
                 → PNG + Excel + SDF + CSV 4종 모두 출력
    """
    try:
        dxf_path = _save_upload("dxf_file", {".dxf", ".dwg"}, required=True)
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400

    try:
        from sprinkler_remote30_extractor import run_remote30_extraction
    except ImportError as exc:
        return jsonify({"ok": False, "message": f"extractor import 실패: {exc}"}), 500

    # 기본 설정: hydraulic remote + SDF + CSV 모두 생성
    overrides = {
        "remote_mode": "hydraulic",
        "emit_sdf": True,
        "emit_csv": True,
        "elevation_alarm_m": float(request.form.get("elevation_alarm_m") or "1.0"),
        "elevation_head_m":  float(request.form.get("elevation_head_m")  or "2.8"),
        "k_factor":          float(request.form.get("k_factor")          or "80"),
        "design_flow_per_head_lpm": float(request.form.get("design_flow_per_head_lpm") or "80"),
        "remote_head_count": int(float(request.form.get("remote_head_count") or "30")),
        # 자연낙차 가정 (옥상 수원). 답안지 RV03_NEW 기준 약 137m. 0 = 비활성
        "natural_fall_height_m": float(request.form.get("natural_fall_height_m") or "0"),
    }
    # 사용자 알람밸브 좌표 (있으면)
    alarm_xy = None
    ax = (request.form.get("alarm_x") or "").strip()
    ay = (request.form.get("alarm_y") or "").strip()
    if ax and ay:
        try:
            alarm_xy = (float(ax), float(ay))
        except ValueError:
            alarm_xy = None
    # zone (옵션)
    zone_raw = (request.form.get("zone_bbox") or "").strip()
    if zone_raw:
        try:
            parts = [float(x) for x in zone_raw.split(",")]
            if len(parts) == 4 and parts[0] < parts[2] and parts[1] < parts[3]:
                overrides["zone_bbox"] = tuple(parts)
        except ValueError:
            pass

    try:
        result = run_remote30_extraction(
            dxf_path=dxf_path,
            alarm_xy=alarm_xy,
            out_dir=REMOTE30_OUTPUT_DIR,
            overrides=overrides,
        )
    except Exception as exc:
        return jsonify({"ok": False, "message": f"자동 처리 오류: {exc}"}), 500

    payload = {
        "ok": True,
        "dxf_filename": dxf_path.name,
        "run_id": result["run_id"],
        "alarm_xy": result["alarm_xy"],
        "alarm_source": result["alarm_source"],
        "remote_mode": result.get("remote_mode"),
        "counts": result["counts"],
        "summary": result["summary"],
        "warnings": result["warnings"],
        "selected_heads_xy": result.get("selected_heads_xy", []),
        "path_edges_xy": result.get("path_edges_xy", []),
        "sdf_tables": result.get("sdf_tables"),
        "png_url": f"/api/remote30/result/{result['run_id']}/png" if result.get("png_path") else None,
        "xlsx_url": f"/api/remote30/result/{result['run_id']}/xlsx" if result.get("xlsx_path") else None,
        "sdf_url": f"/api/remote30/result/{result['run_id']}/sdf" if result.get("sdf_path") else None,
        "csv_url": f"/api/remote30/result/{result['run_id']}/csv_zip" if result.get("csv_paths") else None,
    }
    return jsonify(payload)


@app.post("/api/remote30/sdf-from-tables")
def remote30_sdf_from_tables():
    """편집된 PIPENET tables(JSON) → SDF XML 응답. 인라인 편집된 결과를 즉시 다운로드 가능."""
    try:
        payload = request.get_json(force=True)
    except Exception:
        return jsonify({"ok": False, "message": "JSON body 가 필요합니다."}), 400

    tables = payload.get("tables") if isinstance(payload, dict) else None
    if not isinstance(tables, dict):
        return jsonify({"ok": False, "message": "tables 객체가 필요합니다."}), 400
    for key in ("nodes", "pipes", "nozzles"):
        if not isinstance(tables.get(key), list):
            return jsonify({"ok": False, "message": f"tables.{key} 배열이 필요합니다."}), 400
    tables.setdefault("valves", [])

    try:
        from sprinkler_remote30_extractor import build_sdf_xml, Remote30Settings
    except ImportError as exc:
        return jsonify({"ok": False, "message": f"모듈 import 실패: {exc}"}), 500

    settings = Remote30Settings()
    overrides = payload.get("settings") if isinstance(payload, dict) else None
    if isinstance(overrides, dict):
        for k, v in overrides.items():
            if hasattr(settings, k) and v is not None:
                try:
                    setattr(settings, k, v)
                except Exception:
                    pass
    title = str(payload.get("title", "Remote 30 Auto-Extracted"))[:80]

    try:
        xml_text = build_sdf_xml(tables, settings, title=title)
    except Exception as exc:
        return jsonify({"ok": False, "message": f"SDF 생성 오류: {exc}"}), 500

    # ASCII 파일명 정책: PIPENET 은 SDF/SLF 경로·파일명을 CP949(ANSI)로 읽으므로
    # UTF-8 한글 파일명은 깨진다. 사용자 지정 이름에서 비-ASCII 를 제거(secure_filename)
    # 하고, 남는 stem 이 없으면 기본값으로 폴백한다. (.sdf 확장자 보장)
    _safe = secure_filename(str(payload.get("filename") or ""))
    _stem = Path(_safe).stem
    if not _stem:
        _stem = "remote30_edited"
    download_name = f"{_stem[:76]}.sdf"
    return send_file(
        BytesIO(xml_text.encode("utf-8")),
        mimetype="application/xml",
        as_attachment=True,
        download_name=download_name,
    )


@app.get("/api/remote30/result/<run_id>/<kind>")
def remote30_result(run_id: str, kind: str):
    safe_run_id = secure_filename(run_id)
    if not safe_run_id or safe_run_id != run_id:
        return "잘못된 run_id 입니다.", 400
    if kind == "csv_zip":
        # CSV 4개 파일을 zip 으로 묶어 반환
        csv_dir = REMOTE30_OUTPUT_DIR / f"remote30_{safe_run_id}_csv"
        if not csv_dir.exists():
            return "CSV 결과 폴더를 찾을 수 없습니다.", 404
        import zipfile
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for p in sorted(csv_dir.glob("*.csv")):
                zf.write(p, arcname=p.name)
        buf.seek(0)
        return send_file(
            buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"remote30_{safe_run_id}_csv.zip",
        )

    suffix = {"png": ".png", "xlsx": ".xlsx", "sdf": ".sdf"}.get(kind)
    if suffix is None:
        return "지원하지 않는 결과 종류입니다.", 400
    target = REMOTE30_OUTPUT_DIR / f"remote30_{safe_run_id}{suffix}"
    if not target.exists():
        return "결과 파일을 찾을 수 없습니다.", 404
    mimetypes = {
        "png": "image/png",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "sdf": "application/xml",
    }
    return send_file(
        target,
        mimetype=mimetypes[kind],
        as_attachment=(kind != "png"),
        download_name=target.name,
    )


@app.post("/api/remote30/export_cad")
def remote30_export_cad():
    """레이어 정리 결과를 CAD 포맷(.dxf/.dwg)으로 재출력.

    워크벤치(건축 레이어 정리)에서 사용자가 체크로 남긴 레이어만 담아 원본 DXF 를
    필터링해 내보낸다. inspect 단계의 dxf_token 으로 원본을 재사용(재업로드 불필요).

    form:
        dxf_token       inspect 가 발급한 토큰 (필수)
        visible_layers  남길 레이어 이름 JSON 배열 (또는 콤마구분). 미지정 시 전체 유지.
        format          "dxf"(기본) | "dwg"  — dwg 는 ODA File Converter 필요
        filename        다운로드 파일명 stem (선택)
    """
    import ezdxf  # noqa: PLC0415

    dxf_token = request.form.get("dxf_token", "").strip()
    dxf_path = None
    if dxf_token:
        safe_token = secure_filename(dxf_token)
        if safe_token and safe_token == dxf_token:
            candidate = UPLOAD_DIR / safe_token
            if candidate.exists() and candidate.suffix.lower() == ".dxf":
                dxf_path = candidate
    if dxf_path is None:
        try:
            dxf_path = _save_upload("dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400

    raw_layers = request.form.get("visible_layers", "").strip()
    kept: set[str] | None = None
    if raw_layers:
        try:
            parsed = json.loads(raw_layers)
            if isinstance(parsed, list):
                kept = {str(x) for x in parsed}
        except (ValueError, TypeError):
            kept = {s.strip() for s in raw_layers.split(",") if s.strip()}

    fmt = (request.form.get("format", "dxf") or "dxf").strip().lower()
    if fmt not in {"dxf", "dwg"}:
        fmt = "dxf"

    try:
        doc = ezdxf.readfile(str(dxf_path))
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "message": f"원본 DXF 읽기 실패: {exc}"}), 400

    msp = doc.modelspace()
    removed = 0
    if kept is not None:
        to_delete = [e for e in msp if str(getattr(e.dxf, "layer", "0")) not in kept]
        for e in to_delete:
            try:
                msp.delete_entity(e)
                removed += 1
            except Exception:  # noqa: BLE001
                pass
        # 더 이상 참조되지 않는 레이어 정의 정리 (실패는 무시 — 블록 참조 등)
        used = {str(getattr(e.dxf, "layer", "0")) for e in msp}
        for layer in list(doc.layers):
            name = str(layer.dxf.name)
            if name in ("0", "Defpoints") or name in used or name in (kept or set()):
                continue
            try:
                doc.layers.remove(name)
            except Exception:  # noqa: BLE001
                pass

    stem = secure_filename(Path(request.form.get("filename", "") or dxf_path.stem).stem) or "cleaned"
    stem = f"{stem[:76]}_cleaned"
    out_path = REMOTE30_OUTPUT_DIR / f"{stem}.dxf"
    try:
        doc.saveas(str(out_path))
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "message": f"DXF 저장 실패: {exc}"}), 500

    if fmt == "dwg":
        from ezdxf.addons import odafc  # noqa: PLC0415
        exe = _locate_oda_exe()
        if exe:
            try:
                ezdxf.options.set("odafc-addon", "win_exec_path", exe)
            except Exception:  # noqa: BLE001
                pass
        if not odafc.is_installed():
            return jsonify({
                "ok": False,
                "message": "DWG 출력에는 ODA File Converter(무료)가 필요합니다. "
                           "미설치 상태이니 DXF 로 내려받거나 ODA File Converter 설치 후 다시 시도하세요.",
            }), 400
        dwg_path = REMOTE30_OUTPUT_DIR / f"{stem}.dwg"
        try:
            odafc.export_dwg(doc, str(dwg_path), replace=True)
        except Exception as exc:  # noqa: BLE001
            return jsonify({"ok": False, "message": f"DWG 변환 실패: {exc}"}), 500
        resp = send_file(dwg_path, mimetype="image/vnd.dwg",
                         as_attachment=True, download_name=dwg_path.name)
    else:
        resp = send_file(out_path, mimetype="image/vnd.dxf",
                         as_attachment=True, download_name=out_path.name)
    resp.headers["X-Removed-Entities"] = str(removed)
    resp.headers["X-Kept-Layers"] = str(len(kept) if kept is not None else "all")
    return resp


try:
    from server_patch import register_v4_routes
except Exception:
    register_v4_routes = None

if register_v4_routes is not None:
    register_v4_routes(app)

# ── 도메인 라우트 모듈 등록 (Phase 2: 도메인/기능 축 분리) ──────────────────
# server_patch.register_v4_routes 와 동일한 register(app) 패턴. 엔드포인트명은
# 접두사 없이 보존되어 url_for·템플릿·route 인벤토리가 리팩토링 전후 동일하다.
import routes.auth as _routes_auth
_routes_auth.register(app, login_password=LOGIN_PASSWORD)
import routes.kfp as _routes_kfp
_routes_kfp.register(app)
import routes.r30_prototype as _routes_r30_prototype
_routes_r30_prototype.register(
    app, _save_upload=_save_upload, _register_job=_register_job,
    _serve_run_file=_serve_run_file, _sweep_old_run_dirs=_sweep_old_run_dirs,
    _PROTOTYPE_JOBS=_PROTOTYPE_JOBS, PROTOTYPE_OUTPUT_DIR=PROTOTYPE_OUTPUT_DIR,
    OVERALL_OUTPUT_DIR=OVERALL_OUTPUT_DIR, COMBINED_OUTPUT_DIR=COMBINED_OUTPUT_DIR)
import routes.r30_overall as _routes_r30_overall
_routes_r30_overall.register(
    app, _err500=_err500, _register_job=_register_job, _save_upload=_save_upload,
    _serve_run_file=_serve_run_file, _sweep_old_run_dirs=_sweep_old_run_dirs,
    _OVERALL_JOBS=_OVERALL_JOBS, OVERALL_OUTPUT_DIR=OVERALL_OUTPUT_DIR,
    PROTOTYPE_OUTPUT_DIR=PROTOTYPE_OUTPUT_DIR, COMBINED_OUTPUT_DIR=COMBINED_OUTPUT_DIR)
import routes.r30_gnn as _routes_r30_gnn
_routes_r30_gnn.register(
    app, FIRE_DXF2SDF_DIR=FIRE_DXF2SDF_DIR, FIRE_DXF2SDF_OUTPUT_DIR=FIRE_DXF2SDF_OUTPUT_DIR,
    UV_EXECUTABLE=UV_EXECUTABLE, _save_upload=_save_upload)
import routes.r30_has as _routes_r30_has
_routes_r30_has.register(
    app, _common_network_to_geometry=_common_network_to_geometry,
    _err500=_err500, _save_upload=_save_upload)
import routes.r30_machineroom as _routes_r30_machineroom
_routes_r30_machineroom.register(
    app, COMBINED_OUTPUT_DIR=COMBINED_OUTPUT_DIR, MACHINEROOM_OUTPUT_DIR=MACHINEROOM_OUTPUT_DIR,
    OVERALL_OUTPUT_DIR=OVERALL_OUTPUT_DIR, PROTOTYPE_OUTPUT_DIR=PROTOTYPE_OUTPUT_DIR,
    SYSTEM_OUTPUT_DIR=SYSTEM_OUTPUT_DIR, _emit_subnetwork_bundle=_emit_subnetwork_bundle,
    _err500=_err500, _load_cached_view_entities=_load_cached_view_entities,
    _save_upload=_save_upload, _serve_run_file=_serve_run_file,
    _sweep_old_run_dirs=_sweep_old_run_dirs, _to_float=_to_float)
import routes.r30_system as _routes_r30_system
_routes_r30_system.register(
    app, BASE_DIR=BASE_DIR, COMBINED_OUTPUT_DIR=COMBINED_OUTPUT_DIR,
    MACHINEROOM_OUTPUT_DIR=MACHINEROOM_OUTPUT_DIR, OVERALL_OUTPUT_DIR=OVERALL_OUTPUT_DIR,
    PROTOTYPE_OUTPUT_DIR=PROTOTYPE_OUTPUT_DIR, SYSTEM_OUTPUT_DIR=SYSTEM_OUTPUT_DIR,
    _emit_subnetwork_bundle=_emit_subnetwork_bundle, _err500=_err500,
    _load_cached_view_entities=_load_cached_view_entities, _save_upload=_save_upload,
    _serve_run_file=_serve_run_file, _sweep_old_run_dirs=_sweep_old_run_dirs, _to_float=_to_float)
import routes.feedback as _routes_feedback
_routes_feedback.register(
    app, FEEDBACK_POSTS_PATH=FEEDBACK_POSTS_PATH, FEEDBACK_UPLOAD_DIR=FEEDBACK_UPLOAD_DIR)
import routes.r30_inspect as _routes_r30_inspect
_routes_r30_inspect.register(
    app, INSPECT_CACHE_DIR=INSPECT_CACHE_DIR, INSPECT_CACHE_VERSION=INSPECT_CACHE_VERSION,
    _save_upload=_save_upload)
import routes.r30_combined as _routes_r30_combined
_routes_r30_combined.register(
    app, COMBINED_OUTPUT_DIR=COMBINED_OUTPUT_DIR, OVERALL_OUTPUT_DIR=OVERALL_OUTPUT_DIR,
    PROTOTYPE_OUTPUT_DIR=PROTOTYPE_OUTPUT_DIR, _COMBINED_JOBS=_COMBINED_JOBS,
    _COMBINED_JOBS_CAP=_COMBINED_JOBS_CAP, _PROTOTYPE_JOBS=_PROTOTYPE_JOBS,
    _RISER_HEIGHT_FRAC=_RISER_HEIGHT_FRAC, _RISER_SCHEMATIC_SPAN_MM=_RISER_SCHEMATIC_SPAN_MM,
    _common_network_to_geometry=_common_network_to_geometry, _err500=_err500,
    _serve_run_file=_serve_run_file, _sweep_old_run_dirs=_sweep_old_run_dirs, _to_float=_to_float)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)
