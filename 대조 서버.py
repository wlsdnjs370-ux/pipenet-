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
INSPECT_CACHE_VERSION = "v4"


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
from cad_match import (CAD_SDF_LEARNING_PROFILE_PATH, PIPE_SEGMENTATION_MODEL_CANDIDATES, _AI_MATCH_MAX_EDGES, _pipe_segmentation_engine_status, _load_cad_sdf_learning_profile, _write_cad_sdf_learning_profile, _mark_similar_cad_pipe_entities, _ai_edge_features, _recompute_edge_degrees, _merge_collinear_cad_edges, _compact_cad_graph_for_sdf, _rasterize_edges_for_fft, _fft_shape_similarity, _component_similarity_stats)  # noqa: E501  (Phase2b core)

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
# 보안 하드닝 — fncadnet.com 외부 노출 시 기본 방어선
# ────────────────────────────────────────────────────────────────────────────
# (1) 업로드 크기 제한 — 수GB 파일 업로드 DoS 차단.
#     일반 DXF/SDF/KFP 는 수십 MB 이내. 200 MB 면 큰 통합 도면도 충분.
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

# (2) 세션 쿠키 하드닝
#     SECURE: HTTPS 만 전송 (cloudflared 가 TLS terminate → origin 까지는 HTTP
#       지만, X-Forwarded-Proto: https 를 ProxyFix 로 신뢰 시 Flask 가 https
#       세션 발급. 아래 ProxyFix 와 함께 동작.)
#     HTTPONLY: JS 접근 차단 — XSS 발생 시 쿠키 탈취 막음.
#     SAMESITE=Lax: 외부 사이트의 cross-site 요청에 쿠키 전송 안 함 (CSRF 완화).
app.config["SESSION_COOKIE_SECURE"] = True
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

# (3) ProxyFix — cloudflared 가 보내는 X-Forwarded-Proto/For 헤더를 신뢰해서
#     request.is_secure 가 https 로 인식되도록. 없으면 SESSION_COOKIE_SECURE=True
#     일 때 쿠키 자체가 안 발급되어 로그인이 작동 안 함.
from werkzeug.middleware.proxy_fix import ProxyFix as _ProxyFix
app.wsgi_app = _ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)


# 같은 서버가 https(cloudflared) 와 평문 http(localhost) 를 동시에 받는다.
# Secure 를 전역 고정하면 평문 접속에서 브라우저가 쿠키 저장을 거부해 로그인이
# 무한 반복된다. 요청 스킴별로 판단하면 https 는 Secure 를 그대로 유지한 채
# 로컬 http 만 통과한다 — 외부 노출 방어선 손실 없음.
from flask.sessions import SecureCookieSessionInterface as _SecureCookieSI


class _SchemeAwareSessionInterface(_SecureCookieSI):
    def get_cookie_secure(self, app):  # noqa: D102
        return bool(request.is_secure)


app.session_interface = _SchemeAwareSessionInterface()


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
# 한 줄 비밀번호 폼 → 세션 쿠키. SECRET_KEY·비밀번호 모두 env var(.env) 로 주입.
# 미설정 시 4자리 난수를 발급하고 서버 콘솔에만 찍는다 — 소스에 상수를 두면
# 저장소·배포 zip 을 읽은 사람이 곧바로 게이트를 통과한다.
# 로그인 폼이 pattern="[0-9]*" 라 난수도 숫자여야 브라우저 제출이 된다.
import secrets as _secrets
import os as _os_for_auth
app.secret_key = _os_for_auth.environ.get("FLASK_SECRET_KEY") or _secrets.token_hex(32)
LOGIN_PASSWORD = _os_for_auth.environ.get("LOGIN_PASSWORD")
if not LOGIN_PASSWORD:
    LOGIN_PASSWORD = f"{_secrets.randbelow(10000):04d}"
    print(f"[auth] LOGIN_PASSWORD 미설정 — 이번 기동 한정 임시 비밀번호: {LOGIN_PASSWORD}",
          flush=True)

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


@app.after_request
def _no_store_behind_gate(response):
    """게이트 뒤 응답은 캐시 금지.

    Cloudflare 엣지가 apex 루트를 HIT 로 물고 있어 로그인 후에만 보여야 할
    랜딩 페이지가 비로그인 방문자에게 그대로 서빙됐다. 오리진이 캐시 가능
    응답을 내주는 한 재발하므로 게이트 대상 경로 전체를 no-store 로 못박는다.
    """
    if not request.path.startswith(_AUTH_EXEMPT_PREFIXES):
        response.headers["Cache-Control"] = "no-store, private"
        response.headers.pop("Expires", None)
        response.headers.pop("Pragma", None)
    return response


# 인증 라우트(login/logout)는 routes/auth.py 로 분리 — 파일 끝에서 register().


# ────────────────────────────────────────────────────────────────────────────
# 전역 에러 핸들러 — /api/ 요청에 대해 HTML 500 페이지 대신 JSON 반환
# (클라이언트 fetch 가 await resp.json() 에서 SyntaxError 나는 것 차단)
# ────────────────────────────────────────────────────────────────────────────
    # pages(misc) 도메인 라우트 → routes/pages.py (register 로 등록)

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
    from cad_engine import locate_oda_exe
    exe = locate_oda_exe()
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
    tb_text = traceback.format_exc()
    app.logger.error("Route error (%s): %s\n%s", request.path, exc, tb_text)
    body = {"ok": False,
            "message": message if message is not None else str(exc)[:300]}
    # 외부 노출 환경 — traceback 본문은 서버 로그로만. EXPOSE_TRACEBACK=1 시에만 클라이언트 노출.
    if _os_for_auth.environ.get("EXPOSE_TRACEBACK") == "1":
        body["traceback"] = tb_text[-1500:]
    body.update(extra)
    return jsonify(body), 500






from dxf_geometry import (_to_float, _point_on_polyline, _cad_entity_points, _polyline_length, _normalize_layer_name, _entity_preview_row, _approx_arc_points, _bbox, _norm_point, _norm_xy, _segments_from_points, _edge_points, _edge_length, _edge_angle, _angle_delta, _graph_bbox_from_edges, _node_key)  # noqa: E501  (Phase2b core)


def _fig_to_data_url(fig, *, tight: bool = True) -> str:
    buf = BytesIO()
    save_kwargs = {"format": "png", "dpi": 140}
    if tight:
        save_kwargs["bbox_inches"] = "tight"
    fig.savefig(buf, **save_kwargs)
    plt.close(fig)
    buf.seek(0)
    return "data:image/png;base64," + base64.b64encode(buf.read()).decode("ascii")






















from sdf_analysis import (_sdf_parse_nodes, _sdf_parse_pipes_equipment, _sdf_parse_nozzles, _sdf_build_adjacency, _sdf_av_node, _sdf_dijkstra, _sdf_farthest_heads, _sdf_length_checks, _sdf_bore_reductions, _sdf_branch_nodes, _sdf_fitting_stats, _sdf_vertical_pipes, _sdf_graph_pipes, _analyze_sdf_sprinkler_network)  # noqa: E501  (Phase2b core)


















































































    # feedback 도메인 라우트 → routes/feedback.py (register 로 등록)

    # cad_compare 도메인 라우트 → routes/cad_compare.py (register 로 등록)

























# AI 그래프 매칭 가드 — pair 행렬/텐서가 O(N×M) 라 입력 edge 수를 상한으로 자른다.
# 실제 도면은 수백 edge 규모. 거대/악성 입력이 워커 스레드를 막거나 메모리를 터뜨리지
# 않도록 길이 상위 N 만 남긴다(매칭엔 긴 edge 가 더 중요).



















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
    from remote30_full_network import ProjectContext, emit_full_sdf
    out_dir.mkdir(parents=True, exist_ok=True)
    out_sdf = out_dir / f"{prefix}_{job_id}.sdf"
    emit_full_sdf(net, out_sdf, ctx=ProjectContext.titled(project_title))
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
    _save_upload=_save_upload, UPLOAD_DIR=UPLOAD_DIR)
import routes.r30_combined as _routes_r30_combined
_routes_r30_combined.register(
    app, COMBINED_OUTPUT_DIR=COMBINED_OUTPUT_DIR, OVERALL_OUTPUT_DIR=OVERALL_OUTPUT_DIR,
    PROTOTYPE_OUTPUT_DIR=PROTOTYPE_OUTPUT_DIR, _COMBINED_JOBS=_COMBINED_JOBS,
    _COMBINED_JOBS_CAP=_COMBINED_JOBS_CAP, _PROTOTYPE_JOBS=_PROTOTYPE_JOBS,
    _RISER_HEIGHT_FRAC=_RISER_HEIGHT_FRAC, _RISER_SCHEMATIC_SPAN_MM=_RISER_SCHEMATIC_SPAN_MM,
    _common_network_to_geometry=_common_network_to_geometry, _err500=_err500,
    _serve_run_file=_serve_run_file, _sweep_old_run_dirs=_sweep_old_run_dirs, _to_float=_to_float)
import routes.pages as _routes_pages
_routes_pages.register(
    app, _analyze_sdf_sprinkler_network=_analyze_sdf_sprinkler_network,
    DESIGN_AUTOMATION_PID_PATH=DESIGN_AUTOMATION_PID_PATH,
    DESIGN_AUTOMATION_PORT=DESIGN_AUTOMATION_PORT, DESIGN_AUTOMATION_ROOT=DESIGN_AUTOMATION_ROOT,
    DESIGN_AUTOMATION_SERVER_PATH=DESIGN_AUTOMATION_SERVER_PATH,
    DESIGN_AUTOMATION_STDERR_PATH=DESIGN_AUTOMATION_STDERR_PATH,
    DESIGN_AUTOMATION_STDOUT_PATH=DESIGN_AUTOMATION_STDOUT_PATH,
    EXPORT_SCHEMA=EXPORT_SCHEMA, REMOTE30_OUTPUT_DIR=REMOTE30_OUTPUT_DIR,
    UPDATE_HISTORY_PATH=UPDATE_HISTORY_PATH, UPLOAD_DIR=UPLOAD_DIR,
    _approx_arc_points=_approx_arc_points, _bbox=_bbox,
    _ensure_design_automation_static_layout=_ensure_design_automation_static_layout,
    _entity_preview_row=_entity_preview_row, _fig_to_data_url=_fig_to_data_url,
    _is_local_port_open=_is_local_port_open,
    _load_cad_sdf_learning_profile=_load_cad_sdf_learning_profile,
    _mark_similar_cad_pipe_entities=_mark_similar_cad_pipe_entities,
    _norm_point=_norm_point, _normalize_layer_name=_normalize_layer_name,
    _point_on_polyline=_point_on_polyline, _save_upload=_save_upload,
    _sdf_av_node=_sdf_av_node, _sdf_bore_reductions=_sdf_bore_reductions,
    _sdf_branch_nodes=_sdf_branch_nodes, _sdf_build_adjacency=_sdf_build_adjacency,
    _sdf_dijkstra=_sdf_dijkstra, _sdf_farthest_heads=_sdf_farthest_heads,
    _sdf_fitting_stats=_sdf_fitting_stats, _sdf_graph_pipes=_sdf_graph_pipes,
    _sdf_length_checks=_sdf_length_checks, _sdf_parse_nodes=_sdf_parse_nodes,
    _sdf_parse_nozzles=_sdf_parse_nozzles,
    _sdf_parse_pipes_equipment=_sdf_parse_pipes_equipment,
    _sdf_vertical_pipes=_sdf_vertical_pipes, _to_float=_to_float,
    _write_cad_sdf_learning_profile=_write_cad_sdf_learning_profile)
import routes.cad_compare as _routes_cad_compare
_routes_cad_compare.register(
    app, BASE_DIR=BASE_DIR, UPLOAD_DIR=UPLOAD_DIR, _AI_MATCH_MAX_EDGES=_AI_MATCH_MAX_EDGES,
    _ai_edge_features=_ai_edge_features, _compact_cad_graph_for_sdf=_compact_cad_graph_for_sdf,
    _component_similarity_stats=_component_similarity_stats, _edge_length=_edge_length,
    _recompute_edge_degrees=_recompute_edge_degrees, _save_upload=_save_upload)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)
