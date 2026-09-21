# -*- coding: utf-8 -*-
"""routes/module_f.py 를 성격별 패키지로 가른다. 본문은 원본에서 잘라 옮긴다."""
from __future__ import annotations

import io
import os
import re

SRC = "routes/module_f.py"
OUT = "routes/module_f"
lines = io.open(SRC, encoding="utf-8").read().split("\n")

# ── 최상위 def/class 의 구간을 잡는다 ────────────────────────────
starts = []
for i, ln in enumerate(lines):
    m = re.match(r"^(def |class )([A-Za-z_][A-Za-z0-9_]*)", ln)
    if m:
        starts.append((i, m.group(2)))
span = {}
for k, (i, name) in enumerate(starts):
    j = starts[k + 1][0] if k + 1 < len(starts) else len(lines)
    end = j
    while end - 1 > i and lines[end - 1].strip() == "":
        end -= 1
    span[name] = (i, end)


def grab(*names):
    """함수 본문 + 바로 위에 붙은 주석/구분선을 함께 떠 온다."""
    out = []
    for n in names:
        i, j = span[n]
        k = i
        while k - 1 >= 0 and lines[k - 1].lstrip().startswith("#"):
            k -= 1
        out.append("\n".join(lines[k:j]))
    return "\n\n\n".join(s.strip("\n") for s in out)


def const_block():
    a = next(i for i, ln in enumerate(lines) if ln.startswith("EDITOR_ROOT = "))
    b = span["_boot"][0]
    txt = "\n".join(lines[a:b]).rstrip()
    txt = re.sub(r"\n#\s*─+[^\n]*$", "", txt)
    return txt.rstrip()


def routes_between(first_route, last_route):
    """@app 라우트 구간 — register() 안 들여쓰기를 그대로 유지한다."""
    a = next(i for i, ln in enumerate(lines) if first_route in ln)
    while a - 1 >= 0 and lines[a - 1].lstrip().startswith("#"):
        a -= 1
    if last_route is None:
        b = len(lines)
    else:
        b = next(i for i, ln in enumerate(lines) if last_route in ln)
        while b - 1 >= 0 and lines[b - 1].lstrip().startswith("#"):
            b -= 1
    return "\n".join(lines[a:b]).rstrip()


os.makedirs(OUT, exist_ok=True)


def write(name, text):
    io.open(os.path.join(OUT, name), "w", encoding="utf-8",
            newline="\n").write(text.rstrip() + "\n")
    print("  %-16s %4d줄" % (name, len(text.splitlines())))


Q = '"' * 3

# ── common.py ───────────────────────────────────────────────────
consts = const_block().replace(
    'EDITOR_ROOT = Path(__file__).resolve().parent.parent / "cad_project_editor"',
    'EDITOR_ROOT = Path(__file__).resolve().parents[2] / "cad_project_editor"')
common_head = "\n".join([
    "# -*- coding: utf-8 -*-",
    Q + "모듈 F 공통 바탕 — 경로·상수·부팅·입력 검사.",
    "",
    "여기 있는 것은 «어느 단계에서나 참인 것» 뿐이다. 단계에 딸린 판단은",
    "`world`(찍기) · `graph`(손질) · `remote30`(범위) 이 각자 갖는다.",
    Q,
    "from __future__ import annotations",
    "",
    "import os",
    "import sys",
    "import threading",
    "from pathlib import Path",
    "",
    "from flask import jsonify",
    "",
    "",
])
fail_fn = "\n".join([
    "def _fail(msg, code=400):",
    "    " + Q + "실패 응답 한 벌. 라우트마다 다시 쓰지 않는다." + Q,
    '    return jsonify({"ok": False, "message": msg}), code',
])
write("common.py", common_head + consts + "\n\n\n" + fail_fn + "\n\n\n"
      + grab("_boot", "_r1", "_check_key", "_layer_category"))

# ── jobs.py ─────────────────────────────────────────────────────
jobs_head = "\n".join([
    "# -*- coding: utf-8 -*-",
    Q + "세션과 무거운 작업 — 진행 표시는 «실제로 찍힌 줄» 로만 한다." + Q,
    "from __future__ import annotations",
    "",
    "import sys",
    "import threading",
    "import time",
    "import traceback",
    "import uuid",
    "",
    "from routes.module_f.common import LOG_TAIL, SESSION_TTL_SECONDS",
    "",
    "_SESSIONS: dict[str, dict] = {}",
    "_SESSIONS_LOCK = threading.Lock()",
    "# 무거운 단계(도면 파싱·망 구성·평면 그래프)는 한 번에 하나만 돈다.",
    "# docs/import 캐시와 stdout 을 공유하므로 겹치면 로그가 섞이고 캐시가 깨진다.",
    "_HEAVY_LOCK = threading.Lock()",
    "",
    "",
])
write("jobs.py", jobs_head + grab("_Tee", "_sweep", "_new_session", "_sess",
                                  "_run_job", "_job_view", "_job_running"))

# ── world.py ────────────────────────────────────────────────────
world_head = "\n".join([
    "# -*- coding: utf-8 -*-",
    Q + "찍기 단계가 브라우저로 내려보내는 도면 — 직렬화와 상한만 다룬다." + Q,
    "from __future__ import annotations",
    "",
    "import json",
    "import os",
    "import time",
    "",
    "from routes.module_f.common import (",
    "    MAX_ARCS, MAX_CIRCLES, MAX_SEGS, _layer_category, _r1)",
    "",
    "",
])
write("world.py", world_head + grab("_world_payload", "_pts_bounds", "_saved_keys"))

# ── graph.py ────────────────────────────────────────────────────
graph_head = "\n".join([
    "# -*- coding: utf-8 -*-",
    Q + "손질 망의 셈 — 덩이 통계와 자동 이음.",
    "",
    "«A 처럼 재고, E 처럼 붙인다». 후보를 고르는 것만 여기서 하고, 실제로 붙일지는",
    "모듈 E 의 `board.join` 이 모양을 보고 정한다.",
    Q,
    "from __future__ import annotations",
    "",
    "import math",
    "",
    "from routes.module_f.common import (",
    "    AUTOJOIN_ANG_TOL_DEG, AUTOJOIN_LADDER_MM, AUTOJOIN_MAX_PAIRS,",
    "    AUTOJOIN_PLATEAU, _r1)",
    "",
    "",
])
write("graph.py", graph_head + grab("_body_index", "_body_stat",
                                    "_autojoin_scan", "_autojoin_apply"))

# ── remote30.py ─────────────────────────────────────────────────
r30_head = "\n".join([
    "# -*- coding: utf-8 -*-",
    Q + "모듈 A 에서 빌려온 것 — 최불리 K · 도면 장 나누기 · 범위 제한 · PIPENET." + Q,
    "from __future__ import annotations",
    "",
    "import heapq",
    "import math",
    "import os",
    "",
    "from routes.module_f.common import REMOTE_K_DEFAULT, _r1",
    "",
    "",
])
write("remote30.py", r30_head + grab("_worst_k_heads", "_worst_view",
                                     "_sheet_frames", "_restrict_to_worst",
                                     "_emit_pipenet"))

# ── views.py ────────────────────────────────────────────────────
views_head = "\n".join([
    "# -*- coding: utf-8 -*-",
    Q + "화면이 받는 상태 한 장.",
    "",
    "★망 도형은 «바뀌었을 때만» 싣는다. 무엇이 안 바뀌었는지는 `keep` 으로 알린다 —",
    "빈 배열만으로는 «비었다» 와 구별이 안 된다.",
    Q,
    "from __future__ import annotations",
    "",
    "from routes.module_f.common import _r1",
    "from routes.module_f.graph import _body_stat",
    "from routes.module_f.remote30 import _worst_view",
    "from routes.module_f.world import _pts_bounds",
    "",
    "",
])
write("views.py", views_head + grab("_pick_state", "_net_rev", "_edit_state",
                                    "_autojoin_view"))

print("\n[라우트]")


def api(name, doc, imports, extra, first, last):
    head = "\n".join([
        "# -*- coding: utf-8 -*-",
        Q + doc + Q,
        "from __future__ import annotations",
        "",
        imports,
        "",
        "",
        "def register(app" + extra + "):",
        "",
    ])
    write(name, head + routes_between(first, last))


api("api_open.py",
    "모듈 F 라우트 — 페이지·설명 그림·열기/이어열기·진행·도면.",
    "\n".join([
        "import os",
        "import time",
        "",
        "from flask import jsonify, render_template, request, send_file",
        "",
        "from routes.module_f.common import (",
        "    DIAGRAMS, IMPORT_WORK_ROOT, _boot, _check_key, _fail)",
        "from routes.module_f.jobs import _job_view, _new_session, _run_job, _sess",
        "from routes.module_f.remote30 import _sheet_frames",
        "from routes.module_f.views import _pick_state",
        "from routes.module_f.world import _saved_keys, _world_payload",
    ]),
    ", *, _save_upload",
    '@app.get("/module-f")', '@app.post("/api/module-f/pick/mode")')

api("api_pick.py",
    "모듈 F 라우트 — 1단계 찍기(재료·헤드).",
    "\n".join([
        "import time",
        "",
        "from flask import jsonify, request",
        "",
        "from routes.module_f.common import _fail",
        "from routes.module_f.jobs import _job_running, _run_job, _sess",
        "from routes.module_f.remote30 import _sheet_frames",
        "from routes.module_f.views import _pick_state",
    ]),
    "",
    '@app.post("/api/module-f/pick/mode")', '@app.get("/api/module-f/edit/state")')

api("api_edit.py",
    "모듈 F 라우트 — 2단계 손질(이음·삭제·급수·종류·자동 이음·최불리).",
    "\n".join([
        "import os",
        "import time",
        "",
        "from flask import jsonify, request",
        "",
        "from routes.module_f.common import REMOTE_K_DEFAULT, _fail, _r1",
        "from routes.module_f.graph import _autojoin_apply, _autojoin_scan",
        "from routes.module_f.jobs import _job_running, _run_job, _sess",
        "from routes.module_f.remote30 import _worst_k_heads",
        "from routes.module_f.views import _edit_state",
    ]),
    "",
    '@app.get("/api/module-f/edit/state")', '@app.get("/api/module-f/convert/fields")')

api("api_convert.py",
    "모듈 F 라우트 — 3단계 변환(.kfp/.sdf/.slf)과 내려받기.",
    "\n".join([
        "import os",
        "import zipfile",
        "from pathlib import Path",
        "",
        "from flask import jsonify, request, send_file",
        "",
        "from routes.module_f.common import GROUP_DIAGRAM, _boot, _fail",
        "from routes.module_f.jobs import _job_running, _job_view, _run_job, _sess",
        "from routes.module_f.remote30 import _emit_pipenet, _restrict_to_worst",
    ]),
    ", *, UPLOAD_DIR",
    '@app.get("/api/module-f/convert/fields")', None)

print("\n원본 routes/module_f.py 는 아직 그대로다 — __init__.py 를 쓴 뒤 지운다.")
