# -*- coding: utf-8 -*-
"""모듈 G — 모듈 E 를 복제한 일곱 번째 모듈. 라우트와 초기화면 카드를 붙인다."""
import io

# ── ① 상수: G 의 편집기 트리 ─────────────────────────────────────
p = "routes/pages.py"
s = io.open(p, encoding="utf-8").read()

old_const = '''CAD_EDITOR_ROOT = Path(__file__).resolve().parent.parent / "cad_project_editor"
CAD_EDITOR_MAIN = CAD_EDITOR_ROOT / "main.py"'''
new_const = '''CAD_EDITOR_ROOT = Path(__file__).resolve().parent.parent / "cad_project_editor"
CAD_EDITOR_MAIN = CAD_EDITOR_ROOT / "main.py"
# 모듈 G — E 의 편집기를 통째로 복제한 별개 트리. 편집기가 작업 폴더를 제
# 위치(_APP_ROOT/docs/import)에서 잡으므로, 트리가 다르면 캐시·찍은스펙도
# 저절로 갈라진다. 같은 트리를 두 번 띄우면 그 캐시를 두 프로세스가 함께
# 헤집게 되므로 «복제» 는 카드만이 아니라 소스까지여야 한다.
CAD_EDITOR_G_ROOT = Path(__file__).resolve().parent.parent / "cad_project_editor_g"
CAD_EDITOR_G_MAIN = CAD_EDITOR_G_ROOT / "main.py"'''
assert old_const in s, "편집기 상수를 못 찾음"
s = s.replace(old_const, new_const, 1)

old_proc = '_cad_editor_proc: dict = {"handle": None}'
new_proc = ('_cad_editor_proc: dict = {"handle": None}\n'
            '# G 는 제 핸들을 갖는다 — E 와 핸들을 나눠 쓰면 한쪽이 «이미 실행 중»\n'
            '# 으로 오판해 서로를 못 띄운다.\n'
            '_cad_editor_g_proc: dict = {"handle": None}')
assert old_proc in s, "프로세스 핸들을 못 찾음"
s = s.replace(old_proc, new_proc, 1)

# ── ② 라우트: E 라우트 바로 뒤에 G 를 붙인다 ──────────────────────
tail = '''        response = make_response(page)
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return response

    @app.get("/print-report/<path:filename>")'''
g_route = '''        response = make_response(page)
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return response

    @app.get("/module-g-cad-editor")
    def module_g_cad_editor():
        """모듈 G — 모듈 E 를 복제한 편집기. 트리도 프로세스도 E 와 따로다.

        E 와 동시에 띄울 수 있고, 찍은스펙·표시캐시는 각자의 트리 아래
        `docs/import` 에 쌓이므로 서로를 덮지 않는다.
        """
        if not CAD_EDITOR_G_MAIN.exists():
            return (
                "모듈 G 편집기 프로그램을 찾을 수 없습니다: "
                f"{html_lib.escape(str(CAD_EDITOR_G_MAIN))}",
                500,
            )
        proc = _cad_editor_g_proc.get("handle")
        already_running = proc is not None and proc.poll() is None
        launch_error = ""
        if not already_running:
            try:
                _cad_editor_g_proc["handle"] = subprocess.Popen(
                    [sys.executable, str(CAD_EDITOR_G_MAIN)],
                    cwd=str(CAD_EDITOR_G_ROOT),
                )
            except Exception as exc:
                launch_error = str(exc)

        if launch_error:
            status_line = (
                "<p class=\\"err\\">편집기를 실행하지 못했습니다: "
                f"{html_lib.escape(launch_error)}</p>"
            )
        elif already_running:
            status_line = "<p class=\\"ok\\">모듈 G 편집기가 이미 실행 중입니다. 서버 화면에서 창을 확인하세요.</p>"
        else:
            status_line = "<p class=\\"ok\\">모듈 G 편집기를 실행했습니다. 서버(이 PC) 화면에 창이 열립니다.</p>"

        page = f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <title>Module G · CAD 프로젝트 편집기 (사본)</title>
  <style>
    body {{ margin:0; min-height:100vh; display:flex; align-items:center; justify-content:center;
           background:#0a1228; color:#e5e7eb; font-family:"Malgun Gothic","맑은 고딕",sans-serif; }}
    .box {{ max-width:560px; padding:36px 40px; background:#111827; border:1px solid #1f2937;
            border-radius:16px; box-shadow:0 18px 60px rgba(0,0,0,.45); }}
    .chip {{ display:inline-block; font-size:12px; letter-spacing:.14em; font-weight:800; color:#93c5fd;
             border:1px solid #1d4ed8; border-radius:999px; padding:5px 12px; margin-bottom:14px; }}
    h1 {{ margin:0 0 10px; font-size:22px; }}
    p {{ margin:8px 0; line-height:1.6; font-size:14px; color:#cbd5e1; }}
    .ok {{ color:#86efac; font-weight:700; }}
    .err {{ color:#fca5a5; font-weight:700; }}
    .note {{ font-size:12.5px; color:#94a3b8; }}
    a.btn {{ display:inline-block; margin-top:18px; padding:10px 18px; background:#1d4ed8; color:#fff;
             font-weight:800; text-decoration:none; border-radius:10px; }}
  </style>
</head>
<body>
  <div class="box">
    <span class="chip">MODULE G</span>
    <h1>CAD 프로젝트 편집기 (사본)</h1>
    {status_line}
    <p class="note">모듈 E 를 그대로 복제한 편집기입니다. 소스 트리가 따로라 E 와 동시에 띄울 수 있고, 찍은 스펙·손질 결과도 서로 섞이지 않습니다.</p>
    <p class="note">데스크톱 프로그램이라 브라우저 안에는 표시되지 않고, 서버를 실행 중인 PC 화면에 창으로 열립니다.</p>
    <a class="btn" href="/" onclick="window.close();return false;">이 창 닫기</a>
  </div>
</body>
</html>"""
        response = make_response(page)
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return response

    @app.get("/print-report/<path:filename>")'''
assert tail in s, "E 라우트 끝을 못 찾음"
s = s.replace(tail, g_route, 1)
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("routes/pages.py — 상수·핸들·/module-g-cad-editor 추가")

# ── ③ 초기화면 일곱 번째 카드 ────────────────────────────────────
p2 = "templates/index.html"
t = io.open(p2, encoding="utf-8").read()
anchor = '''            <a class="home-module-btn" href="{{ url_for('module_f_page') }}">
              <span class="module-chip">Module F</span>'''
idx = t.index(anchor)
end = t.index("</a>", idx) + len("</a>")
card = '''
            <a class="home-module-btn" href="{{ url_for('module_g_cad_editor') }}" target="_blank" rel="noopener">
              <span class="module-chip">Module G</span>
              <strong>CAD Project Editor II (Desktop)</strong>
              <span>A full clone of the Module E desktop editor on its own source tree — run it side by side with E without their picked specs, caches or edits ever touching. Same flow: load a DXF, map heads, tidy the network, save a K-solver .kfp.</span>
            </a>'''
t = t[:end] + card + t[end:]
io.open(p2, "w", encoding="utf-8", newline="\n").write(t)
print("templates/index.html — Module G 카드(7번째) 추가")
