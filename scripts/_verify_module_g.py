# -*- coding: utf-8 -*-
"""모듈 G 검증 — 카드·라우트·«E 와 정말 따로 도는가».

G 는 E 의 복제본이다. 확인해야 할 것은 «떴는가» 가 아니라 **E 와 갈라져
있는가** 다 — 트리가 같으면 두 프로세스가 한 캐시를 헤집는다.

띄운 창은 검사 끝에 반드시 닫는다(사람 화면에 Qt 창을 남기지 않는다).
"""
from __future__ import annotations

import importlib.util
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

spec = importlib.util.spec_from_file_location("daejo", os.path.join(ROOT, "대조 서버.py"))
srv = importlib.util.module_from_spec(spec)
sys.modules["daejo"] = srv
spec.loader.exec_module(srv)
app = srv.app
app.config["TESTING"] = True

FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def main() -> int:
    import routes.pages as pages

    print("\n[1] 복제된 소스 트리")
    g_root, e_root = pages.CAD_EDITOR_G_ROOT, pages.CAD_EDITOR_ROOT
    check("G 트리 존재", pages.CAD_EDITOR_G_MAIN.exists(), str(g_root.name))
    check("E 와 다른 트리", str(g_root) != str(e_root),
          f"{e_root.name} ≠ {g_root.name}")
    def rels(root):
        out = set()
        for r, _d, fs in os.walk(root):
            if "docs" in r or "__pycache__" in r:
                continue
            for f in fs:
                if f.endswith(".py"):
                    out.add(os.path.relpath(os.path.join(r, f), root))
        return out

    g_files, e_files = rels(g_root), rels(e_root)
    # ★«개수 동일» 을 요구하면 안 된다. G 는 복제본이지 사본이 아니라서, 제 몫의
    #   기능(수리계산 입력 design/·창·검사)이 붙으면 개수가 커진다. 확인해야 할
    #   것은 «E 의 것이 하나도 빠지지 않았나» 다.
    missing = sorted(e_files - g_files)
    check("E 의 소스가 하나도 빠지지 않음", not missing,
          f"G {len(g_files)}개 ⊇ E {len(e_files)}개"
          + (f" · 빠짐 {missing[:3]}" if missing else ""))
    added = sorted(g_files - e_files)
    if added:
        print(f"      G 고유 {len(added)}개 (이번 작업 산출): "
              f"{', '.join(a.replace(os.sep, '/') for a in added[:3])} …")
    # 편집기는 작업 폴더를 제 위치에서 잡는다 → 트리가 다르면 캐시도 갈라진다.
    check("작업 폴더가 갈라짐",
          not (g_root / "docs").exists() or
          str(g_root / "docs") != str(e_root / "docs"),
          "G 는 제 docs/import 를 새로 만든다")

    print("\n[2] 초기화면 카드 7장")
    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        r = c.get("/")
        html = r.get_data(as_text=True)
        n_cards = html.count('class="module-chip">Module ')
        check("카드 7장", r.status_code == 200 and n_cards == 7, f"{n_cards}장")
        check("Module G 카드", 'class="module-chip">Module G</span>' in html)
        check("G 카드가 G 라우트로 간다", "/module-g-cad-editor" in html)
        check("E 카드는 그대로", "/module-e-cad-editor" in html)

        print("\n[3] G 라우트 — 실제로 띄우고, 띄운 창은 닫는다")
        before = pages._cad_editor_g_proc.get("handle")
        check("띄우기 전 핸들 없음", before is None or before.poll() is not None)
        r = c.get("/module-g-cad-editor")
        body = r.get_data(as_text=True)
        check("G 페이지 200", r.status_code == 200, f"HTTP {r.status_code}")
        check("MODULE G 표기", "MODULE G" in body)
        proc = pages._cad_editor_g_proc.get("handle")
        launched = proc is not None and proc.poll() is None
        check("편집기 프로세스 기동", launched,
              f"pid {getattr(proc, 'pid', None)}")
        # ★E 의 핸들은 건드리지 않았어야 한다 — 핸들을 나눠 쓰면 서로를 막는다.
        check("E 핸들 무영향", pages._cad_editor_proc.get("handle") is None,
              str(pages._cad_editor_proc.get("handle")))
        if launched:
            time.sleep(1.0)
            proc.terminate()
            try:
                proc.wait(timeout=10)
            except Exception:  # noqa: BLE001
                proc.kill()
            print(f"      띄운 창을 닫았습니다 (pid {proc.pid})")

    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("모듈 G 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
