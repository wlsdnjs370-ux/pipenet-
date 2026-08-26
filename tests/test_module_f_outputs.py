# -*- coding: utf-8 -*-
"""[F-4] 산출 3종 체크 — 8가지 조합에서 만들어지는 파일이 정확히 그것뿐인가.

    전체망 .kfp / 최불리 .kfp(_최불리K<n>) / 최불리 .sdf(+.slf)

덤으로 규약 둘: 하나도 안 고르면 400 · 최불리 계열인데 선정이 아직이면
막지 않고 worst_required 로 수리계산 패널을 안내한다.

    python tests/test_module_f_outputs.py
"""
from __future__ import annotations

import importlib.util
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KEY = "B1F 현장조사 소화설비 평면도"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def _wait(c, sid, limit=900):
    for _ in range(limit):
        jb = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if jb.get("state") in ("done", "error", "idle"):
            return jb
        time.sleep(0.3)
    return jb


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    spec = importlib.util.spec_from_file_location(
        "daejo", os.path.join(str(ROOT), "대조 서버.py"))
    srv = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(srv)
    app = srv.app
    app.config["TESTING"] = True
    out_root = ROOT / "data" / "uploads" / "module_f"

    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True

        r = c.post("/api/module-f/reopen", json={"key": KEY})
        sid = r.get_json()["sid"]
        if not check("B1F reopen", _wait(c, sid).get("state") == "done"):
            return 1

        def run(outputs):
            r = c.post("/api/module-f/convert/run",
                       json={"sid": sid, "dto": {}, "outputs": outputs})
            return r

        print("\n[규약] 빈 선택과 선정 이전")
        r = run({})
        check("하나도 안 고르면 400", r.status_code == 400,
              f"HTTP {r.status_code} · {str(r.get_json())[:60]}")
        r = run({"worst_kfp": True})
        j = r.get_json()
        check("최불리 선정 전 — 막지 않고 안내", j.get("code") == "worst_required",
              str(j)[:80])

        # 최불리 선정 + 설계 표 확정 (한 번만 — 이후 8조합이 공유)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 30})
        r = run({"worst_sdf": True})
        j = r.get_json()
        check("표 확정 전 SDF — 막지 않고 안내", j.get("code") == "worst_required",
              str(j)[:80])
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
        if not check("design build", _wait(c, sid).get("state") == "done"):
            return 1

        print("\n[8조합] 만들어지는 파일이 정확히 그것뿐인가")

        def files_for_sid():
            got = set()
            full = out_root / f"{sid}.kfp"
            if full.is_file():
                got.add("full_kfp")
            for p in out_root.glob(f"{sid}_최불리K*.kfp"):
                got.add("worst_kfp")
            d = out_root / f"{sid}_design"
            if any(d.glob("*.sdf")) if d.is_dir() else False:
                got.add("worst_sdf")
            return got

        def clear_files():
            full = out_root / f"{sid}.kfp"
            if full.is_file():
                full.unlink()
            for p in out_root.glob(f"{sid}_최불리K*.kfp"):
                p.unlink()
            d = out_root / f"{sid}_design"
            if d.is_dir():
                for p in d.iterdir():
                    p.unlink()

        for mask in range(1, 8):
            want = {"full_kfp": bool(mask & 1),
                    "worst_kfp": bool(mask & 2),
                    "worst_sdf": bool(mask & 4)}
            clear_files()
            r = run(want)
            if not r.get_json().get("ok"):
                check(f"조합 {mask:03b} 수락", False, str(r.get_json())[:70])
                continue
            jb = _wait(c, sid)
            res = c.get(f"/api/module-f/convert/result?sid={sid}"
                        ).get_json()["result"]
            made = files_for_sid()
            expect = {k for k, v in want.items() if v}
            check(f"조합 {mask:03b} → {sorted(expect)}",
                  jb.get("state") == "done" and res.get("ok")
                  and made == expect,
                  f"실제 {sorted(made)}")

        # 최불리 .kfp 파일명 규약
        wk = sorted(out_root.glob(f"{sid}_최불리K*.kfp"))
        check("최불리 파일명에 K 가 박힌다", bool(wk)
              and "최불리K30" in wk[-1].name, wk[-1].name if wk else "없음")

        # .sdf 만 고른 흐름 — .kfp 가 하나도 안 생긴다(«저장 대화 없음» 의 웹 판).
        clear_files()
        run({"worst_sdf": True})
        _wait(c, sid)
        made = files_for_sid()
        check(".sdf 만 골라도 .kfp 는 안 생긴다", made == {"worst_sdf"},
              str(sorted(made)))

        # 옛 호출부(remote_only) 하위호환
        clear_files()
        r = c.post("/api/module-f/convert/run",
                   json={"sid": sid, "dto": {}, "remote_only": True})
        _wait(c, sid)
        check("옛 remote_only=True → 최불리 .kfp 만",
              files_for_sid() == {"worst_kfp"}, str(sorted(files_for_sid())))

        # 선택 기억
        j = c.get(f"/api/module-f/convert/result?sid={sid}").get_json()
        # (세션 저장 확인은 서버 내부 값이라 결과 요약의 outputs 로 갈음)
        outs = ((j.get("result") or {}).get("summary") or {}).get("outputs")
        check("마지막 선택이 요약에 남는다", outs == {
            "full_kfp": False, "worst_kfp": True, "worst_sdf": False},
            str(outs))

    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("F-4 산출 3종 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
