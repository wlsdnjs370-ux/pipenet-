# -*- coding: utf-8 -*-
"""[최불리규칙복원 §2-3 · 기준 5] 뽑힌 K 안에 못 붙는 헤드가 있으면 정말 막나.

지시서 §2-3. 건강한 도면에서는 이 문이 **안 열린다** — 대명동·B1F 실측에서
뽑힌 K 개는 전부 물이 닿는다(그래서 막을 일이 없다). 그러면 이 길은 한 번도
안 밟히고, 안 밟히는 길은 고장 나 있어도 모른다.

그래서 여기서는 전개 탐침이 「뽑힌 것 중 하나가 안 붙는다」고 답하도록
**바꿔 끼우고** 실제로 `/edit/worst` 를 불러 본다. 판정 코드는 그대로다 —
답만 갈아 끼운다.

    확인하는 것
      · 400 으로 막는가 (표를 만들지 않는가)
      · 그 헤드의 «자리 · 사유 · 할 일» 이 응답에 실리는가
      · 선정을 지웠는가 (안 지우면 옛 선정으로 표가 선다)
      · 다른 헤드로 채우지 않았는가

    python scripts/_probe_block_unattached.py
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from _probe_candidate_drop import _Null, setup                # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default="")
    ap.add_argument("--k", type=int, default=30)
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    ctx = _Null() if args.key else isolated_workdir(prefix="blk_")
    with ctx, srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        sid = setup(c, argparse.Namespace(key=args.key))

        from routes.module_f.jobs import _sess
        from routes.module_f import attach as att
        sess = _sess(sid)
        b = sess["edit"].board
        est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (est.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in est["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/mode",
                   json={"sid": sid, "mode": "급수시작위치"})
            c.post("/api/module-f/edit/click",
                   json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                         "y": (seg[1] + seg[3]) / 2, "max_d": 2000})

        body = {"sid": sid, "k": args.k}
        ok0 = c.post("/api/module-f/edit/worst", json=body).get_json()
        if not ok0.get("ok"):
            print(f"★먼저 정상으로 서야 한다 — {str(ok0)[:180]}")
            return 1
        picked = [int(h) for h in (sess["worst"].get("heads") or ())]
        victim = picked[0]                     # 1등을 «안 붙는다» 로 만든다
        print(f"\n■ §2-3 막음 확인 · {args.key or '대명동'} · K={args.k}")
        print(f"\n[정상]      {len(picked)}개 선정 — 지금은 막지 않는다")

        real = att.wet_heads

        def fake(sess_, es_, selected_source=None):
            out = dict(real(sess_, es_, selected_source=selected_source))
            wet = {int(i) for i in (out.get("wet") or ())}
            wet.discard(victim)
            out["wet"] = wet
            rs = dict(out.get("reason") or {})
            rs[victim] = "no_center"
            out["reason"] = rs
            return out

        att.wet_heads = fake
        try:
            jw = c.post("/api/module-f/edit/worst", json=body).get_json()
        finally:
            att.wet_heads = real

        blocked = jw.get("ok") is False
        na = jw.get("not_attached") or {}
        items = na.get("items") or []
        hit = next((it for it in items if int(it.get("disk", -1)) == victim),
                   None)
        print(f"[막음]      {'예' if blocked else '★아니오 — 그냥 지나갔다'}"
              f" · 막힌 헤드 {na.get('n', 0)}개")
        if hit:
            print(f"[자리·사유] disk {victim} ({hit['xy'][0]:.0f},"
                  f" {hit['xy'][1]:.0f}) · {hit.get('why')}")
            print(f"[할 일]     {hit.get('todo')}")
        else:
            print(f"[자리·사유] ★그 헤드가 응답에 없다 — {str(jw)[:200]}")
        kept = sess.get("worst")
        print(f"[선정]      {'지웠다' if not kept else '★남아 있다 — 옛 선정으로 표가 선다'}")
        print(f"[채움]      {'안 했다' if blocked else '★확인 불가'}"
              f" (막았으므로 채울 기회 자체가 없다)")
        good = blocked and hit is not None and not kept
        print("\n  " + ("★기준 5 가 선다" if good else "★★안 선다"))
        return 0 if good else 2


if __name__ == "__main__":
    raise SystemExit(main())
