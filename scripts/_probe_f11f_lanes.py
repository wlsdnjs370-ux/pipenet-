# -*- coding: utf-8 -*-
"""[F-11f] 세 도면 완주 실측 — 자동 진입 → 원클릭 → 직접 입력 → 생존 → 산출.

지시서 F-11f 가 요구하는 한 바퀴를 도면마다 그대로 돈다:

    자동 진입(지배 띠)  정찰이 임계를 «스스로» 정하고 그 이유를 말하는가
    원클릭              급수 시작 위치 한 번으로 최불리가 서는가
    직접 입력 1건       부속(또는 등가길이) 한 자리를 사람이 채운다
    다시 계산           그 수정이 살아남는가 (F-11d 의 안정 키)
    산출                .sdf + .slf 가 실제로 써지는가

★사람 결정 «횟수» 를 센다. 지시서가 묻는 것은 인식률이 아니라 **완주 가능성**
  이기 때문이다(§0.1). LH306 이 질문 없이 완주하는지가 수용 기준에 있다.

★B1F 는 표 확정 1회가 23분이다(§21). 그래서 확정을 **두 번만** 돈다 —
  기준 1회 + 직접 입력 뒤 1회. 「생존」은 그 두 번으로 증명된다.

    python scripts/_probe_f11f_lanes.py [도면.dxf ...]
    → data/_f11f_lanes.json  (표로 옮길 원자료)
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = [
    ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf",
    ROOT / "samples" / "dxf" / "LH306동_평면도.dxf",
    # ★B1F 는 «같은 이름의 다른 파일» 이 둘 있다 — F-11a·F-11c 실측과 저장
    #   슬롯이 모두 uploads 쪽이므로 여기도 그것을 쓴다(§F-11c 커밋 참조).
    ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf",
]

_T0 = time.time()


def say(msg: str) -> None:
    print(f"  [{time.time() - _T0:7.1f}s] {msg}", flush=True)


def wait(c, sid, limit=30000, tag=""):
    t0, last = time.time(), 0.0
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        el = time.time() - t0
        if tag and el - last >= 60:
            last = el
            say(f"{tag} 진행 {el:.0f}초…")
        time.sleep(0.1)
    return {"state": "timeout"}


def build(c, sid, tag="표 확정"):
    t0 = time.time()
    r = c.post("/api/module-f/design/build", json={"sid": sid})
    d = r.get_json() or {}
    if not d.get("ok"):
        return {"state": "reject", "error": d.get("message") or r.status_code}, 0
    j = wait(c, sid, tag=tag)
    if j.get("state") == "done":
        res = (c.get(f"/api/module-f/convert/result?sid={sid}").get_json()
               or {}).get("result") or {}
        if res.get("error"):
            j = dict(j, state="result-error", error=res["error"])
    return j, time.time() - t0


def look(c, sid):
    d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
    t = d.get("tables") or {}
    un = t.get("unresolved") or {}
    meta = {k: v for k, v in (t.get("meta") or [])}
    return {"un": un, "meta": meta, "missed": d.get("ov_missed") or [],
            "pipes": (d.get("view") or {}).get("pipes") or []}


def run(c, dxf) -> dict:
    rec = {"file": dxf.name, "path": str(dxf),
           "mb": round(dxf.stat().st_size / 1048576, 1), "decisions": 0}
    t0 = time.time()
    with open(dxf, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                   content_type="multipart/form-data")
    jr = r.get_json() or {}
    if "sid" not in jr:
        rec["fail"] = f"열기 실패 — {jr.get('message')}"
        return rec
    sid = jr["sid"]
    wait(c, sid, tag="열기")
    rec["open_s"] = round(time.time() - t0, 1)

    # ── 자동 진입: 임계를 «서버가» 정한다. 사람 결정 0회.
    r2 = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
          .get("recon") or {})
    ad = r2.get("adopt") or {}
    rec.update(bands=r2.get("bands"), cands=r2.get("n"),
               rule=ad.get("rule"), conf_min=ad.get("conf_min"),
               adopt_n=ad.get("n"), why=ad.get("why"))
    say(f"{dxf.name} — 후보 {r2.get('n')} · 규칙 {ad.get('rule')} · "
        f"채택 예정 {ad.get('n')}")
    if ad.get("conf_min") is None:
        rec["fail"] = "규칙이 0 을 냈다 — 게이트가 찍기로 보낸다"
        return rec

    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    c.post("/api/module-f/pick/adopt",
           json={"sid": sid, "materials": True,
                 "heads": {"conf_min": ad["conf_min"]}})
    j = wait(c, sid, tag="채택")
    res = (c.get(f"/api/module-f/convert/result?sid={sid}").get_json()
           or {}).get("result") or {}
    rec["head_applied"] = res.get("head_applied")
    # ★`head_skipped` 는 경로에 따라 «목록» 이기도 «개수» 이기도 하다. 한쪽만
    #   가정하면 도면에 따라 프로브가 죽는다(대명동에서 실제로 그랬다).
    sk = res.get("head_skipped")
    rec["ghost"] = (len(sk) if isinstance(sk, (list, tuple))
                    else int(sk or 0))

    c.post("/api/module-f/pick/commit", json={"sid": sid})
    j = wait(c, sid, tag="조립")
    if j.get("state") != "done":
        rec["fail"] = f"조립 실패 — {j.get('error')}"
        return rec
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    rec["board"] = {"pts": st["counts"]["pts"], "edges": st["counts"]["edges"],
                    "heads": st["counts"]["heads"]}

    # ── 원클릭 — 사람 결정 1회(급수 시작 위치).
    heads = [(float(h[0]), float(h[1])) for h in (st["heads"] or ())]
    hs = heads[::max(1, len(heads) // 60)][:60]
    groups = sorted((g.get("segs") or [] for g in st["body_groups"]),
                    key=len, reverse=True)
    clicks = 0
    for s2 in groups[:18]:
        pts = [(float(s2[i]), float(s2[i + 1]))
               for i in range(0, len(s2) - 3, 4)]
        if len(pts) > 4000:
            pts = pts[::len(pts) // 4000]
        if not (pts and hs):
            continue
        p0, d0 = None, None
        for hx, hy in hs:
            p = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
            d = ((p[0] - hx) ** 2 + (p[1] - hy) ** 2) ** 0.5
            if d0 is None or d < d0:
                p0, d0 = p, d
        if p0 is None or d0 > 2000.0:
            continue
        clicks += 1
        rec["decisions"] += 1
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": p0[0], "y": p0[1]})
        wait(c, sid, tag="원클릭")
        w = (c.get(f"/api/module-f/edit/state?sid={sid}")
             .get_json()["state"].get("worst")) or {}
        if w.get("k"):
            rec["worst"] = {"k": w["k"], "far_m": w.get("far_m")}
            break
    rec["anchor_clicks"] = clicks
    if not rec.get("worst"):
        rec["fail"] = "원클릭으로 최불리를 못 세웠다"
        return rec
    say(f"  원클릭 {clicks}회 → 최불리 {rec['worst']['k']}개 · "
        f"최원 {rec['worst']['far_m']} m")

    # ── 표 확정 ①
    j, dt = build(c, sid)
    if j.get("state") != "done":
        rec["fail"] = f"표 확정 실패 — {j.get('state')} {j.get('error')}"
        return rec
    rec["build1_s"] = round(dt, 1)
    a = look(c, sid)
    rec["unresolved_kind"] = len(a["un"].get("kind_items") or [])
    rec["unresolved_pairs"] = len(a["un"].get("pairs") or [])
    rec["n_pipes"] = len(a["pipes"])
    say(f"  표 확정 {dt:.0f}초 · 배관 {rec['n_pipes']} · "
        f"부속 미해결 {rec['unresolved_kind']}")

    # ── 직접 입력 1건 — 사람 결정 1회. 부속이 없으면 관경으로 대신한다.
    if a["un"].get("kind_items"):
        it = a["un"]["kind_items"][0]
        c.post("/api/module-f/design/fitting-override", json={
            "sid": sid, "kind": [{"node": str(it["node"]),
                                  "pipe": str(it["pipe"]), "kind": "none",
                                  "note": "완주 실측 — 도면에서 확인"}]})
        rec["fill"] = f"부속 {it['pipe']}·노드 {it['node']} → 직선"
        rec["fill_kind"] = "부속"
    else:
        cand = next((p for p in a["pipes"] if p.get("ref")), None)
        if cand is None:
            rec["fail"] = "직접 입력할 자리가 없다"
            return rec
        pa, pb = cand["ref"]
        nd = 80 if int(cand["dia"]) != 80 else 100
        c.post("/api/module-f/design/bore-override", json={
            "sid": sid, "rows": [{"a": pa, "b": pb, "dia": nd,
                                  "note": "완주 실측 — 협의 변경"}]})
        rec["fill"] = f"관경 {cand['label']} {cand['dia']}→{nd}A"
        rec["fill_kind"] = "관경"
    rec["decisions"] += 1

    # ── 표 확정 ② — 생존 확인
    j, dt = build(c, sid, tag="재확정")
    if j.get("state") != "done":
        rec["fail"] = f"재확정 실패 — {j.get('state')} {j.get('error')}"
        return rec
    rec["build2_s"] = round(dt, 1)
    b = look(c, sid)
    rec["meta_fit"] = b["meta"].get("직접 입력 — 부속 판정")
    rec["meta_bore"] = b["meta"].get("직접 입력 — 관경")
    rec["missed"] = len(b["missed"])
    rec["survived"] = (str(rec["meta_fit"]) == "1"
                       if rec["fill_kind"] == "부속"
                       else str(rec["meta_bore"]) == "1")
    say(f"  직접 입력 {rec['fill']} → 살아남음 {rec['survived']} · "
        f"못 적용 {rec['missed']}")

    # ── 산출
    t1 = time.time()
    r3 = c.post("/api/module-f/design/emit", json={"sid": sid}).get_json() or {}
    rec["emit_ok"] = bool(r3.get("ok"))
    rec["emit_s"] = round(time.time() - t1, 1)
    if r3.get("ok"):
        rec["sdf"] = (r3.get("sdf") or {}).get("bytes")
        rec["slf_name"] = (r3.get("slf") or {}).get("name")
    else:
        rec["fail"] = f"저장 실패 — {r3.get('message')}"
    rec["total_s"] = round(time.time() - t0, 1)
    say(f"  산출 {'ok' if rec['emit_ok'] else '★실패'} · "
        f"SDF {rec.get('sdf')}B · 총 {rec['total_s']:.0f}초 · "
        f"사람 결정 {rec['decisions']}회")
    return rec


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    out = []
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        for f in [Path(x) for x in sys.argv[1:]] or DEFAULT:
            if not f.is_file():
                print(f"\n■ {f.name} — 파일 없음")
                out.append({"file": f.name, "fail": "파일 없음"})
                continue
            print(f"\n{'=' * 76}\n■ {f.name}\n{'=' * 76}", flush=True)
            try:
                out.append(run(c, f))
            except Exception as exc:  # noqa: BLE001 — 한 장이 죽어도 나머지는 잰다
                print(f"    ★실패 {type(exc).__name__}: {exc}")
                out.append({"file": f.name,
                            "fail": f"{type(exc).__name__}: {exc}"})
    dst = ROOT / "data" / "_f11f_lanes.json"
    dst.write_text(json.dumps(out, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    print(f"\n■ 원자료 → {dst}")
    print(f"\n{'도면':<26}{'완주':>6}{'사람결정':>9}{'총초':>8}{'확정1':>8}"
          f"{'확정2':>8}")
    for r in out:
        done = "✔" if r.get("emit_ok") and not r.get("fail") else "✗"
        print(f"{r['file'][:25]:<26}{done:>6}{r.get('decisions', 0):>9}"
              f"{r.get('total_s', 0):>8.0f}{r.get('build1_s', 0):>8.0f}"
              f"{r.get('build2_s', 0):>8.0f}")
        if r.get("fail"):
            print(f"    ★{r['fail']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
