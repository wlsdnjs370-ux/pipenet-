# -*- coding: utf-8 -*-
"""[F-11c 수용기준] 관경 «직접 입력» — 덮은 자리만 바뀌는가.

지시서 F-11c 의 수용 기준을 그대로 잰다:

    · 폴백(별표1) 배관 1개를 덮음 → 재확정 → 표·SDF 의 **그 배관만** 바뀌고
      나머지는 비트 동일           (D-F11-1 회귀 문법의 첫 적용)
    · 규칙 값(text) 배관도 덮이고, 원값·원출처가 남는다
    · SLF 에 없는 호칭경(77)은 저장 자체가 거절된다
    · 덮기 0건이면 산출물이 덮기 전과 비트 동일

★비교는 «파일 바이트» 로 한다. 표만 봐서는 emit 이 그 값을 실제로 실었는지
  알 수 없고, 그 둘이 갈리는 것이 이 프로젝트에서 실제로 있었던 일이다.

    python scripts/_probe_f11c_bore.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# ★B1F 는 «같은 이름의 다른 파일» 이 둘 있다. 이 둘을 섞으면 실측이 서로
#   비교가 안 된다 — 실제로 그랬다:
#
#     data/uploads/…dxf   116 MB  후보 3,338 · 채택 3,235 · 최원 416.85 m
#     samples/dxf/…dxf    126 MB  후보 6,688+          · 최원 851.35 m
#
#   F-11a 실측과 저장 슬롯(브라우저 검증·G 시험이 「이어서 열기」로 쓰는 그것)이
#   모두 uploads 쪽이므로 여기도 그것을 쓴다. 45분을 태우고서야 두 배 어긋난
#   수치를 보고 알았다 — 파일을 안 적으면 이 사고는 조용히 되풀이된다.
DEF = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


_T0 = time.time()


def say(msg: str) -> None:
    """★진행을 «시각과 함께» 흘려보낸다.

    B1F 한 바퀴가 수십 분이라, 끝나고서야 로그를 보면 어느 단계가 오래 걸렸는지
    알 길이 없다. 실제로 두 번 그랬다 — 45분을 태우고도 「표 확정 실패」 한 줄만
    남았다. 단계마다 시각을 남기고 즉시 flush 한다.
    """
    print(f"  [{time.time() - _T0:7.1f}s] {msg}", flush=True)


def wait(c, sid, limit=20000, tag=""):
    """잡 하나를 기다린다. 오래 걸리면 «살아 있다» 를 주기적으로 알린다."""
    t0 = time.time()
    last = 0.0
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            if tag and time.time() - t0 > 5:
                say(f"{tag} — {j.get('state')} ({time.time() - t0:.0f}초)")
            return j
        el = time.time() - t0
        if tag and el - last >= 30:      # 30초마다 한 줄
            last = el
            ln = (j.get("lines") or [])
            say(f"{tag} 진행 {el:.0f}초 · {j.get('phase')} · "
                + (str(ln[-1])[:80] if ln else ""))
        time.sleep(0.1)
    return {"state": "timeout"}


def build(c, sid, tag="표 확정"):
    r = c.post("/api/module-f/design/build", json={"sid": sid})
    d = r.get_json() or {}
    if not d.get("ok"):
        return {"state": "reject", "error": d.get("message") or r.status_code}
    j = wait(c, sid, tag=tag)
    # 잡이 done 이어도 «결과» 가 실패일 수 있다(엔진이 ok:False 로 답한다).
    if j.get("state") == "done":
        res = (c.get(f"/api/module-f/convert/result?sid={sid}").get_json()
               or {}).get("result") or {}
        if res.get("error"):
            j = dict(j, state="result-error", error=res["error"])
    return j


def why(j) -> str:
    """★실패를 «실패했다» 로만 적으면 안 된다 — 그 20분이 통째로 버려진다.

    잡의 사유와 마지막 로그 몇 줄을 같이 낸다. 이 프로브가 처음 B1F 를 돌았을
    때 「★표 확정 실패」 한 줄만 남기고 끝나, 왜인지 알려면 45분을 다시 써야
    했다. 진단 없는 실패 보고는 측정을 안 한 것과 같다.
    """
    out = f"{j.get('state')} — {j.get('error')}"
    for ln in (j.get("lines") or [])[-6:]:
        out += f"\n         {str(ln)[:110]}"
    return out


def emit(c, sid):
    """.sdf + .slf 를 쓰고 «바이트» 를 돌려준다."""
    t0 = time.time()
    r = c.post("/api/module-f/design/emit", json={"sid": sid})
    d = r.get_json() or {}
    say(f"저장 — {'ok' if d.get('ok') else d.get('message')} "
        f"({time.time() - t0:.0f}초)")
    if not d.get("ok"):
        return None, None, (d.get("message") or f"HTTP {r.status_code}")
    sdf = Path(_sess_path(sid, "design_sdf_path"))
    slf = Path(_sess_path(sid, "design_slf_path"))
    return sdf.read_bytes(), slf.read_bytes(), None


def _sess_path(sid, key):
    """세션에 적힌 산출 경로 — 다운로드 라우트를 안 거치고 파일을 직접 읽는다."""
    from routes.module_f.jobs import _sess
    return _sess(sid)[key]


def preview(c, sid):
    return c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}


def stand(c, sid):
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    j = wait(c, sid, tag="조립")
    if j.get("state") != "done":
        return j
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    heads = [(float(h[0]), float(h[1])) for h in (st["heads"] or ())]
    groups = sorted((g.get("segs") or [] for g in st["body_groups"]),
                    key=len, reverse=True)
    # ★표본만 쓴다. 원래는 헤드 전부 × 덩이의 점 전부를 돌았는데, B1F 는
    #   헤드 3,235개 × 큰 덩이의 점 수만 개 × 18덩이라 순수 파이썬으로 수십억
    #   번이 된다 — 실측으로 «표 확정이 느린 줄» 알았던 18분이 사실은 여기였다.
    #   앵커 후보는 «큰 배관 덩이 위의 아무 점» 이면 되므로 표본으로 충분하다.
    hs = heads[::max(1, len(heads) // 60)][:60]
    say(f"조립 완료 · 덩이 {len(groups)} · 헤드 {len(heads):,} "
        f"(앵커 탐색 표본 {len(hs)})")
    best_w = None
    for gi, s2 in enumerate(groups[:18]):
        pts = [(float(s2[i]), float(s2[i + 1]))
               for i in range(0, len(s2) - 3, 4)]
        if len(pts) > 4000:                 # 점도 표본으로 — 자리만 찾으면 된다
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
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": p0[0], "y": p0[1]})
        wait(c, sid, tag=f"원클릭 #{gi + 1}")
        st2 = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        w = st2.get("worst")
        say(f"원클릭 #{gi + 1} — 최불리 {(w or {}).get('k')}개 · "
            f"최원 {(w or {}).get('far_m')} m")
        if w and (best_w is None or int(w["k"]) > int(best_w["k"])):
            best_w = w
        if best_w and int(best_w["k"]) >= 30:
            break
    return ({"state": "done", "worst": best_w} if best_w
            else {"state": "no-anchor"})


def diff_lines(a: bytes, b: bytes):
    """다른 줄만 (번호, 앞, 뒤) 로. 길이가 달라지면 그것도 사실이다.

    ★«바뀐 자리» 를 잘라 낸다. 줄 앞부터 자르면 안 된다 — B1F 의 SDF 는 첫
      배관이 스케줄 블록과 한 줄에 붙어 나와서(1,287자), 앞 70자를 보여 주면
      정작 바뀐 `bore` 는 안 보이고 «다른 줄이 바뀐 것처럼» 읽힌다. 실제로 그
      한 줄 때문에 결과를 다시 파야 했다.
    """
    la = a.decode("utf-8", "replace").splitlines()
    lb = b.decode("utf-8", "replace").splitlines()
    out = []
    for i in range(max(len(la), len(lb))):
        x = la[i] if i < len(la) else "<없음>"
        y = lb[i] if i < len(lb) else "<없음>"
        if x == y:
            continue
        # 앞뒤로 같은 부분을 벗겨 «다른 구간» 만 남긴다.
        h = 0
        while h < min(len(x), len(y)) and x[h] == y[h]:
            h += 1
        t = 0
        while (t < min(len(x), len(y)) - h
               and x[len(x) - 1 - t] == y[len(y) - 1 - t]):
            t += 1
        lo = max(0, h - 30)
        out.append((i + 1,
                    ("…" if lo else "") + x[lo:len(x) - t + 30],
                    ("…" if lo else "") + y[lo:len(y) - t + 30]))
    return out


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEF
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    fails = []
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        t0 = time.time()
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid, tag="열기")
        # 어느 파일을 쟀는지 반드시 남긴다 — 이름만으로는 안 갈린다.
        print(f"\n■ {dxf} ({dxf.stat().st_size / 1048576:.0f} MB) "
              f"— 열기 {time.time() - t0:.0f}초", flush=True)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        ad = (rec.get("adopt") or {})
        print(f"    채택 규칙 {ad.get('rule')} · 임계 {ad.get('conf_min')} · "
              f"채택 예정 {ad.get('n')}")
        if ad.get("conf_min") is None:
            print("★규칙이 0 을 냈다 — 이 도면으로는 못 잰다.")
            return 1
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": ad.get("conf_min")}})
        wait(c, sid, tag="채택")
        j = stand(c, sid)
        if j.get("state") != "done":
            print(f"★조립·원클릭 실패 — {j}")
            return 1
        print(f"    최불리 {j['worst']['k']}개 · 최원 {j['worst']['far_m']} m")
        jb = build(c, sid)
        if jb.get("state") != "done":
            print(f"★표 확정 실패 — {why(jb)}")
            return 1

        # ── ① 덮기 0건 — 기준 산출물
        sdf0, slf0, err = emit(c, sid)
        if err:
            print(f"★저장 실패 — {err}")
            return 1
        d = preview(c, sid)
        meta0 = {k: v for k, v in (d["tables"]["meta"] or [])}
        pipes = d["view"]["pipes"]
        print(f"    ① 기준      배관 {len(pipes)} · SDF {len(sdf0):,}B · "
              f"SLF {len(slf0):,}B · meta「직접 입력 — 관경」"
              f"{meta0.get('직접 입력 — 관경')}")

        # ── ② 못 쓰는 호칭경은 그 자리에서 거절
        cand = next((p for p in pipes
                     if p.get("ref") and p.get("src") == "nfpc_fallback"), None)
        if cand is None:
            cand = next((p for p in pipes if p.get("ref")), None)
        if cand is None:
            print("★board 역참조가 있는 배관이 없다 — 이 도면으로는 못 잰다.")
            return 1
        a, b = cand["ref"]
        bad = c.post("/api/module-f/design/bore-override",
                     json={"sid": sid,
                           "rows": [{"a": a, "b": b, "dia": 77}]})
        bj = bad.get_json() or {}
        print(f"    ② 77A 거절  HTTP {bad.status_code} — {bj.get('message')}")
        if bad.status_code < 400 or bj.get("ok"):
            fails.append("규격표에 없는 77A 가 저장됐다")

        # ── ③ 폴백 배관 1개를 덮는다
        new_dia = 80 if int(cand["dia"]) != 80 else 100
        r = c.post("/api/module-f/design/bore-override", json={
            "sid": sid, "rows": [{"a": a, "b": b, "dia": new_dia,
                                  "note": "현장 실측"}]}).get_json() or {}
        print(f"    ③ 덮기      {cand['label']} (노드 {a}–{b}) "
              f"{cand['dia']}A[{cand['src']}] → {new_dia}A · "
              f"재확정 필요 {r.get('needs_rebuild')}")
        jb = build(c, sid)
        if jb.get("state") != "done":
            print(f"★재확정 실패 — {why(jb)}")
            return 1
        d1 = preview(c, sid)
        meta1 = {k: v for k, v in (d1["tables"]["meta"] or [])}
        ov = (d1["tables"].get("bore_overrides") or {})
        row = next((p for p in d1["view"]["pipes"]
                    if str(p["label"]) == str(cand["label"])), None)
        print(f"    ④ 재확정 후 그 배관 {row and row['dia']}A[{row and row['src']}]"
              f" · meta「직접 입력 — 관경」{meta1.get('직접 입력 — 관경')}")
        if not row or int(row["dia"]) != new_dia or row["src"] != "user":
            fails.append("덮은 배관이 안 바뀌었다")
        if str(meta1.get("직접 입력 — 관경")) != "1":
            fails.append(f"meta 가 1 이 아니다 ({meta1.get('직접 입력 — 관경')})")
        got = ov.get(str(cand["label"]))
        if not got:
            fails.append("원값 감사 기록이 없다")
        else:
            print(f"       └ 원값 {got['orig_dia']}A[{got['orig_src']}] · "
                  f"사유 「{got['note']}」")
            if int(got["orig_dia"]) != int(cand["dia"]):
                fails.append("원값이 안 맞는다")

        # ── ⑤ 산출물 — «그 배관만» 달라야 한다
        sdf1, slf1, err = emit(c, sid)
        if err:
            print(f"★저장 실패 — {err}")
            return 1
        dl = diff_lines(sdf0, sdf1)
        print(f"    ⑤ SDF 차이  {len(dl)}줄")
        for n, x, y in dl[:6]:
            print(f"       {n}: {x[:70]}\n          → {y[:70]}")
        if len(dl) > 6:
            print(f"       … 그 외 {len(dl) - 6}줄")
        if not dl:
            fails.append("SDF 가 하나도 안 바뀌었다 — 덮기가 파일에 안 실렸다")
        elif len(dl) > 2:
            fails.append(f"SDF 가 {len(dl)}줄이나 달라졌다 — 그 배관만이 아니다")
        if slf0 != slf1:
            print(f"    · SLF 도 달라졌다 ({len(slf0):,}→{len(slf1):,}B) — "
                  f"{new_dia}A 가 새로 쓰여 스케줄 줄이 붙은 것이면 정상이다")

        # ── ⑥ 규칙 값(text)도 덮인다 — 여기가 부속과 «범위» 가 갈리는 지점이다.
        tp = next((p for p in pipes if p.get("ref") and p.get("src") == "text"
                   and p["ref"] != [a, b]), None)
        if tp is None:
            print("    ⑥ text 근거 배관이 없어 «규칙 값 덮기» 는 못 쟀다")
        else:
            ta, tb = tp["ref"]
            td = 80 if int(tp["dia"]) != 80 else 100
            c.post("/api/module-f/design/bore-override", json={
                "sid": sid,
                "rows": [{"a": a, "b": b, "dia": new_dia, "note": "현장 실측"},
                         {"a": ta, "b": tb, "dia": td, "note": "협의 변경"}]})
            jb = build(c, sid)
            if jb.get("state") != "done":
                print(f"★재확정 실패 — {why(jb)}")
                return 1
            d2 = preview(c, sid)
            ov2 = (d2["tables"].get("bore_overrides") or {})
            g2 = ov2.get(str(tp["label"]))
            print(f"    ⑥ 규칙 값도 덮힘  {tp['label']} {tp['dia']}A[text] "
                  f"→ {g2 and g2['dia']}A · 원출처 {g2 and g2['orig_src']}")
            if not g2 or g2["orig_src"] != "text":
                fails.append("text 근거 배관을 못 덮었거나 원출처가 안 남았다")

        # ── ⑦ 되돌리면 비트 동일 (D-F11-1 회귀 문법)
        c.post("/api/module-f/design/bore-override",
               json={"sid": sid, "rows": []})
        jb = build(c, sid)
        if jb.get("state") != "done":
            print(f"★되돌린 뒤 재확정 실패 — {why(jb)}")
            return 1
        sdf2, slf2, err = emit(c, sid)
        same = (sdf2 == sdf0) and (slf2 == slf0)
        print(f"    ⑦ 되돌림    SDF {'비트 동일' if sdf2 == sdf0 else '★다름'}"
              f" · SLF {'비트 동일' if slf2 == slf0 else '★다름'}")
        if not same:
            fails.append("덮기 0건인데 산출물이 안 돌아왔다")

    print()
    if fails:
        for f in fails:
            print(f"  ★{f}")
        return 1
    print("  [OK] 덮은 자리만 바뀐다 · 원값이 남는다 · 되돌리면 비트 동일")
    return 0


if __name__ == "__main__":
    sys.exit(main())
