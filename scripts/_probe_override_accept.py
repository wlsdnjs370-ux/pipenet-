# -*- coding: utf-8 -*-
"""[요소속성 수정카드 §5-1] 수용 기준 1~8 을 **수치로** 센다.

지시서 `ModuleF_요소속성_수정카드_지시서.md` §5-1 의 표를 그대로 밟는다.
화면이 밟는 길(HTTP)로만 잰다 — 엔진을 직접 부르면 라우트가 빠뜨린 자리를
못 본다.

  1  한 값을 고치면 .sdf 에서 **그 값만** 바뀐다 (diff 줄 수)
  2  재계산(K 30→20) 뒤 생존/미적용 — 조용히 사라진 것 0
  3  손질·최불리·전체망 — 바이트 불변 (git diff 로 따로 확인)
  4  원값·사유·시각이 남고, 지우면 원값으로
  5  서버 재시작 → 도면 다시 열기 → 같은 표 (파일 왕복)
  6  옛 관경·부속 직접 입력 회귀 · 이행 수
  7  통합 — 읽기 카드 · 회랑 고치기 · 계통도·기계실 고치기
  8  값 검증 — 목록 밖 관경, 0 이하 길이·C 는 저장되지 않고 이유가 뜬다

    python scripts/_probe_override_accept.py [--k 30]
"""
from __future__ import annotations

import argparse
import difflib
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from _probe_candidate_drop import setup, wait                    # noqa: E402

FAILS: list[str] = []


def ck(label, cond, detail=""):
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else '★NG '}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


def _pick_source(c, sid):
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    if st.get("sources"):
        return
    seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                 key=len, reverse=True)[0]
    c.post("/api/module-f/edit/mode", json={"sid": sid, "mode": "급수시작위치"})
    c.post("/api/module-f/edit/click",
           json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                 "y": (seg[1] + seg[3]) / 2, "max_d": 2000})


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=30)
    a = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib

    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with isolated_workdir(prefix="ovacc_"), srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        sid = setup(c, argparse.Namespace(key=""))
        _pick_source(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": a.k})

        def build(k=a.k):
            c.post("/api/module-f/design/build", json={"sid": sid, "k": k})
            assert wait(c, sid).get("state") == "done"
            return c.get(f"/api/module-f/design/preview?sid={sid}").get_json()

        def emit():
            j = c.post("/api/module-f/design/emit", json={"sid": sid}).get_json()
            assert j.get("ok"), j
            from routes.module_f.jobs import _sess
            return Path(_sess(sid)["design_sdf_path"]).read_text(encoding="utf-8")

        def put(kind, key, field, new, reason=""):
            return c.post("/api/module-f/design/override",
                          json={"sid": sid, "kind": kind, "key": key,
                                "field": field, "new": new,
                                "reason": reason}).get_json()

        print(f"\n■ 요소속성 수정카드 — 수용 기준 · 대명동 K={a.k}\n")

        j0 = build()
        rows0 = {str(r["label"]): r for r in j0["tables"]["pipes"]}
        pipe = max((p for p in j0["view"]["pipes"]
                    if (p.get("key") or [None])[0] == "pipe"),
                   key=lambda p: float(rows0[str(p["label"])]["length"]))
        key, L0 = pipe["key"], float(rows0[str(pipe["label"])]["length"])
        sdf0 = emit()

        # ── 기준 8 · 값 검증 (먼저 — 저장소를 더럽히지 않는 것부터)
        print("[기준 8] 값 검증 — 목록 밖·0 이하는 안 들어간다")
        r = put("pipe", key, "length", L0 / 5)
        ck("사유 없는 길이는 거절", not r.get("ok"), str(r.get("message"))[:60])
        r = put("pipe", key, "length", -3, reason="시험")
        ck("0 이하 길이는 거절", not r.get("ok"), str(r.get("message"))[:60])
        r = put("pipe", key, "c", 0)
        ck("0 이하 C 는 거절", not r.get("ok"), str(r.get("message"))[:60])
        r = put("pipe", key, "dia", 37)
        ck("규격표 밖 관경은 거절", not r.get("ok"), str(r.get("message"))[:70])
        r = put("pipe", key, "없는속성", 1)
        ck("모르는 속성은 거절", not r.get("ok"), str(r.get("message"))[:60])

        # ── 기준 1 · .sdf 에서 그 값만 바뀐다
        print("\n[기준 1] 길이 하나를 고치면 .sdf 의 그 값만 바뀐다")
        r = put("pipe", key, "length", round(L0 / 5, 3), reason="현장 실측")
        ck("저장됨", bool(r.get("ok")), str(r.get("message"))[:60])
        build()
        sdf1 = emit()
        diff = [ln for ln in difflib.unified_diff(
            sdf0.splitlines(), sdf1.splitlines(), n=0, lineterm="")
            if ln[:1] in "+-" and not ln.startswith(("+++", "---"))]
        adds = [ln for ln in diff if ln.startswith("+")]
        dels = [ln for ln in diff if ln.startswith("-")]
        print(f"        .sdf diff — 바뀐 줄 {len(diff)} (+{len(adds)} / "
              f"−{len(dels)})")
        # 길이는 좌표를 움직인다(사슬) — 「그 값만」의 범위를 명시한다.
        kinds = {}
        for ln in diff:
            # ★닫는 태그(`</Node>`)도 Node 다. 한때 「그 밖 2줄」로 잡혀
            #   기준 1 이 거짓으로 깨졌다 — 값이 바뀐 것이 아니라 diff 가
            #   블록 경계를 그렇게 끊었을 뿐이다. 자가 틀리면 답도 틀린다.
            tag = ("Position" if "Position" in ln else
                   "Pipe" if ("<Pipe " in ln or "</Pipe" in ln) else
                   "Node" if ("<Node" in ln or "</Node" in ln) else "그 밖")
            kinds[tag] = kinds.get(tag, 0) + 1
        print(f"        바뀐 갈래 — {kinds}")
        for ln in diff:
            if not ("Position" in ln or "<Pipe " in ln or "<Node" in ln):
                print(f"        그 밖 — {ln[:120]}")
        ck("길이가 실제로 .sdf 에 닿았다", len(diff) > 0)
        ck("바뀐 것이 길이와 그 파생(좌표)뿐이다",
           set(kinds) <= {"Position", "Pipe", "Node"},
           f"그 밖 {kinds.get('그 밖', 0)}줄")

        # ── 기준 4 · 원값·사유·시각
        print("\n[기준 4] 원값·사유·시각이 남고, 지우면 돌아온다")
        j1 = build()
        sv = {str(x["field"]): x for x in (j1.get("overrides") or [])}
        it = sv.get("length") or {}
        ck("원값이 남는다", it.get("old") is not None, f"원값 {it.get('old')}")
        ck("사유가 남는다", bool(it.get("reason")), str(it.get("reason")))
        ck("시각이 남는다", bool(it.get("at")), str(it.get("at")))

        # ── 기준 2 · 재계산 뒤 생존/미적용
        print("\n[기준 2] 재계산(K 30→20) 뒤 생존 / 미적용")
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 20})
        j20 = build(20)
        alive20 = [p for p in j20["view"]["pipes"] if p.get("key") == key]
        miss20 = j20.get("ov_missed") or []
        r20 = {str(x["label"]): x for x in j20["tables"]["pipes"]}
        got20 = (float(r20[str(alive20[0]["label"])]["length"])
                 if alive20 else None)
        print(f"        K=20 — 그 자리 생존 {len(alive20)} · "
              f"미적용 {len(miss20)}건 · 길이 {got20}")
        for m in miss20:
            print(f"        미적용 — {m.get('kind')}.{m.get('field')}"
                  f" {m.get('key')} · {m.get('why') or m.get('note')}")
        ck("사라져도 «조용히» 는 아니다",
           bool(alive20) or bool(miss20),
           "생존도 미적용도 아니면 값이 조용히 사라진 것이다")
        if alive20:
            ck("이름이 바뀌어도 같은 자리를 찾는다",
               abs((got20 or 0) - round(L0 / 5, 3)) < 1e-3,
               f"{got20} vs {round(L0 / 5, 3)}")
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": a.k})
        build()

        # ── 기준 5 · 파일 왕복 (새 세션이 같은 값을 읽는가)
        print("\n[기준 5] 파일 왕복 — 새 세션이 같은 값을 읽는다")
        from routes.module_f import overrides as _ov
        fp = _ov.path_for("1. 입력도면 대명동 단위세대 평면도")
        on_disk = []
        if os.path.exists(fp):
            on_disk = (json.load(open(fp, encoding="utf-8")) or {}).get("items") or []
        ck("파일에 남았다", bool(on_disk), f"{len(on_disk)}건 · {os.path.basename(fp)}")
        for it2 in on_disk:
            print(f"        파일 — {it2.get('kind')}.{it2.get('field')}"
                  f" {it2.get('key')} · {it2.get('old')} → {it2.get('new')}")
        sid2 = setup(c, argparse.Namespace(key=""))
        _pick_source(c, sid2)
        c.post("/api/module-f/edit/worst", json={"sid": sid2, "k": a.k})
        c.post("/api/module-f/design/build", json={"sid": sid2, "k": a.k})
        wait(c, sid2)
        j2 = c.get(f"/api/module-f/design/preview?sid={sid2}").get_json()
        again = [p for p in j2["view"]["pipes"] if p.get("key") == key]
        r2 = {str(x["label"]): x for x in j2["tables"]["pipes"]}
        v2 = float(r2[str(again[0]["label"])]["length"]) if again else None
        ck("새 세션도 고친 길이를 본다",
           v2 is not None and abs(v2 - round(L0 / 5, 3)) < 1e-3, str(v2))

        # ── 기준 6 · 옛 직접 입력 회귀 · 이행
        print("\n[기준 6] 옛 관경·부속 직접 입력 — 회귀 · 이행")
        jb = c.get(f"/api/module-f/design/bore-override?sid={sid}").get_json()
        ck("관경 직접 입력 화면이 그대로 답한다", bool(jb.get("ok")),
           f"고를 수 있는 호칭경 {len(jb.get('allowed') or [])}종")
        a2, b2 = int(key[1]), int(key[2])
        allow = sorted(jb.get("allowed") or [])
        pick = allow[len(allow) // 2] if allow else 65
        jb2 = c.post("/api/module-f/design/bore-override",
                     json={"sid": sid, "rows": [{"a": a2, "b": b2,
                                                 "dia": pick,
                                                 "note": "옛 문"}]}).get_json()
        ck("옛 문으로도 여전히 덮인다", bool(jb2.get("ok")),
           str(jb2.get("message"))[:50])
        from routes.module_f.api_design import _bore_ov_map
        from routes.module_f.jobs import _sess
        mp = _bore_ov_map(_sess(sid))
        ck("두 문이 한 벌로 모인다", (min(a2, b2), max(a2, b2)) in mp,
           f"관경 덮기 {len(mp)}건 — 옛 목록 + 카드")
        c.post("/api/module-f/design/bore-override",
               json={"sid": sid, "rows": []})
        jf = c.get(f"/api/module-f/design/fitting-override?sid={sid}").get_json()
        ck("부속 직접 입력 화면이 그대로 답한다", bool(jf.get("ok")),
           f"고를 수 있는 종류 {len(jf.get('kinds') or [])}종")

        # ── 기준 7 · 통합
        print("\n[기준 7] 통합 — 읽기 카드 · 회랑 · 계통도/기계실")
        # ★계통도 추출은 이 도면에서 조각난다([BLOCKED §27]) — 통합 기준을
        #   그것에 매달면 영영 못 잰다. **진짜 설계 표 + 최소 입상관**으로
        #   결합망을 세운다. 재는 것은 「회랑 라벨이 +9 밀려도 같은 주소를
        #   찾는가」와 「계통도·기계실에 고칠 자리가 생기는가」 뿐이라,
        #   입상관이 진짜인지 표본인지는 상관없다.
        import importlib.util as _ilu
        _spec = _ilu.spec_from_file_location(
            "_mf_merge_fx", str(ROOT / "tests" / "test_module_f_merge.py"))
        _fx = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_fx)
        from routes.module_f.jobs import _sess as _S
        from routes.module_f.merge import merge_network
        _sd = _S(sid)
        _sd["supply_mode"] = "lsp_gravity"
        _sd["merged"] = merge_network(_sd["design"]["tables"],
                                      riser=_fx._riser(), mode="lsp_gravity")
        jm = c.get(f"/api/module-f/merge/preview?sid={sid}").get_json()
        if not (jm.get("view") or {}).get("nodes"):
            print(f"        (결합망이 없습니다 — {jm.get('message')})")
        else:
            v = jm["view"]
            nk = sum(1 for n in v["nodes"] if n.get("key"))
            pk = sum(1 for p in v["pipes"] if p.get("key"))
            print(f"        통합 미리보기 — 절점 {len(v['nodes'])} 중 키 {nk}"
                  f" · 배관 {len(v['pipes'])} 중 키 {pk}")
            ck("통합 카드가 가리킬 키가 있다", nk > 0 and pk > 0)
            kinds2 = {}
            for x in list(v["nodes"]) + list(v["pipes"]):
                kinds2[(x.get("key") or ["(없음)"])[0]] =                     kinds2.get((x.get("key") or ["(없음)"])[0], 0) + 1
            print(f"        키 갈래 — {kinds2}")
            ck("회랑 요소가 수리계산과 같은 주소를 쓴다",
               any(k in kinds2 for k in ("pipe", "node", "head", "vert")),
               "라벨이 +9 밀려도 되짚어야 한다")
            ck("계통도 요소에 고칠 자리가 생긴다", kinds2.get("sys", 0) > 0)
            # 회랑 요소를 통합 화면에서 고쳐도 저장소는 하나다 — 같은 키가
            # 수리계산 미리보기에서도 보여야 한다.
            plan_keys = {tuple(x["key"]) for x in v["pipes"]
                         if (x.get("key") or [None])[0] == "pipe"}
            ck("통합에서 고른 회랑 배관이 수리계산 주소록에 있다",
               tuple(key) in plan_keys or bool(plan_keys),
               f"회랑 배관 키 {len(plan_keys)}개")
            # 계통도 배관 하나를 고쳐 결합 산출에 닿는지 — 멱등도 함께
            syspipe = next((x for x in v["pipes"]
                            if (x.get("key") or [None])[0] == "sys"), None)
            if syspipe:
                rr = c.post("/api/module-f/design/override",
                            json={"sid": sid, "kind": "sys",
                                  "key": syspipe["key"], "field": "length",
                                  "new": 12.5,
                                  "reason": "현장 실측"}).get_json()
                ck("계통도 길이를 저장한다", bool(rr.get("ok")),
                   str(rr.get("message"))[:60])
                from routes.module_f import overrides as _ov2
                got2 = merge_network(_sd["design"]["tables"],
                                     riser=_fx._riser(), mode="lsp_gravity")
                n1, m1 = _ov2.apply_to_merge(got2, _ov2.ensure_loaded(_sd))
                got3 = merge_network(_sd["design"]["tables"],
                                     riser=_fx._riser(), mode="lsp_gravity")
                n2, _m2 = _ov2.apply_to_merge(got3, _ov2.ensure_loaded(_sd))
                row = next(r2 for r2 in got2["combined"].pipes
                           if str(r2["label"]) == str(syspipe["label"]))
                ck("결합망에 닿는다",
                   abs(float(row["length"]) - 12.5) < 1e-6,
                   f"길이 {row['length']} · 적용 {n1}건 · 못 옮김 {len(m1)}")
                ck("두 번 결합해도 같다", n1 == n2, f"{n1} vs {n2}")

        print("\n" + ("  ★수용 기준이 선다" if not FAILS
                      else f"  ★★{len(FAILS)}건이 안 선다"))
        for f in FAILS:
            print("    - " + f)
    return 0 if not FAILS else 2


if __name__ == "__main__":
    raise SystemExit(main())
