# -*- coding: utf-8 -*-
"""기기표 누락의 «크기» — 몇 m 가 빠져 있나.

수정 가능성 매트릭스(F-11f)에서 「기기표가 통째로 빈다」를 백로그 1순위로 뒀다.
고칠지 정하려면 **얼마나 빠졌는지** 를 알아야 한다. 등가길이가 몇 m 인데
그 망의 총연장이 얼마인지 — 비율이 작으면 미룰 수 있고, 크면 못 미룬다.

권위 레퍼런스(`assets/3-1형_…알람밸브.sdf`)가 어떻게 싣는지와, 우리 산출물이
무엇을 싣는지를 나란히 센다.

    python scripts/_probe_equipment_gap.py [우리.sdf]
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def scan(path: Path) -> dict:
    t = path.read_text(encoding="utf-8", errors="replace")
    eq = re.findall(r'<Equipment description="([^"]*)"'
                    r' equivalent-length="([^"]*)"', t)
    pipes = re.findall(r'<Pipe [^>]*length="([^"]*)"', t)
    total = sum(float(v) for v in pipes if v)
    by: dict = {}
    for desc, m in eq:
        d = by.setdefault(desc, {"n": 0, "m": 0.0})
        d["n"] += 1
        d["m"] += float(m or 0)
    return {"n_pipe": len(pipes), "pipe_m": total, "eq": by,
            "eq_m": sum(d["m"] for d in by.values()),
            "n_nozzle": len(re.findall(r"<Nozzle", t))}


def show(tag: str, path: Path) -> dict:
    if not path.is_file():
        print(f"\n■ {tag} — 파일 없음 {path}")
        return {}
    s = scan(path)
    print(f"\n■ {tag}")
    print(f"   {path.name}")
    print(f"   배관 {s['n_pipe']}개 · 배관 총연장 {s['pipe_m']:,.1f} m · "
          f"노즐 {s['n_nozzle']}")
    if not s["eq"]:
        print("   기기(Equipment) **0개** — 등가길이 0.0 m")
    for desc, d in sorted(s["eq"].items(), key=lambda kv: -kv[1]["m"]):
        print(f"   기기 {desc:<6} {d['n']:>3}개 · 등가길이 합 {d['m']:>8,.1f} m"
              f"  (개당 {d['m'] / max(1, d['n']):,.1f} m)")
    if s["pipe_m"]:
        print(f"   → 기기 등가길이 / 배관 총연장 = "
              f"{s['eq_m'] / s['pipe_m'] * 100:,.0f}%")
    return s


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ref = next(iter(sorted((ROOT / "assets").glob("*.sdf"))), None)
    if ref:
        r = show("권위 레퍼런스 (손으로 만든 것)", ref)
    ours = (Path(sys.argv[1]) if len(sys.argv) > 1
            else next(iter(sorted(
                (ROOT / "data" / "uploads" / "module_f").rglob("*대명동*.sdf"),
                key=lambda p: -p.stat().st_mtime)), None))
    o = show("모듈 F 산출물", ours) if ours else {}

    if r and o and o.get("n_nozzle"):
        # 레퍼런스의 «헤드 하나당 FX» 규약을 우리 망에 대 보면 얼마인가.
        fx = r["eq"].get("FX")
        if fx:
            per = fx["m"] / max(1, fx["n"])
            miss = per * o["n_nozzle"]
            av = r["eq"].get("A/V", {}).get("m", 0.0)
            print(f"\n■ 우리 망에 레퍼런스 규약을 대 보면")
            print(f"   헤드 {o['n_nozzle']}개 × FX {per:,.1f} m = "
                  f"{miss:,.1f} m")
            print(f"   + 알람밸브 {av:,.1f} m = **{miss + av:,.1f} m** 이 빠져 있다")
            print(f"   우리 배관 총연장 {o['pipe_m']:,.1f} m 의 "
                  f"**{(miss + av) / max(1e-9, o['pipe_m']) * 100:,.0f}%**")
            print("\n   ★등가길이는 마찰손실에 배관 길이와 «같은 자격» 으로")
            print("     들어간다. 이만큼이 빠지면 압력 계산이 그만큼 낙관적이 된다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
