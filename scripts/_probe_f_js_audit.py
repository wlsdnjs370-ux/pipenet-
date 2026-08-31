# -*- coding: utf-8 -*-
"""module_f.js 감사 — 조용히 깨지는 종류만 골라 잰다.

4,170줄에 eslint 를 못 건다(설치에 네트워크가 필요하다). 대신 **이 코드에서
실제로 사고를 내는 꼴**만 직접 잰다. 공통점은 전부 «예외를 안 내고 조용히
틀린다» 는 것이다 — 구문 검사도 `node --check` 도 통과한다.

    ① 상태 필드 오타      `S.foo` 를 쓰기만 하고 아무도 안 읽거나 그 반대
    ② 핸들러 덮어쓰기      같은 요소에 `.onclick` 을 두 번 걸면 앞의 것이 죽는다
    ③ 선언 없는 전역       `foo = 1` (const/let 없이) — strict 면 던지지만
                          함수 안에서 조용히 새는 꼴이 있다
    ④ `var` 잔재          블록 스코프가 아니라 함수 스코프 — 루프에서 어긋난다
    ⑤ 느슨한 비교          `==` / `!=`
    ⑥ 같은 상수 두 번      값이 다르면 모순

    python scripts/_probe_f_js_audit.py
"""
from __future__ import annotations

import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
JS = ROOT / "static" / "module_f.js"

BAD = 0


def bad(msg: str) -> None:
    global BAD
    BAD += 1
    print(f"   ★{msg}")


def strip_strings(src: str) -> str:
    """문자열·주석을 지운다 — 그 안의 `==` 나 `S.x` 를 세면 안 된다.

    ★줄 수는 보존한다. 블록 주석을 통째로 지우면 그 뒤의 줄 번호가 전부
      밀려서, 지적한 자리가 엉뚱한 줄을 가리킨다(처음에 그래서 1461·1467 을
      가리켰는데 실제는 1463·1469 였다).
    """
    out = re.sub(r"//[^\n]*", "", src)
    out = re.sub(r"/\*.*?\*/",
                 lambda m: "\n" * m.group(0).count("\n"), out, flags=re.S)
    # ★여기서도 줄 수를 보존한다. 템플릿 리터럴은 여러 줄에 걸치는 일이 잦아,
    #   통째로 ``으로 바꾸면 그 뒤 줄 번호가 밀린다.
    def keep_lines(m):
        return '""' + "\n" * m.group(0).count("\n")

    out = re.sub(r'"(?:\\.|[^"\\])*"', keep_lines, out)
    out = re.sub(r"'(?:\\.|[^'\\])*'", keep_lines, out)
    out = re.sub(r"`(?:\\.|[^`\\])*`", keep_lines, out, flags=re.S)
    return out


def main() -> int:
    global BAD
    sys.stdout.reconfigure(errors="replace")
    raw = JS.read_text(encoding="utf-8")
    src = strip_strings(raw)
    print(f"■ {JS.name} — {len(raw.splitlines()):,}줄")

    # ── ① 상태 필드
    print("\n① 상태 필드 — 쓰기만 하거나 읽기만 하는 것")
    # ★읽기는 **원본** 에서 센다. 문자열을 지운 소스로 세면 템플릿 리터럴
    #   안의 `${S.boreSchedule}` 같은 읽기가 통째로 사라져, 멀쩡한 필드를
    #   「아무도 안 읽는다」고 짚는다(처음에 그래서 10건 중 여럿이 가짜였다).
    #   쓰기는 문자열 안에 있을 수 없으므로 지운 소스로 세도 된다.
    writes = Counter(re.findall(r"\bS\.(\w+)\s*=(?!=)", src))
    reads = Counter(re.findall(r"\bS\.(\w+)", re.sub(r"//[^\n]*", "", raw)))
    # 선언 자체도 `S.x = …` 이므로 읽기 수에서 쓰기 수를 뺀다.
    # ★«검증 내보내기» 블록의 필드는 화면이 안 읽는 것이 **정상**이다.
    #   코드에 표시를 두고 그 블록 안의 대입만 골라 빼낸다 — 목록에서
    #   손으로 지우면 다음에 하나 더 늘 때 또 시끄러워진다.
    exported: set = set()
    mk = raw.find("[검증 내보내기]")
    if mk >= 0:
        blk = raw[mk:raw.find("\n\n", mk)]
        exported = set(re.findall(r"\bS\.(\w+)\s*=", blk))
    only_w = sorted(k for k, n in writes.items()
                    if reads[k] - n <= 0 and k not in exported)
    only_r = sorted(k for k in reads if writes[k] == 0)
    # 초기 상태 리터럴 안에서 정의되는 것은 «쓰기» 로 친다.
    init = set(re.findall(r"^\s*(\w+):", src, re.M))
    only_r = [k for k in only_r if k not in init]
    print(f"   필드 {len(reads)}개")
    for k in only_w:
        bad(f"S.{k} — 쓰기만 하고 아무도 안 읽는다")
    for k in only_r:
        bad(f"S.{k} — 읽기만 하고 아무도 안 쓴다 (오타 의심)")
    if not only_w and not only_r:
        print("   [OK] 모든 필드가 쓰이고 읽힌다")

    # ── ② 핸들러 덮어쓰기
    print("\n② 같은 요소에 같은 핸들러를 두 번 거는가")
    h = defaultdict(list)
    for m in re.finditer(r'\$\("([\w-]+)"\)\.(on\w+)\s*=', raw):
        h[(m.group(1), m.group(2))].append(raw[:m.start()].count("\n") + 1)
    twice = {k: v for k, v in h.items() if len(v) > 1}
    for (eid, ev), locs in sorted(twice.items()):
        bad(f"#{eid}.{ev} 를 {len(locs)}번 건다 — 앞의 것이 죽는다 · 줄 {locs}")
    if not twice:
        print("   [OK] 겹쳐 거는 곳 없음")

    # ── ③ 선언 없는 대입
    print("\n③ 선언 없이 대입하는 이름 (전역 누수)")
    declared = set(re.findall(r"\b(?:const|let|var|function|class)\s+(\w+)",
                              src))
    declared |= set(re.findall(r"\bfunction\s*\w*\s*\(([^)]*)\)", src)[0].split(",")
                    if re.findall(r"\bfunction\s*\w*\s*\(([^)]*)\)", src) else [])
    leaks = []
    for m in re.finditer(r"^\s*(\w+)\s*=(?!=)", src, re.M):
        n = m.group(1)
        if n not in declared and n not in ("S", "window", "document"):
            leaks.append((n, src[:m.start()].count("\n") + 1))
    for n, ln in leaks[:10]:
        bad(f"선언 없이 대입 — {n} (줄 {ln})")
    if not leaks:
        print("   [OK] 없음")

    # ── ④⑤ var · 느슨한 비교
    print("\n④⑤ var 잔재 · 느슨한 비교")
    nvar = len(re.findall(r"\bvar\s+\w", src))
    # ★`x != null` 은 관용구다 — null 과 undefined 를 **한 번에** 거른다.
    #   `!==` 로 바꾸면 undefined 를 놓치므로 오히려 틀린 코드가 된다.
    #   그것까지 지적하면 사람이 목록을 통째로 무시하게 된다.
    loose = [(m.group(1), src[:m.start()].count("\n") + 1)
             for m in re.finditer(r"[^=!<>]([=!]=)(?!=)(?!\s*null\b)", src)]
    print(f"   var {nvar}개 · 느슨한 비교 {len(loose)}개")
    for op, ln in loose[:8]:
        bad(f"느슨한 비교 {op.strip()} (줄 {ln})")
    if nvar:
        bad(f"var 가 {nvar}개 남아 있다 — 블록 스코프가 아니다")

    # ── ⑥ 같은 상수 두 번
    print("\n⑥ 같은 이름의 상수를 두 번 선언")
    consts = defaultdict(list)
    for m in re.finditer(r"^\s*const\s+([A-Z][A-Z0-9_]{2,})\s*=\s*([^\n;]+)",
                         raw, re.M):
        consts[m.group(1)].append(m.group(2).strip())
    for k, v in sorted(consts.items()):
        if len(v) > 1:
            bad(f"{k} 을 {len(v)}번 선언 — {v}")
    if not any(len(v) > 1 for v in consts.values()):
        print(f"   [OK] 상수 {len(consts)}개, 중복 없음")

    print(f"\n{'=' * 58}\n  지적 {BAD}건")
    return 0


if __name__ == "__main__":
    sys.exit(main())
