# -*- coding: utf-8 -*-
"""모듈 F 정합성 검사 — «두 곳이 다른 말을 하는» 자리를 찾는다.

프로파일로 확인했다: 모듈 F 자신의 코드는 완주 시간의 **0.3%** 다(엔진이 67%).
그러니 여기서 할 일은 속도가 아니라 **모순 제거**다. 사람 눈으로는 4,900줄 +
4,170줄에서 모순을 못 찾는다. 기계가 찾을 수 있는 것부터 전부 찾는다:

    ① 화면이 부르는 엔드포인트 ↔ 서버가 등록한 엔드포인트
    ② JS 가 찾는 요소 id ↔ 템플릿에 있는 id      (없으면 조용히 null)
    ③ 템플릿의 id ↔ 아무도 안 쓰는 id            (죽은 마크업)
    ④ 같은 상수가 두 곳에 다른 값으로
    ⑤ 주석·도크스트링이 가리키는 이름이 실재하나
    ⑥ BLOCKED §N 참조가 실재하는 절인가

★②가 이 저장소에서 실제로 사고를 낸 종류다 — `$("없는id")` 는 예외를 안 내고
  null 을 주므로, 구문 검사도 통과하고 화면만 조용히 빈다.

    python scripts/_probe_f_consistency.py
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PY_DIR = ROOT / "routes" / "module_f"
JS = ROOT / "static" / "module_f.js"
HTML = ROOT / "templates" / "module_f.html"
BLOCKED = ROOT / "BLOCKED.md"

BAD = 0


def bad(msg: str) -> None:
    global BAD
    BAD += 1
    print(f"   ★{msg}")


def main() -> int:
    global BAD
    sys.stdout.reconfigure(errors="replace")
    js = JS.read_text(encoding="utf-8")
    html = HTML.read_text(encoding="utf-8")
    pys = {f.name: f.read_text(encoding="utf-8") for f in PY_DIR.glob("*.py")}
    allpy = "\n".join(pys.values())

    # ── ① 엔드포인트
    print("\n① 화면이 부르는 엔드포인트 ↔ 서버가 등록한 것")
    served = set(re.findall(r'@app\.(?:get|post|route)\("([^"]+)"', allpy))
    # 서버 전체(모듈 F 밖 포함)도 본다 — 공용 라우트를 부를 수 있다.
    for f in (ROOT / "routes").rglob("*.py"):
        served |= set(re.findall(r'@app\.(?:get|post|route)\("([^"]+)"',
                                 f.read_text(encoding="utf-8")))
    called = set()
    for m in re.finditer(r'["`](/api/module-f/[^"`?${]+)', js):
        called.add(m.group(1).rstrip("/"))

    def is_served(path: str) -> bool:
        """등록된 규칙과 맞나 — **경로 매개변수를 감안한다.**

        ★`/api/module-f/diagram/<key>` 를 화면은 `/diagram/${g.diagram}` 로
          부르므로, 문자열을 그대로 대면 「서버에 없다」로 잘못 짚는다.
          처음에 그렇게 만들어 멀쩡한 라우트를 모순으로 보고했다.
        """
        if path in served:
            return True
        for rule in served:
            base = rule.split("<")[0].rstrip("/")
            if "<" in rule and base and path.startswith(base):
                return True
        return False

    miss = sorted(c for c in called if not is_served(c))
    print(f"   화면이 부르는 {len(called)}개 · 서버 등록 "
          f"{len([s for s in served if 'module-f' in s])}개")
    for c in miss:
        bad(f"화면이 부르는데 서버에 없다 — {c}")
    if not miss:
        print("   [OK] 화면이 부르는 것이 전부 서버에 있다")

    # ── ②③ 요소 id
    print("\n② JS 가 찾는 id ↔ 템플릿에 있는 id")
    have = set(re.findall(r'id="([\w-]+)"', html))
    want = set(re.findall(r'\$\("([\w-]+)"\)', js))
    want |= set(re.findall(r'getElementById\("([\w-]+)"\)', js))
    ghost = sorted(want - have)
    print(f"   템플릿 id {len(have)}개 · JS 가 찾는 id {len(want)}개")
    for g in ghost:
        bad(f"JS 가 찾는데 템플릿에 없다 — #{g}  (null 이 조용히 돌아온다)")
    if not ghost:
        print("   [OK] JS 가 찾는 id 가 전부 템플릿에 있다")

    print("\n③ 템플릿에 있는데 아무도 안 쓰는 id")
    # ★«안 쓰인다» 를 말하려면 세 가지를 먼저 빼야 한다. 안 빼면 멀쩡한 마크업을
    #   지우게 된다 — 처음에 그렇게 만들어 9건 중 9건이 가짜였다:
    #     · 접이식은 `data-fold="…"` 로 잇고 JS 는 `dataset.fold` 로 읽는다
    #     · 이름을 조립해 부르는 것 — `mark("au-s1", …)` → `#au-s1-mark`
    #     · 항상 보이는 패널 — 단계 목록에 없는 것이 «맞다»(`#panel-log`)
    ALWAYS_ON = {"panel-log"}
    # 이름을 조립해 부르는 두 꼴을 다 본다:
    #     $(id + "-mark")        문자열 이어붙이기
    #     $(`${id}-mark`)        템플릿 리터럴   ← 실제로 쓰이는 쪽
    suffixes = set(re.findall(r'\$\(\s*\w+\s*\+\s*"([\w-]+)"\s*\)', js))
    suffixes |= set(re.findall(r'\$\(\s*`\$\{\w+\}([\w-]+)`\s*\)', js))
    prefixes = set(re.findall(r'"([\w-]+)"', js))
    made = {p + s for s in suffixes for p in prefixes}
    unused = sorted(i for i in have
                    if i not in want and i not in made and i not in ALWAYS_ON
                    and f'"{i}"' not in js
                    and f"'{i}'" not in js and f"#{i}" not in js
                    and f'"{i}"' not in allpy
                    and f'data-fold="{i}"' not in html
                    and f'for="{i}"' not in html
                    and f'aria-controls="{i}"' not in html)
    print(f"   {len(unused)}개" + (" — 죽은 마크업 후보" if unused
                                   else "  [OK]"))
    for i in unused[:20]:
        print(f"      #{i}")

    # ── ④ 같은 상수가 두 곳에
    print("\n④ 같은 이름의 상수가 여러 파일에 (값이 다르면 모순)")
    consts: dict[str, dict] = defaultdict(dict)
    for name, src in pys.items():
        for m in re.finditer(r"^([A-Z][A-Z0-9_]{3,})\s*=\s*([^\n#]+)", src,
                             re.M):
            consts[m.group(1)][name] = m.group(2).strip()
    dup = {k: v for k, v in consts.items() if len(v) > 1}
    for k, v in sorted(dup.items()):
        vals = set(v.values())
        if len(vals) > 1:
            bad(f"{k} 이 파일마다 다르다 — {v}")
        else:
            print(f"   (같은 값 중복) {k} — {sorted(v)}")
    if not dup:
        print("   [OK] 여러 파일에 겹치는 상수 없음")

    # ── ⑤ 주석이 가리키는 이름이 실재하나
    #
    # ★찾는 범위를 «저장소 전체» 로 둔다. 모듈 F 안에서만 찾으면 다른 모듈의
    #   이름(모듈 A 의 `override_flag` 등)을 「없다」고 잘못 짚는다 — 처음에
    #   그렇게 만들어 11건 중 여러 건이 가짜였다.
    print("\n⑤ 주석·도크스트링이 가리키는 이름이 저장소에 실재하나")
    src_all = allpy + js
    for d in ("routes", "core", "cad_project_editor_g", "static",
              "templates", "tests", "pipenet_converter"):
        p = ROOT / d
        if not p.is_dir():
            continue
        for f in p.rglob("*.py"):
            try:
                src_all += "\n" + f.read_text(encoding="utf-8")
            except OSError:
                pass
    for f in (ROOT / "remote30_prototype.py", ROOT / "kfp_sdf_converter.py"):
        if f.is_file():
            src_all += "\n" + f.read_text(encoding="utf-8")
    missing = []
    for m in re.finditer(r"`([a-z_][a-z0-9_]{4,})`", allpy):
        n = m.group(1)
        if n in ("module_f", "cad_import", "remote30_prototype"):
            continue
        if re.search(rf"def {re.escape(n)}\b", src_all):
            continue
        if re.search(rf"\b{re.escape(n)}\b", src_all.replace(f"`{n}`", "")):
            continue
        missing.append(n)
    for n in sorted(set(missing)):
        bad(f"주석이 가리키는 `{n}` 을 저장소에서 못 찾는다")
    if not missing:
        print("   [OK] 주석이 가리키는 이름이 전부 실재한다")

    # ── ⑥ BLOCKED §N
    print("\n⑥ 모듈 F 소스의 BLOCKED §N 참조")
    secs = set(re.findall(r"^## (\d+(?:-\d+)?)\.", BLOCKED.read_text(
        encoding="utf-8"), re.M))
    refs = set(re.findall(r"BLOCKED[^\n]{0,8}§(\d+)", allpy + js + html))
    for r in sorted(refs):
        if r not in secs:
            bad(f"BLOCKED §{r} 을 가리키는데 그런 절이 없다")
    print(f"   참조 {len(refs)}개 · 절 {len(secs)}개"
          + ("  [OK]" if all(r in secs for r in refs) else ""))

    print(f"\n{'=' * 60}")
    print(f"  모순 {BAD}건")
    return 0


if __name__ == "__main__":
    sys.exit(main())
