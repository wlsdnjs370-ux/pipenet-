# -*- coding: utf-8 -*-
"""[H-0] module_f.html 의 JS 를 구문 + **스코프**로 검증한다.

`node --check` 만으로 성공을 선언하면 안 된다(실측 회귀): 구문이 맞아도 다른
스코프의 함수-지역 헬퍼를 부르면 런타임에 ReferenceError 로 죽는다. 그래서
두 단계로 본다.

  ① 구문   — <script> 본문을 떼어 `node --check`
  ② 스코프 — 검사 대상 함수가 부르는 이름이 **같은 최상위 스코프**에 선언돼
             있는가. 선언 위치를 들여쓰기로 판별한다(이 템플릿의 최상위는 2칸).

    python scripts/_verify_module_f_js.py
"""
from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TPL = ROOT / "templates" / "module_f.html"

# [H-0] 슬롯 화면이 부르는 이름들. 하나라도 다른 스코프에 있으면 클릭하는 순간
# ReferenceError 다 — 서버 테스트로는 절대 안 잡힌다.
TARGETS = ("renderSlots", "loadSlots", "switchSlot", "renderBoreLegend",
           "drawDesign", "renderSubPanel", "renderSubPicks", "subClick",
           "armSub", "subExtract", "renderSubSummary", "drawSubPicks",
           "loadSub", "loadWorldRaw", "subSpec",
           "loadMergeModes", "setMergeMode", "loadMergeState",
           "renderMergeSummary", "loadMerge",
           "drawZones", "renderZones",
           "renderSteps", "gotoStage", "stageFlow", "stageReachable",
           "setStage")
CALLEES = ("api", "post", "busy", "say", "setStage", "loadEdit", "loadWorld",
           "draw", "$", "kv", "sx", "sy", "fit", "buildLayers", "renderCats")

FAILS: list[str] = []


def check(label: str, cond: bool, detail: str = "") -> bool:
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


def _script_body(html: str) -> str:
    """가장 긴 <script> 본문 — 이 템플릿의 앱 코드다."""
    bodies = re.findall(r"<script[^>]*>(.*?)</script>", html, re.S)
    if not bodies:
        raise SystemExit("module_f.html 에 <script> 가 없습니다.")
    return max(bodies, key=len)


def _top_level_names(body: str) -> set[str]:
    """최상위(2칸 들여쓰기) 선언 이름. 중첩 함수 안의 것은 세지 않는다."""
    names: set[str] = set()
    pat = re.compile(
        r"^  (?:async\s+)?function\s+([A-Za-z_$][\w$]*)"
        r"|^  (?:const|let|var)\s+([A-Za-z_$][\w$]*)")
    for line in body.splitlines():
        m = pat.match(line)
        if m:
            names.add(m.group(1) or m.group(2))
    return names


def _body_of(body: str, name: str) -> str:
    """최상위 함수 하나의 본문 — 닫는 중괄호까지.

    ★«다음 최상위 선언까지» 로 끊으면 안 된다. 이 템플릿의 최상위에는 선언이
      아닌 것도 온다(`$("x").onclick = async () => {…}`). 그것을 만나지 못한
      채 흘러가면 **다음 블록을 함께 삼켜** 남의 이름을 이 함수 것으로 보고한다
      (실측: renderMergeSummary 가 뒤따르는 onclick 의 `async`·`of` 를 물었다).

      이 파일의 최상위 함수는 정확히 2칸 들여쓴 `}` 로 닫힌다 — 그것을 끝으로 본다.
    """
    lines = body.splitlines()
    start = None
    pat = re.compile(rf"^  (?:async\s+)?function\s+{re.escape(name)}\s*\(")
    for i, line in enumerate(lines):
        if start is None:
            if pat.match(line):
                start = i
            continue
        if line == "  }":
            return "\n".join(lines[start:i + 1])
    return "\n".join(lines[start:]) if start is not None else ""


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    html = TPL.read_text(encoding="utf-8")
    body = _script_body(html)
    print(f"module_f.html · <script> {len(body.splitlines())} 줄")

    # ① 구문
    node = shutil.which("node")
    if node is None:
        check("node --check", False, "node 를 찾을 수 없습니다 (구문 검사 생략됨)")
    else:
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False,
                                         encoding="utf-8") as fh:
            fh.write(body)
            tmp = fh.name
        try:
            r = subprocess.run([node, "--check", tmp], capture_output=True,
                               text=True, encoding="utf-8", errors="replace")
            check("node --check 구문", r.returncode == 0,
                  (r.stderr or "").strip().splitlines()[-1] if r.returncode else "")
        finally:
            Path(tmp).unlink(missing_ok=True)

    # ② 스코프
    top = _top_level_names(body)
    for name in TARGETS:
        check(f"{name} 이 최상위에 있다", name in top,
              "" if name in top else "다른 스코프에 갇혔거나 없음")
    for name in CALLEES:
        check(f"callee {name} 이 같은 최상위에 있다", name in top,
              "" if name in top else "슬롯 함수가 부르면 ReferenceError")

    # 대상 함수가 부르는 이름이 전부 최상위에 있는가 — 오탈자까지 잡는다.
    known = top | {
        "document", "window", "console", "Promise", "Set", "Map", "JSON",
        "Object", "Array", "String", "Number", "Math", "fetch", "FormData",
        "EventSource", "setTimeout", "clearTimeout", "setInterval",
        "clearInterval", "encodeURIComponent", "parseFloat", "parseInt",
        "isNaN", "Error", "requestAnimationFrame", "alert", "confirm",
        # 문자열 안의 CSS 함수 — 호출이 아니다(`"rgba(248,113,113,.45)"`).
        # 문자열 리터럴을 정규식으로 가르려다 더 틀리느니 이름으로 뺀다.
        "rgba", "rgb", "hsl", "hsla", "url", "var", "calc", "translate",
    }
    call = re.compile(r"\b([A-Za-z_$][\w$]*)\s*\(")
    for name in TARGETS:
        src = _body_of(body, name)
        if not src:
            check(f"{name} 본문 추출", False, "함수를 찾지 못함")
            continue
        unknown = sorted({
            m.group(1) for m in call.finditer(src)
            if m.group(1) not in known
            and not re.search(rf"\b(?:function|const|let|var)\s+{m.group(1)}\b", src)
            # 메서드 호출(.foo(...))·예약어는 뺀다
            and not re.search(rf"\.\s*{m.group(1)}\s*\($", src[:m.start(1) + len(m.group(1)) + 1])
            and m.group(1) not in ("if", "for", "while", "switch", "catch",
                                   "return", "typeof", "function", "await",
                                   "async", "of", "in", "new", "delete",
                                   "void", "yield", "throw")
        })
        # 메서드 호출은 앞에 점이 붙는다 — 줄 단위로 다시 걸러낸다.
        real = [u for u in unknown
                if not re.search(rf"[.\w]\s*\.\s*{re.escape(u)}\s*\(", src)]
        check(f"{name} 이 부르는 이름이 전부 해석된다", not real,
              f"미해석: {', '.join(real)}" if real else "")

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  - " + f)
        return 1
    print("PASS — 구문·스코프 모두 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
