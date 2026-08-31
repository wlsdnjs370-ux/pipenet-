# -*- coding: utf-8 -*-
"""모듈 F 정합성 — «두 곳이 다른 말을 하는» 자리를 영구히 막는다.

리팩터링 사전 조사에서 프로파일이 방향을 정했다: 모듈 F 자신의 코드는 완주
시간의 **0.3%** 다(엔진 67%). 그러니 여기서 지킬 것은 속도가 아니라 **정합성**
이다. 조사 스크립트(`scripts/_probe_f_consistency.py`)로 한 번 훑는 것으로는
부족하다 — 사람이 부를 때만 돌기 때문이다. 시험으로 옮겨 늘 돌게 한다.

★특히 ②가 이 저장소에서 실제로 사고를 낸 종류다. `$("없는id")` 는 예외를 안
  내고 null 을 준다 — 구문 검사도, `node --check` 도 통과하고 화면만 조용히
  빈다. 그래서 «이름이 맞는가» 를 시험이 봐야 한다.
"""
from __future__ import annotations

import os
import re

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_PY_DIR = os.path.join(_ROOT, "routes", "module_f")

# 단계 목록에 없는 것이 «맞는» 패널 — 늘 보여야 하므로 켜고 끄지 않는다.
ALWAYS_ON_IDS = {"panel-log"}


def _read(*parts) -> str:
    return open(os.path.join(_ROOT, *parts), encoding="utf-8").read()


def _js() -> str:
    return _read("static", "module_f.js")


def _html() -> str:
    return _read("templates", "module_f.html")


def _module_f_py() -> str:
    out = []
    for name in sorted(os.listdir(_PY_DIR)):
        if name.endswith(".py"):
            out.append(open(os.path.join(_PY_DIR, name),
                            encoding="utf-8").read())
    return "\n".join(out)


def _served() -> set:
    """서버가 등록한 라우트 규칙 — 모듈 F 밖의 공용 라우트도 센다."""
    rules = set()
    for root, _dirs, files in os.walk(os.path.join(_ROOT, "routes")):
        for f in files:
            if not f.endswith(".py"):
                continue
            src = open(os.path.join(root, f), encoding="utf-8").read()
            rules |= set(re.findall(
                r'@app\.(?:get|post|route)\("([^"]+)"', src))
    return rules


def test_화면이_부르는_엔드포인트가_전부_서버에_있다():
    """오타 하나면 그 단추만 조용히 404 가 된다 — 시험이 아니면 못 잡는다."""
    served = _served()

    def ok(path: str) -> bool:
        if path in served:
            return True
        # ★경로 매개변수를 감안한다. `/diagram/<key>` 를 화면은
        #   `/diagram/${g.diagram}` 로 부른다 — 문자열을 그대로 대면
        #   멀쩡한 라우트를 «없다» 고 짚는다(조사 때 실제로 그랬다).
        for rule in served:
            base = rule.split("<")[0].rstrip("/")
            if "<" in rule and base and path.startswith(base):
                return True
        return False

    called = {m.rstrip("/") for m in
              re.findall(r'["`](/api/module-f/[^"`?${]+)', _js())}
    assert called, "화면이 부르는 엔드포인트를 하나도 못 찾았다 — 검사가 죽었다"
    missing = sorted(c for c in called if not ok(c))
    assert not missing, f"화면이 부르는데 서버에 없다: {missing}"


def test_JS_가_찾는_id_가_전부_템플릿에_있다():
    """★`$("없는id")` 는 예외가 아니라 **null** 이다.

    구문 검사도 `node --check` 도 통과하고 화면만 조용히 빈다. 이 저장소가
    「프론트 JS 런타임 검증 필수」를 규약으로 둔 이유가 이것이다.
    """
    have = set(re.findall(r'id="([\w-]+)"', _html()))
    js = _js()
    want = set(re.findall(r'\$\("([\w-]+)"\)', js))
    want |= set(re.findall(r'getElementById\("([\w-]+)"\)', js))
    assert want, "JS 가 찾는 id 를 하나도 못 찾았다 — 검사가 죽었다"
    ghost = sorted(want - have)
    assert not ghost, f"JS 가 찾는데 템플릿에 없다: {ghost}"


def test_템플릿에_죽은_id_가_없다():
    """안 쓰는 마크업은 «남은 것» 이지 «있는 것» 이 아니다.

    ★«안 쓰인다» 를 말하려면 세 가지를 먼저 빼야 한다 — 안 빼면 멀쩡한 것을
      지우게 된다(조사 때 9건 중 9건이 가짜였다):
        · 접이식은 `data-fold="…"` 로 잇고 JS 는 `dataset.fold` 로 읽는다
        · 이름을 조립해 부르는 것 — ``$(`${id}-mark`)``
        · 늘 보이는 패널 — 단계 목록에 없는 것이 맞다
    """
    html, js, py = _html(), _js(), _module_f_py()
    have = set(re.findall(r'id="([\w-]+)"', html))
    want = set(re.findall(r'\$\("([\w-]+)"\)', js))
    want |= set(re.findall(r'getElementById\("([\w-]+)"\)', js))
    suffixes = set(re.findall(r'\$\(\s*\w+\s*\+\s*"([\w-]+)"\s*\)', js))
    suffixes |= set(re.findall(r'\$\(\s*`\$\{\w+\}([\w-]+)`\s*\)', js))
    prefixes = set(re.findall(r'"([\w-]+)"', js))
    made = {p + s for s in suffixes for p in prefixes}
    dead = sorted(
        i for i in have
        if i not in want and i not in made and i not in ALWAYS_ON_IDS
        and f'"{i}"' not in js and f"'{i}'" not in js and f"#{i}" not in js
        and f'"{i}"' not in py
        and f'data-fold="{i}"' not in html and f'for="{i}"' not in html
        and f'aria-controls="{i}"' not in html)
    assert not dead, f"아무도 안 쓰는 id: {dead}"


def test_같은_상수가_파일마다_다르지_않다():
    """한 값이 두 곳에 있으면 언젠가 한쪽만 고쳐진다."""
    seen: dict = {}
    for name in sorted(os.listdir(_PY_DIR)):
        if not name.endswith(".py"):
            continue
        src = open(os.path.join(_PY_DIR, name), encoding="utf-8").read()
        for m in re.finditer(r"^([A-Z][A-Z0-9_]{3,})\s*=\s*([^\n#]+)", src,
                             re.M):
            seen.setdefault(m.group(1), {})[name] = m.group(2).strip()
    clash = {k: v for k, v in seen.items()
             if len(v) > 1 and len(set(v.values())) > 1}
    assert not clash, f"같은 이름의 상수가 파일마다 다르다: {clash}"


def test_소스가_가리키는_BLOCKED_절이_실재한다():
    """번호를 정리하고 나면 참조가 깨진다 — 그 사실을 시험이 잡는다."""
    blocked = _read("BLOCKED.md")
    secs = set(re.findall(r"^## (\d+(?:-\d+)?)\.", blocked, re.M))
    text = _module_f_py() + _js() + _html()
    refs = set(re.findall(r"BLOCKED[^\n]{0,8}§(\d+)", text))
    missing = sorted(r for r in refs if r not in secs)
    assert not missing, f"가리키는 BLOCKED 절이 없다: §{missing}"


def test_BLOCKED_번호에_중복이_없다():
    """[F-11e-3] 한 번 정리했으니 다시 흔들리지 않게 못 박는다."""
    nums = re.findall(r"^## (\d+(?:-\d+)?)\.", _read("BLOCKED.md"), re.M)
    dup = sorted({n for n in nums if nums.count(n) > 1})
    assert not dup, f"BLOCKED 번호가 겹친다: {dup}"


# ══════════════════════════════════════════ module_f.js — 조용히 깨지는 꼴
def _js_stripped() -> str:
    """문자열·주석을 지운 소스. **줄 수는 보존한다.**

    통째로 지우면 그 뒤 줄 번호가 밀려 지적이 엉뚱한 줄을 가리킨다 —
    조사할 때 실제로 두 줄씩 어긋났다.
    """
    def keep(m):
        return '""' + "\n" * m.group(0).count("\n")

    out = re.sub(r"//[^\n]*", "", _js())
    out = re.sub(r"/\*.*?\*/",
                 lambda m: "\n" * m.group(0).count("\n"), out, flags=re.S)
    out = re.sub(r'"(?:\\.|[^"\\])*"', keep, out)
    out = re.sub(r"'(?:\\.|[^'\\])*'", keep, out)
    return re.sub(r"`(?:\\.|[^`\\])*`", keep, out, flags=re.S)


def test_상태에_읽히지_않는_필드가_없다():
    """`S.x` 를 담아 두기만 하고 아무도 안 읽으면, 다음 사람이 그것을
    «신뢰할 수 있는 최신값» 으로 오해한다.

    ★검증용으로 일부러 내보내는 것은 뺀다 — 화면이 안 읽는 것이 정상이다.
      코드의 「[검증 내보내기]」 표시를 보고 가른다. 목록을 시험에 손으로
      적으면 하나 늘 때마다 시험이 시끄러워진다.
    """
    raw, src = _js(), _js_stripped()
    writes = {}
    for m in re.finditer(r"\bS\.(\w+)\s*=(?!=)", src):
        writes[m.group(1)] = writes.get(m.group(1), 0) + 1
    # 읽기는 **원본** 에서 센다 — 템플릿 리터럴 안의 읽기를 놓치면 안 된다.
    no_comment = re.sub(r"//[^\n]*", "", raw)
    exported: set = set()
    mk = raw.find("[검증 내보내기]")
    if mk >= 0:
        exported = set(re.findall(r"\bS\.(\w+)\s*=",
                                  raw[mk:raw.find("\n\n", mk)]))
    dead = sorted(
        k for k, n in writes.items()
        if k not in exported
        and len(re.findall(rf"\bS\.{re.escape(k)}\b", no_comment)) - n <= 0)
    assert not dead, f"쓰기만 하고 아무도 안 읽는 상태 필드: {dead}"


def test_같은_요소에_같은_핸들러를_두_번_걸지_않는다():
    """`$("x").onclick` 을 두 번 걸면 **앞의 것이 조용히 죽는다.**"""
    seen: dict = {}
    for m in re.finditer(r'\$\("([\w-]+)"\)\.(on\w+)\s*=', _js()):
        seen.setdefault((m.group(1), m.group(2)), []).append(
            _js()[:m.start()].count("\n") + 1)
    twice = {f"#{k[0]}.{k[1]}": v for k, v in seen.items() if len(v) > 1}
    assert not twice, f"핸들러를 겹쳐 건다(앞의 것이 죽는다): {twice}"


def test_JS_에_var_와_느슨한_비교가_없다():
    """`var` 는 블록 스코프가 아니라 루프에서 어긋난다.

    ★`x != null` 은 예외다 — null 과 undefined 를 한 번에 거르는 관용구이고,
      `!==` 로 바꾸면 undefined 를 놓쳐 **오히려 틀린 코드**가 된다.
    """
    src = _js_stripped()
    assert not re.findall(r"\bvar\s+\w", src), "var 가 남아 있다"
    loose = [src[:m.start()].count("\n") + 1
             for m in re.finditer(r"[^=!<>]([=!]=)(?!=)(?!\s*null\b)", src)]
    assert not loose, f"느슨한 비교(줄): {loose}"
