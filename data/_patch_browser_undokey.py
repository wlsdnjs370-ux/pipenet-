# -*- coding: utf-8 -*-
"""브라우저 검증에 «Ctrl+C 한 박자 되돌리기» 검사를 넣는다."""
import io

p = "scripts/_verify_module_f_browser.py"
lines = io.open(p, encoding="utf-8").read().split("\n")

i = next(k for k, ln in enumerate(lines) if "[5-B] 자동 이음" in ln)
block = '''    print("[5-A3] Ctrl+C 한 박자 되돌리기")
    # ① 되돌릴 것이 없을 때 — 핸들러가 붙어 있으면 안내가 뜬다.
    #    실제 키를 눌러 확인한다(핸들러 유무는 런타임에서만 드러난다).
    page.keyboard.press("Control+c")
    page.wait_for_timeout(700)
    msg_empty = page.inner_text("#status")
    print("   되돌릴 것 없을 때:", msg_empty[:46])
    if "되돌" not in msg_empty:
        bad(f"Ctrl+C 가 되돌리기로 가지 않았다: {msg_empty[:60]}")

    # ② 입력칸 안에서는 복사를 빼앗으면 안 된다(상태표 수치를 복사하는 일이 있다).
    hijack = page.evaluate("""() => {
      const inp = document.createElement('input');
      inp.type = 'text'; inp.value = 'abc';
      document.body.appendChild(inp); inp.focus();
      const ev = new KeyboardEvent('keydown', {key: 'c', ctrlKey: true,
                                               bubbles: true, cancelable: true});
      window.dispatchEvent(ev);
      const p = ev.defaultPrevented;
      inp.remove();
      return p;
    }""")
    print("   입력칸 안에서 가로챘나:", hijack)
    if hijack:
        bad("입력칸 안에서 Ctrl+C 가 복사를 빼앗았다")

    # ③ 진짜로 한 박자 되돌아가나 — 급수원을 껐다가 Ctrl+C 로 되살린다.
    #    (망 도형을 안 건드리므로 뒤 단계가 흔들리지 않고, 스스로 복구된다)
    src0 = page.evaluate("() => window.__mf.edit.sources.length")
    page.click('.emode[data-mode="급수시작위치"]')
    page.wait_for_timeout(400)
    spot = page.evaluate("""() => {
      const S = window.__mf, s = S.edit.sources[0];
      return { px: S.toScreenX(s[0]), py: S.toScreenY(s[1]) };
    }""")
    cbox = page.query_selector("#cv").bounding_box()
    page.mouse.click(cbox["x"] + spot["px"], cbox["y"] + spot["py"])
    page.wait_for_timeout(900)
    src1 = page.evaluate("() => window.__mf.edit.sources.length")
    page.keyboard.press("Control+c")
    page.wait_for_timeout(900)
    src2 = page.evaluate("() => window.__mf.edit.sources.length")
    print(f"   급수원 {src0} → 끔 {src1} → Ctrl+C {src2}")
    if src1 != src0 - 1:
        bad(f"급수원이 꺼지지 않아 되돌리기를 시험하지 못했다 ({src0}→{src1})")
    elif src2 != src0:
        bad(f"Ctrl+C 가 한 박자 되돌리지 못했다 ({src1}→{src2}, 기대 {src0})")
    page.click('.emode[data-mode="이음"]')
    page.wait_for_timeout(400)

'''.split("\n")
lines[i:i] = block
io.open(p, "w", encoding="utf-8", newline="\n").write("\n".join(lines))
print("[5-A3] Ctrl+C 검사 추가 (빈상태·복사보존·실되돌림)")
