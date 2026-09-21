# -*- coding: utf-8 -*-
"""되돌리기 단축키를 Ctrl+C → Ctrl+Z 로 바로잡는다."""
import io

# ── ① 템플릿 ────────────────────────────────────────────────────
p = "templates/module_f.html"
s = io.open(p, encoding="utf-8").read()

old = '''  // ── 되돌리기 단축키 ────────────────────────────────────────────
  // Ctrl+C 한 번 = 한 박자 되돌리기. 단계에 맞는 «되돌리기» 단추를 **그대로
  // 누른다** — 여기서 API 를 따로 부르면 단추와 단축키의 동작이 갈라진다
  // (한쪽만 고치는 사고가 난다).
  //
  // ★복사를 빼앗지 않는다. 글자를 끌어 놓았거나 입력칸·선택칸에 있으면
  //   Ctrl+C 는 원래 뜻(복사)대로 둔다 — 상태표의 수치나 로그를 복사하는 일이
  //   실제로 있다. 그때는 preventDefault 도 하지 않는다.
  window.addEventListener("keydown", (e) => {
    if ((e.key || "").toLowerCase() !== "c") return;
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;

    const el = document.activeElement;
    const tag = el ? el.tagName : "";
    if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT"
        || (el && el.isContentEditable)) return;
    const sel = window.getSelection();
    if (sel && String(sel).length) return;   // 진짜 복사다 — 건드리지 않는다

    e.preventDefault();'''
new = '''  // ── 되돌리기 단축키 ────────────────────────────────────────────
  // Ctrl+Z 한 번 = 한 박자 되돌리기. 단계에 맞는 «되돌리기» 단추를 **그대로
  // 누른다** — 여기서 API 를 따로 부르면 단추와 단축키의 동작이 갈라진다
  // (한쪽만 고치는 사고가 난다).
  //
  // ★입력칸 안에서는 손대지 않는다. 변환 폼에 숫자를 치다 Ctrl+Z 를 누르면
  //   그건 «글자 되돌리기» 지 «손질 되돌리기» 가 아니다 — 브라우저에 맡긴다.
  //   Ctrl+Shift+Z(다시 실행)도 여기서 다루지 않는다(shiftKey 로 걸러진다).
  window.addEventListener("keydown", (e) => {
    if ((e.key || "").toLowerCase() !== "z") return;
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;

    const el = document.activeElement;
    const tag = el ? el.tagName : "";
    if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT"
        || (el && el.isContentEditable)) return;

    e.preventDefault();'''
assert old in s, "템플릿 키 핸들러를 못 찾음"
s = s.replace(old, new, 1)

n = s.count('title="한 박자 되돌리기 (Ctrl+C)"')
assert n == 2, f"단추 툴팁이 2개가 아님: {n}"
s = s.replace('title="한 박자 되돌리기 (Ctrl+C)"',
              'title="한 박자 되돌리기 (Ctrl+Z)"')

io.open(p, "w", encoding="utf-8", newline="\n").write(s)

# ── ② 브라우저 검증 ─────────────────────────────────────────────
p2 = "scripts/_verify_module_f_browser.py"
t = io.open(p2, encoding="utf-8").read()

t = t.replace('print("[5-A3] Ctrl+C 한 박자 되돌리기")',
              'print("[5-A3] Ctrl+Z 한 박자 되돌리기")', 1)
t = t.replace('page.keyboard.press("Control+c")', 'page.keyboard.press("Control+z")')
t = t.replace('''    # ② 입력칸 안에서는 복사를 빼앗으면 안 된다(상태표 수치를 복사하는 일이 있다).''',
              '''    # ② 입력칸 안에서는 브라우저의 «글자 되돌리기» 를 빼앗으면 안 된다.''', 1)
t = t.replace("""      const ev = new KeyboardEvent('keydown', {key: 'c', ctrlKey: true,""",
              """      const ev = new KeyboardEvent('keydown', {key: 'z', ctrlKey: true,""", 1)
t = t.replace('        bad("입력칸 안에서 Ctrl+C 가 복사를 빼앗았다")',
              '        bad("입력칸 안에서 Ctrl+Z 가 글자 되돌리기를 빼앗았다")', 1)
t = t.replace('        bad(f"Ctrl+C 가 되돌리기로 가지 않았다: {msg_empty[:60]}")',
              '        bad(f"Ctrl+Z 가 되돌리기로 가지 않았다: {msg_empty[:60]}")', 1)
t = t.replace('    print(f"   급수원 {src0} → 클릭 {src1} → Ctrl+C {src2}")',
              '    print(f"   급수원 {src0} → 클릭 {src1} → Ctrl+Z {src2}")', 1)
t = t.replace('        bad(f"Ctrl+C 가 한 박자 되돌리지 못했다 ({src1}→{src2}, 기대 {src0})")',
              '        bad(f"Ctrl+Z 가 한 박자 되돌리지 못했다 ({src1}→{src2}, 기대 {src0})")', 1)

assert "Control+c" not in t and "Ctrl+C" not in t, "검증에 Ctrl+C 잔재가 남음"
io.open(p2, "w", encoding="utf-8", newline="\n").write(t)

print("Ctrl+C -> Ctrl+Z 로 교체 (템플릿 핸들러/툴팁 2곳 + 브라우저 검증)")
