# -*- coding: utf-8 -*-
"""모듈 F — Ctrl+C 한 박자 되돌리기."""
import io

p = "templates/module_f.html"
s = io.open(p, encoding="utf-8").read()

# ── ① 전역 키 핸들러 ────────────────────────────────────────────
anchor = '''  // ── 그리기 ─────────────────────────────────────────────────────
  function draw() {'''
block = '''  // ── 되돌리기 단축키 ────────────────────────────────────────────
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

    e.preventDefault();
    // 무거운 작업이 도는 중에 망을 되돌리면 그 작업과 엇갈린다.
    if (!$("busy").classList.contains("hidden")) {
      say("작업이 끝난 뒤에 되돌릴 수 있습니다.", "warn");
      return;
    }
    const btn = S.stage === "pick" ? $("pk-undo")
              : S.stage === "edit" ? $("ed-undo") : null;
    if (btn) btn.click();
    else say("이 단계에는 되돌릴 것이 없습니다.", "warn");
  });

''' + anchor
assert anchor in s
s = s.replace(anchor, block, 1)

# ── ② 단추에 단축키를 적어 둔다 — 숨은 기능은 없는 기능이다 ────────
old_pk = '<button id="pk-undo">되돌리기</button>'
new_pk = '<button id="pk-undo" title="한 박자 되돌리기 (Ctrl+C)">되돌리기</button>'
assert old_pk in s
s = s.replace(old_pk, new_pk, 1)

old_ed = '<button id="ed-undo">되돌리기</button>'
new_ed = '<button id="ed-undo" title="한 박자 되돌리기 (Ctrl+C)">되돌리기</button>'
assert old_ed in s
s = s.replace(old_ed, new_ed, 1)

io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("Ctrl+C 되돌리기 추가 · 단추 툴팁 2곳")
