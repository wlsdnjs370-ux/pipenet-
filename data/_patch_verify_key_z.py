# -*- coding: utf-8 -*-
"""브라우저 검증의 Ctrl+C 잔재를 Ctrl+Z 로 마저 바꾼다."""
import io

p = "scripts/_verify_module_f_browser.py"
t = io.open(p, encoding="utf-8").read()

# 문구부터 — 뒤에 오는 일괄 치환이 이 문장을 건드리지 못하게 먼저 바꾼다.
t = t.replace(
    "    # ② 입력칸 안에서는 복사를 빼앗으면 안 된다(상태표 수치를 복사하는 일이 있다).",
    "    # ② 입력칸 안에서는 브라우저의 «글자 되돌리기» 를 빼앗으면 안 된다.", 1)
t = t.replace('bad("입력칸 안에서 Ctrl+C 가 복사를 빼앗았다")',
              'bad("입력칸 안에서 Ctrl+Z 가 글자 되돌리기를 빼앗았다")', 1)
t = t.replace("{key: 'c', ctrlKey: true,", "{key: 'z', ctrlKey: true,", 1)
t = t.replace('page.keyboard.press("Control+c")', 'page.keyboard.press("Control+z")')
t = t.replace("Ctrl+C", "Ctrl+Z")

left = [ln for ln in t.split("\n") if "Ctrl+C" in ln or "Control+c" in ln]
assert not left, f"잔재: {left}"
io.open(p, "w", encoding="utf-8", newline="\n").write(t)
print("브라우저 검증의 Ctrl+C 잔재 제거 완료")
