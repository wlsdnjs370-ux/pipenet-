# -*- coding: utf-8 -*-
import io
p = "scripts/_verify_module_f_browser.py"
lines = io.open(p, encoding="utf-8").read().split("\n")
i = next(k for k, ln in enumerate(lines) if 'page.click("#ed-aj-apply")' in ln)
block = '''    # 가림막은 «화면 전체» 를 덮어야 한다. 캔버스만 덮으면 옆 패널 단추가
    # 작업 중에도 눌려 같은 작업이 두 번 돈다(서버도 막지만 화면이 1차 방벽).
    page.wait_for_timeout(400)
    cover = page.evaluate("""() => {
      const btn = document.getElementById('ed-aj-apply');
      const r = btn.getBoundingClientRect();
      const hit = document.elementFromPoint(r.left + r.width / 2,
                                            r.top + r.height / 2);
      const busy = document.getElementById('busy');
      return { hidden: busy.classList.contains('hidden'),
               covered: !!hit && (hit === busy || busy.contains(hit)),
               hit: hit ? (hit.id || hit.className || hit.tagName) : null };
    }""")
    print("   작업 중 가림막:", cover)
    if not cover["hidden"] and not cover["covered"]:
        bad(f"작업 중인데 «모두 잇기» 단추가 노출돼 있다 (hit={cover['hit']})")'''.split("\n")
lines[i + 1:i + 1] = block
io.open(p, "w", encoding="utf-8").write("\n".join(lines))
print("가림막 검사 추가")
