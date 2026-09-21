# -*- coding: utf-8 -*-
import io

p = "scripts/_verify_module_f.py"
lines = io.open(p, encoding="utf-8").read().split("\n")


def find(pred, start=0):
    for i in range(start, len(lines)):
        if pred(lines[i]):
            return i
    raise SystemExit(f"못 찾음: {pred}")


# ① 입력 방어 — "없는 sid" 검사 바로 뒤
i = find(lambda ln: "없는 sid" in ln and "check(" in ln)
block1 = '''        # 도면 키는 파일 이름이지 경로가 아니다. E 는 키를 경로에 그대로 끼워
        # 넣으므로(정규화하는 것은 handoff_path 뿐) 문 앞에서 막아야 한다.
        EVIL_KEYS = ["../../../../Users/admin/Desktop/PWNED",
                     "..\\..\\..\\Windows\\Temp\\x",
                     "..", ".", "a/b", "C:\\Windows\\x", "tab\tkey", ""]
        for evil in EVIL_KEYS:
            r = c.post("/api/module-f/reopen", json={"key": evil})
            if not check(f"경로 키 거절 {evil[:24]!r}", r.status_code == 400,
                         f"HTTP {r.status_code}"):
                break'''.split("\n")
lines[i + 1:i + 1] = block1

# ② 잡 중복 거절 — autojoin/apply 수락 검사 바로 뒤
i = find(lambda ln: '"자동 이음 수락"' in ln)
block2 = '''        # ★가림막이 캔버스만 덮던 시절엔 옆 패널 단추가 작업 중에도 눌렸다.
        #   두 번 들어가면 낡은 후보로 또 붙어 되돌리기 한 번으로 복구가 안 된다.
        r2 = c.post("/api/module-f/edit/autojoin/apply", json={"sid": sid2})
        check("작업 중 중복 제출 거절", r2.status_code == 409,
              f"HTTP {r2.status_code}")'''.split("\n")
lines[i + 1:i + 1] = block2

# ③ 저장 응답에 서버 경로 없음 — [3-A] 절 바로 앞
i = find(lambda ln: "[3-A]" in ln)
block3 = '''        r = c.post("/api/module-f/edit/save", json={"sid": sid2})
        j = r.get_json()
        check("손질 저장", j.get("ok"), str(j.get("message"))[:70])
        check("응답에 서버 경로가 없다",
              "path" not in j and ":" not in str(j.get("file", "")),
              f"file={j.get('file')}")
'''.split("\n")
lines[i:i] = block3

io.open(p, "w", encoding="utf-8").write("\n".join(lines))
print("검증 3종 추가 완료")
