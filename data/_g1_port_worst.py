# -*- coding: utf-8 -*-
"""[G1] `_worst_k_heads` · `_sheet_frames` 를 design/worst.py 로 그대로 이식.

지시서 §5: «코드를 옮겨 오되 로직은 고치지 않는다». 본문을 손으로 다시 쓰지
않고 원본에서 잘라 옮긴다 — 이름만 공개 계약(§1)에 맞춘다.
"""
import io
import re

SRC = "routes/module_f/remote30.py"
DST = "cad_project_editor_g/services/cad_import/design/worst.py"
lines = io.open(SRC, encoding="utf-8").read().split("\n")


def span(name):
    i = next(k for k, ln in enumerate(lines) if ln.startswith(f"def {name}("))
    j = len(lines)
    for k in range(i + 1, len(lines)):
        if re.match(r"^(def |class |# ─)", lines[k]):
            j = k
            break
    while j - 1 > i and lines[j - 1].strip() == "":
        j -= 1
    return "\n".join(lines[i:j])


worst = span("_worst_k_heads")
sheet = span("_sheet_frames")

# 공개 이름으로만 바꾼다(§1 시그니처). 본문 로직은 건드리지 않는다.
worst = worst.replace("def _worst_k_heads(", "def worst_k_heads(", 1)
worst = worst.replace("k=REMOTE_K_DEFAULT", "k=REMOTE_K_DEFAULT", 1)
sheet = sheet.replace("def _sheet_frames(", "def sheet_frames(", 1)
# 모듈 F 의 화면 문구를 G 의 것으로 — 로직이 아니라 로그 문자열이다.
sheet = sheet.replace("[손질] 도면 장", "[G] 도면 장")

head = '''# -*- coding: utf-8 -*-
"""[G1] 최불리 K 선정 — 앵커 방식(지시서 D1).

`routes/module_f/remote30.py` 에서 **로직 변경 없이** 옮겨 왔다. 순수 그래프
함수라 Qt·Flask 의존이 없다. 모듈 F 의 결과와 앵커·헤드 집합·far_m·max_load 가
완전히 일치해야 이식이 성공한 것이다(§G1 수용 기준).

「먼 순서 K개」가 아니라 앵커 방식인 이유는 worst_k_heads 의 docstring 에 있다.
"""
from __future__ import annotations

import heapq
import math

# NFPC 103 이 요구하는 «가장 불리한 헤드 K개». 기본 30.
REMOTE_K_DEFAULT = 30


'''
io.open(DST, "w", encoding="utf-8", newline="\n").write(
    head + worst.strip("\n") + "\n\n\n" + sheet.strip("\n") + "\n")
print(f"worst.py 작성 — {len((head + worst + sheet).splitlines())}줄")
