# -*- coding: utf-8 -*-
"""수정 상태로 남은 .pptx 가 «내용» 이 바뀐 건가, 열었다 저장만 한 건가.

PPTX 는 zip 이다. 파워포인트는 파일을 열기만 해도 docProps 의 시각·개정번호를
새로 써서 바이트가 달라진다 — 그것을 «수정» 으로 읽고 커밋하면 잡음이 되고,
반대로 진짜 편집을 «잡음» 으로 보고 버리면 남의 작업을 지운다.

그래서 zip 안을 풀어 항목별로 대 본다. 슬라이드가 그대로면 열었다 닫은 것이다.
"""
from __future__ import annotations

import io
import subprocess
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PATH = "scripts/특허도면(한백수정본V1).pptx"

head = subprocess.run(["git", "show", f"HEAD:{PATH}"],
                      capture_output=True).stdout
work = open(PATH, "rb").read()
print(f"HEAD {len(head):,}B · 작업트리 {len(work):,}B "
      f"(차이 {len(work) - len(head):+,}B)")

za = zipfile.ZipFile(io.BytesIO(head))
zb = zipfile.ZipFile(io.BytesIO(work))
na, nb = set(za.namelist()), set(zb.namelist())

print(f"\n항목 수 : HEAD {len(na)} · 작업트리 {len(nb)}")
if na - nb:
    print("  사라진 항목:", sorted(na - nb)[:8])
if nb - na:
    print("  새 항목    :", sorted(nb - na)[:8])

changed = []
for n in sorted(na & nb):
    if za.read(n) != zb.read(n):
        changed.append(n)

print(f"\n내용이 다른 항목 {len(changed)}개")
for n in changed:
    print(f"  {n}  ({len(za.read(n)):,} → {len(zb.read(n)):,}B)")

# 슬라이드·그림·글씨가 그대로면 «열었다 저장» 이다.
real = [n for n in changed
        if not n.startswith("docProps/")
        and "app.xml" not in n and "core.xml" not in n]
print()
if not real and not (na ^ nb):
    print("→ 슬라이드 내용은 그대로 · docProps(열어본 기록)만 다르다")
else:
    print("→ ★내용이 바뀌었다 — 사람이 편집한 파일일 수 있다:", real[:10])
