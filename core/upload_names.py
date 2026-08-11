# -*- coding: utf-8 -*-
"""업로드 파일명 위생 — 한글을 살리면서 경로 조작만 막는다.

werkzeug 의 `secure_filename` 은 비-ASCII 를 통째로 버려 '평면도.dxf' 와
'계통도.dxf' 가 똑같이 'dxf' 가 된다. 모듈 A 는 평면도·계통도·기계실 세 도면을
한 세션에서 올리므로 그 충돌이 곧 앞 도면 덮어쓰기이고, 토큰이 가리키는 실체가
조용히 바뀐다.
"""
import re
import unicodedata
from pathlib import PurePosixPath

# 유니코드 단어문자(한글 포함)·점·하이픈·밑줄·공백·괄호만 남긴다.
_UNSAFE = re.compile(r"[^\w.\- ()]", re.UNICODE)
_MAX_STEM = 120


def sanitize_upload_name(name: str) -> str:
    """저장·조회에 쓸 파일명. 살릴 수 없으면 빈 문자열."""
    text = unicodedata.normalize("NFC", str(name or ""))
    text = PurePosixPath(text.replace("\\", "/")).name
    text = _UNSAFE.sub("_", text)
    path = PurePosixPath(text)
    stem, suffix = path.stem[:_MAX_STEM], path.suffix
    return f"{stem}{suffix}".strip(" .")
