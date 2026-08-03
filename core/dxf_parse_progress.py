# -*- coding: utf-8 -*-
"""진행률을 측정할 수 있는 DXF 파싱.

`ezdxf.readfile` 은 진행 콜백이 없어 141MB 도면에서 30초 넘게 아무 신호도 주지
않는다. 여기서는 readfile 이 내부적으로 하는 일을 두 단계로 갈라 각각을 실측한다.

  read  — 텍스트 태그 구조 로드(`load_dxf_structure`). 소비한 바이트로 잰다.
  build — entity 바인딩(`_load_section_dict`). entitydb 증가로 잰다.

바이트만으로 재면 안 되는 이유: LH 지하층배관도(141MB) 기준 스트림은 18.6초에
끝나지만 전체 파싱은 35.9초다. 나머지 절반이 build 단계이고, 그 구간을 안 재면
막대가 100% 에서 멈춘 채 한참 기다리는 예전 증상이 그대로 재현된다.
"""
from __future__ import annotations

import io
import os


class _CountingFile(io.FileIO):
    """읽은 바이트 수를 보고하는 raw 파일.

    텍스트 층(readline)이 아니라 raw 층에서 센다 — DXF 는 태그당 두 줄이라
    readline 을 감싸면 16MB 도면에서만 240만 번의 파이썬 호출이 끼어들어
    파싱이 1.8배 느려진다. raw 는 버퍼 크기(수 KB)마다 한 번만 불린다.
    """

    def __init__(self, filename, on_pos):
        super().__init__(filename, "r")
        self._on_pos = on_pos
        self._pos = 0

    def readinto(self, b):  # type: ignore[override]
        n = super().readinto(b)
        if n:
            self._pos += n
            self._on_pos(self._pos)
        return n


def parse_dxf_with_progress(path, state: dict):
    """DXF 를 파싱하며 `state` 를 갱신하고 `Drawing` 을 돌려준다.

    state["stage"]/["done"]/["total"] 이 갱신되며, build 단계에서는 분자를
    state["db"](entitydb) 의 길이로 읽어야 한다 — 바인딩이 도중에 진행률을
    보고할 방법이 그것뿐이다. 바이너리 DXF 등 분해가 불가능한 입력은
    `ezdxf.readfile` 로 넘기고 total=0 으로 두어 호출자가 불확정 표시로 넘어가게 한다.
    """
    import ezdxf
    from ezdxf.document import Drawing
    from ezdxf.filemanagement import dxf_file_info
    from ezdxf.lldxf import loader as dxf_loader
    from ezdxf.lldxf.tagger import ascii_tags_loader, tag_compiler
    from ezdxf.lldxf.validator import is_binary_dxf_file

    filename = str(path)
    if is_binary_dxf_file(filename):
        state.update(stage="read", done=0, total=0)
        return ezdxf.readfile(filename)

    size = os.path.getsize(filename)
    info = dxf_file_info(filename)
    state.update(stage="read", done=0, total=size)

    def _on_pos(pos):
        state["done"] = min(pos, size)

    raw = _CountingFile(filename, _on_pos)
    with io.TextIOWrapper(io.BufferedReader(raw), encoding=info.encoding,
                          errors="surrogateescape") as fp:
        tagger = tag_compiler(ascii_tags_loader(fp))
        sections = dxf_loader.load_dxf_structure(tagger)
    sections.pop("THUMBNAILIMAGE", None)

    doc = Drawing()
    total_records = sum(len(v) for v in sections.values())
    state.update(stage="build", done=0, total=total_records, db=doc.entitydb)
    doc.is_loading = True
    doc._load_section_dict(sections)  # ezdxf 내부 API — Drawing._load 의 2단계
    doc.filename = filename
    state.update(stage="build", done=total_records)
    return doc
