"""Read native DXF annotation transforms without constructing drawing geometry.

This is an ephemeral, filtered read, never a rewritten source drawing. Header,
tables, block boundaries, INSERT transforms and native text tags are preserved;
ezdxf still owns decoding, OCS, base points and nested block transformations.
"""
from __future__ import annotations

from io import StringIO
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ezdxf.document import Drawing


def read_annotation_document(path: str | Path, *, extended: bool = False) -> Drawing:
    """Load an ezdxf document containing text and inserts, not pipe geometry."""
    import ezdxf
    from ezdxf.filemanagement import dxf_file_info
    from ezdxf.lldxf.validator import is_binary_dxf_file

    if is_binary_dxf_file(str(path)):
        return ezdxf.readfile(path)
    info = dxf_file_info(str(path))
    output = StringIO()
    section = ""
    section_header = False
    keep = True
    linked_insert = False
    with open(path, encoding=info.encoding, errors="surrogateescape") as source:
        while True:
            code = source.readline()
            if not code:
                break
            value = source.readline()
            if not value:
                raise ValueError("DXF 태그 값이 누락되었습니다.")
            number = int(code.strip())
            name = value.strip()
            if number == 0:
                section_header = name == "SECTION"
                keep = (section not in {"ENTITIES", "BLOCKS"} or name in
                        {"SECTION", "ENDSEC", "EOF", "BLOCK", "ENDBLK", "TEXT", "MTEXT", "INSERT"})
                if extended and name in {'ATTDEF', 'LEADER'}:
                    keep = True
                if section == 'OBJECTS':
                    keep = name in {'SECTION','ENDSEC','EOF','DICTIONARY','DICTIONARYWDFLT','LAYOUT'}
                # Retain INSERT's ATTRIB/SEQEND chain, but not a removed
                # POLYLINE's VERTEX/SEQEND chain (ezdxf validates both).
                if name in {'ATTRIB', 'SEQEND'} and linked_insert:
                    keep = True
                if name not in {'ATTRIB'}:
                    linked_insert = name == 'INSERT'
                if name == "ENDSEC":
                    section = ""
            elif section_header and number == 2:
                section = name
                section_header = False
            if keep:
                output.write(code)
                output.write(value)
    output.seek(0)
    return ezdxf.read(output)
