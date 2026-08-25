# -*- coding: utf-8 -*-
"""손질 오픈용 자동망 스냅샷. 화면 없음.

오너 찍기/유저손질 JSON 을 덮지 않는다. pts/edges 등 편집 기준망이므로
DXF/스펙 + pipeline 소스까지 바뀌면 무효화한다.
"""
import json
import os

from services.cad_import import kinds
from services.cad_import.pipeline import expand, flow, heads, stage1, stage45
from services.cad_import.pipeline.handoff import import_write_root

_DISP_CACHE_VER = 5
_DISP_CACHE_DIR = import_write_root()


def _disp_cache_dir():
    return _DISP_CACHE_DIR


def _disp_cache_path(key):
    return os.path.join(_disp_cache_dir(), f"_edit_disp_cache_{key}.json")


def _file_stamp(path, content_hash=False):
    """mtime_ns+size. content_hash=True 이면 작은 스펙용 sha256 추가."""
    path = os.path.normpath(path)
    try:
        st = os.stat(path)
    except OSError:
        return {"path": path, "mtime_ns": None, "size": None, "sha256": None}
    rec = {"path": path, "mtime_ns": st.st_mtime_ns, "size": st.st_size,
           "sha256": None}
    if content_hash:
        import hashlib
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        rec["sha256"] = h.hexdigest()
    return rec


def _disp_cache_inputs(key):
    spec = expand._spec_path(key)
    dxf = expand.dxf_path_for(key) or os.path.join(expand.DWG, f"{key}.dxf")
    pipeline_modules = (expand, flow, stage1, heads, stage45, kinds)
    return {
        "ver": _DISP_CACHE_VER,
        "key": key,
        "dxf": _file_stamp(dxf, content_hash=False),
        "spec": _file_stamp(spec, content_hash=True),
        "code": [_file_stamp(module.__file__, content_hash=True)
                 for module in pipeline_modules],
    }


def _disp_cache_load(key):
    """히트 시 EditBoard 용 data dict, 아니면 None (보수적 MISS)."""
    path = _disp_cache_path(key)
    if not os.path.isfile(path):
        return None
    try:
        blob = json.load(open(path, encoding="utf-8"))
    except Exception:
        return None
    want = _disp_cache_inputs(key)
    got = blob.get("inputs")
    if got != want:
        return None
    if want["dxf"]["mtime_ns"] is None or want["spec"]["mtime_ns"] is None:
        return None
    data = blob.get("data") or {}
    need = ("pts", "edges", "edges1", "hcov", "hnodes", "ups", "head_kinds",
            "ho")
    if any(k not in data for k in need):
        return None
    return data


def _disp_cache_save(key, result):
    """EditBoard 기준망 필드만 저장. 실패해도 본 흐름은 계속."""
    inputs = _disp_cache_inputs(key)
    if inputs["dxf"]["mtime_ns"] is None or inputs["spec"]["mtime_ns"] is None:
        return
    data = {
        "pts": [list(p) for p in result["pts"]],
        "edges": [list(e) for e in result["edges"]],
        "edges1": [list(e) for e in result["edges1"]],
        "hcov": [list(h) for h in result["hcov"]],
        "hnodes": [list(n) for n in result["hnodes"]],
        "ups": [list(u) for u in (result.get("ups") or ())],
        "head_kinds": list(result.get("head_kinds") or ()),
        "ho": flow.ho_from_spots(result.get("spots")),
    }
    path = _disp_cache_path(key)
    try:
        os.makedirs(_disp_cache_dir(), exist_ok=True)
        json.dump({"inputs": inputs, "data": data},
                  open(path, "w", encoding="utf-8"), ensure_ascii=False)
        print(f"[손질] 표시 캐시 저장 · {path}")
    except Exception as exc:
        print(f"[손질] 표시 캐시 저장 실패: {exc}")
