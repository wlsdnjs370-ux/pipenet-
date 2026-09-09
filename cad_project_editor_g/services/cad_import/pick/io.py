# -*- coding: utf-8 -*-
"""찍기 입출력 — DXF 열기·스펙/클릭 저장. 화면 없음."""
import json
import os
import time

from services.cad_import.pick.board import Board, heads_from_spec
from services.cad_import.pipeline.expand import CASES
from services.cad_import.pipeline import handoff
from services.cad_import.pipeline.handoff import MIN_PREP_SECONDS, save_world
from services.cad_import.pipeline import stage1 as s1

def new_dir():
    """찍은스펙·자동백업 폴더 — **부를 때** 정한다.

    종전에는 `NEW_DIR = handoff.OUT_DIR` 로 import 시점에 값을 복사했다.
    서버가 부팅 때 쓰기 루트를 바꿔도 이 복사본은 안 따라오므로, 한 프로세스
    안에서 모듈마다 다른 폴더를 붙들었다 — 오류 없이 파일만 엉뚱한 데 생긴다.
    """
    return handoff.pick_out_dir()


def std_dir():
    """0단계 표준샘플 폴더.

    ★종전 값은 cwd 상대경로 "docs/import/0단계_표준샘플" 이었고, 부팅 때
      절대경로로 고정하는 목록에서 **아예 빠져 있었다**. 데스크톱 G 는 cwd 가
      편집기 폴더라 맞지만, 웹서버는 cwd 가 프로젝트 루트라 같은 상대경로가
      엉뚱한 곳을 가리킨다. 소스 실행에서는 `import_write_root()` 가
      "docs/import" 라 **값이 예전과 같다** — 서버에서만 달라지고, 그게
      원래 의도한 동작이다.
    """
    return os.path.join(handoff.import_write_root(), "0단계_표준샘플")


def display_key_for(source_path):
    """알려진 샘플 파일명은 CASES 키로, 그 외는 파일 이름."""
    base = os.path.basename(source_path)
    for k, fname in CASES.items():
        if str(fname).lower() == base.lower():
            return k
    return os.path.splitext(base)[0]


def resolve_dxf(key_or_path):
    """키·경로 → 절대 DXF 경로. 없으면 ValueError('dxf_path')."""
    if os.path.isfile(key_or_path) and str(key_or_path).lower().endswith(".dxf"):
        return os.path.abspath(key_or_path)
    if key_or_path in CASES:
        cand = os.path.join(s1.DWG_DIR, CASES[key_or_path])
        if os.path.isfile(cand):
            return os.path.abspath(cand)
    fname = (key_or_path if str(key_or_path).lower().endswith(".dxf")
             else f"{key_or_path}.dxf")
    cand = os.path.join(s1.DWG_DIR, os.path.basename(fname))
    if os.path.isfile(cand):
        return os.path.abspath(cand)
    raise ValueError("dxf_path")


def open_dxf(source_path, knobs=None):
    """DXF → (world, display_key, knobs)."""
    source_path = os.path.abspath(source_path)
    if not os.path.isfile(source_path):
        raise ValueError("dxf_path")
    key = display_key_for(source_path)
    t0 = time.perf_counter()
    ltab, ents, bdefs = s1.read_dxf(source_path)
    w, _hid = s1.explode(ltab, ents, bdefs)
    w._source_path = source_path
    w._prep_seconds = time.perf_counter() - t0
    kn = dict(s1.DEFAULT_KNOBS)
    if knobs:
        kn.update(knobs)
    return w, key, kn


def _ensure_dir(out_dir):
    os.makedirs(out_dir, exist_ok=True)


def write_pick(key, world, board, out_dir=None):
    """스펙+클릭 정식 저장. 그림·정합 보고 없음."""
    out_dir = out_dir or handoff.pick_out_dir()
    _ensure_dir(out_dir)
    sp = board.spec()
    src = getattr(world, "_source_path", None)
    if src:
        sp["source_dxf"] = os.path.abspath(src)
    p_clicks = os.path.join(out_dir, f"{key}_찍은클릭.json")
    p_spec = os.path.join(out_dir, f"{key}_찍은스펙.json")
    with open(p_clicks, "w", encoding="utf-8") as f:
        json.dump(board.clicks, f, ensure_ascii=False, indent=1)
    with open(p_spec, "w", encoding="utf-8") as f:
        json.dump(sp, f, ensure_ascii=False, indent=1)
    prep_seconds = getattr(world, "_prep_seconds", 0.0)
    if prep_seconds >= MIN_PREP_SECONDS:
        try:
            save_world(key, world._source_path, world)
        except Exception:
            pass
    return p_spec


def write_backup(key, board, out_dir=None):
    """찍을 때마다 자동백업. 정식 파일은 안 건드린다."""
    out_dir = out_dir or handoff.pick_out_dir()
    _ensure_dir(out_dir)
    with open(os.path.join(out_dir, f"{key}_찍은클릭_자동백업.json"),
              "w", encoding="utf-8") as f:
        json.dump(board.clicks, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out_dir, f"{key}_찍은스펙_자동백업.json"),
              "w", encoding="utf-8") as f:
        json.dump(board.spec(), f, ensure_ascii=False, indent=1)


def load_existing(key, board, out_dir=None):
    """이어 찍기 — 찍은스펙·찍은클릭을 불러 계속한다. 반환: 픽 수."""
    out_dir = out_dir or handoff.pick_out_dir()
    p_spec = os.path.join(out_dir, f"{key}_찍은스펙.json")
    p_clicks = os.path.join(out_dir, f"{key}_찍은클릭.json")
    n = 0
    if os.path.exists(p_spec):
        with open(p_spec, encoding="utf-8") as f:
            sp = json.load(f)
        board.mat = [tuple(t) for t in sp.get("material_picks") or []]
        board.heads = heads_from_spec(sp)
        for h in board.heads:
            h.pop("mark_bundle", None)
        board.mat_done = bool(board.mat)
        n = len(board.mat) + len(board.heads)
    if os.path.exists(p_clicks):
        try:
            with open(p_clicks, encoding="utf-8") as f:
                raw = json.load(f)
            board.clicks = [cl for cl in raw
                            if (cl.get("모드") or cl.get("mode")) != "기타"
                            and cl.get("동작") != "문양대기"]
        except Exception:  # noqa: BLE001
            board.clicks = []
    return n


def make_board(world, knobs):
    return Board(world, knobs)
