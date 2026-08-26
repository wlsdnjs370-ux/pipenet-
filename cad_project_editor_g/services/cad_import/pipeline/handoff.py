# -*- coding: utf-8 -*-
"""찍기에서 만든 선택 전 DXF 세계를 stage1에 넘기는 폐기 가능 handoff.

원본은 DXF와 찍은 스펙이다. 이 파일은 SQLite 숫자/문자 열만 사용하며
pickle이나 Python 객체 역직렬화를 하지 않는다. 화면 없음.
"""
import hashlib
import importlib
import inspect
import os
import re
import sqlite3
import sys
import time


FORMAT = "stage1-world-sqlite-v2"


def import_write_root():
    """찍기·캐시·손질 JSON 쓰기 루트.

    소스 실행은 저장소의 docs/import. 빌드본은 Program Files 에 못 쓰므로
    로그·라이브러리와 같은 %LOCALAPPDATA%\\K-Fire 아래.
    """
    if getattr(sys, "frozen", False):
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        return os.path.join(base, "K-Fire", "cad_import")
    return os.path.join("docs", "import")


def default_edits_dir():
    """유저손질 JSON 폴더. DXF 샘플 읽기 경로(DWG_DIR)와 분리한다."""
    return os.path.join(import_write_root(), "DWG")


def pick_out_dir():
    """찍은스펙·자동백업 폴더."""
    return os.path.join(import_write_root(), "0단계_새찍기")


OUT_DIR = pick_out_dir()
_PREP_NAMES = ("_pairs", "read_dxf", "World", "explode")
# 이보다 빠른 DXF 준비는 SQLite 저장·검증 비용과 비슷해 총 대기가 줄지 않았다.
# 도면 내용/이름이 아닌, 찍기에서 이미 잰 실제 중복 비용만으로 산출 여부를 정한다.
MIN_PREP_SECONDS = 0.3


def _sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _prep_digest(module_name):
    mod = importlib.import_module(module_name)
    h = hashlib.sha256()
    for name in _PREP_NAMES:
        h.update(inspect.getsource(getattr(mod, name)).encode("utf-8"))
        h.update(b"\0")
    return h.hexdigest()


def _compatible_prep_digest():
    """찍기와 1단계가 같은 stage1 준비를 쓴다. 옛 poc6/_tmp 지문은 폐기."""
    return _prep_digest("services.cad_import.pipeline.stage1")


def _source_meta(source_path):
    source_path = os.path.normcase(os.path.abspath(source_path))
    st = os.stat(source_path)
    return {
        "source_path": source_path,
        "source_size": str(st.st_size),
        "source_mtime_ns": str(st.st_mtime_ns),
        "source_sha256": _sha256_file(source_path),
    }


def _circ_rows(world):
    """원 + 호. 호의 sa/sweep 는 있을 때만. 없으면 NULL (추정 안 함)."""
    rows = []
    for i, (layer, color, x, y, r) in enumerate(world.circles):
        rows.append((0, i, layer, color, x, y, r, None, None))
    angs = getattr(world, "arc_ang", ())
    for i, (layer, color, x, y, r) in enumerate(world.arcs):
        ang = angs[i] if i < len(angs) else None
        sa = sw = None
        if ang:
            sa, sw = float(ang[0]), float(ang[1])
        rows.append((1, i, layer, color, x, y, r, sa, sw))
    return rows


def handoff_path(key):
    safe = re.sub(r"[^0-9A-Za-z가-힣_.-]+", "_", str(key)).strip("._")
    if not safe:
        safe = "drawing"
    suffix = hashlib.sha256(str(key).encode("utf-8")).hexdigest()[:10]
    return os.path.join(OUT_DIR, f"{safe}_{suffix}_stage1_world.sqlite3")


def save_world(key, source_path, world):
    """기존 World를 원자적으로 저장한다. 실패해도 원본/찍은 스펙은 그대로다."""
    os.makedirs(OUT_DIR, exist_ok=True)
    path = handoff_path(key)
    tmp = f"{path}.tmp-{os.getpid()}-{time.time_ns()}"
    meta = {
        "format": FORMAT,
        "prep_sha256": _compatible_prep_digest(),
        **_source_meta(source_path),
        "n_segs": str(len(world.segs)),
        "n_raw_segs": str(len(world.raw_segs)),
        "n_circles": str(len(world.circles)),
        "n_arcs": str(len(world.arcs)),
        "n_texts": str(len(world.texts)),
    }
    try:
        con = sqlite3.connect(tmp)
        try:
            con.executescript(
                "PRAGMA journal_mode=OFF;"
                "PRAGMA synchronous=OFF;"
                "CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);"
                "CREATE TABLE seg (kind INTEGER NOT NULL, ord INTEGER NOT NULL,"
                " layer TEXT NOT NULL, color INTEGER NOT NULL,"
                " x1 REAL NOT NULL, y1 REAL NOT NULL,"
                " x2 REAL NOT NULL, y2 REAL NOT NULL,"
                " PRIMARY KEY(kind, ord));"
                "CREATE TABLE circ (kind INTEGER NOT NULL, ord INTEGER NOT NULL,"
                " layer TEXT NOT NULL, color INTEGER NOT NULL,"
                " x REAL NOT NULL, y REAL NOT NULL, r REAL NOT NULL,"
                " sa REAL, sweep REAL,"
                " PRIMARY KEY(kind, ord));"
                "CREATE TABLE txt (ord INTEGER PRIMARY KEY, layer TEXT NOT NULL,"
                " color INTEGER NOT NULL, x REAL NOT NULL, y REAL NOT NULL,"
                " h REAL NOT NULL, text TEXT NOT NULL);"
            )
            con.executemany("INSERT INTO meta VALUES (?, ?)", meta.items())
            con.executemany(
                "INSERT INTO seg VALUES (?,?,?,?,?,?,?,?)",
                ((kind, i, layer, color, a[0], a[1], b[0], b[1])
                 for kind, rows in enumerate((world.segs, world.raw_segs))
                 for i, (layer, color, a, b) in enumerate(rows)),
            )
            con.executemany(
                "INSERT INTO circ VALUES (?,?,?,?,?,?,?,?,?)",
                _circ_rows(world),
            )
            con.executemany(
                "INSERT INTO txt VALUES (?,?,?,?,?,?,?)",
                ((i, *row) for i, row in enumerate(world.texts)),
            )
            con.commit()
        finally:
            con.close()
        os.replace(tmp, path)
        return path
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def _meta_matches(meta, source_path):
    """캐시가 이 원본의 것인가.

    ★mtime 은 «빠른 길» 이지 판정 권한이 아니다. 예전에는 mtime 불일치를 곧
      거부로 썼는데, 같은 파일을 다시 올리기만 해도 내용이 한 바이트도 안
      바뀐 채 mtime 만 새로 찍힌다 — 그러면 멀쩡한 캐시를 통째로 버렸다.

      실측(B1F): size·sha256 이 모두 같은데 mtime 만 174초 어긋나 캐시가
      기각됐고, 그 안에 있던 치수 텍스트 3,168행(치수로 읽히는 것 533개)이
      함께 사라져 관경이 **100% 별표1 폴백**이 됐다. 화면에는 아무 말도
      나오지 않는다 — 관경표가 조용히 규약값으로만 채워질 뿐이다.

      내용 동일성의 근거는 sha256 하나면 충분하다. mtime 이 같으면 그것으로
      115MB 해싱을 건너뛰고(빠른 길), 다르면 해싱해서 내용으로 판정한다.
    """
    if meta.get("format") != FORMAT:
        return False
    if meta.get("prep_sha256") != _compatible_prep_digest():
        return False
    source_path = os.path.normcase(os.path.abspath(source_path))
    st = os.stat(source_path)
    if meta.get("source_path") != source_path:
        return False
    if meta.get("source_size") != str(st.st_size):
        return False
    if meta.get("source_mtime_ns") == str(st.st_mtime_ns):
        return True          # 빠른 길 — 크기·mtime 이 같으면 해싱하지 않는다
    return meta.get("source_sha256") == _sha256_file(source_path)


def load_world(key, source_path, world_type):
    """검증된 handoff를 World로 복원한다. 불일치/손상이면 None."""
    path = handoff_path(key)
    if not os.path.exists(path):
        return None
    try:
        con = sqlite3.connect(f"file:{os.path.abspath(path)}?mode=ro", uri=True)
        try:
            con.execute("PRAGMA query_only=ON")
            meta = dict(con.execute("SELECT key, value FROM meta"))
            if not _meta_matches(meta, source_path):
                print("[handoff] 입력/준비 지문 불일치 — 기존 stage1 사용")
                return None
            world = world_type()
            segs = (world.segs, world.raw_segs)
            for kind, _ord, layer, color, x1, y1, x2, y2 in con.execute(
                    "SELECT kind,ord,layer,color,x1,y1,x2,y2"
                    " FROM seg ORDER BY kind,ord"):
                segs[kind].append((layer, color, (x1, y1), (x2, y2)))
            circs = (world.circles, world.arcs)
            for kind, _ord, layer, color, x, y, r, sa, sweep in con.execute(
                    "SELECT kind,ord,layer,color,x,y,r,sa,sweep"
                    " FROM circ ORDER BY kind,ord"):
                circs[kind].append((layer, color, x, y, r))
                if kind == 1:
                    if sa is not None and sweep is not None:
                        world.arc_ang.append((float(sa), float(sweep)))
                    else:
                        world.arc_ang.append(None)
            world.texts.extend(
                (layer, color, x, y, h, text)
                for _ord, layer, color, x, y, h, text in con.execute(
                    "SELECT ord,layer,color,x,y,h,text FROM txt ORDER BY ord")
            )
        finally:
            con.close()
        counts = (
            len(world.segs), len(world.raw_segs), len(world.circles),
            len(world.arcs), len(world.texts),
        )
        expected = tuple(int(meta[f"n_{name}"]) for name in
                         ("segs", "raw_segs", "circles", "arcs", "texts"))
        if counts != expected:
            raise ValueError(f"행 수 불일치: {counts} != {expected}")
        return world
    except Exception as exc:
        print(f"[handoff] 손상/읽기 실패 — 기존 stage1 사용: {exc}")
        return None
