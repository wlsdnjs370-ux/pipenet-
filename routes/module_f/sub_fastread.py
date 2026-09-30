# -*- coding: utf-8 -*-
"""[오너 2026-09-22] 계통도 · 기계실 칸 — 도면 «읽기» 를 빠르게 (결과는 그대로).

그림 37 · 38 · 39 (docs/images/iso_plan_preserve).

B1F 현장조사 평면도(116 MB)를 계통도 칸에 올리면 약 30초(서버 PC 약 38초)가
걸렸다. 그중 20.3초가 `ezdxf.readfile` 이다. 까닭은 파일의 74%(85.8 MB)가
블록 정의이고, 그 1,532개 중 1,246개(83.5 MB)는 도면 어디에도 놓이지 않는데도
ezdxf 가 전부 CAD 객체로 만들기 때문이다. 모듈 A 의 `parse_dxf_for_view` 는
도면 본체에서 출발해 «놓인» 블록만 펼치므로, 안 쓰는 블록은 읽어도 결과에
한 글자도 들어가지 않는다.

여기서 하는 일은 둘이다. **모듈 A 의 코드는 한 줄도 고치지 않는다.**

① 안 쓰는 블록 정의는 «속을 비운 사본» 으로 읽는다
   파일을 글자로만 훑어(0.9초) 도면이 쓰는 블록을 찾는다:
     · 도면 본체(ENTITIES)의 도형이 이름(그룹 코드 2)이나 번호(340~349 · 1005)로
       가리키는 블록 → 그 블록 안의 도형이 또 가리키는 블록 → … 끝까지
     · 치수 스타일(DIMSTYLE)이 가리키는 화살표 블록, 멀티지시선 스타일
       (MLEADERSTYLE)이 가리키는 화살표 · 내용 블록
     · 도면 틀(*Model_Space · *Paper_Space) 블록과 화살표 블록(이름이 «_» 로
       시작 — _ClosedFilled 등)은 늘 남긴다. 멀티지시선은 화살표를 따로 안
       적으면 ezdxf 가 _ClosedFilled 를 **이름으로** 찾아 쓴다 — 파일 안에 그
       블록을 가리키는 글자가 없다(합성 도면 시험에서 잡았다).
   나머지 블록은 시작(BLOCK) · 끝(ENDBLK) 표시만 남기고 속을 비운 임시 사본을
   만들어 A 의 파서에 **그대로** 넘긴다 — 블록이 없어지지 않고 속만 빈다.
   ★모르면 비우지 않는다. 모르는 도형 종류(3D 솔리드 · 프록시 객체 …)가 도면
     본체나 쓰는 블록에 하나라도 있으면, 블록 짝이 안 맞으면, 줄끝이 이상하면,
     이진 DXF 면 — 전부 «원본 그대로» 읽는다. 사본으로 읽다 실패해도 원본으로.
   ★실측(2026-09-22, 도면 75장 · 1.5 GB): 사본으로 읽은 도형 목록 · 레이어 표 ·
     범위가 원본과 한 글자도 같음 73장 · 모르는 종류가 있어 원본 2장 · 다름 0장.
     A 가 원본을 읽으며 연 블록은 전부 «쓰는 블록» 안에 있었다.
     B1F 읽기 21.9초 → 8.5초.

② 같은 도면은 기억한다 (data/parse_cache — 저장소에 안 올라가는 폴더)
   열쇠 = 파일 내용의 지문(blake2b) + 코드 판번호(관련 소스 파일 내용의 지문 ·
   ezdxf · 파이썬 판). 같은 파일이면 이름이 달라도 같은 열쇠다.
   ★코드가 바뀌면 판번호가 달라져 기억을 쓰지 않고 새로 읽는다 — 옛 규칙의
     결과가 섞이지 않는다. 판번호는 이 모듈을 부를 때(서버가 뜰 때) 한 번 잰다.
   ★기억이 깨졌거나 못 읽으면 조용히 새로 읽는다 — 속도만 손해다.
"""
from __future__ import annotations

import hashlib
import os
import pickle
import re
import sys
import tempfile
import threading
import time
import zlib
from collections import defaultdict
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]

# 켜고 끄는 스위치 — 급할 때 코드를 안 고치고 종전 방식으로 돌릴 수 있게.
_STRIP_ON = os.environ.get("MODULE_F_SUB_STRIP", "1").strip() != "0"
_CACHE_ON = os.environ.get("MODULE_F_SUB_CACHE", "1").strip() != "0"
CACHE_DIR = Path(os.environ.get("MODULE_F_SUB_CACHE_DIR")
                 or (_ROOT / "data" / "parse_cache"))
CACHE_VERSION = 1

# 비울 만한가 — 블록 정의가 파일의 15% 도 안 되면 훑는 값도 아깝다. 훑어 보고
# 비울 몫이 10% 도 안 되면 사본을 만들지 않는다(원본을 그대로 읽는다).
STRIP_MIN_BLOCKS_SHARE = 0.15
STRIP_MIN_DROP_SHARE = 0.10
# 기억할 만한가 — 이보다 빨리 끝나는 도면은 기억 파일을 만들지 않는다.
CACHE_MIN_PARSE_S = 1.5
CACHE_MIN_TRACE_S = 1.0
CACHE_KEEP = 40            # 종류(도형 목록 · ★추적)마다 남길 기억 파일 수

# A 의 `parse_dxf_for_view` 가 블록을 여는 길을 아는 도형 종류 — 이 밖의 종류가
# 도면 본체나 쓰는 블록에 있으면 비우지 않는다.
#   INSERT        이름(2)으로 블록을 연다
#   DIMENSION 류  이름(2)의 치수 블록을 연다
#   LEADER        화살표 블록 — 치수 스타일(341~344) 또는 덧붙은 값(1005)
#   MULTILEADER   화살표 · 내용 블록 — 도형의 340~345 · 멀티지시선 스타일
#   ACAD_TABLE    이름(2) · 번호(34x)
#   그 밖         블록을 열지 않는다
_KNOWN_TYPES = frozenset({
    b"LINE", b"ARC", b"CIRCLE", b"LWPOLYLINE", b"POLYLINE", b"VERTEX", b"SEQEND",
    b"POINT", b"INSERT", b"ATTRIB", b"ATTDEF", b"TEXT", b"MTEXT", b"SPLINE",
    b"ELLIPSE", b"HATCH", b"SOLID", b"3DFACE", b"TRACE", b"DIMENSION",
    b"ARC_DIMENSION", b"LARGE_RADIAL_DIMENSION", b"LEADER", b"MULTILEADER",
    b"MLEADER", b"ACAD_TABLE", b"RAY", b"XLINE", b"WIPEOUT", b"VIEWPORT", b"IMAGE",
    b"MLINE"})

# DXF 는 한 줄에 그룹 코드, 다음 줄에 값이다. 아래 틀은 «그룹 코드 줄» 을 찾고
# 값 줄은 앞보기(lookahead)로만 읽는다 — 값 줄을 먹어 버리면, 값이 "0" 인 줄
# 바로 뒤의 진짜 코드 줄을 놓친다(실측: B1F 에서 도형 392개가 그렇게 빠졌다).
# 찾은 줄이 정말 코드 줄인지는 줄 번호의 짝홀로 가른다(`_sweep`).
_RE_BLK = re.compile(rb"\n[ \t]*0[ \t]*(?=\r?\n(BLOCK|ENDBLK)[ \t]*\r?\n)")
_RE_ENT = re.compile(rb"\n[ \t]*0[ \t]*(?=\r?\n([^\r\n]*))")
# 가리키는 값: 이름(2) · 번호(340~349) · 덧붙은 값의 번호(1005) · 글자(1000 —
# R12 치수 덮어쓰기는 화살표 블록을 이름으로 적는다).
_RE_REF = re.compile(rb"\n[ \t]*(2|34[0-9]|1005|1000)[ \t]*(?=\r?\n([^\r\n]*))")


def _sweep(raw: bytes, rx, a: int, b: int, base: int = 0) -> list:
    """`raw[a:b]` 에서 `rx` 에 걸린 것 중 **코드 줄** 인 것만.

    `base` 는 코드 줄이 시작하는 자리(또는 파일 첫머리 0)다. 그 뒤로 짝수 번째
    줄이 코드 줄, 홀수 번째 줄이 값 줄이다 — 값이 "0" 이나 "2" 인 줄을 코드로
    잘못 읽지 않는다.
    """
    out = []
    pos, n = base, 0
    cnt = raw.count
    for m in rx.finditer(raw, a, b):
        p = m.start() + 1
        n += cnt(b"\n", pos, p)
        pos = p
        if not n & 1:
            out.append(m)
    return out


def _tags_at(raw: bytes, pos: int, stop: int):
    """코드 줄 `pos` 부터 (코드, 값, 코드 줄 자리) — 짧은 구간(머리말)에만 쓴다."""
    find = raw.find
    while pos < stop:
        e1 = find(b"\n", pos, stop)
        if e1 < 0:
            return
        e2 = find(b"\n", e1 + 1, stop)
        if e2 < 0:
            e2 = stop
        yield raw[pos:e1].strip(), raw[e1 + 1:e2].strip(), pos
        pos = e2 + 1


def _sections(raw: bytes) -> dict | None:
    """{구역 이름: (본문 첫 자리, ENDSEC 코드 줄 자리)} — 짝이 안 맞으면 None."""
    marks = []
    for word in (b"SECTION", b"ENDSEC"):
        i = raw.find(b"\n" + word)
        while i >= 0:
            j = i + 1 + len(word)
            if raw[j:j + 1] in (b"\r", b"\n", b" ", b"\t"):
                ps = raw.rfind(b"\n", 0, i)
                if raw[ps + 1:i].strip() == b"0":
                    marks.append((ps + 1, word))
            i = raw.find(b"\n" + word, i + 1)
    marks.sort()
    pos, n = 0, 0
    out: dict = {}
    cur = None
    for code_start, word in marks:
        n += raw.count(b"\n", pos, code_start)
        pos = code_start
        if n & 1:
            continue                       # 값 줄이 "0" 이었다 — 코드 줄이 아니다
        if word == b"SECTION":
            tags = list(_tags_at(raw, code_start, min(len(raw), code_start + 512)))
            if len(tags) < 2 or tags[1][0] != b"2":
                return None
            e = code_start
            for _ in range(4):             # "0 / SECTION / 2 / 이름" 다음 줄
                e = raw.find(b"\n", e) + 1
                if e <= 0:
                    return None
            cur = (tags[1][1].upper(), e)
        elif cur is not None:
            if cur[0] in out:
                return None
            out[cur[0]] = (cur[1], code_start)
            cur = None
    return out


def plan_strip(raw: bytes) -> tuple[dict | None, str]:
    """비울 블록을 정한다 → (계획 | None, 까닭).

    계획::

        {"blocks": [[이름(소문자), 머리 자리, 속 첫 자리, ENDBLK 자리, 주인 번호]…],
         "drop": [비울 블록…], "keep": 쓰는 블록 이름 수, "drop_bytes": 비울 바이트}
    """
    if raw[:22] == b"AutoCAD Binary DXF\r\n\x1a\x00":
        return None, "이진 DXF"
    if raw.count(b"\r") != raw.count(b"\r\n"):
        return None, "줄끝이 섞여 있음"
    secs = _sections(raw)
    if not secs or b"BLOCKS" not in secs or b"ENTITIES" not in secs:
        return None, "구역을 못 가름"
    b0, b1 = secs[b"BLOCKS"]
    if (b1 - b0) < STRIP_MIN_BLOCKS_SHARE * len(raw):
        return None, "블록 정의가 작음"
    bm = _sweep(raw, _RE_BLK, b0 - 1, b1)
    if len(bm) % 2:
        return None, "블록 짝이 안 맞음"
    blocks = []
    for k in range(0, len(bm), 2):
        mb, me = bm[k], bm[k + 1]
        if mb.group(1) != b"BLOCK" or me.group(1) != b"ENDBLK":
            return None, "블록 짝이 안 맞음"
        hs, es = mb.start() + 1, me.start() + 1
        name = owner = None
        body = es
        first = True
        for code, val, at in _tags_at(raw, hs, es):
            if first:
                first = False
            elif code == b"0":
                body = at                  # 머리말 다음 첫 도형
                break
            elif code == b"2" and name is None:
                name = val
            elif code == b"330" and owner is None:
                owner = val
        if name is None:
            return None, "이름 없는 블록"
        blocks.append([name.lower(), hs, body, es, owner.upper() if owner else None])
    names = {bk[0] for bk in blocks}
    h2n = {bk[4]: bk[0] for bk in blocks if bk[4]}
    roots: set = set()
    if b"TABLES" in secs:
        a, b = secs[b"TABLES"]
        tl = _sweep(raw, _RE_ENT, a - 1, b)
        dimstyles = []
        for k, m in enumerate(tl):
            ty = m.group(1).strip()
            if ty not in (b"BLOCK_RECORD", b"DIMSTYLE"):
                continue
            end = tl[k + 1].start() + 1 if k + 1 < len(tl) else b
            tags = [(c, v) for c, v, _ in _tags_at(raw, m.start() + 1, end)]
            if ty == b"BLOCK_RECORD":
                hnd = next((v for c, v in tags if c == b"5"), None)
                nm = next((v for c, v in tags if c == b"2"), None)
                if hnd and nm:
                    h2n[hnd.upper()] = nm.lower()
            else:
                dimstyles.append(tags)
        # 치수 스타일은 블록 기록표를 다 읽은 뒤에 본다(표 순서상 DIMSTYLE 이 앞선다).
        for tags in dimstyles:             # 화살표 블록 — R12 는 이름, 그 뒤는 번호
            for _c, v in tags[1:]:
                if v.lower() in names:
                    roots.add(v.lower())
                elif v.upper() in h2n:
                    roots.add(h2n[v.upper()])

    def refs(ms) -> set:
        out = set()
        for m in ms:
            v = m.group(2).strip()
            if m.group(1) in (b"2", b"1000"):
                v = v.lower()
                if v in names:
                    out.add(v)
            if m.group(1) != b"2":
                w = h2n.get(v.upper())
                if w is not None:
                    out.add(w)
        return out

    e0, e1 = secs[b"ENTITIES"]
    roots |= refs(_sweep(raw, _RE_REF, e0 - 1, e1))
    if b"OBJECTS" in secs:                 # 멀티지시선 스타일의 화살표 · 내용 블록
        a, b = secs[b"OBJECTS"]
        if raw.find(b"MLEADERSTYLE", a, b) >= 0:
            ol = _sweep(raw, _RE_ENT, a - 1, b)
            for k, m in enumerate(ol):
                if m.group(1).strip() == b"MLEADERSTYLE":
                    end = ol[k + 1].start() + 1 if k + 1 < len(ol) else b
                    roots |= refs(_sweep(raw, _RE_REF, m.start(), end,
                                         base=m.start() + 1))
    for bk in blocks:
        if bk[0].startswith((b"*model_space", b"*paper_space", b"_")):
            roots.add(bk[0])
    by_name = defaultdict(list)
    for bi, bk in enumerate(blocks):
        by_name[bk[0]].append(bi)
    reach: set = set()
    stack = list(roots)
    while stack:
        nm = stack.pop()
        if nm in reach:
            continue
        reach.add(nm)
        for bi in by_name.get(nm, ()):
            bk = blocks[bi]
            if bk[3] > bk[2]:
                ms = _sweep(raw, _RE_REF, bk[2] - 1, bk[3], base=bk[2])
                stack.extend(r for r in refs(ms) if r not in reach)
    types = {m.group(1).strip() for m in _sweep(raw, _RE_ENT, e0 - 1, e1)}
    for bk in blocks:
        if bk[0] in reach and bk[3] > bk[2]:
            types |= {m.group(1).strip()
                      for m in _sweep(raw, _RE_ENT, bk[2] - 1, bk[3], base=bk[2])}
    bad = sorted(t for t in types if t not in _KNOWN_TYPES)
    if bad:
        return None, "모르는 도형 종류 " + ", ".join(
            t.decode("ascii", "replace") for t in bad[:4])
    drop = [bk for bk in blocks if bk[0] not in reach and bk[3] > bk[2]]
    drop_bytes = sum(bk[3] - bk[2] for bk in drop)
    return {"blocks": blocks, "drop": drop, "keep": len(reach),
            "drop_bytes": drop_bytes}, "ok"


def strip_bytes(raw: bytes, plan: dict) -> bytes:
    """비울 블록의 속(머리말 다음 ~ ENDBLK 앞)만 빼고 나머지는 한 바이트도 안 바꾼다."""
    parts = []
    last = 0
    for bk in sorted(plan["drop"], key=lambda b: b[1]):
        parts.append(raw[last:bk[2]])
        last = bk[3]
    parts.append(raw[last:])
    return b"".join(parts)


# ───────────────────────────────────────────── 기억(캐시)

_VIEW_SOURCES = ("remote30_prototype.py", "sprinkler_remote30_extractor.py",
                 "core/remote30_constants.py", "core/remote30_graph.py",
                 "core/fitting_rules.py", "routes/module_f/sub_fastread.py")
_TRACE_SOURCES = _VIEW_SOURCES + ("routes/module_f/sub_trace.py",)


def _fingerprint(sources) -> str:
    h = hashlib.blake2b(digest_size=8)
    h.update(f"v{CACHE_VERSION}|py{sys.version_info[0]}.{sys.version_info[1]}".encode())
    try:
        import ezdxf
        h.update(f"|ezdxf{getattr(ezdxf, '__version__', '?')}".encode())
    except Exception:  # noqa: BLE001
        h.update(b"|ezdxf?")
    for rel in sources:
        p = _ROOT / rel
        try:
            h.update(rel.encode() + b"=" + p.read_bytes())
        except OSError:
            h.update(rel.encode() + b"=missing")
    return h.hexdigest()


# 판번호는 이 모듈을 부를 때 한 번 잰다 — 서버가 뜰 때 읽은 코드와 같은 것을
# 가리키게 하려는 것이다(api_slot 이 서버 시작 때 이 모듈을 부른다).
FP_VIEW = _fingerprint(_VIEW_SOURCES)
FP_TRACE = _fingerprint(_TRACE_SOURCES)
FP_H_VIEW = _fingerprint(_VIEW_SOURCES + (
    "src/pipenet_converter/render/dxf_source.py",
    "src/pipenet_converter/render/dxf_text.py",
))
_cache_lock = threading.Lock()


def content_key(raw: bytes) -> str:
    return hashlib.blake2b(raw, digest_size=16).hexdigest()


def _cache_path(kind: str, key: str, fp: str) -> Path:
    return CACHE_DIR / f"sub{kind}_v{CACHE_VERSION}_{key}_{fp}.pkl.z"


def _cache_load(kind: str, key: str | None, fp: str):
    if not (_CACHE_ON and key):
        return None
    p = _cache_path(kind, key, fp)
    try:
        if not p.is_file():
            return None
        return pickle.loads(zlib.decompress(p.read_bytes()))
    except Exception as exc:  # noqa: BLE001 — 깨진 기억은 버리고 새로 읽는다
        print(f"  (기억 파일을 못 읽어 새로 읽습니다: {p.name} — {exc})")
        return None


def _cache_store(kind: str, key: str | None, fp: str, obj) -> bool:
    if not (_CACHE_ON and key):
        return False
    p = _cache_path(kind, key, fp)
    tmp = p.with_name(p.name + f".{os.getpid()}.{threading.get_ident()}.tmp")
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        data = zlib.compress(pickle.dumps(obj, protocol=pickle.HIGHEST_PROTOCOL), 1)
        tmp.write_bytes(data)
        os.replace(tmp, p)
    except Exception as exc:  # noqa: BLE001 — 기억 못 해도 결과는 그대로다
        print(f"  (기억 파일을 못 남겼습니다 — 다음에도 새로 읽습니다: {exc})")
        try:
            tmp.unlink()
        except OSError:
            pass
        return False
    _prune(kind, key, p)
    return True


def _prune(kind: str, key: str, keep: Path) -> None:
    """같은 도면의 옛 판 기억과, 개수 상한을 넘는 오래된 기억을 지운다."""
    def age(q):
        try:
            return q.stat().st_mtime
        except OSError:
            return 0.0
    with _cache_lock:
        try:
            files = sorted(CACHE_DIR.glob(f"sub{kind}_v*_*.pkl.z"), key=age,
                           reverse=True)
        except OSError:
            return
        stale = [q for q in files if q != keep and f"_{key}_" in q.name]
        rest = [q for q in files if q not in stale]
        for q in stale + rest[CACHE_KEEP:]:
            try:
                q.unlink()
            except OSError:
                pass


# ───────────────────────────────────────────── 읽기

def read_view(dxf_path, *, source_display: bool = False) -> dict:
    """계통도 · 기계실 도면 한 장 → A 의 `parse_dxf_for_view` 와 **같은** 결과.

    반환::

        {"entities": [...], "parsed": {...},      # parse_dxf_for_view 그대로
         "key": 파일 내용 지문 | None,
         "how": "cache" | "strip" | "plain",
         "why": 사본을 안 쓴 까닭(plain) | None,
         "keep": 쓰는 블록 수, "drop": 비운 블록 수, "drop_mb": 비운 MB,
         "t_plan": 훑기 초, "seconds": 전체 초}
    """
    from remote30_prototype import parse_dxf_for_view
    cache_kind, fingerprint = ("hview", FP_H_VIEW) if source_display else ("view", FP_VIEW)

    def parse(path):
        if not source_display:
            return parse_dxf_for_view(path, include_hidden_layers=True)
        import ezdxf
        from src.pipenet_converter.render.dxf_source import source_display as display
        document = ezdxf.readfile(path)
        # Keep calculation entities equivalent. H display is a separate
        # projection of the SAME loaded document, never a graph input.
        parsed = parse_dxf_for_view(path, include_hidden_layers=True, document=document)
        parsed["source_display"] = display(document)
        return parsed
    t0 = time.perf_counter()
    info: dict = {"key": None, "how": "plain", "why": None, "keep": 0, "drop": 0,
                  "drop_mb": 0.0, "t_plan": 0.0}
    try:
        raw = Path(dxf_path).read_bytes()
    except OSError:
        raw = None                         # 원본 파서가 같은 오류를 그대로 낸다
    if raw is not None:
        info["key"] = content_key(raw)
        got = _cache_load(cache_kind, info["key"], fingerprint)
        if got is not None:
            entities, meta = got
            parsed = dict(meta)
            parsed["entities"] = entities
            info.update(entities=entities, parsed=parsed, how="cache",
                        seconds=time.perf_counter() - t0)
            return info
    tmp = None
    if raw is not None and _STRIP_ON:
        t = time.perf_counter()
        try:
            plan, why = plan_strip(raw)
        except Exception as exc:  # noqa: BLE001 — 훑다 막히면 원본 그대로
            plan, why = None, f"훑기 실패 {exc}"
        info["t_plan"] = time.perf_counter() - t
        if plan is not None and plan["drop_bytes"] < STRIP_MIN_DROP_SHARE * len(raw):
            plan, why = None, "비울 것이 적음"
        info["why"] = None if plan is not None else why
        if plan is not None:
            out = strip_bytes(raw, plan)
            fd, tmp = tempfile.mkstemp(prefix="fncad_sub_", suffix=".dxf")
            try:
                with os.fdopen(fd, "wb") as fh:
                    fh.write(out)
            except OSError as exc:
                info["why"] = f"임시 사본을 못 씀 {exc}"
                _unlink(tmp)
                tmp = None
            else:
                info.update(keep=plan["keep"], drop=len(plan["drop"]),
                            drop_mb=plan["drop_bytes"] / 1e6)
            del out
    del raw                                # 파서가 메모리를 쓰기 전에 놓는다
    t = time.perf_counter()
    parsed = None
    if tmp is not None:
        try:
            parsed = parse(tmp)
            info["how"] = "strip"
        except Exception as exc:  # noqa: BLE001 — 사본이 안 되면 원본으로
            info["why"] = f"사본 읽기 실패 {exc}"
            info.update(keep=0, drop=0, drop_mb=0.0)
            parsed = None
        finally:
            _unlink(tmp)
    if parsed is None:
        parsed = parse(dxf_path)
        info["how"] = "plain"
    t_parse = time.perf_counter() - t
    entities = parsed.get("entities") or []
    if t_parse >= CACHE_MIN_PARSE_S:
        meta = {k: v for k, v in parsed.items() if k != "entities"}
        _cache_store(cache_kind, info["key"], fingerprint, (entities, meta))
    info.update(entities=entities, parsed=parsed,
                seconds=time.perf_counter() - t0)
    return info


def _unlink(path) -> None:
    try:
        os.remove(path)
    except OSError:
        pass


def load_trace(key: str | None):
    """기억해 둔 (레이어 표 행, ★추적 계획) — 없으면 None."""
    got = _cache_load("trace", key, FP_TRACE)
    if not (isinstance(got, tuple) and len(got) == 2):
        return None
    return got


def store_trace(key: str | None, rows, tr, seconds: float) -> bool:
    """★추적 계획을 기억한다 — 금방 끝난 도면은 남기지 않는다."""
    if seconds < CACHE_MIN_TRACE_S or not tr:
        return False
    keep = {k: v for k, v in tr.items() if k != "payload"}
    return _cache_store("trace", key, FP_TRACE, (rows, keep))


def describe(info: dict) -> str | None:
    """작업 기록에 남길 한 줄 — 무엇으로 읽었나."""
    if info.get("how") == "cache":
        return f"같은 도면을 전에 읽어 둔 것을 씁니다 ({info['seconds']:.1f}s)"
    if info.get("how") == "strip":
        return (f"도면에 안 쓰이는 블록 정의 {info['drop']:,}개"
                f"({info['drop_mb']:.1f} MB)는 비우고 읽었습니다 — "
                f"쓰는 블록 {info['keep']:,}개 (훑기 {info['t_plan']:.1f}s)")
    why = info.get("why")
    if why and why not in ("블록 정의가 작음", "비울 것이 적음"):
        return f"블록을 비우지 않고 원본 그대로 읽었습니다 — {why}"
    return None
