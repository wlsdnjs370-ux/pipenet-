# -*- coding: utf-8 -*-
"""[요소속성 수정카드] 사람이 덮은 값 한 저장소 — 안정 키로 담고 파일로 남긴다.

지시서 `ModuleF_요소속성_수정카드_지시서.md` · 그림 14·15.

■ 규칙 여섯 (그림 15)

  1. **값만 고친다** — 노드·간선을 만들거나 지우지 않는다. 위상은 손질이 주인.
  2. **안정 키로 저장한다** — kfp 이름(N12·P7)은 재계산마다 바뀐다.
     실측(대명동 K 30→20): 살아남은 배관 111개가 **전부** 이름이 바뀌었다.
     board 노드쌍·disk 번호로 담고 빌드마다 번역한다.
  3. **정해진 자리에서만 적용한다** — 손질·선정·planar·engine 은 이 값을
     모른다. 되먹임 없음.
  4. **못 옮긴 수정은 말한다** — 조용히 버리지 않는다(`missed`).
  5. **원값을 지우지 않는다** — 카드가 「도면/규칙 값 → 사람 값 · 사유 · 시각」.
  6. **이미 있는 셋을 흡수한다** — 관경(F-11c) · 부속(§18) · 읽기 카드(F-12).

■ ★적용 자리가 **셋**인 이유 (오너 2026-09-14 확정)

  §2 계측으로 드러났다 — 속성이 두 곳에 산다. 표에 칸이 있는 것(관경·길이·C·
  관종·표고)과 kfp 메타에만 있는 것(거칠기·등가길이·K 값·필요압력)이 갈린다.
  거기에 오너의 한 줄이 하나를 더 만들었다:

      「20m 배관을 2m로 바꾸었는데, 아이소가 그대로인건 말이안되니까」

  회랑 좌표는 사슬로 **만든다** — `p(자식) = p(부모) + L·u`(사슬좌표 지시서
  §1-1). 곧 **길이가 좌표를 정한다.** 그러니 덮은 길이를 사슬에 넣으면 좌표가
  따라 움직이고 아이소가 따라 변한다. 그래서:

      ① 사슬 전   길이(평면 배관·세로 토막) → chain_coords 가 이 길이로
                  좌표를 만든다 → **아이소가 따라 변한다**
      ② 표 전     kfp 메타에만 있는 것(거칠기·등가길이·K·필요압력·C·관종)
      ③ 표 직후   표 칸에만 있는 것(관경·표고 등)

  ★이로써 지시서 D4 의 「덮은 길이는 declared_pipes 와 같은 부류(좌표 거리 ≠
    표 길이가 정상)」는 **철회**된다. 사슬이 덮은 길이를 그대로 쓰므로 좌표
    거리 = 표 길이가 유지된다 — 사슬좌표 C2 의 예외가 아니라 **그대로 통과**다.
"""
from __future__ import annotations

import json
import math
import os
import uuid
from collections import defaultdict, deque

VERSION = 1

# 고칠 수 있는 속성 — 값 검증과 함께. 여기 없는 이름은 저장하지 않는다.
#   (kind, field) → (파서, 검사, 사람이 읽는 이름)
KINDS = ("pipe", "node", "head", "vert", "sys", "mr")


def _num(v):
    return float(v)


FIELDS = {
    # 배관 — 평면 배관(board 노드쌍 키)
    ("pipe", "length"): (_num, lambda v: v > 0, "길이(m)"),
    ("pipe", "dia"): (int, lambda v: v > 0, "관경(호칭 mm)"),
    ("pipe", "type"): (str, lambda v: bool(v), "관종"),
    ("pipe", "c"): (_num, lambda v: v > 0, "C 값"),
    ("pipe", "roughness_mm"): (_num, lambda v: v >= 0, "거칠기(mm)"),
    ("pipe", "equivalent_length"): (_num, lambda v: v >= 0, "등가길이(m)"),
    # 절점
    ("node", "elevation"): (_num, lambda v: True, "표고(m)"),
    # 헤드(노즐)
    ("head", "k_factor_si"): (_num, lambda v: v > 0, "K 값"),
    ("head", "required_pressure_bar"): (_num, lambda v: v >= 0, "필요압력(bar)"),
    # ★표고는 헤드도 고칠 수 있다. 상하향 **종류**는 손질이 주인이지만(그것은
    #   위상에 가깝다), 표고는 표의 한 칸이다. 다른 절점은 다 고치는데 노즐만
    #   못 고치면 그것은 규칙이 아니라 구멍이다.
    ("head", "elevation"): (_num, lambda v: True, "표고(m)"),
    # 세로 토막 ①②③④ · 가지 상승 · 알람밸브
    ("vert", "length"): (_num, lambda v: v > 0, "토막 길이(m)"),
    ("vert", "c"): (_num, lambda v: v > 0, "C 값"),
    ("vert", "roughness_mm"): (_num, lambda v: v >= 0, "거칠기(mm)"),
    # 통합 — 계통도 · 기계실 (라벨 키 · §7-1 오너 승인 2026-09-14)
    ("sys", "length"): (_num, lambda v: v > 0, "길이(m)"),
    ("sys", "dia"): (int, lambda v: v > 0, "관경(호칭 mm)"),
    ("sys", "c"): (_num, lambda v: v > 0, "C 값"),
    ("sys", "elevation"): (_num, lambda v: True, "표고(m)"),
    ("mr", "length"): (_num, lambda v: v > 0, "길이(m)"),
    ("mr", "dia"): (int, lambda v: v > 0, "관경(호칭 mm)"),
    ("mr", "c"): (_num, lambda v: v > 0, "C 값"),
    ("mr", "elevation"): (_num, lambda v: True, "표고(m)"),
}

# 사유가 **필수**인 속성 — 오너 2026-09-14: 길이만. 전부 필수로 하면 사람이
# 아무 글자나 적게 되고, 그러면 사유 칸이 뜻을 잃는다.
REASON_REQUIRED = {("pipe", "length"), ("vert", "length"),
                   ("sys", "length"), ("mr", "length")}


# ─────────────────────────────────────────────── 안정 키
def key_pipe(board_i, board_j):
    a, b = int(board_i), int(board_j)
    return ("pipe", min(a, b), max(a, b))


def key_node(board_index):
    return ("node", int(board_index))


def key_head(disk_index):
    return ("head", int(disk_index))


def key_vert(root_kind, root_index, role):
    """세로 토막 — 뿌리(헤드 disk 또는 board 티)와 «역할».

    ★역할은 z 순서로 매긴다. 한 헤드에 토막이 하나뿐이면 role=1 로 끝이지만
      (실측 대명동 26/26 · B1F 28/28), **상하향식은 한 헤드에 ①②③④ 가 걸린다.**
      그 도면이 아직 없다고 나중으로 미루면, 오는 날 조용히 엉뚱한 토막을
      덮는다 — 오너 지시(2026-09-14)대로 지금 넣어 둔다.
      순서는 (아래 z, 위 z) 오름차순 — 엔진이 결정적으로 만드므로 재현된다.
    """
    return ("vert", str(root_kind), int(root_index), int(role))


def key_merge(part, label):
    """통합의 계통도·기계실 요소 — 그쪽 빌더가 붙인 라벨이 유일한 식별자다.

    ★회랑 요소는 board 키를 쓴다(`part=="plan"`). 계통도·기계실은 board 가
      없어 라벨뿐이고, 라벨은 `to_head_tables` 가 오프셋만 더해 옮긴다 —
      **같은 추출 결과 안에서는 안정**이다. 추출을 다시 하면 바뀔 수 있으므로
      그때는 «적용 못 한 수정» 으로 올라간다(규칙 4).
    """
    return (str(part), str(label))


def key_to_json(k):
    return list(k)


def key_from_json(v):
    if not isinstance(v, (list, tuple)) or not v:
        return None
    k = list(v)
    try:
        if k[0] == "pipe":
            return key_pipe(k[1], k[2])
        if k[0] == "node":
            return key_node(k[1])
        if k[0] == "head":
            return key_head(k[1])
        if k[0] == "vert":
            return key_vert(k[1], k[2], k[3])
        if k[0] in ("sys", "mr"):
            return key_merge(k[0], k[1])
        if k[0] == "add":
            return key_added(k[1])          # [§3-3] 사람이 만든 요소
    except (IndexError, TypeError, ValueError):
        return None
    return None


# ─────────────────────────────────────────────── 저장소
def load(sess) -> list:
    return list(sess.get("element_overrides") or ())


def save(sess, rows) -> None:
    sess["element_overrides"] = list(rows)


def put(rows, key, kind, field, new, *, old=None, reason="", at=""):
    """한 항목을 넣거나 갈아 끼운다 — 같은 (키, 속성)은 하나뿐이다."""
    kj = key_to_json(key)
    out = [r for r in rows
           if not (r.get("key") == kj and r.get("field") == field)]
    out.append({"key": kj, "kind": kind, "field": field,
                "old": old, "new": new, "reason": reason, "at": at})
    return out


def drop(rows, key, field):
    kj = key_to_json(key)
    return [r for r in rows
            if not (r.get("key") == kj and r.get("field") == field)]


def validate(kind, field, raw):
    """값 검증 — 서버에서 한다. 밖이면 저장하지 않고 이유를 돌려준다(기준 8)."""
    spec = FIELDS.get((kind, field))
    if spec is None:
        return None, f"고칠 수 없는 속성입니다: {kind}.{field}"
    parse, ok, label = spec
    try:
        v = parse(raw)
    except (TypeError, ValueError):
        return None, f"{label} 값을 읽지 못했습니다: {raw!r}"
    # 이름표에 이미 「값」이 든 것이 있다(「C 값」) — 그대로 붙이면 «C 값 값이»
    # 가 된다. 이름표는 화면과 한 벌이라 여기서 고치지 않고, 문장을 맞춘다.
    lab = label if label.endswith("값") else f"{label} 값"
    if isinstance(v, float) and not math.isfinite(v):
        return None, f"{lab}이 수가 아닙니다."
    if not ok(v):
        return None, f"{lab}이 허용 범위를 벗어납니다: {v}"
    return v, None


def field_label(kind, field):
    spec = FIELDS.get((kind, field))
    return spec[2] if spec else field


# ─────────────────────────────────────────────── 파일 (D2 · write-through)
def path_for(key: str) -> str:
    from services.cad_import.pipeline.handoff import default_edits_dir
    return os.path.join(default_edits_dir(), f"{key}_수리계산수정.json")


def write_file(key: str, rows, ops=None) -> str:
    """값 수정과 위상 수정을 **한 파일**에 — 형제 목록 둘(§3-3).

    `ops=None` 이면 파일에 있던 위상 수정을 그대로 둔다. 값만 고치는 기존
    호출자(라우트 넷)가 위상 수정을 모르고 지우는 일을 막는 자리다 — 한
    저장소를 둘이 쓰면 늘 여기서 사고가 난다.
    """
    p = path_for(key)
    if ops is None:
        ops = read_file_ops(key)
    if not rows and not ops:
        # ★«수정 없음» 은 **빈 파일이 아니라 파일 없음**이다. 0건짜리 파일을
        #   남겨 두면 작업폴더에 정체 모를 파일이 쌓이고, 사람이 그것을 보고
        #   「뭔가 덮여 있나」 하고 뒤진다. 마지막 수정을 지우면 자취도 지운다.
        try:
            os.remove(p)
        except OSError:
            pass
        return p
    os.makedirs(os.path.dirname(p) or ".", exist_ok=True)
    # 위상 수정이 하나도 없으면 **version 1 모양 그대로** 쓴다 — 읽는 자가
    # 옛 판이어도 터지지 않고, 기존 도면의 파일이 이유 없이 바뀌지 않는다.
    body = {"version": VERSION, "items": list(rows)}
    if ops:
        body = {"version": OPS_VERSION, "items": list(rows), "ops": list(ops)}
    with open(p, "w", encoding="utf-8") as f:
        json.dump(body, f, ensure_ascii=False, indent=2)
    return p


def _read_raw(key: str):
    p = path_for(key)
    if not os.path.exists(p):
        return None
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError) as exc:
        print(f"[수정] ★{os.path.basename(p)} 을 읽지 못했습니다 — {exc}")
        return None


def read_file(key: str) -> list:
    raw = _read_raw(key)
    if raw is None:
        return []
    items = raw.get("items") if isinstance(raw, dict) else raw
    return [r for r in (items or []) if isinstance(r, dict)]


def read_file_ops(key: str) -> list:
    """위상 수정 목록. **version 1 파일은 `ops` 없음으로 읽는다**(§3-3).

    읽자마자 2 로 올려 쓰지 않는다 — 열어만 보고 아무것도 안 고쳤는데 파일이
    바뀌어 있으면, 사람은 무엇이 달라졌는지 알 길이 없다.
    """
    raw = _read_raw(key)
    if not isinstance(raw, dict):
        return []
    return [r for r in (raw.get("ops") or []) if isinstance(r, dict)]


# ─────────────────────────────────────────────── 빌드마다 «키 ↔ 이번 이름»
def _ends(pr):
    return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")


def _disk_near(board, vid, disks):
    """board 절점이 어느 헤드 원 안에 있나 — 원의 반지름으로 잰다.

    `hcov` 한 줄은 (x, y, r) 이다. «가장 가까운 원» 이 아니라 «그 원 안» 인지를
    묻는다 — 옆 헤드가 더 가까운 일은 없지만, 아무 원에도 안 들면 None 이어야
    한다(억지로 붙이면 엉뚱한 헤드의 K 값을 덮는다).
    """
    pts = getattr(board, "pts", None) or ()
    try:
        px, py = float(pts[int(vid)][0]), float(pts[int(vid)][1])
    except (IndexError, TypeError, ValueError):
        return None
    best, bd = None, 1e18
    for i, d in enumerate(disks):
        r = float(d[2]) if len(d) > 2 else 0.0
        dd = math.hypot(px - float(d[0]), py - float(d[1]))
        if dd <= r + 1.0 and dd < bd:
            best, bd = i, dd
    return best


def build_index(got, board) -> dict:
    """안정 키 → 이번 빌드의 이름. 번역이지 판정이 아니다(§18 주석 그대로).

    반환 `{"pipe": {key: pid}, "node": {key: nid}, "head": {key: nid},
           "vert": {key: pid}, "roots": {pid: key}}`
    """
    kfp = (got or {}).get("kfp") or {}
    nd = kfp.get("nodes_meta_runtime") or {}
    pr = kfp.get("pipe_data") or {}
    eref = {str(k): v for k, v in ((got or {}).get("edge_ref") or {}).items()
            if v is not None}
    nref = {str(k): int(v) for k, v in ((got or {}).get("node_ref") or {}).items()}
    o = (got or {}).get("origin_mm") or (0.0, 0.0)
    ox, oy = float(o[0]) - 1000.0, float(o[1]) - 1000.0
    disks = list(getattr(board, "disks", None) or ())

    def xyz(nid):
        c = (nd.get(str(nid)) or {}).get("coords") or (0.0, 0.0, 0.0)
        return (float(c[0]), float(c[1]),
                float(c[2]) if len(c) > 2 else 0.0)

    def mm(nid):
        p = xyz(nid)
        return (p[0] * 1000.0 + ox, p[1] * 1000.0 + oy)

    idx = {"pipe": {}, "node": {}, "head": {}, "vert": {}, "roots": {}}

    # ★[§3-3] 쪼갠 조각은 **원래 조각의 board 쌍을 함께 쓴다**(관경이 같아야
    #   하니까). 그러면 한 키에 두 이름이 걸리고, 나중에 들어온 쪽이 이겨
    #   저장된 값 수정이 옆 조각으로 옮겨간다 — 그것이 §5 가 금지하는 바로
    #   그 사고다. 사람이 만든 조각은 이 주소록에서 **뺀다**(그쪽 주소는
    #   `("add", id)` 다).
    born = {str(p) for p in ((got or {}).get("ops_added") or {}).get("pipes") or ()}
    for pid, ref in eref.items():
        if str(pid) in pr and str(pid) not in born:
            idx["pipe"][key_pipe(ref[0], ref[1])] = str(pid)
    # [§3-3-1 ⑤] 사람이 만든 요소 — 그것을 만든 op 의 id 가 주소다.
    _born = (got or {}).get("ops_added") or {}
    for op_id, (_kind, nid) in (_born.get("by_id") or {}).items():
        if str(nid) in nd:
            idx["node"][key_added(op_id)] = str(nid)
    for op_id, pid in (_born.get("pipe_by_id") or {}).items():
        if str(pid) in pr:
            idx["pipe"][key_added(op_id)] = str(pid)
    for nid, vid in nref.items():
        if str(nid) in nd:
            idx["node"][key_node(vid)] = str(nid)

    # 헤드(노즐) — kfp 헤드 절점을 disk 번호로 되짚는다.
    #
    #   ★**좌표로 찾지 않는다.** 회랑 좌표는 사슬로 «만든» 것이라 도면 mm 와
    #   어긋난다 — 실측(대명동 K=30)으로 헤드가 가장 가까운 원에서 최소
    #   199mm · 중앙 913mm · 최대 1998mm 떨어져 있었다. 60mm 자를 대면
    #   30개가 전부 떨어져 나가 헤드 키가 **하나도** 안 생긴다.
    #
    #   대신 이음을 따라간다: kfp 절점 → `node_ref` → board 절점 →
    #   `board.hnodes[i]`(원 i 에 붙은 board 절점들) → 원 번호. 같은 실측에서
    #   30/30 이 이 길로 닿았다. 좌표가 아니라 «누가 누구에 붙어 있나» 라서
    #   사슬이 점을 옮겨도 흔들리지 않는다.
    hn = getattr(board, "hnodes", None)
    if not hn:
        try:
            hn = board._head_nodes()
        except Exception:
            hn = ()
    disk_of = {}
    for i, nodes in enumerate(hn or ()):
        for n in (nodes or ()):
            disk_of.setdefault(int(n), i)

    adj = defaultdict(list)
    for pid, p in pr.items():
        s, e = _ends(p)
        if s is None or e is None:
            continue
        adj[str(s)].append(str(e))
        adj[str(e)].append(str(s))

    head_of = {}
    for nid, m in nd.items():
        if str((m or {}).get("type_id") or "") != "head":
            continue
        vid = nref.get(str(nid))
        if vid is None:                 # 세로 스텁 끝이면 뿌리를 한 칸 되짚는다
            for nb in adj.get(str(nid), ()):
                if nb in nref:
                    vid = nref[nb]
                    break
        bi = disk_of.get(int(vid)) if vid is not None else None
        if bi is None and vid is not None:
            # `hnodes` 에 없는 board 절점 — §2-2 가 «헤드를 관통하던 관» 을
            #   헤드 중심에서 쪼개 만든 자리가 그렇다(그때 이미 셈해 둔
            #   목록에는 없다). board 좌표는 **도면 그대로**라 원과 바로
            #   맞댈 수 있다 — 사슬이 옮긴 kfp 좌표와 달리 흔들리지 않는다.
            #   실측(대명동 K=30): 이 한 줄이 없어 헤드 키가 29/30 이었다.
            bi = _disk_near(board, vid, disks)
        if bi is None:                  # 마지막 수단 — 그때만 kfp 좌표에 자를 댄다
            q = mm(nid)
            bd = 1e18
            for i, d in enumerate(disks):
                dd = math.hypot(q[0] - float(d[0]), q[1] - float(d[1]))
                if dd < bd:
                    bi, bd = i, dd
            if bd > 60.0:
                bi = None
        if bi is not None:
            head_of[str(nid)] = bi
            idx["head"][key_head(bi)] = str(nid)

    # 세로 토막 — 뿌리(헤드 disk 또는 board 절점)와 z 순서 역할
    #
    #   ★뿌리는 **이어 달려 찾는다.** 토막이 하나뿐이면 그 한쪽 끝이 곧
    #     뿌리지만, 상하향식은 한 헤드에 ①②③④ 가 층층이 걸린다 — 그때
    #     두 번째 토막부터는 양 끝이 둘 다 «엔진이 만든 절점» 이라 뿌리를
    #     못 찾고 **키가 없는 채로 떨어진다**(실측: 2단 스택에서 아래 토막이
    #     통째로 빠졌다). 세로 간선만 따라 걸어 올라가 뿌리에 닿는다.
    vpipes = {}                        # pid → (s, e, lo, hi)
    vadj = defaultdict(list)
    for pid, p in pr.items():
        s, e = _ends(p)
        if s is None or e is None:
            continue
        a, b = mm(s), mm(e)
        if math.hypot(a[0] - b[0], a[1] - b[1]) > 1.0:
            continue                    # 세로 토막이 아니다
        za, zb = xyz(s)[2], xyz(e)[2]
        vpipes[str(pid)] = (str(s), str(e), min(za, zb), max(za, zb))
        vadj[str(s)].append((str(e), str(pid)))
        vadj[str(e)].append((str(s), str(pid)))

    roots_at = {}                       # 절점 → 뿌리 키
    for nid in vadj:
        if nid in head_of:
            roots_at[nid] = ("head", head_of[nid])
        elif nid in nref:
            roots_at[nid] = ("tee", nref[nid])

    by_root = defaultdict(list)
    taken = set()
    # 뿌리 순서를 고정한다 — 두 뿌리가 같은 토막에 닿을 수 있고(티와 티 사이의
    # 수직 주행), 그때 먼저 잡는 쪽이 판마다 달라지면 키가 흔들린다.
    for nid in sorted(roots_at, key=lambda n: (str(roots_at[n]), n)):
        root = roots_at[nid]
        stack, seen_n = [nid], {nid}
        while stack:
            cur = stack.pop()
            for nxt, pid in vadj.get(cur, ()):
                if pid in taken:
                    continue
                taken.add(pid)
                by_root[root].append((*vpipes[pid][2:], pid))
                if nxt not in seen_n and nxt not in roots_at:
                    seen_n.add(nxt)
                    stack.append(nxt)
    for root, lst in by_root.items():
        for role, (_lo, _hi, pid) in enumerate(sorted(lst), start=1):
            k = key_vert(root[0], root[1], role)
            idx["vert"][k] = pid
            idx["roots"][pid] = k
    return idx


def label_keys(got, board, tbl):
    """이번 표의 **라벨 → 안정 키**. 카드가 「이 줄」을 고칠 때 쓰는 주소다.

    `build_index` 는 kfp 이름(N12·P7)까지만 번역한다. 화면이 누르는 것은 표
    라벨(노드 3·배관 P587)이라 한 칸을 더 건너야 한다. 그 건너뛰기를 **한 곳**
    에 둔다 — 수리계산 화면과 통합 화면이 각자 셈하면 두 화면이 다른 자리를
    가리키는 날이 온다(두 화면 선정일치에서 이미 겪었다).

      · 배관: **표가 다시 매긴 이름**으로 잇는다(§3-6 · D7). 종전에는 「표
        라벨 = kfp 배관 이름」이라 그대로 이었는데, 이름을 물 흐르는 순서로
        다시 매기면서 그 말이 거짓이 됐다. ★더 나쁜 것은 두 이름이 **같은
        글자꼴**(P1·P2…)이라는 점이다 — 그대로 두면 「P11」이 조용히 다른
        배관을 가리키고, 사람이 적어 둔 값이 옆 배관으로 옮겨간다(§5 금지).
      · 절점: 표는 BFS 로 번호를 새로 매기므로 **좌표**로 잇는다.
        ★표의 노드 좌표는 mm 정수(§T3 — 노드 좌표만 mm)이고 kfp 는 m 이다.
          m 끼리 맞추려 들면 한 칸도 안 맞는다(실측 0/241). x·y 만 보면
          세로로 쌓인 절점(헤드 스텁·가지 상승)이 같은 자리라 서로 덮어쓴다
          — 표고까지 넣어야 1:1 이다(실측 241/241).
        ★헤드를 **나중에** 얹는다. 라벨이 겹칠 때 노즐이 밀려나면 K·방출압을
          고칠 칸이 사라진다(실측: 헤드 키 30개 중 1개가 그렇게 없어졌다).

    반환 `{"pipe": {label: key}, "node": {label: key}, "nid": {label: kfp 이름}}`
    — 키는 JSON 으로 그대로 실을 수 있게 list 다.
    """
    idx = (build_index(got, board) if board is not None
           else {"pipe": {}, "node": {}, "head": {}, "vert": {}})
    of_pipe, of_node, of_head = {}, {}, {}
    # pid → 이번 표의 이름. 표가 안 주면(옛 호출자) pid 를 그대로 쓴다.
    lab_of_pid = dict(getattr(tbl, "pipe_labels", None) or {})

    def _plab(pid):
        return str(lab_of_pid.get(str(pid), pid))

    for k, pid in (idx.get("pipe") or {}).items():
        of_pipe[_plab(pid)] = list(k)
    for k, pid in (idx.get("vert") or {}).items():
        of_pipe.setdefault(_plab(pid), list(k))
    for k, nid in (idx.get("node") or {}).items():
        of_node[str(nid)] = list(k)
    for k, nid in (idx.get("head") or {}).items():
        of_head[str(nid)] = list(k)

    kmeta = ((got or {}).get("kfp") or {}).get("nodes_meta_runtime") or {}
    lab_by_xyz = {}
    for n in (getattr(tbl, "nodes", None) or ()):
        lab_by_xyz[(int(n.get("x") or 0), int(n.get("y") or 0),
                    round(float(n.get("elevation") or 0), 3))] = \
            str(n.get("label"))

    def _lab(nid):
        c = (kmeta.get(str(nid)) or {}).get("coords")
        if not c:
            return None
        return lab_by_xyz.get((
            int(round(float(c[0]) * 1000.0)),
            int(round(float(c[1]) * 1000.0)),
            round(float(c[2]) if len(c) > 2 else 0.0, 3)))

    key_of_lab, nid_of_lab = {}, {}
    for src in (of_node, of_head):
        for nid, kk in src.items():
            lab = _lab(nid)
            if lab:
                key_of_lab[lab] = kk
    for nid in kmeta:
        lab = _lab(nid)
        if lab:
            nid_of_lab.setdefault(lab, str(nid))
    return {"pipe": of_pipe, "node": key_of_lab, "nid": nid_of_lab}


def resolve(rows, idx):
    """저장된 항목 → (적용할 것, 못 옮긴 것). 규칙 4 — 조용히 안 버린다.

    ★돌려주는 것은 **원래 dict 그대로**다(사본이 아니다). 적용하는 쪽이
      `r["old"]` 에 원값을 채워 넣으면 그것이 저장소에 그대로 남아야 하기
      때문이다 — 사본을 주면 채운 원값이 조용히 버려지고, 카드는 「원값 없음」
      만 보인다(규칙 5 가 깨진다). 한 번 그렇게 만들었다가 고쳤다.
      이름은 곁들여 준다 — 항목에 `_name` 을 박으면 파일에도 새어 나간다.
    """
    applied, missed = [], []
    for r in rows:
        k = key_from_json(r.get("key"))
        kind = str(r.get("kind") or "")
        if k is None:
            missed.append({**r, "why": "키를 읽지 못했습니다"})
            continue
        if kind in ("sys", "mr"):
            applied.append((k, r, None))   # 통합 단계에서 라벨로 찾는다
            continue
        name = (idx.get(kind) or {}).get(k)
        if name is None:
            missed.append({**r, "why": "그 자리가 이번 계산 범위에 없습니다"})
            continue
        applied.append((k, r, name))
    return applied, missed


def length_overrides(rows, idx) -> dict:
    """★사슬에 넣을 «길이» 만 — 좌표가 이 값으로 놓인다(아이소가 따라 변한다).

    반환 `{pid: 길이(m)}`. 평면 배관(pipe.length)과 세로 토막(vert.length)
    둘 다 담는다.
    """
    out = {}
    for r in rows:
        if str(r.get("field")) != "length":
            continue
        kind = str(r.get("kind") or "")
        if kind not in ("pipe", "vert"):
            continue
        k = key_from_json(r.get("key"))
        pid = (idx.get(kind) or {}).get(k)
        if pid is None:
            continue
        try:
            v = float(r.get("new"))
        except (TypeError, ValueError):
            continue
        if v > 0:
            out[str(pid)] = v
    return out


# ─────────────────────────────────────────────── 적용 ① kfp 메타 · 길이 재배치
_META_PIPE = {"c": "C", "roughness_mm": "roughness_mm",
              "equivalent_length": "equivalent_length", "type": "type"}
_META_HEAD = {"k_factor_si": ("k_factor_si", "k_factor"),
              "required_pressure_bar": ("required_pressure_bar",)}


def _keep_old(r, cur, new):
    """원값을 적되 **덮인 값으로 덮어쓰지 않는다** (규칙 5).

    빌드마다 엔진이 새 망을 내므로 `cur` 은 보통 규칙이 낸 원값이고, 그것을
    다시 적는 편이 낫다(도면을 고치면 원값도 따라 고쳐진다). 그러나 같은 망에
    두 번 적용하면 `cur` 은 **이미 사람 값**이다 — 그때 적으면 카드가
    「원값 100 → 100」 을 보인다. 그 한 경우만 건너뛴다.
    """
    if cur == new:
        # 지금 값이 이미 사람 값이다 — 이것을 원값이라 적으면 「원값 100 → 100」
        # 이 된다. 처음 본 값이 None 이어서 아직 원값이 안 찬 경우에도 마찬가지다
        # (그 구멍으로 실제로 새 값이 원값 자리에 들어갔다).
        return
    r["old"] = cur


def apply_to_kfp(got, board, rows):
    """표를 만들기 **전에** kfp 메타를 덮고, 길이가 바뀌면 좌표를 다시 놓는다.

    ★여기가 «아이소가 따라 변하는» 자리다. 길이를 덮으면 `relay_from_lengths`
      가 사슬을 다시 걸어 좌표를 옮긴다 — 표·`.sdf`·그림이 한꺼번에 맞는다.

    반환: (applied_n, missed, report)
    """
    from services.cad_import.design.chain import relay_from_lengths

    idx = build_index(got, board)
    applied, missed = resolve(rows, idx)
    kfp = (got or {}).get("kfp") or {}
    nd = kfp.get("nodes_meta_runtime") or {}
    pr = kfp.get("pipe_data") or {}
    n_meta = 0
    for key, r, name in applied:
        kind, field, new = str(r.get("kind")), str(r.get("field")), r.get("new")
        if name is None:
            continue
        # ★길이의 «원값» 은 좌표를 다시 놓기 **전에** 적어 둔다 — 재배치가
        #   `length_m` 을 덮어쓰므로 뒤에서는 원값을 영영 못 읽는다.
        if field == "length" and kind in ("pipe", "vert"):
            row = pr.get(str(name))
            if row is not None and r.get("old") is None:
                r["old"] = round(float(row.get("length_m") or 0), 6)
        if kind in ("pipe", "vert") and field in _META_PIPE:
            row = pr.get(str(name))
            if row is not None:
                _keep_old(r, row.get(_META_PIPE[field]), new)
                row[_META_PIPE[field]] = new
                n_meta += 1
        elif kind == "head" and field in _META_HEAD:
            m = nd.get(str(name))
            if m is not None:
                keys = _META_HEAD[field]
                _keep_old(r, m.get(keys[0]), new)
                for kk in keys:
                    m[kk] = new
                n_meta += 1
    # ★길이 — 좌표를 다시 놓는다 (아이소 반영)
    lov = length_overrides(rows, idx)
    relay = relay_from_lengths(kfp, length_overrides=lov)
    if relay.get("applied"):
        print(f"[수정] 길이 덮기 {relay['applied']}건 — 좌표를 다시 놓았습니다"
              f" (절점 {relay['moved']}개 이동 · 아이소가 따라 변합니다)")
    if n_meta:
        print(f"[수정] kfp 메타 {n_meta}건을 덮었습니다"
              f" (거칠기·등가길이·K·필요압력·C·관종)")
    return n_meta + int(relay.get("applied") or 0), missed, {
        "index": idx, "relay": relay, "applied": applied,
        # ★이번 판에서 **원값을 이미 적어 둔** (키, 속성). 적용 ③ 이 같은 칸의
        #   원값을 다시 쓰면, 그때 표는 이미 덮인 값을 들고 있을 수 있어
        #   「원값 100 → 100」 같은 거짓이 남는다(규칙 5). 먼저 적은 쪽이 이긴다.
        "old_filled": {(json.dumps(r.get("key")), str(r.get("field")))
                       for _k, r, _n in applied
                       if r.get("old") is not None},
    }


# ─────────────────────────────────────────────── 적용 ② 표 칸
_TBL_PIPE = {"dia": "dia", "length": "length", "c": "c", "type": "type"}


def apply_to_tables(tbl, got, board, rows, report=None):
    """표가 선 **뒤에** 표에만 있는 칸을 덮는다 (관경·표고 등).

    ★길이·C·관종은 kfp 메타와 표 양쪽에 있다 — 메타를 덮으면 표가 그 값을
      받아 오지만, 표가 제 규칙으로 다시 채우는 칸(관경·표고)은 여기서 덮는다.
      두 자리가 **서로 다른 것을 덮는다** — 같은 값을 두 번 덮지 않는다.
    """
    idx = (report or {}).get("index") or build_index(got, board)
    applied, missed = resolve(rows, idx)
    done = (report or {}).get("old_filled") or set()
    kfp = (got or {}).get("kfp") or {}
    nd = kfp.get("nodes_meta_runtime") or {}
    # 관경의 «원값» 은 표 행에서 못 읽는다 — §6 으로 두 문을 한 벌로 모은 뒤로
    # `decide_bores` 가 **이미** 사람 값으로 덮어서 들어오기 때문이다. 규칙이
    # 낸 값은 그 함수가 따로 적어 둔다(`orig_dia`) — 그것이 원값이다.
    bore_ov = dict(getattr(tbl, "bore_overrides", None) or {})
    # ★[§3-6] **pid → 표 행.** 표 이름은 다시 매겨지므로(P1·P2…) kfp 이름
    #   (P36)으로는 못 찾는다. 더 나쁜 것은 두 이름이 같은 글자꼴이라는
    #   점이다 — 그대로 두면 사람이 덮은 관경이 **조용히 옆 배관**에 실린다.
    #   `bore_overrides` 는 여전히 pid 가 주소라 그쪽은 안 건드린다.
    _lab_of_pid = {str(k): str(v) for k, v
                   in (getattr(tbl, "pipe_labels", None) or {}).items()}
    row_of_pid = {}
    for _row in (getattr(tbl, "pipes", None) or ()):
        row_of_pid[str(_row.get("label"))] = _row      # 이름이 곧 주소인 옛 표
    for _pid, _lab in _lab_of_pid.items():
        for _row in (getattr(tbl, "pipes", None) or ()):
            if str(_row.get("label")) == _lab:
                row_of_pid[str(_pid)] = _row
                break
    lab_of = {}
    for n in (getattr(tbl, "nodes", None) or ()):
        lab_of[str(n.get("label"))] = n
    # kfp 절점 이름 → 표 행: 표는 BFS 순서로 새 라벨을 매기므로 좌표로 잇는다.
    #
    #   ★**단위를 맞춰야 한다.** 표의 노드 좌표는 mm 정수(§T3 — 노드 좌표만
    #     mm)이고 kfp 는 m 이다. 한때 m 끼리 맞추려 들어 한 칸도 안 맞았고
    #     (실측 0/241), 그런데도 `resolve` 는 성공이라 **못 옮긴 목록에도 안
    #     떴다** — 표고 수정이 조용히 사라졌다(규칙 4 위반).
    #   ★x·y 만 보면 세로로 쌓인 절점(헤드 스텁·가지 상승)이 같은 자리라 서로
    #     덮어쓴다. 표고까지 넣어야 1:1 이다(실측 241/241).
    by_xyz = {}
    for n in (getattr(tbl, "nodes", None) or ()):
        by_xyz[(int(n.get("x") or 0), int(n.get("y") or 0),
                round(float(n.get("elevation") or 0), 3))] = n
    n_tbl = 0
    for key, r, name in applied:
        kind, field, new = str(r.get("kind")), str(r.get("field")), r.get("new")
        if name is None:
            continue
        if kind in ("pipe", "vert") and field in _TBL_PIPE:
            if field == "length":
                # ★길이는 여기서 안 덮는다. 적용 ① 이 kfp 길이를 바꾸고
                #   표가 그 값을 그대로 받아 온다 — 여기서 또 덮으면 같은
                #   값을 두 번 덮게 되고, `old` 가 «이미 덮인 값» 으로
                #   채워져 원값을 잃는다.
                continue
            row = row_of_pid.get(str(name))
            if row is None:
                # 조용히 버리지 않는다 — 배관이 지워졌을 수도 있다(규칙 4).
                missed.append({**r, "why": "표에서 그 배관 행을 못 찾았습니다"})
                continue
            if field == "dia":
                # ★실측(대명동): 도면 글씨로 50A 이던 배관을 100A 로 덮었더니
                #   카드가 「원값 100 → 100」 을 보였다. `decide_bores` 가
                #   이미 100 을 넣은 표를 여기서 원값이라 읽은 탓이다.
                #   규칙이 낸 값은 `orig_dia` 에만 남아 있다.
                o = bore_ov.get(str(name)) or {}
                if "orig_dia" in o:
                    r["old"] = o["orig_dia"]     # 규칙이 낸 값 — 권위 있다
                else:
                    _keep_old(r, row.get("dia"), new)
                row["dia"] = new
                # 관경은 «무엇이 정했나» 가 행에 남는다 — 집계만으로는
                # 도면 텍스트에서 온 것인지 사람이 넣은 것인지 모른다.
                row["dia_src"] = "사람"
                n_tbl += 1
                continue
            if (json.dumps(r.get("key")), field) not in done:
                r["old"] = row.get(field)
            row[field] = new
            n_tbl += 1
        elif kind in ("node", "head") and field == "elevation":
            m = nd.get(str(name))
            if m is None:
                continue
            c = m.get("coords") or (0, 0, 0)
            row = by_xyz.get((
                int(round(float(c[0]) * 1000.0)),
                int(round(float(c[1]) * 1000.0)),
                round(float(c[2]) if len(c) > 2 else 0.0, 3)))
            if row is None:
                # 조용히 버리지 않는다 — 못 찾았으면 못 찾았다고 말한다(규칙 4).
                missed.append({**r, "why": "표에서 그 절점 행을 못 찾았습니다"})
                continue
            if (json.dumps(r.get("key")), field) not in done:
                r["old"] = row.get("elevation")
            row["elevation"] = new
            n_tbl += 1
    if n_tbl:
        print(f"[수정] 표 칸 {n_tbl}건을 덮었습니다 (관경·표고)")
    return n_tbl, missed


_MERGE_PART = {"sys": "system", "mr": "machineroom"}
_MERGE_PIPE = ("length", "dia", "c")


def merge_parts(got) -> dict:
    """통합망의 라벨이 어느 도면 것인가 — `{label: "plan"|"system"|"machineroom"}`.

    절점은 `parts` 가 그대로 알려 준다. 배관은 목록에 없으므로 **양 끝으로**
    정한다 — 두 끝이 같은 도면이면 그 도면 것이고, 갈리면 이음매(seam)다.
    이음매는 어느 쪽도 아니므로 고치지 않는다(어느 도면을 고친 것인지 말할
    수 없는 값은 두지 않는다).
    """
    parts = (got or {}).get("parts") or {}
    of = {}
    for kind in ("system", "machineroom", "plan"):
        for lab in (parts.get(kind) or ()):
            of[str(lab)] = kind
    pipe_of = {}
    c = (got or {}).get("combined")
    for r in (getattr(c, "pipes", None) or ()):
        ka = of.get(str(r.get("in")))
        kb = of.get(str(r.get("out")))
        pipe_of[str(r.get("label"))] = ka if (ka and ka == kb) else "seam"
    return {"node": of, "pipe": pipe_of}


def apply_to_merge(got, rows):
    """통합망의 **계통도·기계실** 요소를 덮는다 (§4 · 오너 승인 2026-09-14).

    회랑(평면도) 요소는 여기서 손대지 않는다 — 그것은 이미 설계 표에서 덮여
    통합으로 흘러든다. 여기서 또 덮으면 같은 값을 두 번 덮게 되고, 한쪽만
    고치는 날 두 화면이 갈린다.

    ★길이를 바꿔도 라이저 **좌표는 안 움직인다.** 그것이 이 자리의 규약이다:
      입상관 막대는 표 길이에 비례하지 않고 **균등 간격**으로 눕는다(오너
      2026-09-08 「이전 디자인이 더 좋아」 — 0.017 m 짜리 구간이 사라져
      막대가 한쪽으로 뭉쳤다). 길이의 권위는 표이지 그림이 아니다. 회랑은
      반대다 — 거기는 사슬이 길이로 좌표를 만들어 아이소가 따라 변한다.
      카드가 이 차이를 말해야 사람이 «안 먹혔다» 고 읽지 않는다.

    반환: (덮은 수, 못 옮긴 것)
    """
    c = (got or {}).get("combined")
    if c is None:
        return 0, [dict(r) for r in rows
                   if str(r.get("kind")) in _MERGE_PART]
    where = merge_parts(got)
    pipe_by = {str(r.get("label")): r for r in (getattr(c, "pipes", None) or ())}
    node_by = {str(n.get("label")): n for n in (getattr(c, "nodes", None) or ())}

    n_ok, missed = 0, []
    for r in rows:
        kind = str(r.get("kind"))
        key = key_from_json(r.get("key"))
        field = str(r.get("field"))
        want = _MERGE_PART.get(kind)
        # ★«이 화면 것인가» 는 갈래(kind)가 정한다 — 다만 **사람이 만든 요소**는
        #   갈래가 아니라 **어느 망에서 만들어졌나**가 정한다. 통합에서 만든
        #   것이면 통합에서 고치고, 회랑에서 만든 것이면 설계 표에서 이미
        #   고쳐져 흘러든다. 그래서 add 키는 갈래로 거르지 않고 아래에서 본다.
        _is_add = bool(key) and str(key[0]) == "add"
        if want is None and not _is_add:
            continue                      # 회랑 요소 — 설계 표에서 이미 덮였다
        if key is None or len(key) < 2:
            missed.append({**r, "why": "키를 읽지 못했습니다"})
            continue
        lab = str(key[1])
        if _is_add:
            # [§3-3-1 ⑤] 사람이 **만든** 요소 — 그것을 만든 op 의 id 가 주소다.
            #   라벨은 결합할 때마다 새로 매겨지므로 id 로 되짚는다.
            _born = (got or {}).get("ops_added") or {}
            src = (_born.get("merge_pipe_by_id") if field in _MERGE_PIPE
                   else _born.get("merge_by_id")) or {}
            lab = str(src.get(lab) or "")
            if not lab:
                # 통합이 만든 것이 아니다 — 회랑에서 만들어 표로 흘러든
                # 요소이므로 여기서 손대지 않는다(그쪽이 이미 덮었다).
                continue
            want = where["pipe" if field in _MERGE_PIPE else "node"].get(lab)
        if field in _MERGE_PIPE:
            row, got_part = pipe_by.get(lab), where["pipe"].get(lab)
        else:
            row, got_part = node_by.get(lab), where["node"].get(lab)
        if row is None:
            missed.append({**r, "why": "그 라벨이 이번 결합망에 없습니다 "
                                       "— 계통도·기계실을 다시 뽑으면 "
                                       "라벨이 바뀝니다"})
            continue
        if got_part != want:
            # 라벨이 살아 있어도 «다른 도면의 것» 이면 덮지 않는다.
            missed.append({**r, "why": f"그 라벨은 이제 {got_part or '알 수 없음'} "
                                       f"쪽입니다 (수정은 {want} 것)"})
            continue
        col = {"length": "length", "dia": "dia", "c": "c",
               "elevation": "elevation"}.get(field)
        if col is None:
            missed.append({**r, "why": f"통합에서 고칠 수 없는 속성입니다: {field}"})
            continue
        _keep_old(r, row.get(col), r.get("new"))
        row[col] = r.get("new")
        n_ok += 1
    if n_ok or missed:
        print(f"[수정] 통합 요소 {n_ok}건을 덮었습니다"
              + (f" · 못 옮김 {len(missed)}건" if missed else ""))
    return n_ok, missed


def ensure_loaded(sess) -> list:
    """도면의 저장된 수정을 **한 번** 올린다 (D2 · 기준 5 파일 왕복).

    ★부르는 자리를 셋(열기·찍기·자동)으로 두지 않는다 — 하나라도 빠뜨리면
      그 길로 연 도면만 수정이 안 올라오고, 그것은 «가끔 사라지는 값» 이 된다.
      표를 세울 때와 카드를 열 때 이 한 줄이 지킨다. 키가 바뀌면(다른 도면을
      열면) 다시 읽는다.
    """
    key = str(sess.get("key") or "")
    if not key:
        return load(sess)
    if sess.get("_ov_loaded_for") == key:
        return load(sess)
    rows = read_file(key)
    sess["_ov_loaded_for"] = key
    if rows:
        # 세션에 이미 있는 것이 우선 — 사람이 이번에 고친 것을 파일이 덮으면 안 된다.
        cur = load(sess)
        have = {(str(r.get("key")), str(r.get("field"))) for r in cur}
        add = [r for r in rows
               if (str(r.get("key")), str(r.get("field"))) not in have]
        if add:
            save(sess, cur + add)
            print(f"[수정] 저장된 요소 수정 {len(add)}건을 올렸습니다"
                  f" — {os.path.basename(path_for(key))}")
    return load(sess)


# ══════════════════════════════════════════════════════════════════════════
# 위상 수정 (§3-3) — 값이 아니라 **망의 모양**을 고친다
#
# 값 수정(`items`)과 나란히 두되 섞지 않는다. 값은 「이 배관의 길이를 2 m 로」
# 이고 위상은 「이 배관을 지운다」다 — 같은 저장소에 담되 적용하는 함수가
# 다르고, 적용 순서도 정해져 있다(§3-3-1: 값 → 삭제 → 노드추가 → 기기추가 →
# 새 요소의 값).
#
# ★망을 직접 고쳐 세션에 눌러 담지 않는다(§5 금지). 목록에 쌓고, 표를 만들
#   때마다 **처음 망에** 다시 적용한다. 그래서 목록을 비우면 원래 망으로
#   정확히 돌아온다(M3).
# ══════════════════════════════════════════════════════════════════════════
OPS_VERSION = 2
OP_KINDS = ("delete", "add_node", "add_equip")

# 적용 순서 — 이 순서가 «멱등» 의 근거다. 같은 목록을 몇 번 적용해도 같은 망이
# 나오려면 순서가 목록의 배열이 아니라 **규칙**으로 정해져 있어야 한다.
_OP_ORDER = {"delete": 0, "add_node": 1, "add_equip": 2}


def key_added(op_id):
    """새로 생긴 요소의 안정 키 — 그 요소를 만든 op 의 id 가 곧 주소다.

    board 에 대응할 자리가 없어(사람이 만든 요소다) 좌표로도 라벨로도 못
    가리킨다. 만든 자의 id 로 가리키는 것이 유일하게 안정된 길이다.
    """
    return ("add", str(op_id))


def new_op_id(ops) -> str:
    """서버가 매기는 8자리 — 목록 안에서 안 겹칠 때까지."""
    have = {str(r.get("id")) for r in (ops or ())}
    while True:
        cand = uuid.uuid4().hex[:8]
        if cand not in have:
            return cand


def load_ops(sess) -> list:
    return list(sess.get("element_ops") or ())


def save_ops(sess, ops) -> None:
    sess["element_ops"] = list(ops)


def put_op(ops, op_id, op, target, payload, *, reason="", at=""):
    """위상 수정 하나를 넣거나 갈아 끼운다 — 같은 id 는 하나뿐이다."""
    out = [r for r in (ops or ()) if str(r.get("id")) != str(op_id)]
    out.append({"id": str(op_id), "op": str(op),
                "target": key_to_json(target), "payload": dict(payload or {}),
                "reason": str(reason or ""), "at": str(at or "")})
    return out


def drop_op(ops, op_id):
    return [r for r in (ops or ()) if str(r.get("id")) != str(op_id)]


def ops_sorted(ops) -> list:
    """적용 순서대로 — 같은 종류 안에서는 목록에 들어온 차례.

    ★`sorted` 는 안정 정렬이라 같은 종류의 상대 순서가 보존된다. 그것이
      「여러 번 눌러도 결과가 같다」(§4 M3)의 조건이다.
    """
    known = [r for r in (ops or ()) if str(r.get("op")) in _OP_ORDER]
    return sorted(known, key=lambda r: _OP_ORDER[str(r.get("op"))])


def ensure_ops_loaded(sess) -> list:
    """저장된 위상 수정을 **한 번** 올린다 — `ensure_loaded` 와 같은 규약."""
    key = str(sess.get("key") or "")
    if not key:
        return load_ops(sess)
    if sess.get("_ops_loaded_for") == key:
        return load_ops(sess)
    rows = read_file_ops(key)
    sess["_ops_loaded_for"] = key
    if rows:
        cur = load_ops(sess)
        have = {str(r.get("id")) for r in cur}
        add = [r for r in rows if str(r.get("id")) not in have]
        if add:
            save_ops(sess, cur + add)
            print(f"[수정] 저장된 위상 수정 {len(add)}건을 올렸습니다"
                  f" — {os.path.basename(path_for(key))}")
    return load_ops(sess)


# ─────────────────────────────────────────────── 위상 수정 — 망을 보는 한 눈
class _TopoNet:
    """위상 수정이 보는 «망». 두 자리가 **같은 규칙**을 쓰게 하는 얇은 껍데기.

    회랑은 kfp dict(`nodes_meta_runtime`·`pipe_data`)이고 통합은 표 행 목록
    (`combined.nodes`·`combined.pipes`)이다. 모양이 전혀 다른데 규칙은 하나여야
    한다(§3-3-2·§3-3-3) — 규칙을 두 벌로 적으면 한쪽만 고쳐지는 날이 온다.
    그래서 «무엇을 할 수 있는가» 만 여기에 적고, «어떻게» 는 자식이 채운다.
    """

    def node_ids(self) -> list:
        raise NotImplementedError

    def pipe_ids(self) -> list:
        raise NotImplementedError

    def ends(self, pid):
        raise NotImplementedError

    def length(self, pid) -> float:
        raise NotImplementedError

    def xyz(self, nid):
        raise NotImplementedError

    def root(self):
        raise NotImplementedError

    def is_head(self, nid) -> bool:
        return False

    def protected(self, nid):
        """지울 수 없는 절점이면 그 **이유**를, 지울 수 있으면 None."""
        if str(nid) == str(self.root()):
            return "급수원(Input)은 망의 정의입니다"
        return None

    def seam(self, pid) -> bool:
        return False

    def bore(self, pid):
        return None

    def split(self, pid, new_pid, new_nid, xyz, l1, l2, up, dn, op_id):
        raise NotImplementedError

    def absorb(self, keep, drop, far_a, far_b, length_m):
        raise NotImplementedError

    def drop_pipe(self, pid):
        raise NotImplementedError

    def drop_node(self, nid):
        raise NotImplementedError

    def next_node_name(self, op_id):
        raise NotImplementedError

    def next_pipe_name(self, op_id):
        raise NotImplementedError


def _adj_of(net) -> dict:
    adj: dict = {}
    for pid in net.pipe_ids():
        a, b = net.ends(pid)
        if a is None or b is None:
            continue
        adj.setdefault(str(a), []).append((str(b), pid))
        adj.setdefault(str(b), []).append((str(a), pid))
    return adj


def _reach(net, adj, *, skip_pipe=None, skip_node=None) -> set:
    """급수원에서 걸어 닿는 절점 — 배관/절점 하나를 빼고 걸어 볼 수 있다."""
    root = str(net.root() or "")
    if not root or root == str(skip_node):
        return set()
    seen, q = {root}, deque([root])
    while q:
        u = q.popleft()
        for v, pid in adj.get(u, ()):
            if pid == skip_pipe or v == str(skip_node) or v in seen:
                continue
            seen.add(v)
            q.append(v)
    return seen


def _depth_of(net, adj) -> dict:
    """급수원에서의 걸음 수 — «상류/하류» 를 가르는 유일한 자(표의 in/out 과 같다)."""
    root = str(net.root() or "")
    if not root:
        return {}
    d = {root: 0}
    q = deque([root])
    while q:
        u = q.popleft()
        for v, _pid in sorted(adj.get(u, ()), key=lambda t: str(t[1])):
            if v in d:
                continue
            d[v] = d[u] + 1
            q.append(v)
    return d


def _lerp_xyz(a, b, t):
    return tuple(float(a[i]) + (float(b[i]) - float(a[i])) * float(t)
                 for i in range(3))


def split_lengths(total, t):
    """길이 한 개 → 두 개. **합이 원래와 같아야 한다**(§4 M5).

    ★반올림을 두 번 하면 합이 어긋난다(0.333+0.667 ≠ 1.000 인 경우가 실제로
      난다). 한쪽을 반올림하고 **나머지는 빼서** 만든다. 지시서는 「잔차를 긴
      쪽에 얹는다」고 했으므로, 긴 쪽을 «빼서 만드는 쪽» 으로 둔다.
    """
    L = round(float(total), 3)
    a = round(L * float(t), 3)
    b = round(L - a, 3)
    if a > b:                      # 긴 쪽이 잔차를 받는다
        b = round(L * (1.0 - float(t)), 3)
        a = round(L - b, 3)
    return max(a, 0.0), max(b, 0.0)


def run_ops(net, ops, resolve, *, report=None):
    """목록을 망에 적용한다 — §3-3-2(삭제) · §3-3-3(추가) 규칙 한 벌.

    `resolve(target_key) -> (what, name)` 는 안정 키를 **이번 망의 이름**으로
    옮긴다. `what` 은 "pipe" 또는 "node". 못 옮기면 (None, None).

    반환: (적용 건수, 못 한 것, 보고)
      보고 `{"equip": [...], "loop_pass": n, "heads_removed": n,
             "added": {"nodes": {op_id: name}, "pipes": {op_id: name}}}`
    """
    rep = report if report is not None else {}
    rep.setdefault("equip", [])
    rep.setdefault("loop_pass", 0)
    rep.setdefault("heads_removed", 0)
    rep.setdefault("added", {"nodes": {}, "pipes": {}})
    missed, n_ok = [], 0

    def fail(r, why):
        missed.append({**r, "why": why})

    for r in ops_sorted(ops):
        op = str(r.get("op"))
        op_id = str(r.get("id") or "")
        key = key_from_json(r.get("target"))
        if key is None:
            fail(r, "키를 읽지 못했습니다")
            continue
        what, name = resolve(key)
        if name is None:
            fail(r, "그 자리가 이번 계산 범위에 없습니다")
            continue
        adj = _adj_of(net)
        if op == "delete":
            ok, why = _do_delete(net, adj, what, name, rep)
        elif op == "add_node":
            ok, why = _do_add_node(net, adj, what, name, r, op_id, rep)
        elif op == "add_equip":
            ok, why = _do_add_equip(net, what, name, r, op_id, rep)
        else:
            ok, why = False, f"모르는 수정입니다: {op}"
        if ok:
            n_ok += 1
        else:
            fail(r, why)
    return n_ok, missed, rep


def _do_delete(net, adj, what, name, rep):
    if what == "pipe":
        if str(name) not in {str(p) for p in net.pipe_ids()}:
            return False, "그 배관은 이미 없습니다"
        if net.seam(name):
            return False, ("이음매 배관은 지울 수 없습니다 — "
                           "어느 도면의 것도 아닙니다")
        before = _reach(net, adj)
        after = _reach(net, adj, skip_pipe=name)
        lost = before - after
        if lost:
            # ★끊기는 쪽이 **절점 하나** 면 길이 있다 — 그 절점을 지우면 된다.
            #   「안 됩니다」로 끝내지 않고 할 수 있는 길을 말한다.
            tail = ("" if len(lost) > 1 else
                    " — 그 절점을 지우면 배관도 함께 지워집니다")
            return False, (f"이 배관을 지우면 아래쪽 {len(lost)}개 절점이 "
                           f"급수원에서 끊깁니다{tail}")
        net.drop_pipe(name)
        # 고리가 있어 안 갈라진 삭제 — 드문 일이라 세어 둔다(⑩).
        rep["loop_pass"] = int(rep.get("loop_pass") or 0) + 1
        return True, None

    nid = str(name)
    if nid not in {str(n) for n in net.node_ids()}:
        return False, "그 절점은 이미 없습니다"
    why = net.protected(nid)
    if why:
        return False, why
    links = list(adj.get(nid, ()))
    if len(links) >= 3:
        return False, (f"이 절점에는 배관이 {len(links)}개 붙어 있습니다 — "
                       "무엇을 남길지 프로그램이 정할 수 없습니다")
    if len(links) <= 1:
        # ★거절될 수 있는 검사를 **다 지난 뒤에** 센다 — 거절된 삭제까지 세면
        #   표 메타의 「사용자가 지운 헤드 N」이 거짓이 되고, 그 수로 기준개수를
        #   읽는 사람이 잘못 읽는다.
        # ★그리고 **지우기 전에** 센다 — 지우고 나면 「헤드였나」를 물을 자리가
        #   사라진다(결합 어댑터는 노즐 행을 함께 지운다).
        _count_head(net, nid, rep)
        for _far, pid in links:
            net.drop_pipe(pid)
        net.drop_node(nid)
        return True, None

    # 차수 2 — 두 배관을 하나로 합친다. 살리는 쪽은 **상류**다(관경·관종·C 를
    # 상류 것으로 가져가라는 규칙이 그 쪽을 권위로 둔다).
    (fa, p1), (fb, p2) = links[0], links[1]
    if str(fa) == str(fb):
        return False, "고리의 두 끝이 같은 절점입니다 — 합칠 수 없습니다"
    depth = _depth_of(net, adj)
    if depth.get(str(fb), 1 << 30) < depth.get(str(fa), 1 << 30):
        (fa, p1), (fb, p2) = (fb, p2), (fa, p1)
    # ★합을 **표가 읽는 자리에서** 지킨다. 6자리로 더한 뒤 표가 3자리로
    #   자르면 두 행을 각각 자른 옛 합과 0.001 m 어긋난다(실측: 대명동에서
    #   174.296 → 174.295). 먼저 자르고 더하면 그 틈이 없다(§4 M5).
    _count_head(net, nid, rep)
    L = round(round(float(net.length(p1) or 0.0), 3)
              + round(float(net.length(p2) or 0.0), 3), 3)
    net.absorb(p1, p2, fa, fb, L)
    net.drop_node(nid)
    return True, None


def _count_head(net, nid, rep):
    """지운 것이 헤드였으면 센다 — 기준개수 K 가 그만큼 준다(D6)."""
    if net.is_head(nid):
        rep["heads_removed"] = int(rep.get("heads_removed") or 0) + 1


def _do_add_node(net, adj, what, name, r, op_id, rep):
    if what != "pipe":
        return False, "노드는 **배관 위에만** 더할 수 있습니다"
    if str(name) not in {str(p) for p in net.pipe_ids()}:
        return False, "그 배관이 이번 계산 범위에 없습니다"
    try:
        t = float((r.get("payload") or {}).get("t"))
    except (TypeError, ValueError):
        return False, "위치(t)를 읽지 못했습니다"
    if not (0.0 < t < 1.0):
        return False, f"위치(t)는 0 과 1 사이여야 합니다: {t}"
    a, b = net.ends(name)
    depth = _depth_of(net, adj)
    if depth.get(str(b), 1 << 30) < depth.get(str(a), 1 << 30):
        # t 는 **표의 in 쪽**에서 잰다 — 상류가 in 이다.
        t = 1.0 - t
        a, b = b, a
    L = float(net.length(name) or 0.0)
    l1, l2 = split_lengths(L, t)
    xyz = _lerp_xyz(net.xyz(a), net.xyz(b), t)
    new_nid = net.next_node_name(op_id)
    new_pid = net.next_pipe_name(op_id)
    net.split(name, new_pid, new_nid, xyz, l1, l2, a, b, op_id)
    rep["added"]["nodes"][str(op_id)] = str(new_nid)
    rep["added"]["pipes"][str(op_id)] = str(new_pid)
    return True, None


def _do_add_equip(net, what, name, r, op_id, rep):
    from services.cad_import.design.fitting import resolve_eq_len

    if what != "pipe":
        return False, "기기는 **배관 위에만** 붙습니다 (⑥·⑪)"
    if str(name) not in {str(p) for p in net.pipe_ids()}:
        return False, "그 배관이 이번 계산 범위에 없습니다"
    pay = r.get("payload") or {}
    lib_id = str(pay.get("lib_id") or "")
    if not lib_id:
        return False, "어떤 기기인지(lib_id)가 없습니다"
    try:
        t = float(pay.get("t", 0.5))
    except (TypeError, ValueError):
        return False, "위치(t)를 읽지 못했습니다"
    t = min(1.0, max(0.0, t))
    dia = net.bore(name)
    eq, why = resolve_eq_len(lib_id, dia)
    rep["equip"].append({
        "op_id": str(op_id), "pipe": str(name), "lib_id": lib_id,
        "desc": str(pay.get("desc") or equip_label(lib_id)),
        # ★못 구하면 **0 이 아니라 None** 이다(§4 M8). 0 은 「손실이 없다」는
        #   주장이라, 모르는 것을 0 으로 채우면 계산이 조용히 낙관적이 된다.
        "eq_len": eq, "why": why, "rel_pos": round(t, 4), "dia": dia,
    })
    return True, None


# ─────────────────────────────────────────────── 고를 수 있는 기기 (§3-3-3)
#
# ★목록을 화면에 박지 않는다. **서버가 라이브러리에서 뽑아** 내려보낸다 —
#   라이브러리에 항목이 늘면 화면이 저절로 따라오고, 화면만 아는 이름이
#   생기지 않는다(등가길이를 못 찾는 조용한 0 의 원인이 그것이다).
EQUIP_PICK = ("VALVE_GATE", "VALVE_BUTTERFLY", "VALVE_SWING_CHECK",
              "VALVE_ALARM", "VALVE_GLOBE", "VALVE_ANGLE", "VALVE_BALL",
              "VALVE_FLOW_SWITCH", "VALVE_PREACTION", "VALVE_DRY")
EQUIP_CATEGORY_PICK = ("strainer",)


def equip_catalog() -> list:
    """`[{"id", "name", "category"}]` — 밸브 10종 + 스트레이너(§3-3-3).

    스트레이너는 id 가 uuid 라 이름으로 박을 수가 없다 — **분류(category)**
    로 고른다. 라이브러리가 id 를 다시 매겨도 따라온다.
    """
    from pathlib import Path
    from services.cad_import.design import fitting as _f

    p = Path(getattr(_f, "_G_ROOT")) / "fittings_library_v3.json"
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        print(f"[수정] ★부속 라이브러리를 읽지 못했습니다 — {exc}")
        return []
    out, seen = [], set()
    items = list(data.get("items") or ())
    for want in EQUIP_PICK:
        for r in items:
            if str(r.get("id")) == want and want not in seen:
                seen.add(want)
                out.append({"id": want, "name": r.get("name_ko") or want,
                            "category": r.get("category")})
    for r in items:
        rid = str(r.get("id"))
        if str(r.get("category")) in EQUIP_CATEGORY_PICK and rid not in seen:
            seen.add(rid)
            out.append({"id": rid, "name": r.get("name_ko") or rid,
                        "category": r.get("category")})
    return out


def equip_label(lib_id) -> str:
    for r in equip_catalog():
        if str(r.get("id")) == str(lib_id):
            return str(r.get("name"))
    return str(lib_id)


# ─────────────────────────────────────────────── 회랑(kfp) 어댑터
class _KfpNet(_TopoNet):
    """회랑 망 — `got["kfp"]` 를 그 자리에서 고친다(표를 만들기 **전**)."""

    def __init__(self, got):
        self.got = got or {}
        kfp = self.got.get("kfp") or {}
        self.nodes = kfp.setdefault("nodes_meta_runtime", {})
        self.pipes = kfp.setdefault("pipe_data", {})
        self.eref = self.got.setdefault("edge_ref", {})
        self.nref = self.got.setdefault("node_ref", {})
        self._root = None

    def node_ids(self):
        return list(self.nodes)

    def pipe_ids(self):
        return list(self.pipes)

    def ends(self, pid):
        pr = self.pipes.get(str(pid)) or {}
        return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")

    def length(self, pid):
        return float((self.pipes.get(str(pid)) or {}).get("length_m") or 0.0)

    def xyz(self, nid):
        c = (self.nodes.get(str(nid)) or {}).get("coords") or (0.0, 0.0, 0.0)
        return (float(c[0]), float(c[1]), float(c[2]) if len(c) > 2 else 0.0)

    def root(self):
        if self._root is None:
            from services.cad_import.design.anchor import require_anchor
            try:
                self._root = require_anchor(self.nodes, what="위상 수정")
            except Exception:                           # noqa: BLE001
                self._root = next(iter(self.nodes), None)
        return self._root

    def is_head(self, nid):
        return str((self.nodes.get(str(nid)) or {}).get("type_id")) == "head"

    def bore(self, pid):
        # ★여기서는 **모른다**. 표의 호칭경은 `decide_bores` 가 이 뒤에 정하고,
        #   kfp 의 `nominal_mm` 은 전개가 심어 둔 다른 값이다. 둘을 섞으면
        #   기기 등가길이가 표와 다른 관경으로 계산된다 — 못 구한 것은 못
        #   구한 대로 두고, 표가 선 뒤 `apply_ops_to_tables` 가 다시 푼다.
        return None

    def next_node_name(self, op_id):
        return f"UX{op_id}"

    def next_pipe_name(self, op_id):
        return f"PX{op_id}"

    def split(self, pid, new_pid, new_nid, xyz, l1, l2, up, dn, op_id):
        pr = self.pipes[str(pid)]
        self.nodes[str(new_nid)] = {
            "id": str(new_nid),
            "coords": [float(xyz[0]), float(xyz[1]), float(xyz[2])],
            "elevation_m": float(xyz[2]),
            "type": "기본", "type_id": "base", "category_id": "",
            "k_factor_si": None, "head_spec_name": None,
            "required_pressure_bar": 0.0,
        }
        new = dict(pr)
        new["start"], new["end"] = str(new_nid), str(dn)
        new["length_m"] = float(l2)
        # 부속·등가길이는 **원래 조각에 그대로 둔다**(§3-3-3). 나눠 옮기면
        # 어느 쪽에 있던 것인지 아무도 모르게 된다.
        new["equivalent_length"] = 0.0
        new["fittings"] = []
        self.pipes[str(new_pid)] = new
        pr["start"], pr["end"] = str(up), str(new_nid)
        pr["length_m"] = float(l1)
        # ★같은 board 선분 위의 두 조각이다 — 역참조를 함께 쓰면 관경이
        #   같아진다(도면 치수·별표1 이 같은 답을 낸다). 주소록이 헷갈리지
        #   않도록 새 조각은 `ops_added` 에 적어 둔다(`build_index` 가 뺀다).
        ref = self.eref.get(str(pid))
        if ref is not None:
            self.eref[str(new_pid)] = ref
        born = self.got.setdefault("ops_added", {})
        born.setdefault("pipes", []).append(str(new_pid))
        born.setdefault("nodes", []).append(str(new_nid))
        # ★[§3-3-1 ⑤] 「새 요소의 값 수정」이 서려면 `("add", id)` 도 **주소록에
        #   있어야** 한다. 만든 이름을 여기 적어 두고 `build_index` 가 싣는다 —
        #   안 그러면 방금 만든 절점의 표고를 고칠 때 「그 자리가 이번 계산
        #   범위에 없습니다」가 뜬다(실제로 그랬다).
        born.setdefault("by_id", {})[str(op_id)] = ("node", str(new_nid))
        born.setdefault("pipe_by_id", {})[str(op_id)] = str(new_pid)

    def absorb(self, keep, drop, far_a, far_b, length_m):
        p1 = self.pipes[str(keep)]
        p2 = self.pipes.get(str(drop)) or {}
        p1["start"], p1["end"] = str(far_a), str(far_b)
        p1["length_m"] = float(length_m)
        p1["equivalent_length"] = (float(p1.get("equivalent_length") or 0.0)
                                   + float(p2.get("equivalent_length") or 0.0))
        fit = list(p1.get("fittings") or []) + list(p2.get("fittings") or [])
        if fit:
            p1["fittings"] = fit
        self.pipes.pop(str(drop), None)
        self.eref.pop(str(drop), None)
        a_ref, b_ref = self.nref.get(str(far_a)), self.nref.get(str(far_b))
        if a_ref is not None and b_ref is not None:
            self.eref[str(keep)] = (a_ref, b_ref)
        else:
            self.eref.pop(str(keep), None)

    def drop_pipe(self, pid):
        self.pipes.pop(str(pid), None)
        self.eref.pop(str(pid), None)

    def drop_node(self, nid):
        self.nodes.pop(str(nid), None)
        self.nref.pop(str(nid), None)


def apply_ops_to_kfp(got, board, ops):
    """[§3-3-1 적용 ★] 회랑 대상 위상 수정 — 표를 만들기 **전에** 망을 고친다.

    ★이 자리여야 ⑨ 가 선다. 여기서 고치면 04 의 표·`.sdf` 와, 그 표가 흘러든
      통합 `.sdf` 가 **함께** 바뀐다 — 두 산출이 같은 망을 말한다. 표가 선
      뒤에 고치면 통합은 옛 망을 들고 간다.

    반환: (적용 건수, 못 한 것, 보고)
    """
    ops = [r for r in (ops or ())
           if str((r.get("target") or [None])[0]) not in ("sys", "mr")]
    if not ops:
        return 0, [], {}
    net = _KfpNet(got)
    idx = build_index(got, board)
    added: dict = {}

    def resolve(key):
        kind = str(key[0])
        if kind == "add":
            return added.get(str(key[1]), (None, None))
        if kind in ("pipe", "vert"):
            return "pipe", (idx.get(kind) or {}).get(key)
        if kind in ("node", "head"):
            return "node", (idx.get(kind) or {}).get(key)
        return None, None

    rep: dict = {}
    n, missed, rep = run_ops(net, ops, resolve, report=rep)
    for op_id, nid in (rep.get("added", {}).get("nodes") or {}).items():
        added[str(op_id)] = ("node", nid)
    if n or missed:
        print(f"[수정] 회랑 위상 {n}건을 적용했습니다"
              + (f" · 못 함 {len(missed)}건" if missed else "")
              + (f" · 지운 헤드 {rep.get('heads_removed')}개"
                 if rep.get("heads_removed") else "")
              + (f" · 고리 덕분에 통과한 삭제 {rep.get('loop_pass')}건"
                 if rep.get("loop_pass") else ""))
    return n, missed, rep


def apply_ops_to_tables(tbl, rep, *, missed=None):
    """표가 선 **뒤에** 기기 행을 단다 — 배관 이름과 호칭경이 그때 정해진다.

    ★등가길이는 여기서 «다시» 푼다. 망을 고칠 때는 표의 호칭경을 아직 모르고
      (`decide_bores` 가 그 뒤에 돈다), 0 으로 채우면 손실이 조용히 사라진다.
    """
    from services.cad_import.design.fitting import resolve_eq_len

    rows = list((rep or {}).get("equip") or ())
    if not rows:
        return 0
    lab_of = dict(getattr(tbl, "pipe_labels", None) or {})
    by_label = {str(r.get("label")): r for r in (tbl.pipes or ())}
    n = 0
    for r in rows:
        lab = lab_of.get(str(r.get("pipe")), str(r.get("pipe")))
        prow = by_label.get(str(lab))
        if prow is None:
            if missed is not None:
                missed.append({**r, "why": "그 배관이 표에 없습니다"})
            continue
        dia = r.get("dia")
        eq, why = r.get("eq_len"), r.get("why")
        if dia is None:
            dia = prow.get("dia")
            eq, why = resolve_eq_len(str(r.get("lib_id")), dia)
        tbl.equipment.append({
            "pipe": str(lab), "in": prow.get("in"), "out": prow.get("out"),
            "label": str(len(tbl.equipment) + 1),
            "desc": str(r.get("desc") or r.get("lib_id")),
            # ★못 구했으면 None 을 그대로 싣는다 — 세는 자리가 그것을
            #   「미해결」로 읽는다(§4 M8). 0 으로 메우지 않는다.
            "eq_len": eq, "rel_pos": r.get("rel_pos"),
            "lib": str(r.get("lib_id")),
            # ★누가 단 기기인가. `.kfp` 는 **사람이 더한 것만** 등가길이에
            #   더한다(§3-3-3) — 자동이 단 FX·A/V 는 종전 동작 그대로 둔다.
            "op_id": str(r.get("op_id") or ""),
        })
        n += 1
        if eq is None and missed is not None:
            missed.append({**r, "why": f"등가길이를 못 구했습니다 "
                                       f"(호칭경 {dia}) — 라이브러리에 값이 "
                                       f"없습니다"})
        elif why:
            r["why"] = why
    return n


# ─────────────────────────────────────────────── 통합(결합표) 어댑터
class _MergeNet(_TopoNet):
    """결합망 — 표 행 목록을 그 자리에서 고친다(`combined`).

    회랑과 달리 좌표가 **표 mm** 그대로다(정규화가 없다 · 1-B ④). 그래서
    노드를 더할 때 좌표를 내분하면 그대로 화면 자리가 된다.
    """

    def __init__(self, got):
        self.got = got or {}
        c = self.got.get("combined")
        self.c = c
        # ★`or []` 로 받으면 안 된다. 빈 목록은 falsy 라 **새 리스트**가 생기고,
        #   거기 담은 것은 결합표에 영영 안 닿는다 — 기기 목록은 보통 비어
        #   있으므로(자동이 단 것이 없으면) 사람이 단 밸브가 「적용했습니다」
        #   라고 보고된 채 산출물에서 사라진다. 없으면 **객체에 붙여** 둔다.

        def _lst(name):
            cur = getattr(c, name, None)
            if cur is None:
                cur = []
                setattr(c, name, cur)
            return cur

        self.nodes = _lst("nodes")
        self.pipes = _lst("pipes")
        self.nozzles = _lst("nozzles")
        self.fittings = _lst("fittings")
        self.equipment = _lst("equipment")
        self._n = {str(r.get("label")): r for r in self.nodes}
        self._p = {str(r.get("label")): r for r in self.pipes}
        self._where = merge_parts(self.got)
        self._heads = {str(r.get("in")) for r in self.nozzles}

    def node_ids(self):
        return [str(r.get("label")) for r in self.nodes]

    def pipe_ids(self):
        return [str(r.get("label")) for r in self.pipes]

    def ends(self, pid):
        r = self._p.get(str(pid)) or {}
        return r.get("in"), r.get("out")

    def length(self, pid):
        return float((self._p.get(str(pid)) or {}).get("length") or 0.0)

    def xyz(self, nid):
        r = self._n.get(str(nid)) or {}
        return (float(r.get("x") or 0.0), float(r.get("y") or 0.0),
                float(r.get("elevation") or 0.0))

    def root(self):
        for r in self.nodes:
            if str(r.get("io_node")) == "Input":
                return str(r.get("label"))
        return str((self.nodes[0] or {}).get("label")) if self.nodes else None

    def is_head(self, nid):
        return str(nid) in self._heads

    def protected(self, nid):
        from routes.module_f.merge import ANCHOR_LABEL
        if str(nid) == str(self.root()):
            return "급수원(Input)은 망의 정의입니다"
        if str(nid) == str(ANCHOR_LABEL):
            return (f"기준점(라벨 {ANCHOR_LABEL})은 세 도면이 만나는 "
                    "자리입니다 — 지울 수 없습니다")
        return None

    def seam(self, pid):
        return self._where["pipe"].get(str(pid)) == "seam"

    def bore(self, pid):
        try:
            return int((self._p.get(str(pid)) or {}).get("dia"))
        except (TypeError, ValueError):
            return None

    def next_node_name(self, op_id):
        # ⑧ 「맨 끝 번호 + 1」 — 숫자로 읽히는 라벨만 센다.
        top = 0
        for lab in self.node_ids():
            try:
                top = max(top, int(lab))
            except (TypeError, ValueError):
                continue
        return str(top + 1)

    def next_pipe_name(self, op_id):
        base, i = f"P{op_id}", 0
        have = set(self.pipe_ids())
        name = base
        while name in have:
            i += 1
            name = f"{base}_{i}"
        return name

    def _elev(self, row):
        a, b = row.get("in"), row.get("out")
        return round(self.xyz(b)[2] - self.xyz(a)[2], 3)

    def split(self, pid, new_pid, new_nid, xyz, l1, l2, up, dn, op_id):
        row = self._p[str(pid)]
        nrow = {"label": str(new_nid),
                "x": int(round(xyz[0])), "y": int(round(xyz[1])),
                "elevation": round(float(xyz[2]), 3), "io_node": "No"}
        self.nodes.append(nrow)
        self._n[str(new_nid)] = nrow
        new = dict(row)
        new["label"] = str(new_pid)
        new["in"], new["out"] = str(new_nid), str(dn)
        new["length"] = float(l2)
        self.pipes.append(new)
        self._p[str(new_pid)] = new
        row["in"], row["out"] = str(up), str(new_nid)
        row["length"] = float(l1)
        row["elev"], new["elev"] = self._elev(row), self._elev(new)
        # [§3-3-1 ⑤] 회랑과 **같은 규약** — 만든 이름을 id 로 적어 둔다.
        born = self.got.setdefault("ops_added", {})
        born.setdefault("merge_by_id", {})[str(op_id)] = str(new_nid)
        born.setdefault("merge_pipe_by_id", {})[str(op_id)] = str(new_pid)
        # 기기는 t 를 기준으로 갈라 보내고 위치를 다시 센다(§3-3-3).
        cut = (l1 / (l1 + l2)) if (l1 + l2) > 0 else 0.5
        for e in self.equipment:
            if str(e.get("pipe")) != str(pid):
                continue
            rp = float(e.get("rel_pos") or 0.0)
            if rp <= cut:
                e["rel_pos"] = round(rp / cut, 4) if cut > 0 else 0.0
                e["out"] = str(new_nid)
            else:
                e["pipe"] = str(new_pid)
                e["rel_pos"] = round((rp - cut) / (1.0 - cut), 4) if cut < 1 else 0.0
                e["in"], e["out"] = str(new_nid), str(dn)

    def absorb(self, keep, drop, far_a, far_b, length_m):
        p1 = self._p[str(keep)]
        p2 = self._p.get(str(drop)) or {}
        L1 = float(p1.get("length") or 0.0)
        p1["in"], p1["out"] = str(far_a), str(far_b)
        p1["length"] = float(length_m)
        p1["elev"] = self._elev(p1)
        cut = (L1 / length_m) if length_m > 0 else 0.5
        for e in self.equipment:
            if str(e.get("pipe")) == str(drop):
                e["pipe"] = str(keep)
                e["rel_pos"] = round(
                    cut + float(e.get("rel_pos") or 0.0) * (1.0 - cut), 4)
            elif str(e.get("pipe")) == str(keep):
                e["rel_pos"] = round(float(e.get("rel_pos") or 0.0) * cut, 4)
            else:
                continue
            e["in"], e["out"] = p1.get("in"), p1.get("out")
        for f in self.fittings:
            if str(f.get("pipe")) == str(drop):
                f["pipe"] = str(keep)
        self.pipes[:] = [r for r in self.pipes
                         if str(r.get("label")) != str(drop)]
        self._p.pop(str(drop), None)
        _ = p2

    def drop_pipe(self, pid):
        self.pipes[:] = [r for r in self.pipes
                         if str(r.get("label")) != str(pid)]
        self.fittings[:] = [r for r in self.fittings
                            if str(r.get("pipe")) != str(pid)]
        self.equipment[:] = [r for r in self.equipment
                             if str(r.get("pipe")) != str(pid)]
        self._p.pop(str(pid), None)

    def drop_node(self, nid):
        self.nodes[:] = [r for r in self.nodes
                         if str(r.get("label")) != str(nid)]
        # 노즐은 절점에 붙어 있다 — 절점을 지우면 그 행도 함께 지운다.
        self.nozzles[:] = [r for r in self.nozzles
                           if str(r.get("in")) != str(nid)]
        self._n.pop(str(nid), None)
        self._heads.discard(str(nid))


def apply_ops_to_merge(got, ops):
    """[§3-3-1 적용 ★] 계통도·기계실 대상 위상 수정 — 결합에서만 적용한다.

    회랑 요소는 여기서 손대지 않는다 — 이미 설계 표에서 고쳐져 흘러들었다
    (D8). 여기서 또 고치면 같은 수정이 두 번 먹는다.

    반환: (적용 건수, 못 한 것, 보고)
    """
    ops = [r for r in (ops or ())
           if str((r.get("target") or [None])[0]) in ("sys", "mr")]
    if not ops:
        return 0, [], {}
    if (got or {}).get("combined") is None:
        return 0, [{**r, "why": "결합망이 없습니다"} for r in ops], {}
    net = _MergeNet(got)
    where = net._where
    added: dict = {}

    def resolve(key):
        kind = str(key[0])
        if kind == "add":
            return added.get(str(key[1]), (None, None))
        lab = str(key[1])
        # 라벨은 절점과 배관이 따로 센다 — 둘 다 있으면 «절점» 이 먼저다.
        # (결합표의 배관 라벨은 P12·r1 처럼 생겨 절점 번호와 안 겹친다.)
        if lab in set(net.node_ids()):
            got_part = where["node"].get(lab)
            return ("node", lab) if got_part == _MERGE_PART.get(kind) else (None, None)
        if lab in set(net.pipe_ids()):
            got_part = where["pipe"].get(lab)
            # ★이음매는 «못 찾음» 이 아니라 **거절** 이다. 없는 셈 치면 화면이
            #   「범위 밖」이라고 말하는데, 사람은 그 배관을 눈으로 보고 있다 —
            #   왜 안 되는지를 말해야 한다(§3-3-2).
            if got_part == "seam":
                return "pipe", lab
            return ("pipe", lab) if got_part == _MERGE_PART.get(kind) else (None, None)
        return None, None

    rep: dict = {}
    n, missed, rep = run_ops(net, ops, resolve, report=rep)
    for op_id, nid in (rep.get("added", {}).get("nodes") or {}).items():
        added[str(op_id)] = ("node", nid)
    # 기기는 결합표에 바로 단다 — 여기서는 호칭경을 이미 안다.
    for r in (rep.get("equip") or ()):
        prow = net._p.get(str(r.get("pipe")))
        if prow is None:
            missed.append({**r, "why": "그 배관이 결합망에 없습니다"})
            continue
        net.equipment.append({
            "pipe": str(r.get("pipe")), "in": prow.get("in"),
            "out": prow.get("out"),
            "label": str(len(net.equipment) + 1),
            "desc": str(r.get("desc") or r.get("lib_id")),
            "eq_len": r.get("eq_len"), "rel_pos": r.get("rel_pos"),
            "lib": str(r.get("lib_id")),
            "op_id": str(r.get("op_id") or ""),
        })
        if r.get("eq_len") is None:
            missed.append({**r, "why": f"등가길이를 못 구했습니다 "
                                       f"(호칭경 {r.get('dia')})"})
    if n or missed:
        print(f"[수정] 통합 위상 {n}건을 적용했습니다"
              + (f" · 못 함 {len(missed)}건" if missed else ""))
    return n, missed, rep
