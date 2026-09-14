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
from collections import defaultdict

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


def write_file(key: str, rows) -> str:
    p = path_for(key)
    if not rows:
        # ★«수정 없음» 은 **빈 파일이 아니라 파일 없음**이다. 0건짜리 파일을
        #   남겨 두면 작업폴더에 정체 모를 파일이 쌓이고, 사람이 그것을 보고
        #   「뭔가 덮여 있나」 하고 뒤진다. 마지막 수정을 지우면 자취도 지운다.
        try:
            os.remove(p)
        except OSError:
            pass
        return p
    os.makedirs(os.path.dirname(p) or ".", exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"version": VERSION, "items": list(rows)}, f,
                  ensure_ascii=False, indent=2)
    return p


def read_file(key: str) -> list:
    p = path_for(key)
    if not os.path.exists(p):
        return []
    try:
        with open(p, encoding="utf-8") as f:
            raw = json.load(f)
    except (OSError, ValueError) as exc:
        print(f"[수정] ★{os.path.basename(p)} 을 읽지 못했습니다 — {exc}")
        return []
    items = raw.get("items") if isinstance(raw, dict) else raw
    return [r for r in (items or []) if isinstance(r, dict)]


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

    for pid, ref in eref.items():
        if str(pid) in pr:
            idx["pipe"][key_pipe(ref[0], ref[1])] = str(pid)
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

      · 배관: 표 라벨 = kfp 배관 이름이라 그대로 잇는다.
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
    for k, pid in (idx.get("pipe") or {}).items():
        of_pipe[str(pid)] = list(k)
    for k, pid in (idx.get("vert") or {}).items():
        of_pipe.setdefault(str(pid), list(k))
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
                r["old"] = row.get(_META_PIPE[field])
                row[_META_PIPE[field]] = new
                n_meta += 1
        elif kind == "head" and field in _META_HEAD:
            m = nd.get(str(name))
            if m is not None:
                keys = _META_HEAD[field]
                r["old"] = m.get(keys[0])
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
        "index": idx, "relay": relay, "applied": applied}


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
    kfp = (got or {}).get("kfp") or {}
    nd = kfp.get("nodes_meta_runtime") or {}
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
            for row in (getattr(tbl, "pipes", None) or ()):
                if str(row.get("label")) != str(name):
                    continue
                if field == "length":
                    # ★길이는 여기서 안 덮는다. 적용 ① 이 kfp 길이를 바꾸고
                    #   표가 그 값을 그대로 받아 온다 — 여기서 또 덮으면 같은
                    #   값을 두 번 덮게 되고, `old` 가 «이미 덮인 값» 으로
                    #   채워져 원값을 잃는다.
                    break
                r["old"] = row.get(field)
                row[field] = new
                if field == "dia":
                    # 관경은 «무엇이 정했나» 가 행에 남는다 — 집계만으로는
                    # 도면 텍스트에서 온 것인지 사람이 넣은 것인지 모른다.
                    row["dia_src"] = "사람"
                n_tbl += 1
                break
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
        want = _MERGE_PART.get(kind)
        if want is None:
            continue                      # 회랑 요소 — 설계 표에서 이미 덮였다
        key = key_from_json(r.get("key"))
        field = str(r.get("field"))
        if key is None or len(key) < 2:
            missed.append({**r, "why": "키를 읽지 못했습니다"})
            continue
        lab = str(key[1])
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
        r["old"] = row.get(col)
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
