"""헤드 종류 정규화 SSOT. UI 없음.

권위: 1-1 분류 → 편집 kind_overrides. 찍기 heads[].kind 는 권위 아님.
미지정 가정(전부 상향) 금지.
"""
import math

KIND_OK = ("상향식", "하향식", "상하향식", "미지정")
CONFIRMED_KINDS = ("상향식", "하향식", "상하향식")
# 찍기 칸 — 스펙 JSON 저장 문자열. 화면 글자(상향식/하향식)와 다르다.
SLOT_UPDOWN = "상향하향"
SLOT_COMBO = "상하향"
SLOT_OK = (SLOT_UPDOWN, SLOT_COMBO)


def disk_key(hx, hy, hr):
    return (round(float(hx), 1), round(float(hy), 1), round(float(hr), 1))


def normalize_head_slot(label):
    """찍기 칸 — 구스펙·화면 글자 호환 [2026-08-08 오너].

    `하향`/`상향`/`측벽`/`상향식/하향식` → 상향하향.
    `상하향`/`상하향식` → 상하향.
    """
    s = str(label or "").strip()
    if s in ("상하향", "상하향식"):
        return SLOT_COMBO
    return SLOT_UPDOWN


def normalize_head_kind(kind):
    """판정·라벨 동의어 → 상하향식|하향식|상향식|미지정."""
    s = str(kind or "").strip()
    if s in ("하향", "하향식"):
        return "하향식"
    if s in ("상향", "상향식"):
        return "상향식"
    if "상하향" in s:
        return "상하향식"
    if s == "미지정":
        return "미지정"
    return s or "미지정"


def require_head_kinds(hcov, head_kinds):
    """평면/편집 그래프 — hcov 디스크마다 kind 필수.

    분류 결과가 없으면 미지정. 상향식 가정 금지.
    """
    out, by = [], {}
    for rec in head_kinds or ():
        if not isinstance(rec, dict):
            continue
        rec = dict(rec)
        rec["kind"] = normalize_head_kind(rec.get("kind"))
        if rec["kind"] not in KIND_OK:
            rec["kind"] = "미지정"
        out.append(rec)
        if "c" not in rec:
            continue
        c = rec["c"]
        if "head_r" in rec:
            by[disk_key(c[0], c[1], rec["head_r"])] = rec
        elif rec.get("tri_side"):
            by[disk_key(c[0], c[1],
                        float(rec["tri_side"]) / math.sqrt(3.0))] = rec
    for disk in hcov or ():
        hx, hy, hr = float(disk[0]), float(disk[1]), float(disk[2])
        k = disk_key(hx, hy, hr)
        if k in by:
            continue
        rec = {"c": (hx, hy), "head_r": hr, "kind": "미지정"}
        out.append(rec)
        by[k] = rec
    return out


def disk_kind_list(disks, head_kinds):
    """hcov 디스크마다 kind. disk_key 우선, 없으면 좌표만."""
    by = {}
    by_xy = {}
    for rec in head_kinds or ():
        if not isinstance(rec, dict) or "c" not in rec:
            continue
        c = rec["c"]
        kind = normalize_head_kind(rec.get("kind"))
        if kind not in KIND_OK:
            kind = "미지정"
        if "head_r" in rec:
            by[disk_key(c[0], c[1], rec["head_r"])] = kind
        by_xy[(round(float(c[0]), 1), round(float(c[1]), 1))] = kind
    out = []
    for d in disks or ():
        hx, hy, hr = float(d[0]), float(d[1]), float(d[2])
        kind = by.get(disk_key(hx, hy, hr))
        if kind is None:
            kind = by_xy.get((round(hx, 1), round(hy, 1)), "미지정")
        out.append(kind)
    return out
