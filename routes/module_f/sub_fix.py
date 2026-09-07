# -*- coding: utf-8 -*-
"""[§27 후속] 뽑힌 계통도·기계실 배관을 **사람이 고친다**.

■ 왜 필요한가 — 실측이 정했다

  대명동 계통도를 뽑으면 배관 53개가 나오는데 **관경이 53개 모두 «추측 150A»**
  다(`dia_source="default"` · 도면 치수 텍스트 매치 0건). 표고 근거도 전부
  기본값이다. 그런데 이 값들은 그대로 `merge_network` 로 흘러 최종 SDF 의
  입상관이 된다. 사람이 볼 자리도, 고칠 자리도 없었다.

  그리고 그래프는 조각나 있어 프로그램이 200mm~10m «다리» 를 61~66개 놓아
  잇는다. 그 다리 길이도 배관 연장에 섞인다(경로의 92~95%가 다리를 하나
  이상 밟는다 · 연장의 1~3%, 나쁠 땐 그 이상). **그 자동 판정을 늘리는 대신
  사람이 고칠 수 있게 한다** — 이 저장소가 §27 에서 두 번 확인한 방향이다.

■ 규약 세 가지

  ⑴ **자리는 좌표로 가리킨다.** 라벨(`r1`·`m3`)은 경로 순서로 매겨져, 다시
     뽑으면 같은 이름이 다른 배관을 가리킨다(D-F11-4 가 겪은 그 사고). 그래서
     키는 «정렬된 두 끝 좌표» 다 — 다시 뽑아도 같은 구간에 다시 붙는다.
  ⑵ **원래 값을 지운 적이 없다.** 덮은 자리는 `orig_*` 로 남기고 화면이 늘
     같이 보여 준다. 덮기를 지우면 원래 값으로 정확히 돌아온다.
  ⑶ **못 붙인 것은 조용히 버리지 않는다.** 다시 뽑아 그 구간이 사라졌으면
     몇 개가 못 붙었는지 세어 보고한다.
"""
from __future__ import annotations

SLOT_KEY = {"system": "riser", "machineroom": "machineroom"}
MAX_ROWS = 2000
NOTE_MAX = 200


def _pt(node) -> tuple:
    return (int(round(float(node.get("x") or 0))),
            int(round(float(node.get("y") or 0))))


def _key(a: tuple, b: tuple) -> tuple:
    """방향을 지운 두 끝 좌표. 뽑는 방향이 바뀌어도 같은 자리를 가리킨다."""
    return (a, b) if a <= b else (b, a)


def pipe_keys(got: dict) -> dict:
    """배관 라벨 → 좌표 키. 절점 표에서 좌표를 끌어온다."""
    at = {str(n.get("label")): _pt(n) for n in (got.get("nodes") or ())}
    out = {}
    for p in (got.get("pipes") or ()):
        a, b = at.get(str(p.get("in"))), at.get(str(p.get("out")))
        if a is not None and b is not None:
            out[str(p.get("label"))] = _key(a, b)
    return out


def _as_key(row) -> tuple:
    a = (int(round(float(row["a"][0]))), int(round(float(row["a"][1]))))
    b = (int(round(float(row["b"][0]))), int(round(float(row["b"][1]))))
    return _key(a, b)


def parse_rows(rows) -> list:
    """화면이 보낸 덮기 목록을 검사해 정규형으로. 문제가 있으면 ValueError."""
    if rows is None:
        rows = []
    if not isinstance(rows, list):
        raise ValueError("덮기 목록이 올바르지 않습니다.")
    if len(rows) > MAX_ROWS:
        raise ValueError(f"덮기 목록이 너무 깁니다 (최대 {MAX_ROWS}).")
    clean, seen = [], set()
    for r in rows:
        if not isinstance(r, dict):
            raise ValueError("덮기 항목은 객체여야 합니다.")
        try:
            key = _as_key(r)
        except (KeyError, TypeError, ValueError, IndexError):
            raise ValueError(
                "덮기 항목에 a, b 두 끝 좌표가 있어야 합니다 (예: a=[x,y]).")
        if key in seen:
            raise ValueError("같은 배관을 두 번 덮었습니다.")
        seen.add(key)
        item = {"a": list(key[0]), "b": list(key[1])}
        if r.get("dia") not in (None, ""):
            try:
                dia = int(r["dia"])
            except (TypeError, ValueError):
                raise ValueError(f"호칭경이 숫자가 아닙니다: {r.get('dia')!r}")
            if not (6 <= dia <= 3000):
                raise ValueError(f"호칭경이 범위를 벗어납니다: {dia}")
            item["dia"] = dia
        if r.get("length") not in (None, ""):
            try:
                ln = float(r["length"])
            except (TypeError, ValueError):
                raise ValueError(f"길이가 숫자가 아닙니다: {r.get('length')!r}")
            # ★0 을 허용하면 PIPENET 이 그 배관을 못 푼다. 지우려면 덮기를
            #   빼야지, 길이를 0 으로 만드는 길을 열어 두면 안 된다.
            if not (0.001 <= ln <= 100000.0):
                raise ValueError(f"길이가 범위를 벗어납니다: {ln} m")
            item["length"] = round(ln, 3)
        if "dia" not in item and "length" not in item:
            raise ValueError("관경이나 길이 중 하나는 있어야 합니다.")
        note = str(r.get("note") or "").strip()
        if len(note) > NOTE_MAX:
            raise ValueError(f"사유가 너무 깁니다 ({NOTE_MAX}자).")
        item["note"] = note
        clean.append(item)
    return clean


def apply_overrides(got: dict, rows) -> dict:
    """추출 결과에 덮기를 **다시** 입힌다 — 늘 원래 값에서 출발한다.

    ★«덮은 위에 덮기» 를 하면 원래 값이 한 번 씻겨 나간다. 그래서 먼저 전부
      원상 복구하고 나서 지금 목록을 입힌다. 덮기를 빼면 정확히 되돌아온다.
    """
    pipes = list((got or {}).get("pipes") or ())
    want = {_as_key(r): r for r in (rows or ())}
    keys = pipe_keys(got or {})
    applied = 0
    for p in pipes:
        # ① 원상 복구
        for fld in ("dia", "length"):
            ok = f"orig_{fld}"
            if ok in p:
                p[fld] = p.pop(ok)
        if "orig_dia_source" in p:
            p["dia_source"] = p.pop("orig_dia_source")
        p.pop("fix_note", None)
        p.pop("fixed", None)
        # ② 지금 목록 입히기
        r = want.get(keys.get(str(p.get("label"))))
        if not r:
            continue
        applied += 1
        if "dia" in r and int(r["dia"]) != int(p.get("dia") or 0):
            p["orig_dia"] = p.get("dia")
            p["orig_dia_source"] = p.get("dia_source")
            p["dia"] = int(r["dia"])
            p["dia_source"] = "user_fix"
        if "length" in r and float(r["length"]) != float(p.get("length") or 0):
            p["orig_length"] = p.get("length")
            p["length"] = float(r["length"])
        p["fixed"] = True
        if r.get("note"):
            p["fix_note"] = r["note"]
    if pipes:
        got["pipes"] = pipes
        got["total_length_m"] = round(
            sum(float(p.get("length") or 0.0) for p in pipes), 3)
    return {"applied": applied, "given": len(want),
            "unmatched": max(0, len(want) - applied)}


def rows_for_view(got: dict) -> list:
    """화면이 그릴 배관표 — 덮은 자리는 원래 값을 **같이** 싣는다."""
    keys = pipe_keys(got or {})
    out = []
    for p in ((got or {}).get("pipes") or ()):
        lab = str(p.get("label"))
        k = keys.get(lab)
        row = {
            "label": lab, "in": p.get("in"), "out": p.get("out"),
            "dia": p.get("dia"), "length": p.get("length"),
            "elev": p.get("elev"), "dia_source": p.get("dia_source"),
            "a": list(k[0]) if k else None, "b": list(k[1]) if k else None,
        }
        for fld in ("orig_dia", "orig_length", "orig_dia_source",
                    "fix_note", "fixed"):
            if fld in p:
                row[fld] = p[fld]
        out.append(row)
    return out
