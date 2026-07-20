#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""SDF 관경(bore) 치환 도구 — 원인 확정 실험용 독립 스크립트.

배경
    자동 추출 SDF와 수작업 SDF의 PIPENET 계산 결과 차이가 "관경"에서 오는지
    확정하려는 통제 실험용. 관경은 하겐-윌리엄스 마찰손실에서 d^-4.87 로 작용해
    압력 계산을 지배하므로, 다른 변수(토폴로지/길이/표고)를 고정한 채 관경만
    바꿔보면 원인을 분리할 수 있다.

설계 원칙
    - 스키마 무손실: <Pipe> 태그의 bore 속성값만 인라인 문자열 치환한다.
      DOCTYPE / Graphics / Attributes / 들여쓰기까지 나머지는 바이트 그대로 보존
      (ElementTree 재직렬화로 DOCTYPE·포맷이 깨지는 재발버그 회피).
    - 독립 실행: 프로젝트 모듈 import 없음. 표준 라이브러리만 사용.
    - 한글 경로/파일명 지원: 입출력 모두 UTF-8.
    - SDF bore 단위는 '미터'(0.1 = 100mm). 사다리는 호칭경 mm 기준.

모드 (택1)
    --from-sdf REF.sdf   레퍼런스 SDF의 관경을 대상 SDF의 각 배관에 이식(graft)
        --match {nodepair,label,midpoint,auto}   (기본 auto = nodepair→label 폴백)
        --mid-tol M      midpoint 매칭 허용 반경(도면 단위, 기본 500)
    --set-mm N           전 배관을 N mm 균일 관경으로 강제 (변수 통제 실험)
    --bump K             각 배관을 사다리에서 K 단계 이동(+상향/-하향), K 정수
    --scale F            각 배관 내경 × F 후 사다리 스냅

공통
    -o OUT.sdf           출력 경로(미지정 시 <입력>_bore.sdf)
    --report-only        파일 쓰지 않고 변경 예정 내역만 출력
    --limit N            변경 로그 상위 N행만 출력(기본 25, 0=전체)

예시
    # 수작업 관경을 자동 SDF에 이식 (원인 확정)
    python sdf_bore_swap.py auto.sdf --from-sdf "2. 출력 배관망_수작업.sdf" -o auto_grafted.sdf
    # 두 SDF를 100mm 균일로 강제 후 각각 PIPENET 재계산 비교
    python sdf_bore_swap.py auto.sdf --set-mm 100 -o auto_100.sdf
    python sdf_bore_swap.py "2. 출력 배관망_수작업.sdf" --set-mm 100 -o hand_100.sdf
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

# 호칭경 사다리 (mm) — 표준 스케줄. 독립 유지 위해 로컬 정의.
LADDER_MM = [25, 32, 40, 50, 65, 80, 100, 125, 150, 200]

_ATTR_RE = re.compile(r'([\w:-]+)\s*=\s*"([^"]*)"')
# <Pipe ...>  또는  <Pipe ... />.  뒤에 공백을 강제(?=\s)해 <Pipe-set>/<Pipe-type>
# (Pipe 뒤 '-') 를 배제한다. 속성엔 '>' 가 없으므로 [^>]* 안전.
_PIPE_TAG_RE = re.compile(r'<Pipe(?=\s)[^>]*?/?>')
_BORE_RE = re.compile(r'(bore\s*=\s*")([^"]*)(")')


def _fmt_m(mm: float) -> str:
    """mm → 미터 문자열 (SDF bore 표기). 0.08, 0.125, 0.1 처럼 간결."""
    return f"{mm / 1000.0:g}"


def _snap_ladder_mm(mm: float) -> int:
    """가장 가까운(같거나 큰) 사다리 호칭경. 상한 200 클램프."""
    for d in LADDER_MM:
        if d >= mm - 1e-9:
            return d
    return LADDER_MM[-1]


def _parse_attrs(tag: str) -> dict:
    return {k: v for k, v in _ATTR_RE.findall(tag)}


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _iter_pipe_attrs(text: str):
    """(match_obj, attrs_dict) for every <Pipe> opening tag in text."""
    for m in _PIPE_TAG_RE.finditer(text):
        yield m, _parse_attrs(m.group(0))


def _node_coords(text: str) -> dict[str, tuple[float, float]]:
    """label → (x, y).  <Node label=..><Position x=.. y=../> 구조."""
    coords: dict[str, tuple[float, float]] = {}
    for nm in re.finditer(r'<Node\b[^>]*>(.*?)</Node>', text, re.S):
        block = nm.group(0)
        a = _parse_attrs(nm.group(0).split('>', 1)[0] + '>')
        lbl = a.get("label")
        pos = re.search(r'<Position\b[^>]*>', block)
        if lbl and pos:
            pa = _parse_attrs(pos.group(0))
            try:
                coords[lbl] = (float(pa.get("x", "0")), float(pa.get("y", "0")))
            except ValueError:
                pass
    return coords


# ── 레퍼런스 SDF에서 관경 룩업 구축 ──────────────────────────────────────────
def build_ref_lookup(ref_text: str):
    by_nodepair_dir: dict[tuple[str, str], str] = {}
    by_nodepair_und: dict[frozenset, str] = {}
    by_label: dict[str, str] = {}
    mids: list[tuple[float, float, str]] = []  # (mx, my, bore)
    coords = _node_coords(ref_text)
    for _, a in _iter_pipe_attrs(ref_text):
        bore = a.get("bore")
        if bore is None:
            continue
        i, o, lbl = a.get("input"), a.get("output"), a.get("label")
        if i and o:
            by_nodepair_dir[(i, o)] = bore
            by_nodepair_und[frozenset((i, o))] = bore
            if i in coords and o in coords:
                mx = (coords[i][0] + coords[o][0]) / 2.0
                my = (coords[i][1] + coords[o][1]) / 2.0
                mids.append((mx, my, bore))
        if lbl:
            by_label[lbl] = bore
    return by_nodepair_dir, by_nodepair_und, by_label, mids


def _nearest_mid(mx: float, my: float, mids, tol: float):
    best_b, best_d = None, tol
    for rx, ry, b in mids:
        d = ((rx - mx) ** 2 + (ry - my) ** 2) ** 0.5
        if d <= best_d:
            best_d, best_b = d, b
    return best_b, (best_d if best_b is not None else None)


def compute_new_bore(a: dict, mode: dict, ref, tgt_coords):
    """이 배관의 새 bore(미터 문자열) 또는 None(변경 없음/미매칭). + 매칭방법 태그."""
    old = a.get("bore")
    if mode["kind"] == "set_mm":
        return _fmt_m(mode["mm"]), "set"
    if mode["kind"] == "bump":
        cur_mm = float(old) * 1000.0 if old else 100.0
        cur = _snap_ladder_mm(cur_mm)
        idx = LADDER_MM.index(cur)
        nidx = max(0, min(len(LADDER_MM) - 1, idx + mode["k"]))
        return _fmt_m(LADDER_MM[nidx]), "bump"
    if mode["kind"] == "scale":
        cur_mm = float(old) * 1000.0 if old else 100.0
        return _fmt_m(_snap_ladder_mm(cur_mm * mode["f"])), "scale"
    if mode["kind"] == "graft":
        dir_, und, lbl, mids = ref
        i, o, label = a.get("input"), a.get("output"), a.get("label")
        match = mode["match"]
        # auto 폴백 체인: 노드쌍(방향→무방향) → 라벨 → 미매칭.
        # ⚠ 라벨 tier 는 두 SDF의 라벨 공간이 독립이면 물리적으로 다른 배관끼리의
        #   우연 충돌일 수 있다. 리포트에서 nodepair 와 분리 표기(<label> 경고).
        if match in ("nodepair", "auto"):
            if i and o and (i, o) in dir_:
                return dir_[(i, o)], "nodepair"
            if i and o and frozenset((i, o)) in und:
                return und[frozenset((i, o))], "nodepair~"  # 방향 반전 매칭
        if match in ("label", "auto"):
            if label and label in lbl:
                return lbl[label], "label"
        if match == "midpoint" and i in tgt_coords and o in tgt_coords:
            mx = (tgt_coords[i][0] + tgt_coords[o][0]) / 2.0
            my = (tgt_coords[i][1] + tgt_coords[o][1]) / 2.0
            b, _d = _nearest_mid(mx, my, mids, mode["mid_tol"])
            if b is not None:
                return b, "midpoint"
        return None, "UNMATCHED"
    return None, "none"


def main(argv=None):
    ap = argparse.ArgumentParser(description="SDF 관경(bore) 치환 도구")
    ap.add_argument("target", help="대상 SDF (이 파일의 관경을 바꿈)")
    ap.add_argument("-o", "--out", help="출력 SDF 경로")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--from-sdf", help="레퍼런스 SDF의 관경을 이식(graft)")
    g.add_argument("--set-mm", type=float, help="전 배관을 N mm 균일 관경으로")
    g.add_argument("--bump", type=int, help="사다리 K단계 이동(+상향/-하향)")
    g.add_argument("--scale", type=float, help="내경 × F 후 사다리 스냅")
    ap.add_argument("--match", choices=["nodepair", "label", "midpoint", "auto"],
                    default="auto",
                    help="graft 매칭 키. auto=노드쌍→라벨→미매칭리포트(기본). "
                         "라벨 tier 는 우연충돌 가능(리포트에서 분리 표기).")
    ap.add_argument("--mid-tol", type=float, default=500.0, help="midpoint 허용 반경")
    ap.add_argument("--report-only", action="store_true", help="파일 안 쓰고 리포트만")
    ap.add_argument("--limit", type=int, default=25, help="변경 로그 상위 N행(0=전체)")
    args = ap.parse_args(argv)

    tgt_path = Path(args.target)
    if not tgt_path.is_file():
        ap.error(f"대상 SDF 없음: {tgt_path}")
    text = _read_text(tgt_path)

    if args.from_sdf is not None:
        ref_path = Path(args.from_sdf)
        if not ref_path.is_file():
            ap.error(f"레퍼런스 SDF 없음: {ref_path}")
        ref = build_ref_lookup(_read_text(ref_path))
        mode = {"kind": "graft", "match": args.match, "mid_tol": args.mid_tol}
        print(f"[REF] {ref_path.name}: nodepair={len(ref[0])} label={len(ref[2])} midpoints={len(ref[3])}")
    elif args.set_mm is not None:
        ref, mode = None, {"kind": "set_mm", "mm": args.set_mm}
    elif args.bump is not None:
        ref, mode = None, {"kind": "bump", "k": args.bump}
    else:
        ref, mode = None, {"kind": "scale", "f": args.scale}

    tgt_coords = _node_coords(text) if (ref and args.match == "midpoint") else {}

    # 각 Pipe 태그를 순회하며 새 bore 계산 → 인라인 치환.
    changes = []          # (label, in, out, old, new, method)
    method_counts: dict[str, int] = {}
    total = 0

    def _sub(m: re.Match) -> str:
        nonlocal total
        total += 1
        tag = m.group(0)
        a = _parse_attrs(tag)
        new, method = compute_new_bore(a, mode, ref, tgt_coords)
        method_counts[method] = method_counts.get(method, 0) + 1
        old = a.get("bore")
        if new is None or new == old:
            if new is None:
                changes.append((a.get("label"), a.get("input"), a.get("output"),
                                old, None, method))
            return tag
        changes.append((a.get("label"), a.get("input"), a.get("output"),
                        old, new, method))
        return _BORE_RE.sub(lambda bm: bm.group(1) + new + bm.group(3), tag, count=1)

    new_text = _PIPE_TAG_RE.sub(_sub, text)

    changed = [c for c in changes if c[4] is not None]
    unmatched = [c for c in changes if c[4] is None]
    print(f"[SCAN] pipes={total}  changed={len(changed)}  "
          f"unchanged/unmatched={total - len(changed)}")
    print(f"[MATCH-METHOD] {method_counts}")

    lim = args.limit if args.limit > 0 else 10 ** 9

    def _dump(rows, tag):
        print(f"  [{tag}] {len(rows)}건")
        for k, (lbl, i, o, old, new, method) in enumerate(rows):
            if k >= lim:
                print(f"      ... (+{len(rows) - k} more)")
                break
            arrow = f"{old}->{new}" if new is not None else f"{old} (원본 유지)"
            print(f"      pipe {lbl!s:>5} [{i}->{o}]  bore {arrow}  <{method}>")

    if mode['kind'] == 'graft':
        # tier 분리: 노드쌍(신뢰) / 라벨(⚠ 우연충돌 가능) / 미매칭.
        np_rows = [c for c in changed if c[5] in ("nodepair", "nodepair~")]
        lb_rows = [c for c in changed if c[5] == "label"]
        mp_rows = [c for c in changed if c[5] == "midpoint"]
        if total:
            print(f"[GRAFT-RATE] node쌍={len(np_rows)}  라벨={len(lb_rows)}"
                  + (f"  midpoint={len(mp_rows)}" if mp_rows else "")
                  + f"  미매칭={len(unmatched)}  (총 {total}, 이식 {100.0*len(changed)/total:.1f}%)")
        _dump(np_rows, "TIER1 노드쌍 (신뢰)")
        if lb_rows:
            print("  ⚠ 라벨 tier: 두 SDF 라벨 공간이 독립이면 물리적으로 다른 배관 간"
                  " 우연 충돌일 수 있음 — 아래 목록을 직접 검증하세요.")
            _dump(lb_rows, "TIER2 라벨 (⚠ 검증 필요)")
        if mp_rows:
            _dump(mp_rows, "TIER3 midpoint")
        _dump(unmatched, "미매칭 리포트 (원본 관경 유지)")
    else:
        _dump(changes[:], "변경")

    if args.report_only:
        print("[REPORT-ONLY] 파일을 쓰지 않았습니다.")
        return 0

    out_path = Path(args.out) if args.out else tgt_path.with_name(tgt_path.stem + "_bore.sdf")
    out_path.write_text(new_text, encoding="utf-8")
    # 스키마 무손실 확인: 재파싱 가능 + bore 개수 보존
    import xml.etree.ElementTree as ET
    n_pipe = sum(1 for _ in ET.parse(out_path).getroot().iter("Pipe"))
    print(f"[WRITE] {out_path}  ({out_path.stat().st_size}B, pipes={n_pipe})  재파싱 OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
