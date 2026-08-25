# -*- coding: utf-8 -*-
"""전체망 .kfp 회귀 기준선 — 이번 작업 전/후로 **비트 동일**해야 한다(지시서 §4).

`.kfp` 는 솔버가 전체망을 받아 설계구역을 스스로 고르는 경로다. 이번 작업은
SDF 쪽에만 손대므로, 이 파일이 한 바이트라도 달라지면 어떤 항목도 완료가 아니다.

    python tests/_kfp_baseline.py make    기준선 생성(작업 전 1회)
    python tests/_kfp_baseline.py check   현재 산출과 대조
"""
from __future__ import annotations

import hashlib
import json
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
os.chdir(_ROOT)          # 편집기는 작업 폴더를 cwd 기준으로 잡는다

KEY = "B1F 현장조사 소화설비 평면도"
OUT_DIR = _ROOT / "tests" / "_out"
BASE = _ROOT / "tests" / "_kfp_baseline.json"


def build_kfp() -> tuple[Path, dict]:
    """손질 저장본 → 전체망 .kfp. 기존 경로를 그대로 탄다."""
    from services.cad_import.edit.session import EditSession
    from services.cad_import.convert.engine import convert_to_kfp, ensure_planar
    from services.cad_import.dto import default_dto, dto_to_convert_kwargs

    from services.cad_import.convert.planar import pick_convert_sources

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    # 이 저장본은 급수원이 둘이다. 변환은 하나를 지정해야 하므로 첫 번째(Z1)를
    # 못박는다 — 기준선이므로 «매번 같은 것» 이 값 자체보다 중요하다.
    srcs = payload.get("sources") or ()
    if len(srcs) > 1:
        picked, err = pick_convert_sources(srcs, srcs[0].get("tag"))
        if err:
            raise SystemExit(f"급수원 선택 실패: {err}")
        payload["sources"] = picked
    payload = ensure_planar(payload)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "baseline_full.kfp"
    res = convert_to_kfp(payload, str(out), **dto_to_convert_kwargs(default_dto()))
    if not res["ok"]:
        raise SystemExit(f"변환 실패: {res.get('blockers')}")
    b = es.board
    res["_board"] = {"pts": len(b.pts), "edges": len(b.edges),
                     "disks": len(b.disks), "sources": len(b.sources)}
    return out, res


def digest(path: Path) -> dict:
    """구조 지문 — 바이트가 아니라 «망의 모양» 을 잰다.

    ★sha256 으로는 못 잰다. `build_planar_graph` 의 노드 번호(N1145 …)가 집합
    순회 순서를 타서 **같은 코드·같은 입력에도 실행마다 달라진다**(실측: 41,892
    줄 중 97줄이 다르고 전부 노드 id, 크기는 같다). 바이트 비교는 코드 회귀가
    아니라 번호 뽑기를 재는 셈이다.

    대신 이름에 안 기대는 것만 본다 — 노드/배관 수, 배관 길이·호칭경의 정렬된
    목록. 코드가 망을 바꾸면 이 셋 중 하나는 반드시 움직인다.
    """
    raw = path.read_bytes()
    kfp = json.loads(raw.decode("utf-8"))
    pipes = kfp.get("pipe_data") or {}
    lens = sorted(round(float((q or {}).get("length_m") or 0.0), 3)
                  for q in pipes.values())
    dias = sorted(int((q or {}).get("nominal_mm") or 0) for q in pipes.values())
    shape = json.dumps({"lens": lens, "dias": dias}, sort_keys=True)
    return {"shape": hashlib.sha256(shape.encode()).hexdigest(),
            "bytes": len(raw)}


def main() -> int:
    mode = sys.argv[1] if len(sys.argv) > 1 else "check"
    out, res = build_kfp()
    cur = digest(out)
    kfp = res["kfp"]
    cur["nodes"] = len(kfp.get("nodes_meta_runtime") or {})
    cur["pipes"] = len(kfp.get("pipe_data") or {})
    cur["board"] = res.get("_board")

    if mode == "make":
        BASE.write_text(json.dumps(cur, ensure_ascii=False, indent=2),
                        encoding="utf-8")
        print(f"기준선 생성 · 노드 {cur['nodes']} · 배관 {cur['pipes']} · "
              f"{cur['bytes']:,} bytes\n  구조 {cur['shape'][:16]}…")
        return 0

    if not BASE.exists():
        print("!! 기준선이 없다 — 먼저 `make` 로 만들 것")
        return 1
    old = json.loads(BASE.read_text(encoding="utf-8"))
    # ★board 가 다르면 그것은 «입력이 바뀐 것» 이지 코드 회귀가 아니다.
    #   둘을 같은 빨간불로 알리면 진짜 회귀가 났을 때 그 경고를 안 믿게 된다(B6).
    board_same = (old.get("board") is None or old.get("board") == cur.get("board"))
    same = (old.get("shape") == cur["shape"]
            and old["nodes"] == cur["nodes"] and old["pipes"] == cur["pipes"])
    if not board_same:
        print(f"  기준선 board {old.get('board')}")
        print(f"  현재   board {cur.get('board')}")
        print("\n[정보] 입력(board)이 달라졌다 — 코드 회귀가 아니다."
              "\n       작업 폴더의 표시 캐시가 다시 만들어진 것이다(BLOCKED B6)."
              "\n       코드를 검증하려면 `make` 로 기준선을 다시 뜬 뒤 비교하라.")
        return 2
    print(f"  기준선 노드 {old['nodes']} 배관 {old['pipes']} "
          f"모양 {str(old.get('shape'))[:16]}…")
    print(f"  현재   노드 {cur['nodes']} 배관 {cur['pipes']} "
          f"모양 {cur['shape'][:16]}…")
    print("\n[OK  ] 전체망 .kfp 구조 동일" if same
          else "\n[FAIL] 전체망 .kfp 가 달라졌다 — 어떤 항목도 완료가 아니다")
    return 0 if same else 1


if __name__ == "__main__":
    sys.exit(main())
