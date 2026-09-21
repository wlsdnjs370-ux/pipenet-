# -*- coding: utf-8 -*-
"""회귀 기준선을 «바이트» 에서 «구조» 비교로 바꾼다.

sha256 으로는 못 잰다: `build_planar_graph` 의 노드 번호가 집합 순회 순서를 타서
같은 코드·같은 입력에도 실행마다 달라진다(실측 41,892줄 중 97줄이 전부 노드 id).
바이트 비교는 코드 회귀가 아니라 번호 뽑기를 재는 셈이다.
"""
import io

p = "cad_project_editor_g/tests/_kfp_baseline.py"
lines = io.open(p, encoding="utf-8").read().split("\n")

# ── digest() 교체 ────────────────────────────────────────────────
i = next(k for k, ln in enumerate(lines) if ln.startswith("def digest("))
j = next(k for k in range(i + 1, len(lines)) if lines[k].startswith("def "))
new_digest = '''def digest(path: Path) -> dict:
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

'''.split("\n")
lines[i:j] = new_digest

src = "\n".join(lines)
src = src.replace(
    """              f"{cur['bytes']:,} bytes\\n  sha256 {cur['sha256'][:16]}…")""",
    """              f"{cur['bytes']:,} bytes\\n  구조 {cur['shape'][:16]}…")""")
src = src.replace(
    """    same = old["sha256"] == cur["sha256"]""",
    """    same = (old.get("shape") == cur["shape"]
            and old["nodes"] == cur["nodes"] and old["pipes"] == cur["pipes"])""")
src = src.replace(
    """    print(f"  기준선 노드 {old['nodes']} 배관 {old['pipes']} {old['bytes']:,}B "
          f"{old['sha256'][:16]}…")
    print(f"  현재   노드 {cur['nodes']} 배관 {cur['pipes']} {cur['bytes']:,}B "
          f"{cur['sha256'][:16]}…")
    print("\\n[OK  ] 전체망 .kfp 비트 동일" if same""",
    """    print(f"  기준선 노드 {old['nodes']} 배관 {old['pipes']} "
          f"모양 {str(old.get('shape'))[:16]}…")
    print(f"  현재   노드 {cur['nodes']} 배관 {cur['pipes']} "
          f"모양 {cur['shape'][:16]}…")
    print("\\n[OK  ] 전체망 .kfp 구조 동일" if same""")

io.open(p, "w", encoding="utf-8", newline="\n").write(src)
print("기준선 구조 비교 전환 완료")
