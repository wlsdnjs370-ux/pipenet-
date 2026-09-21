# -*- coding: utf-8 -*-
"""회귀 기준선이 «입력 변화» 와 «코드 회귀» 를 구분하게 한다.

지금은 board 가 바뀌기만 해도 「.kfp 가 달라졌다 — 어떤 항목도 완료가 아니다」로
운다(B6). 그러면 진짜 코드 회귀가 났을 때 그 경고를 믿지 않게 된다.
board 지문을 함께 기록해 원인을 갈라 말하게 만든다.
"""
import io

p = "cad_project_editor_g/tests/_kfp_baseline.py"
s = io.open(p, encoding="utf-8").read()

# ── build_kfp 가 board 지문도 돌려주도록 ────────────────────────────
old = '''    res = convert_to_kfp(payload, str(out), **dto_to_convert_kwargs(default_dto()))
    if not res["ok"]:
        raise SystemExit(f"변환 실패: {res.get('blockers')}")
    return out, res'''
new = '''    res = convert_to_kfp(payload, str(out), **dto_to_convert_kwargs(default_dto()))
    if not res["ok"]:
        raise SystemExit(f"변환 실패: {res.get('blockers')}")
    b = es.board
    res["_board"] = {"pts": len(b.pts), "edges": len(b.edges),
                     "disks": len(b.disks), "sources": len(b.sources)}
    return out, res'''
assert old in s
s = s.replace(old, new, 1)

old2 = '''    cur["nodes"] = len(kfp.get("nodes_meta_runtime") or {})
    cur["pipes"] = len(kfp.get("pipe_data") or {})'''
new2 = '''    cur["nodes"] = len(kfp.get("nodes_meta_runtime") or {})
    cur["pipes"] = len(kfp.get("pipe_data") or {})
    cur["board"] = res.get("_board")'''
assert old2 in s
s = s.replace(old2, new2, 1)

# ── 비교: board 가 다르면 «입력 변화» 로 갈라 말한다 ─────────────────
old3 = '''    same = (old.get("shape") == cur["shape"]
            and old["nodes"] == cur["nodes"] and old["pipes"] == cur["pipes"])'''
new3 = '''    # ★board 가 다르면 그것은 «입력이 바뀐 것» 이지 코드 회귀가 아니다.
    #   둘을 같은 빨간불로 알리면 진짜 회귀가 났을 때 그 경고를 안 믿게 된다(B6).
    board_same = (old.get("board") is None or old.get("board") == cur.get("board"))
    same = (old.get("shape") == cur["shape"]
            and old["nodes"] == cur["nodes"] and old["pipes"] == cur["pipes"])
    if not board_same:
        print(f"  기준선 board {old.get('board')}")
        print(f"  현재   board {cur.get('board')}")
        print("\\n[정보] 입력(board)이 달라졌다 — 코드 회귀가 아니다."
              "\\n       작업 폴더의 표시 캐시가 다시 만들어진 것이다(BLOCKED B6)."
              "\\n       코드를 검증하려면 `make` 로 기준선을 다시 뜬 뒤 비교하라.")
        return 2'''
assert old3 in s
s = s.replace(old3, new3, 1)

io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("기준선이 입력 변화와 코드 회귀를 구분하게 함")
