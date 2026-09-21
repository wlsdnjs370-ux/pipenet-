# -*- coding: utf-8 -*-
"""[G2 · B4 1안] select_and_expand() — 붙일 수 있는 헤드로 좁혀 선정 후 제한 전개."""
import io

p = "cad_project_editor_g/services/cad_import/design/restrict.py"
s = io.open(p, encoding="utf-8").read()

tail = '''

def select_and_expand(payload: dict, board, *, k=None, only_heads=None,
                      selected_source=None, key: str | None = None) -> dict:
    """최불리 선정 → corridor 제한 전개를 한 번에. G7 창이 부르는 진입점.

    ★선정 후보를 **전개가 붙일 수 있는 헤드**로 먼저 좁힌다(BLOCKED B4 · 1안).
    선정은 board 도달로 헤드를 세고 전개는 더 엄격한 규칙을 쓰므로, 좁히지
    않으면 최불리 K 가 통째로 전개 밖으로 떨어져 빈 망이 나온다.

    `only_heads`(도면 장 제한)가 함께 오면 **교집합**을 쓴다 — 장 제한이 먼저고
    그 안에서 붙일 수 있는 것만 남긴다.

    반환의 `excluded_heads` 는 «붙지 못해 후보에서 뺀 헤드 수» 다. 이 값이 크면
    그 도면의 배관이 끊겨 있다는 뜻이므로 **화면이 반드시 보여야 한다** —
    조용히 빼면 더 불리한 헤드를 못 본 채 수리계산이 나간다.
    """
    from services.cad_import.design.worst import REMOTE_K_DEFAULT, worst_k_heads

    k = REMOTE_K_DEFAULT if k is None else k
    probe = attachable_heads(payload, selected_source=selected_source, key=key)
    if not probe["ok"]:
        return {"ok": False, "error": probe.get("error")}

    wet = probe["wet"]
    cand = wet if only_heads is None else (set(only_heads) & wet)
    if not cand:
        return {"ok": False,
                "error": ("전개가 배관에 붙일 수 있는 헤드가 없습니다. "
                          "손질 단계에서 배관을 먼저 이어 주세요.")}

    b = board
    worst = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources,
                          k=k, only_heads=cand)
    if not worst.get("heads"):
        return {"ok": False, "error": "급수원에서 닿는 헤드가 없습니다."}

    got = expand_worst(payload, b, worst,
                       selected_source=selected_source, key=key)
    if not got.get("ok"):
        return got

    got["worst"] = worst
    got["excluded_heads"] = probe["dropped"]
    got["candidate_heads"] = len(cand)
    got["total_heads"] = probe["total"]
    if probe["dropped"]:
        print(f"[G2] 전개가 붙이지 못한 헤드 {probe['dropped']}개는 후보에서 "
              f"제외했습니다 — 그만큼 배관이 끊겨 있을 수 있습니다 "
              f"(후보 {len(cand)} / 도면 {probe['total']}).")
    return got
'''
io.open(p, "a", encoding="utf-8", newline="\n").write(tail)
print("select_and_expand() 추가")
