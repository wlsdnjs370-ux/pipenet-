# -*- coding: utf-8 -*-
"""[G2 · B4 1안] 최불리를 «전개가 붙일 수 있는 헤드» 안에서 고른다.

선정(board 도달)과 전개(헤드 중심이 물길 안)가 «물 닿음» 을 다르게 보므로,
최불리 K 가 통째로 전개 밖으로 떨어질 수 있다(B1F 실측: 30개 중 0개 생존).

이식한 `worst_k_heads` 는 손대지 않는다 — 이미 있는 `only_heads` 인자에
«전개가 인정한 헤드» 를 넘겨 후보를 좁힐 뿐이다. 그래서 G1 의 「모듈 F 와 완전
일치」는 그대로 유지된다(그 검사는 only_heads 없이 부른다).

★이것은 증상 완화다. 붙지 못한 헤드가 있다는 것은 그 도면의 배관이 끊겨 있다는
뜻이므로, 몇 개를 제외했는지 반드시 밖으로 드러낸다 — 조용히 빼면 「더 불리한
헤드가 있는데 못 본 채로」 수리계산이 나간다.
"""
import io

p = "cad_project_editor_g/services/cad_import/design/restrict.py"
s = io.open(p, encoding="utf-8").read()

old_sig = '''def expand_worst(payload: dict, board, worst: dict, *,
                 selected_source=None, key: str | None = None) -> dict:'''
new_sig = '''def attachable_heads(payload: dict, *, selected_source=None,
                     key: str | None = None) -> dict:
    """전개가 **배관에 붙일 수 있는** 헤드 번호. 전체망 전개를 한 번 돌려 얻는다.

    선정은 board 그래프 도달로 헤드를 세지만, 전개는 «헤드 중심 노드가 물길
    필터 안» 이라야 인정한다. 후자가 더 엄격해 B1F 실측으로 868 대 619 다.
    이 차이를 모른 채 최불리를 고르면 30개가 통째로 전개 밖으로 떨어진다.

    반환: {"ok", "wet": set(hcov 번호), "total": 도면 헤드 수, "dropped": 제외 수}
    """
    from services.cad_import.convert.planar import build_planar_graph

    built = build_planar_graph(
        key or payload.get("key") or "probe",
        write=False,
        selected_source=selected_source or payload.get("selected_source"),
        pts=payload.get("pts"),
        edges=payload.get("edges"),
        hcov=payload.get("hcov"),
        ups=payload.get("ups"),
        head_kinds=payload.get("head_kinds"),
        user_sources=payload.get("sources"),
        ho=payload.get("ho"),
    )
    if not built.get("ok"):
        return {"ok": False,
                "error": built.get("error") or "전개가 실패했습니다.",
                "wet": set(), "total": 0, "dropped": 0}
    wet = set(built.get("wet_head_idx") or [])
    total = len(payload.get("hcov") or [])
    return {"ok": True, "wet": wet, "total": total,
            "dropped": max(0, total - len(wet))}


def expand_worst(payload: dict, board, worst: dict, *,
                 selected_source=None, key: str | None = None) -> dict:'''
assert old_sig in s, "expand_worst 시그니처를 못 찾음"
s = s.replace(old_sig, new_sig, 1)
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("attachable_heads() 추가")
