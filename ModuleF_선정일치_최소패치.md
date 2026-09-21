# 최소 패치 — 두 화면이 같은 헤드를 그리게 (이것만 한다)

작성 2026-09-10 · `ModuleF_두화면_선정일치_지시서.md` 의 **§2-1 만** 떼어낸 것

---

## 범위

**고치는 파일은 `routes/module_f/api_design.py` 한 개다.**
다른 파일은 열지 않는다. 지시서의 §2-2 · §2-3 · §2-4 · §2-5 는 **이번에 하지 않는다.**

---

## 왜 이것만으로 되나

평면 보기는 `sess["worst"]`(손질 선정)를 그리고, 표는 `got["worst"]`(백필된 선정)로
만들어진다. `/design/build` 가 `sess["worst"]` 를 안 고치기 때문에 두 화면이 갈린다.
**실제로 쓰인 선정을 세션에 돌려놓으면** 평면 보기가 그것을 그린다 — 화면 코드도,
`_worst_view` 도, JS 도 손댈 필요가 없다.

---

## 고칠 자리

`routes/module_f/api_design.py` 의 `module_f_design_build` 안 `job()`.
`got` 이 성공한 뒤, 아래 두 줄 **바로 다음**이다.

    got["handoff"] = _worst_handoff_note(got, picked, k_use, int(cfg["k"]))
    got["_picked"] = picked      # 표를 보고 다시 셀 때 쓴다(§2-3)

## 붙일 코드

    # ★[두 화면 일치] 실제로 계산에 들어간 선정을 세션에 돌려놓는다.
    #
    #   평면 보기(`drawEdit`)는 `sess["worst"]` 를, 표는 `got["worst"]` 를
    #   그린다. 백필(위 `filled`)이 일어나면 두 집합이 갈라지는데 종전에는
    #   그것을 되돌려 놓는 자리가 없어, 같은 도면에서 두 화면이 다른 헤드를
    #   그렸다(실측: 30개 중 27개만 겹침 · 4개 교체).
    w_fin = got.get("worst") or {}
    if w_fin.get("heads") is not None:
        prev = dict(sess.get("worst") or {})
        merged = dict(w_fin)
        # ★손질만 아는 칸은 그대로 물려준다. `worst_k_heads` 는 이 다섯을
        #   내지 않으므로(`api_edit.py` 가 나중에 얹는다) 안 옮기면 화면의
        #   영역 상자·급수원 이름·도면 장이 사라진다.
        for _k in ("zones", "sheet", "source_tag", "source_index", "candidates"):
            if _k in prev:
                merged[_k] = prev[_k]
        if prev and set(prev.get("heads") or ()) != set(merged.get("heads") or ()):
            sess["worst_edit"] = prev     # 손질 원본은 버리지 않는다
        sess["worst"] = merged
        # `views.py` 의 지문 비교를 무효화한다 — 안 그러면 화면이 옛 corridor 를
        # 그대로 들고 있는다(`views.py` 의 `worst_rev`).
        sess["worst_rev"] = None

## 건드리지 않는 것

- `sess["worst_cand"]` — 다음 백필의 후보 범위다. 손질 것 그대로 둔다.
- `sess["worst_k"]` · `sess["worst_zones"]` — 손질 것 그대로 둔다.
- `_worst_view` · `views.py` · `remote30.py` · `static/module_f.js` — **열지 않는다.**
- 백필 로직(`api_design.py` 의 `filled` 블록) — **그대로 둔다.**
- `_worst_handoff_note` · `_handoff_after_table` 의 문구 — 이번에는 그대로 둔다.

---

## 확인

1. 같은 도면에서 영역 지정 → 최불리 → 표 확정.
2. 「평면에서 보기」를 켰다 껐다 하며 헤드가 하나하나 겹치는지 본다.
3. 빨간 점선(최원 경로)이 두 화면에서 같은 줄인지 본다.

수치로는:

    len(sess["worst"]["heads"]) == 표 노즐 수
    set(sess["worst"]["heads"]) 의 disk 좌표 ↔ 표 노즐 좌표가 1:1 (6~19 mm 안)

## 수용 기준

| | 기준 |
|---|---|
| 1 | 표 확정 뒤 두 화면의 헤드 집합이 좌표로 1:1 |
| 2 | 두 화면의 최원 경로가 같은 줄 |
| 3 | 백필이 없던 세션 — 산출·화면 **불변** |
| 4 | 화면에 영역 상자·급수원 이름이 그대로 뜬다 (다섯 칸 이관 확인) |
| 5 | 손질에서 「최불리 선정」을 다시 누르면 그것이 최신이 된다 |
| 6 | 기존 시험 전부 통과 · 골든 재생성 없음 |

**이 여섯이 되면 멈춘다. 더 고치지 않는다.**
