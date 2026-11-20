# 모듈 F — 세 도면 공통 노드의 좌표가 통일되지 않는 건 수정 지시서

작성 2026-09-08 · 대상 저장소 `PycharmProjects/JupyterProject`
증상: 통합 배관망에서 **평면도 알람밸브 ↔ 계통도 알람밸브**, **계통도 펌프 ↔ 기계실 입상관** 의 좌표가 일치하지 않아 선이 화면을 가로질러 꼬인다

> 선행 조건: `ModuleF_위상손실_수정지시서.md` 의 **D1 은 이미 반영되어 있다**(merge.py 에 `_riser_input_label` 확인). 이 문서는 그 뒤에 남은 건이다.

---

## 0. 먼저 확정된 것 — 굽는 식은 무죄다

`bake_combined_iso` 는 부위마다 다른 식을 쓴다.

    plan        : rot(x,y) + (0, (e − e_ref)·lift)      lift = 1000·zs
    system      : a_iso + (x − ax, y − ay)              회전 없음
    machineroom : rot(x,y) + shift                      shift = new_pj − rot(pj)

**두 이음매에서 두 식이 같은 답을 내는지 계산하면 항등적으로 0 이다.**

| 이음매 | 한쪽 식 | 다른쪽 식 | 차 |
|---|---|---|---|
| 라벨 `"10"` | plan → `rot(a) + (0,(e₁₀−e_ref)·lift)` = `a_iso` (e₁₀ ≡ e_ref) | system → `a_iso + (a−a)` = `a_iso` | **0** |
| `pump_junction` (`"1"`) | system → `a_iso + (pj−a)` = `new_pj` | machineroom → `rot(pj) + shift` = `rot(pj) + new_pj − rot(pj)` = `new_pj` | **0** |

**그러므로 좌표가 안 맞으면 굽기 전에 이미 안 맞는 것이다.** `bake_combined_iso` 의 식을 건드리지 말 것 — 고칠 곳은 `stitch_riser_and_heads` 와 그 뒤다.

---

## 1. E1 (주범) — 펌프 토출 노드가 「평면」으로 분류된다

### 무슨 일인가

`merge_network` 의 순서는 이렇다.

    combined = stitch_riser_and_heads(...)          ← 여기까지가 세 도면의 배치
    if is_pump and pump:
        combined = insert_source_pump(...)          ← ★노드를 하나 더 만든다
    ...
    parts = {"plan": [ht.nodes 라벨], "system": riser_labels, "machineroom": sorted(mr)}
                                                     ← ★추가 «전» 표로 만든다

`insert_source_pump`(core/remote30_full_network.py) 가 하는 일:

    disch = {"label": f"{src_label}_pd",
             "x": int(src.get("x", 0)) + 400,
             "y": int(src.get("y", 0)),
             "elevation": float(src.get("elevation", 0.0)), "io_node": "No"}
    combined.nodes.append(disch)
    for p in combined.pipes:
        if str(p.get("in")) == src_label:
            p["in"] = disch_label        # 수원에서 나가던 배관을 **전부** 옮긴다

펌프 가압 모드에서 수원(Input)은 기계실의 `m1` 이다. 따라서 새 노드는 `"m1_pd"` 이고 **기계실 소속**이다. 그런데 `parts` 세 목록 중 어디에도 없다.

`bake_combined_iso` 와 `api_merge` 의 미리보기가 둘 다 이렇게 읽는다:

    kind = of.get(lab, "plan")        # ★기본값이 plan

→ **기계실 노드가 평면 식으로 굽힌다.**

### 그래서 얼마나 틀어지나

평면 식은 `rot(x,y) + (0, (e − e_ref)·lift)` 이므로 두 가지가 동시에 어긋난다.

1. **`shift` 가 빠진다** — 기계실 군집 전체를 라이저 끝에 붙이는 평행이동이 이 노드에만 적용되지 않는다.
2. **표고 lift 가 붙는다** — 기계실 식에는 없는 항이다. 펌프 가압이면 수원 표고가 `라이저 최저 − source_drop` 이라 음의 큰 값이고, `lift = 1000·zs` 이므로 **1 m 당 1,000 단위**로 튕겨 나간다.

그리고 수원에서 나가던 배관이 **전부** 이 노드로 옮겨졌으므로 그 선들이 화면을 가로지른다. 사용자가 보는 「꼬임」이다.

### 부수 증상

- 미리보기에서 그 노드가 **평면 색**으로 칠해진다(`rec["part"]`)
- 그 노드와 기계실 노드를 잇는 배관이 `ka != kb` 라 **거짓 seam** 으로 표시된다
- 즉 화면이 「여기가 이음매」라고 잘못 가리킨다

### 발생 조건

`is_pump and pump` — 즉 **펌프 가압 + 펌프 제원 입력** 일 때만. 자연낙차·감압 모드에서는 안 생긴다. 기계실이 있는 도면에서만 보이는 이유가 이것이다.

### 조치 E1

**⑴ `insert_source_pump` 가 만든 라벨을 원래 노드와 같은 부위에 넣는다.**

`insert_source_pump` 가 새 라벨을 돌려주게 하거나(시그니처 변경이 부담이면 호출 뒤 `combined.nodes[-1]["label"]` 로 잡아도 된다), `merge_network` 에서 `parts` 를 만들 때 **수원 라벨이 속한 목록에 같이 넣는다.**

    · 수원 라벨(src_label)이 machineroom 목록에 있으면 → 새 라벨도 machineroom
    · system 에 있으면 → system
    · 그 밖 → plan (지금과 같음)

**⑵ 기본값 «plan» 을 없앤다 — 이것이 구조적 조치다.**

`of.get(lab, "plan")` 은 **`parts` 에 없는 노드를 조용히 평면으로 만든다.** `_pd` 는 그 사례 하나일 뿐이고, 앞으로 `stitch` 뒤에 노드를 추가하는 코드가 생기면 같은 일이 반복된다.

    · `bake_combined_iso` 와 `api_merge` 미리보기 양쪽에서, `parts` 에 없는 라벨을
      **세어서 보고**한다 — `unclassified: [라벨…]`
    · 굽기는 종전대로 plan 식으로 하되(동작 호환), **0 이 아니면 화면과 로그가 말한다**
    · `check_combined` 에도 같은 값을 싣는다

두 곳 다 고칠 것. 한 곳만 고치면 화면과 파일이 갈린다.

### 수용 기준 E1

| | 기준 |
|---|---|
| 펌프 가압 + 기계실 도면에서 | `unclassified` 가 **0** |
| 같은 도면에서 | 펌프 토출 노드가 미리보기에서 **기계실 색**으로 나온다 |
| 같은 도면에서 | 그 노드에 붙은 배관이 더 이상 `seam` 으로 표시되지 않는다 |
| 자연낙차 모드 | 산출·화면 **바이트 불변** (`insert_source_pump` 가 안 도니까) |

---

## 2. E2 — `stitch_riser_and_heads` 의 조용한 좌표 폴백 둘

배치가 실패하면 **그 부위가 원 DXF 좌표로 남는데 아무 기록이 없다.**

    # ① 계통도
    if head_av_node is not None and head_xs and head_ys:
        try:
            translated_riser_nodes = _layout_riser_as_schematic(...)
        except (KeyError, TypeError, ValueError):
            translated_riser_nodes = list(true_riser_nodes)   # ← raw DXF (수십 m)
    else:
        translated_riser_nodes = list(true_riser_nodes)       # ← raw DXF

    # ② 기계실
    if mr_nodes and pump_junction_label is not None:
        try:
            mr_rel, plan_rel = _layout_machine_room_plan(...)
        except (KeyError, TypeError, ValueError):
            mr_rel, plan_rel = [], []                         # ← 뒤에서 raw 로 남는다

①이 걸리면 **계통도 전체가** 원 좌표에 남고, 평면망의 `"10"` 은 riser 쪽 사본이 살아남으므로 **평면 배관이 통째로 계통도 좌표까지 늘어난다.** 정확히 「평면 AV ↔ 계통 AV 가 안 맞는다」로 보인다.

②가 걸리면 기계실이 원 좌표에 남는다(D1 과 같은 그림).

### 조치 E2

**폴백을 없애지 말 것.** 예외를 그대로 올리면 결합이 통째로 실패한다. 대신 **사실을 기록한다.**

    · 각 폴백 자리에서 무엇이 왜 안 됐는지 담는다
        {"riser_layout": "skipped" | "raised:<예외명>" | "ok",
         "mr_layout":    "skipped" | "raised:<예외명>" | "ok"}
    · `CombinedTables` 에 실어 `merge_network` 가 `steps` 에 한 줄 남긴다
        「★계통도 배치 실패 — 원 도면 좌표로 남았습니다(<이유>)」
    · `check_combined` 의 반환에도 싣는다

`else` 가지(`head_av_node is None`)는 **왜 None 인지**를 함께 남긴다 — `av_lbl` 값과 평면 표에 있는 라벨 앞 8개.

---

## 3. E3 — `check_combined` 에 좌표 검사가 없다

지금 보는 것: 연결성분 · 고아 참조 · Input 개수. **좌표는 하나도 안 본다.**

### 조치 E3 — 이음매 좌표 검사를 넣는다 (보고만, 예외 아님)

`check_combined` 에 아래를 추가한다. **굽기 전 좌표**(`combined.nodes`)에서 잰다.

| 검사 | 내용 |
|---|---|
| `anchor_gap_mm` | 라벨 `"10"` 이 combined 에 몇 개인가 · riser 쪽과 plan 쪽 원좌표의 거리(mm). **원래 0 이어야 한다** |
| `pump_seam_mm` | `pump_junction` 노드와, 그 노드에 붙은 기계실 배관의 반대쪽 노드 사이 거리(mm) · 그 배관의 표 length(m). **좌표 거리 ÷ 1000 ≈ 표 length 여야 한다** |
| `bbox_span_mm` | combined 전체 x·y 폭. **평면망 폭의 몇 배인지** 함께 낸다 — 어느 부위가 raw 로 남으면 여기서 폭발한다 |
| `part_bbox` | plan / system / machineroom 각각의 bbox. 셋이 서로 몇 배 차이인지 |
| `unclassified` | `parts` 어디에도 없는 라벨 (E1) |
| `layout_status` | E2 의 두 값 |

**판정은 하지 않는다.** 값만 낸다. 화면이 읽고 사람이 본다.

### 굽은 뒤 검사도 하나 넣는다

`bake_combined_iso` 가 굽고 나서, **이음매 양쪽을 각 부위 식으로 따로 굽었을 때 같은 점인지** 확인한다(§0 이 항등적으로 0 임을 보였으므로, 0 이 아니면 식이 바뀐 것이다 — 회귀 감지기다).

    seam_check = {"anchor": |plan식(10) − system식(10)|,
                  "pump":   |system식(pj) − mr식(pj)|}

둘 다 **1e-6 미만**이어야 한다. 아니면 로그에 남긴다.

---

## 4. 계측 스크립트

`scripts/_merge_seam_probe.py` — 결합 세션 하나를 받아 위 값을 전부 표로 낸다.

    · merge_network 결과(got)를 받아
    ·   parts 세 목록의 크기 · unclassified 목록
    ·   부위별 bbox 와 서로의 배율
    ·   anchor_gap_mm · pump_seam_mm (표 length 와 대조)
    ·   layout_status
    · bake_combined_iso 를 돌린 뒤 seam_check 두 값
    · 마지막으로 combined 배관 중 **화면 거리 ÷ 1000 이 표 length 와 5 % 넘게 다른** 배관 목록
      ← 좌표가 어긋난 배관은 여기서 전부 잡힌다

**대명동 3장(평면 + 계통 + 기계실 · 펌프 가압) 으로 조치 전/후를 나란히 낸다.**

---

## 5. 금지 사항

- `bake_combined_iso` 의 세 식을 건드리지 말 것 — §0 에서 이음매 연속성이 증명됐다
- `_layout_riser_as_schematic` · `_layout_machine_room_plan` 의 배치 규칙을 바꾸지 말 것
- 폴백을 예외로 승격하지 말 것 (결합이 통째로 실패한다). **기록만 추가**
- `insert_source_pump` 의 물리 모델(수원 → 펌프 → 토출노드)을 바꾸지 말 것. 고칠 것은 **그 노드가 어느 부위인지**뿐이다
- 좌표가 안 맞는다고 노드를 옮겨 맞추지 말 것 — 원인을 덮는다
- 골든 재생성 금지
- 리팩터링·파일 이동 금지

## 6. 보고 형식

    [계측표]  조치 전 / E1 후 / E2·E3 후

    [E1] unclassified __ → 0 | 펌프 토출 노드 부위 plan → machineroom
         거짓 seam __건 → 0 | 자연낙차 모드 산출 해시 동일 여부
    [E2] layout_status = {riser_layout: __, mr_layout: __}
         폴백이 걸린 적이 있으면 그 이유와 그때의 bbox
    [E3] anchor_gap_mm __ (기대 0) | pump_seam_mm __ vs 표 length __ m
         bbox_span_mm __ | 부위별 배율 plan 1 : system __ : machineroom __
         seam_check anchor __ · pump __ (기대 <1e-6)
    [좌표↔length 불일치 배관] __건 (5 % 초과)

    [막힌 것] 못 고친 것 · 이유 · 필요한 판단
