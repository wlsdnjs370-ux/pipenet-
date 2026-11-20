# 모듈 F 위상 손실 수정 지시서

작성 2026-09-08 · 대상 저장소 `PycharmProjects/JupyterProject`

---

## 0. 이 작업의 범위

평면도·계통도·기계실 **추출 세 개는 정상이다.** 깨지는 곳은 두 경계다.

1. 평면도 → 수리계산 입력 변환
2. 평면도·계통도·기계실 통합 (S700)

아래 **다섯 건만** 손댄다. 그 밖의 어떤 것도 바꾸지 않는다.

### 절대 하지 말 것

- **골든 재생성 금지.** `combined_build__plane_daemyeong` 등은 BLOCKED.md §2·§7·§8·§10·§13 으로 이미 보류 중이다. 불일치가 나도 재생성하지 말고 보고만 한다.
- **루프·트리 정책을 건드리지 말 것.** `on_residual_cycle="force_tree"`, `ortho=True`, union-find 게이트, `planarize_edges` 는 전부 **사용자 지시로 채택된 설계**다(BLOCKED §13). 위상 손실처럼 보여도 결함이 아니다.
- **리팩터링·파일 이동·이름 변경 금지.** 진단한 자리만 최소로 고친다.
- **추측으로 메우지 말 것.** 이 저장소의 S340 원칙이다 — 임의로 채우지 않고 «못 했다» 고 올린다.
- **산출 파일 형식(SDF/KFP 스키마)을 바꾸지 말 것.**

### 라인 번호에 의존하지 말 것

아래의 줄 번호는 참고다. **반드시 코드 조각(앵커)으로 찾아서** 고친다.

---

## 1. 작업 순서 — 계측이 먼저다

```
0단계  계측 (코드 수정 없음, 스크립트만)
1단계  D2 · D3 을 «세기» 로 켠다        ← 예외 아님. 숫자만 낸다
2단계  다시 계측 → 실제 손상 건수 확정
3단계  D1 수정                          ← 가장 큰 건
4단계  다시 계측 → D1 이 얼마를 없앴는지 확인
5단계  D4 수정 (계측에서 개명 0 이어도 예방으로 넣는다)
6단계  D5 결합 후 검사 추가 (보고만, 예외 아님)
7단계  최종 계측 + 보고
```

**1단계를 3단계보다 먼저** 하는 이유: D1 을 먼저 고치면 D2·D3 가 잡았어야 할 건수가 사라져, 남은 손상이 얼마인지 영영 알 수 없다.

---

## 2. 계측 스크립트 (0 · 2 · 4 · 7 단계에서 같은 것을 돌린다)

`scripts/_f_topology_probe.py` 로 새로 만든다. 대상 도면 두 장:

- `26F 대명동` — 알려진 정상 사례(기준선)
- 기계실이 붙는 도면 한 장 — 사용자에게 어느 것인지 확인받는다

`merge_network()` 를 부른 뒤 `combined` 에서 아래를 **전부** 찍는다.

| 지표 | 왜 보는가 |
|---|---|
| `len(nodes)` · `len(pipes)` · `len(nozzles)` · `len(fittings)` · `len(equipment)` | 기본 규모 |
| **좌표 bbox** `min/max x, y` 와 그 span | **D1 의 결정적 지표.** 기계실이 원 DXF 좌표에 남으면 span 이 수십 m 로 뛴다 |
| `emit_sdf` 의 `_scale` (= `3000 / _longest`) | span 이 커진 만큼 작아진다 — 망이 한 점으로 뭉치는 정도 |
| **연결성분 수** (nodes+pipes 로 무향 그래프) | 1 이 아니면 통합 실패 |
| `pipes` 중 `in` 또는 `out` 이 `"?"` 인 행 수 | D2 가 통과시키는 것 |
| `nodes` 중 `x==0 and y==0` 인 행 수 | D3 가 만드는 것 |
| `stitch` 에서 개명된 배관 수 (`label_2` 형태) | D4 의 조건 성립 여부 |
| 개명 때문에 고아가 된 `fittings`/`equipment` 행 수 | D4 의 실제 피해 |
| `riser_av.x/y` 와 `head_av.x/y` 의 **거리(mm)** 와 `elevation` 차 | 두 "10" 이 정말 같은 점인지 |
| `machine_room_labels` 중 `combined.nodes` 에서 좌표가 원본과 같은 것의 수 | D1 이 발생했는지의 직접 증거 |

출력은 도면별 한 줄 표로, **매 단계 같은 형식**으로 낸다. 단계 간 비교가 이 작업의 근거다.

---

## 3. 결함별 조치

### D1 — 기계실이 통합 좌표계로 옮겨지지 않는다 (항상)

**파일** `routes/module_f/merge.py` (주 수정) · `core/remote30_full_network.py` (확인용)

**앵커 — `core/remote30_full_network.py`, `stitch_riser_and_heads` 안 (≈1352·1421)**

```python
mr_set = {str(l) for l in (machine_room_labels or [])}
true_riser_nodes = [n for n in riser.nodes if str(n["label"]) not in mr_set]
...
mr_laid = list(mr_nodes)
if mr_rel:
    pump_node = next((n for n in translated_riser_nodes
                      if str(n["label"]) == str(pump_junction_label)), None)
```

**앵커 — `routes/module_f/merge.py`, `merge_network` 안 (≈309)**

```python
combined = stitch_riser_and_heads(
    rt, ht,
    machine_room_labels=mr_labels or None,
    pump_junction_label=(str(rt.nodes[0].get("label"))
                         if (attached and rt.nodes) else None),
```

**원인**

`prepend_machine_room_to_riser` 는 `RiserTables(nodes=new_mr_nodes + new_riser_nodes, ...)` 로 돌려준다. 그래서 `rt.nodes[0]` 은 기계실 수원 `"m1"` 이다. 그런데 `stitch` 는 그 라벨을 `translated_riser_nodes` 에서 찾고, 그 목록은 `mr_set`(기계실 라벨 전체)을 이미 제외했다. → `pump_node` 는 **항상 `None`** → `mr_laid` 가 원본 그대로 남아 기계실 노드가 **원 DXF 좌표에 방치**된다. `plan_laid` 도 빈 채다.

`stitch` 의 주석은 `펌프 junction("1")` 을 기대한다. 즉 넘겨야 할 것은 라이저의 Input 노드이지 기계실 수원이 아니다.

**조치 — `merge.py` 에서, `prepend` 를 부르기 전에 라이저 Input 라벨을 잡아 둔다**

```python
# ★prepend 뒤의 rt.nodes[0] 은 기계실 수원(m1)이다. stitch 가 찾는 것은
#   기계실 평면이 실제로 붙는 자리 = 라이저의 Input 노드다.
#   선택 규칙을 prepend_machine_room_to_riser 와 **같게** 둔다 — 두 곳이
#   다른 노드를 고르면 평면이 엉뚱한 데 붙는다.
_riser_input_label = next(
    (str(n.get("label")) for n in rt.nodes
     if str(n.get("io_node", "")).lower() == "input"), None)
if _riser_input_label is None:
    _riser_input_label = next(
        (str(n.get("label")) for n in rt.nodes
         if str(n.get("label")) == "1"), None)
```

그리고 `stitch_riser_and_heads(...)` 호출과 반환 dict 의 `"pump_junction"` 두 자리 모두에서
`str(rt.nodes[0].get("label"))` → `_riser_input_label` 로 바꾼다.

**금지** — `stitch_riser_and_heads` 쪽 탐색 범위를 넓히는 우회는 쓰지 않는다. `mr_set` 제외는 «기계실은 막대로 뭉개지 않는다» 는 의도적 분리이고, 거기에 기계실 라벨을 도로 넣으면 그 의도가 깨진다.

**수용 기준**

- 통합망 좌표 bbox span 이 **평면도 bbox 수준**으로 돌아온다 (계측표에서 확인)
- `plan_laid` 가 비어 있지 않다 (기계실 평면 형상이 렌더된다)
- 기계실 노드 중 좌표가 원본 그대로인 것이 **0**
- 대명동(기계실 없는 경우)의 산출은 **바이트 단위로 불변**

---

### D2 — 고아 참조 검사가 고아의 표식을 예외로 둔다

**파일** `routes/module_f/merge.py`, `_check_anchor` (≈190)

**앵커**

```python
v = str(r.get(k))
if v not in labels and v not in ("?", "None"):
    raise MergeError(...)
```

**원인**

표를 만들 때 라벨을 못 찾은 배관 끝점은 `label_of.get(a, "?")` 로 `"?"` 가 된다
(`cad_project_editor_g/services/cad_import/design/tables.py`, `pipe_row`).
그 `"?"` 를 고아 검사가 **명시적으로 면제**한다. `_shift("?")` 도 숫자가 아니라 그대로 통과한다.
→ 끝점 없는 배관이 검사 셋을 전부 지나 SDF 까지 간다.

**조치 (1단계 · 세기)**

`_check_anchor` 를 예외로 바꾸지 말고, **집계해서 돌려준다.**

```python
def _check_anchor(ht: HeadTables) -> dict:
    ...
    dangling = []      # (표이름, 행라벨, 칸, 값)
    for name, rows, keys in (...):
        for r in rows:
            for k in keys:
                v = str(r.get(k))
                if v in ("?", "None"):
                    dangling.append((name, r.get("label") or r.get("pipe"), k, v))
                elif v not in labels:
                    raise MergeError(...)      # 기존 동작 유지
    return {"dangling": dangling}
```

`to_head_tables` 가 이 결과를 받아 `merge_network` 의 `steps` 에 **«끝점 없는 배관 n건»** 으로 싣는다. 0 건이면 그 줄을 넣지 않는다.

**조치 (7단계 이후 · 사용자 판단)**

건수를 보고한 뒤, 예외로 승격할지는 **사용자가 정한다.** 임의로 raise 로 바꾸지 않는다 — 지금 통과하던 실행이 통째로 실패하게 되는 변경이다.

**수용 기준** — 산출 파일은 한 바이트도 바뀌지 않는다. 화면·로그에 건수만 는다.

---

### D3 — 배관표에 있는데 메타에 없는 노드가 원점 (0,0) 에 생긴다

**파일** `cad_project_editor_g/services/cad_import/design/tables.py`, `build_design_tables` (≈169–213)

**앵커**

```python
def xy(nid):
    c = (meta_nodes.get(nid) or {}).get("coords") or (0.0, 0.0, 0.0)
    return float(c[0]), float(c[1])

def z(nid):
    c = (meta_nodes.get(nid) or {}).get("coords") or (0.0, 0.0, 0.0)
    return float(c[2]) if len(c) > 2 else 0.0
```

**원인**

`bfs_order` 는 `pipe_data` 의 끝점으로 순서를 만들고 노드표는 그 순서를 그대로 쓴다.
그런데 좌표는 `nodes_meta_runtime` 에서 읽으며, 없으면 조용히 `(0,0,0)` 이 된다.
→ 배관에는 있는데 메타에 없는 노드가 **표고 0 · 좌표 원점**의 노드가 되어 실제 배관에 연결된다.

**조치 (1단계 · 세기)**

폴백을 즉시 예외로 바꾸지 말고 **모아서 한 번에 보고**한다. 첫 건에서 죽으면 전체 규모를 못 본다.

```python
_missing_meta: set = set()

def xy(nid):
    m = meta_nodes.get(nid)
    if m is None or not m.get("coords"):
        _missing_meta.add(nid)
        return 0.0, 0.0
    ...
```

표를 다 만든 뒤 `_missing_meta` 가 비어 있지 않으면 `tbl.meta` 에
`{"missing_node_meta": sorted(_missing_meta)[:20], "missing_node_meta_count": len(_missing_meta)}`
를 싣고 `print` 로도 한 줄 낸다.

**조치 (7단계 이후 · 사용자 판단)** — 예외 승격 여부는 D2 와 함께 사용자가 정한다.

**수용 기준** — 산출 불변. 대명동에서 `missing_node_meta_count` 는 **0 이어야 한다.** 0 이 아니면 그 자체가 새 발견이니 즉시 보고하고 멈춘다.

---

### D4 — 배관 개명이 부속·기기표를 데려가지 않는다

**파일** `core/remote30_full_network.py`, `stitch_riser_and_heads`

**앵커**

```python
combined_pipes.append({**p, "label": new_lbl})
...
combined_fittings = list(head_tables.fittings)
```

**원인**

결합 시 배관 라벨이 겹치면 두 번째부터 `label_2` 로 개명한다(주석: 같은 라벨 두 배관을 한 pid 로 접으면 KFP 토폴로지가 붕괴). 그런데 `fittings`·`equipment` 는 옛 라벨 그대로다.
`emit_sdf` 는 `fittings_by_pipe[str(f["pipe"])]` 로 색인하므로 **개명된 배관의 부속·기기는 어느 배관에도 안 붙고 사라진다.** 고아 검사는 결합 **전**에만 돈다.

**조치**

개명 루프에서 `{옛 라벨: 새 라벨}` 을 만들고, 부속·기기에 같이 먹인다.

```python
renamed: dict[str, str] = {}
...
    renamed[lbl] = new_lbl        # 개명한 자리에서 기록
    combined_pipes.append({**p, "label": new_lbl})
...
combined_fittings = [
    ({**f, "pipe": renamed[str(f["pipe"])]} if str(f.get("pipe")) in renamed else f)
    for f in head_tables.fittings
]
```

`equipment` 도 같은 방식으로 옮긴다 (`head_tables.equipment` 를 쓰는 자리 전부).

**주의** — 개명은 «두 번째 이후» 에만 일어난다. 라이저 배관(`r*`)이 먼저 들어가므로, 겹치면 개명되는 쪽은 **평면도 배관**이다. 즉 `renamed` 의 키는 평면도 pid 다.

**수용 기준**

- 계측에서 «개명 때문에 고아가 된 부속/기기» 가 **0**
- 개명이 애초에 0 건인 도면에서는 산출이 **바이트 불변**

---

### D5 — 결합 후 검사가 하나도 없다 (신규)

**파일** `routes/module_f/merge.py` — 새 함수 `check_combined(got) -> dict`

현재 `api_merge.py` 는 `check_supply_mode` 만 부른다. 결합망이 성립했는지 확인하는 코드가 없다.

**만들 검사 (전부 «보고», 예외 아님)**

| 검사 | 판정 |
|---|---|
| 연결성분 수 | 1 이 아니면 `components: n` 과 각 성분 크기 |
| 모든 `pipe.in/out` 이 노드표에 있는가 | 없는 것의 목록(최대 20) |
| 모든 `fitting.pipe` · `equipment.pipe` 가 배관표에 있는가 | 없는 것의 목록 |
| `io_node == "Input"` 노드 개수 | 정확히 1 이어야 한다 |
| 두 AV(`riser_av`, `head_av`) 좌표 거리 mm · `elevation` 차 m | 값 그대로 보고 |
| 좌표 bbox span x·y (mm) | 값 그대로 보고 |

`merge_network` 의 반환에 `"checks"` 로 싣고, `combined_summary` 가 그대로 통과시킨다.
`api_merge.py` 응답에도 실어 화면이 볼 수 있게 한다.

**두 AV 좌표 거리**는 `stitch_riser_and_heads` 안에서만 알 수 있으므로, 거기서 계산해
`CombinedTables.meta` 나 반환 dict 에 담아 올린다 — **`stitch` 의 기존 동작은 바꾸지 않는다.**

**수용 기준** — 산출 파일 불변. 응답에 `checks` 키가 는다.

---

## 4. 최종 보고 형식

```
[계측표]  0단계 / 2단계 / 4단계 / 7단계 를 같은 표에 나란히

[D1] 조치: …  |  bbox span 전 ____mm → 후 ____mm  |  기계실 미변환 노드 __ → 0
[D2] 끝점 없는 배관: 대명동 __건 · <기계실도면> __건
[D3] missing_node_meta: 대명동 __건 · <기계실도면> __건
[D4] 개명 __건 · 고아 부속 __ → 0
[D5] 연결성분 __개 · Input 노드 __개 · 두 AV 거리 __mm · 표고차 __m

[산출 불변 확인]  대명동 SDF/KFP 해시 전후 비교 — D1 은 기계실 없는 도면에서 불변이어야 함
[골든]            불일치 목록만. 재생성하지 않음
[막힌 것]         고치지 못한 것 · 그 이유 · 필요한 판단
```

**끝까지 못 간 항목이 있으면 그 자리에서 멈추고 보고한다.** 우회로 메우지 않는다.
