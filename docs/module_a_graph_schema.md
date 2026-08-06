# 모듈 A 배관망 그래프 스키마 — 모듈 C 개조판이 지켜야 할 계약

- 목적: 「모듈 C 설계 자동화 개조 지시서 v2」 §0.1 의 단 하나의 설계 기준을 실측으로 고정한다.
  > 모듈 A는 "완성된 소방도면 → 배관망 그래프" **역방향**,
  > 모듈 C 개조판은 "건축도면 → 배관망 그래프" **순방향**,
  > **두 파이프라인의 산출 그래프는 동일한 자료구조여야 한다.**
- 이 문서가 승인되기 전에는 C5(라우팅, PR-9)를 구현하지 않는다 (지시서 §0.1).
- 조사 기준 커밋: `8b0bac5` (main). 라인 번호는 이 커밋 기준이다.

---

## 0. 지시서 §0.2 확인 명령 결과

| # | 명령 | 결과 | 판정 |
|---|---|---|---|
| 1 | `grep -n "elevation_m\|display_z_m\|edge_len\|_edge_load" core/remote30_graph.py remote30_prototype.py` | `core/remote30_graph.py` 4건 — **전부 `edge_len`** (`_dijkstra_from:93,103`, `_shortest_path:110,134`). `remote30_prototype.py` — `edge_len` 163건, `_edge_load` 5건, `elevation_m` 1건, **`display_z_m` 0건** | **[문서정합] 대상 어긋남.** `elevation_m`/`display_z_m` 의 권위적 정의는 두 파일 어디에도 없고 `kfp_sdf_converter.py` 에 있다 (§2 참조). 지시서 §0.2 의 grep 대상 파일을 `kfp_sdf_converter.py` 로 정정해야 한다 |
| 2 | `grep -n "compute_edge_load\|_downstream_heads\|prune_by_load" -A 20 remote30_prototype.py` | `compute_edge_load:3284`, `prune_by_load:3358`, `_downstream_heads:5845` | 존재. §3 에 규약 정리 |
| 3 | `grep -n "Tee\|Elbow\|fitting" core/*.py` | 매핑표 `core/has_converter.py:84-117`. 실제 **생성**은 `core/*.py` 가 아니라 `remote30_prototype.py:5922-5956` | **`core/*.py` 만 보면 생성 규칙을 못 찾는다.** §4 에 정리 |
| 4 | `grep -n "emit_full_sdf\|_emit_kfp\|_emit_has\|round_trip_check" routes/r30_combined.py` | `584,588,597` (SDF) · `610` (KFP) · `618` (HAS) · `772` · `1303`. **`round_trip_check` 0건** | 방출 체인 존재. 왕복 검증은 이 라우트에서 호출되지 않는다 (`kfp_sdf_converter.py:2528`, `core/has_converter.py:877` 에 함수는 있음). §5 참조 |
| 5 | `sed -n '1,200p' docs/load_extraction.md` | 8개 절. §4 에 비성립 조건 4행 | 존재. §3.4 에 전재 |

---

## 1. 두 겹의 자료구조 — 어느 것이 계약인가

모듈 A 는 그래프를 **두 단계**로 표현한다. 모듈 C 가 맞춰야 할 것은 앞의 것이다.

```
DXF ─(모듈 A: 인식/추출)→ PipeTables ─(stitch)→ CombinedTables ─→ emit_full_sdf ─→ SDF
                                                       │
                                         parse_sdf ────┘
                                                       ↓
                                                 CommonNetwork ─→ emit_kfp / emit_has
```

- **`CombinedTables`** (`core/remote30_full_network.py:635`) = 모듈 A 파이프라인의 **최종 산출물**이자 `emit_full_sdf` 의 유일한 입력. **모듈 C 개조판이 생성해야 할 자료구조는 이것이다.**
- **`CommonNetwork` / `CommonNode` / `CommonPipe`** (`kfp_sdf_converter.py:129/81/107`) = SDF 를 다시 읽어 KFP·HAS 로 넘기는 **하류 중립 표현**. 모듈 C 는 이것을 직접 만들지 않는다. 단, `elevation_m` ↔ `display_z_m` 분리 규약(§2)은 여기서 강제되므로 반드시 이해해야 한다.

> **결론:** 모듈 C 의 C570(그래프 방출)은 `CombinedTables` 를 **필드 단위로 동일하게** 채워야 한다. `emit_full_sdf` 는 각 리스트의 dict 를 거의 그대로 통과시키므로, 키 하나만 달라도 SDF 가 조용히 비거나 "Unset" 이 된다.

---

## 2. 노드/엣지 필드 표

### 2.1 `CombinedTables` — 모듈 C 가 채워야 할 것 (`core/remote30_full_network.py:635-647`)

| 필드 | 자료형 | 내용 | 라인 |
|---|---|---|---|
| `nodes` | `list[dict]` | `label`, `elevation`(**m**), `io_node`(`"Input"`/`"No"`), `x`,`y`(**mm**), 선택 `pressure_pa` | 638 |
| `pipes` | `list[dict]` | `label`, `in`, `out`(노드 label), `type`(`"KSD 3507"`), `dia`(**호칭경 mm**), `length`(**m**), `elev`(낙차 **m**), `c`(HW 계수), `status`, `group` | 639 |
| `nozzles` | `list[dict]` | `label`, `in`(노드), `out`(`@`/idx), `lib`, `flow_lmin`, `flow_m3s` | 640 |
| `fittings` | `list[dict]` | `pipe`(파이프 label 참조), `in`, `out`, `type`(`elbow`/`elbow-45`/`tee`/...), `count` | 641 |
| `equipment` | `list[dict]` | `desc`(`"A/V"` 등), `pipe`, `in`, `out` | 642 |
| `pumps` | `list[dict]` | `label`, `in`, `out`, `library_pump`, `efficiency`, `status`, `percentage_open`, 성능곡선(`rated_q`,`rated_h`,`shutoff_h`,`peak_q`,`peak_h`) | 643 |
| `valves` | `list[dict]` | PRV. `label`, `in`, `out`, `target_value`(**Pa**), `type`(`"output"`) | 644 |
| `meta` | `list[tuple[str,str]]` | 메타 | 645 |
| `machine_room_plan_edges` | `list[list[float]]` | `[[x1,y1,x2,y2],...]` **시각화 전용, SDF 미포함** | 647 |

**단위 규약 (모듈 C 가 틀리기 쉬운 지점):**
- `nodes.x`, `nodes.y` = **mm**, `nodes.elevation` = **m** — 같은 dict 안에서 단위가 다르다.
- `pipes.length`, `pipes.elev` = **m**, `pipes.dia` = **호칭경 mm**(내경 아님).
- 압력 = **Pa**.

### 2.2 `CommonNode` — 하류 중립 표현 (`kfp_sdf_converter.py:81-95`)

| 필드 | 자료형 | 내용 | 라인 |
|---|---|---|---|
| `id` | `str` | KFP `"N5"` 원형 보존 (SDF `label="5"` 와 양방향 매핑) | 85 |
| `x`, `y` | `float` | **m** (CombinedTables 의 mm 와 다름) | 86-87 |
| `elevation_m` | `float` | **수리 실표고 (m). 권위값.** | 88 |
| `kind` | `str` | `base`/`nozzle`/`wt`/`valve`/`pump` | 89 |
| `k_factor_si` | `float\|None` | 헤드 K (L/min·bar^-0.5) | 90 |
| `pressure_bar` | `float\|None` | 수원 boundary 압력 | 91 |
| `pump_curve` | `dict\|None` | 펌프 곡선 | 92 |
| `valve_type` | `str\|None` | alarm/check/gate/... | 93 |
| `is_check_valve` | `bool` | | 94 |
| `raw` | `dict` | 변환 손실 방지 원본. **`raw["display_z_m"]` 가 여기 산다** | 95 |

### 2.3 `CommonPipe` (`kfp_sdf_converter.py:107-126`)

| 필드 | 자료형 | 내용 | 라인 |
|---|---|---|---|
| `id` | `str` | KFP `"P7"` | 111 |
| `start`, `end` | `str` | 노드 id | 112-113 |
| `length_m` | `float` | **m** | 114 |
| `diameter_inner_mm` | `float` | **내경 mm** | 115 |
| `nominal_mm` | `int` | **호칭경 mm** | 116 |
| `c_factor` | `float` | 기본 120.0 | 117 |
| `roughness_mm` | `float` | 기본 0.15 | 118 |
| `pipe_type_label` | `str` | `"KSD 3507"` 등 | 119 |
| `fittings` | `list[CommonFitting]` | `type_id`,`count`,`l_over_d` (`:98-104`) | 120 |
| `equivalent_length_m` | `float` | 등가길이 추가분 m | 121 |
| `waypoints` | `list[tuple[float,float]]` | 폴리라인 중간 꺾임점 (**m**). ㄷ자 본관용 | 125 |
| `raw` | `dict` | | 126 |

### 2.4 ★ `elevation_m` ↔ `display_z_m` 분리 — 확인 완료

지시서 §15 PR-0 완료 기준의 핵심 항목. **분리는 실재하며 의도적이다.**

| 구분 | `elevation_m` | `raw["display_z_m"]` |
|---|---|---|
| 정의 | 수리 실표고 (m) | 표시 전용 z (m) |
| 위치 | `CommonNode.elevation_m` (`:88`) | `CommonNode.raw` 안 (`:296`, `:1242`) |
| 생성 | `parse_*` 시 meta 의 `elevation_m` 또는 `xyz[2]` (`:282`) | `cn.raw["display_z_m"] = float(xyz[2])` (`:296`) |
| 소비 | 수리계산·SDF `elevation` 속성 | `emit_kfp` 좌표 z |
| 방출 | `"elevation_m": cn.elevation_m` (`:604`) | `"coords": [cn.x, cn.y, disp_z]` (`:601`) |

선택 로직 (`kfp_sdf_converter.py:579-580`):

```python
disp_z = (cn.raw.get("display_z_m", cn.elevation_m)
          if display_geometry else cn.elevation_m)
```

- `display_geometry=True` — 스키매틱 표시좌표(라이저=수직 기둥). 미리보기 비율 일치용.
- `display_geometry=False` — 실표고. **`routes/r30_combined.py:610` 의 KFP 방출은 `display_geometry=False`** 이므로 단독 KFP 는 실표고를 쓴다.
- 어느 경로든 `length_m` · `elevation_m` 는 실값(수리 권위값)으로 보존된다 (`:512` 주석).
- `INTERNAL_RAW_KEYS` (`:567`) 에 `"display_z_m"` 이 포함되어 SDF 재방출 시 raw 로 새지 않는다.

**모듈 C 에 대한 함의:** 하향식 헤드의 수직 드롭은 **표시 z 로만 존재하고 `elevation_m` 에는 반영되지 않는다** (`:202` 주석). 모듈 C 가 헤드 드롭을 `elevation` 에 넣으면 수리계산 낙차가 이중 계상된다. C570 은 **실표고만** `nodes.elevation` 에 넣고, 스키매틱 표현은 별도 채널로 넘겨야 한다.

---

## 3. 담당 헤드 수(부하) 산출 규약 — C570 이 같은 방식을 써야 함

### 3.1 `compute_edge_load` (`remote30_prototype.py:3284`)

```python
def compute_edge_load(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    source: tuple | None,
    heads: list,
    penalty_keys: set | None = None,
    max_attach_mm: float = HEAD_DROP_MAX_MM,
    parents: dict[tuple, tuple] | None = None,
    unreachable_out: list | None = None,
) -> dict[tuple, int]:
```

| 항목 | 규약 |
|---|---|
| 간선 키 | **`(min(u,v), max(u,v))`** — 무방향 정규화. 모듈 C 도 동일해야 조회가 맞는다 |
| 순회 | 소스 루트 최단경로 트리 위 **후위순회 1회 O(V+E)** (`:3340-3354`) |
| 반환 | `graph` 의 **모든** 간선 키. 트리 외 간선(사이클 폐쇄·타 컴포넌트) 부하는 **0** (`:3308-3309`) |
| `source is None` | 전 헤드를 `unreachable_out` 에 수집하고 부하 전부 0 반환 (`:3315-3320`) — 조용히 버리지 않는다 |
| `parents` | 이미 구한 최단경로 부모 맵을 넘기면 재계산 회피 (`:3321-3322`) |
| 도메인 정체 | 누적값 = **담당 헤드 수** = NFPC 103 별표1 '가'칸 최소 호칭경 입력 (`:3302-3303`) |

### 3.2 `prune_by_load` (`remote30_prototype.py:3358`)

- 절단 기준은 **위상이 아니라 "물이 흐르는가"** — `load[k] == 0` 인 간선만 제거 (`:3388-3389`).
- 사이클: 부하 최솟값 0 이면 절단, 1 이상이면 격자 배관 보존 (`:3377-3380`). 별도 사이클 열거 불필요.
- `on_residual_cycle="preserve"` 기본, `"force_tree"` 로 명시 절단 가능.

### 3.3 `_downstream_heads` (`remote30_prototype.py:5845`)

```python
def _downstream_heads(a, b) -> int:
    return _edge_load.get((min(a, b), max(a, b)), 0)
```

중복 구현이 **아니다.** 이미 계산된 `_edge_load` 를 조회만 하는 얇은 래퍼이고, 용도는 NFPC 최소 호칭경 결정(`:5860`). 즉 **추출 소속 판정과 법정 최소 호칭경이 같은 양을 쓴다** (`docs/load_extraction.md` §6). 모듈 C 도 이 단일화를 깨지 말아야 한다.

### 3.4 부하 기준이 성립하지 않는 경우 (`docs/load_extraction.md` §4)

| 경우 | 왜 안 되는가 | 폴백 |
|---|---|---|
| 다중 급수원 | 루트가 둘 이상이면 "상류"가 정의되지 않는다 | `load_mode` 를 켜지 않는다. `force_spanning_tree` 유지 |
| 완전 격자식 | 최단경로 트리는 사각 루프마다 변 하나를 비트리로 남기고 그 부하는 정의상 0 → 격자 미보존 (`BLOCKED.md` §11) | `prune_by_load` 는 부하 dict 를 인자로 받으므로 다중경로 유량 분배 정의를 넣으면 동작. 범위 밖 |
| 환상 배관 | 양방향 급수가 설계 의도인데 트리 부하는 한 방향만 센다 | 동일. 필요 시 `on_residual_cycle="force_tree"` + audit 기록 |
| 헤드 > K(기본 30) | 별표1 최소 호칭경은 선정 top-K 만, 파이프라인 부하는 승인 전체 헤드를 센다 | 두 값을 통합하지 않는다 (`BLOCKED.md` §12) |

**모듈 C 에 대한 함의:** 모듈 C 는 설계 순방향이므로 급수원이 **정의상 단일**이고 배치도 트리로 만든다 — 위 4행 중 1~3 행은 원천적으로 회피된다. 4행(top-K vs 전체 헤드)은 모듈 C 에서도 동일하게 남는다. C570 은 부하를 **전체 헤드 기준**으로 계산하고, 기준개수 절단은 별도 경로로 유지해야 한다.

---

## 4. 부속류(Tee/Elbow) 위상 재구성 규약 — C5 가 같은 규약을 따라야 함

**생성 위치는 `core/*.py` 가 아니라 `remote30_prototype.py:5922-5956` 이다.**

### 4.1 Elbow — 흡수된 꺾임에서 복원

Collinear merge (`:4896-4954`) 가 degree-2 노드를 흡수하면서 각도를 기록해 둔다:

```python
diff = math.degrees(abs(((ang2 - ang1 + math.pi) % (2*math.pi)) - math.pi))
if diff <= COLLINEAR_TOL_DEG:            # 직선 흡수 — 부속류 미기록
elif diff <= ELBOW_MERGE_TOL_DEG and (l_an + l_nb) <= 2 * SHORT_SEG_MM:
    edge_elbows[new_key] = [(n, diff)]   # :4952 기록
```

그 뒤 fitting 으로 변환 (`:5931-5938`):

```python
if 43.5 <= angle_deg <= 46.5:   ftype = "elbow-45"
elif angle_deg >= 70:           ftype = "elbow"
else:                           continue
```

| 각도 구간 | 처리 |
|---|---|
| `diff ≤ COLLINEAR_TOL_DEG` | 직선으로 흡수, 부속류 없음 |
| `43.5° ≤ diff ≤ 46.5°` | `elbow-45` |
| `diff ≥ 70°` | `elbow` |
| 그 외 (46.5°~70°) | fitting 미생성 — 흡수만 되고 사라진다 |
| `diff > ELBOW_MERGE_TOL_DEG` | 흡수 안 함, edge 로 보존 |

### 4.2 Tee — 차수로 판정 (`:5943-5956`)

```python
for p in tables.pipes:
    node_degrees[p["in"]] += 1
    node_degrees[p["out"]] += 1
for p in tables.pipes:
    if node_degrees[p["in"]] >= 3:
        tables.fittings.append({"pipe": p["label"], "in": p["in"], "out": p["out"],
                                "type": "tee", "count": "1"})
```

- 규칙: **파이프의 `in` 노드 차수 ≥ 3 → tee**.
- **순수 기하 판정이다. 유량 방향과 무관하다.** 모듈 C 가 "분기 방향"을 유량으로 정하려 하면 모듈 A 와 갈라진다.
- 주의: 차수 3 노드에 파이프가 3개 붙으면 `in` 이 그 노드인 파이프마다 tee 가 하나씩 생긴다 — 카운트 의미가 "노드당 1개"가 아니라 "파이프당 1개"다. 모듈 C 도 동일하게 채워야 SDF 부속류 수량이 일치한다.

### 4.3 HAS 매핑 (`core/has_converter.py:84-117`)

| `type` | HAS 카운트 필드 |
|---|---|
| `tee` | `CntDivideTee` |
| `elbow` / `elbow-90` | `Cnt90DegreeElbow` |
| `elbow-45` | `Cnt45DegreeElbow` |

여기 표에 없는 `type` 문자열을 모듈 C 가 만들면 HAS 방출에서 **조용히 누락**된다.

---

## 5. 솔버 방출 체인 (`routes/r30_combined.py`)

```
584  from remote30_full_network import emit_full_sdf
585  from remote30_prototype import emit_kfp as _emit_kfp, emit_has as _emit_has
588  emit_full_sdf(net_obj, b_sdf, project_title=title)          # CombinedTables → SDF
597  emit_full_sdf(z_net, z_sdf, project_title=title)            # iso 변형
610  _emit_kfp(z_sdf, b_kfp, coord_scale=..., display_geometry=False)
618  _emit_has(z_sdf, b_has, isometric=True, iso_z_scale=...)
```

- KFP·HAS 는 **SDF 를 다시 읽어** 만든다. 즉 `CombinedTables → SDF` 가 병목이자 단일 계약면이다. 모듈 C 가 SDF 를 올바로 만들면 KFP·HAS 는 공짜로 따라온다.
- `_emit_kfp(display_geometry=False)` → 통합 KFP 좌표 z 는 **실표고**.
- `round_trip_check` 는 이 라우트에서 **호출되지 않는다.** 함수는 `kfp_sdf_converter.py:2528`, `core/has_converter.py:877` 에 존재하며 테스트 경로에서만 쓰인다. 모듈 C 의 왕복 검증(지시서 §13)은 이 두 함수를 직접 불러야 한다.

---

## 6. 모듈 C 개조판이 지켜야 할 계약 — 요약

1. **C570 의 산출물은 `CombinedTables` 다.** `CommonNetwork` 를 직접 만들지 않는다.
2. **단위를 섞지 마라.** `nodes.x/y`=mm, `nodes.elevation`=m, `pipes.length/elev`=m, `pipes.dia`=호칭경 mm, 압력=Pa.
3. **`elevation` 에는 실표고만 넣는다.** 하향식 드롭 등 표시 전용 z 는 `display_z_m` 채널이며 수리 낙차가 아니다.
4. **간선 키는 `(min(u,v), max(u,v))`.**
5. **부하는 담당 헤드 수 하나로 단일화한다.** 추출 판정과 법정 최소 호칭경이 같은 값을 쓴다.
6. **Tee 는 차수 ≥ 3 기하 판정, Elbow 는 각도 구간 판정.** 유량 방향으로 분기를 정하지 않는다.
7. **fitting `type` 문자열은 `core/has_converter.py:84-117` 표 안의 값만 쓴다.**
8. **왕복 검증은 `round_trip_check` 를 직접 호출해야 한다.** 방출 라우트에는 없다.

---

## 7. [문서정합] — 지시서와 구현의 불일치 보고 (지시서 §16-3)

| # | 지시서 기술 | 실제 | 제안 |
|---|---|---|---|
| 1 | §0.2 grep #1 대상 = `core/remote30_graph.py remote30_prototype.py` | `elevation_m` 은 `kfp_sdf_converter.py:88`, `display_z_m` 은 `:296`/`:1242` 에만 존재. 지정 두 파일에는 `display_z_m` 0건 | grep 대상에 `kfp_sdf_converter.py` 추가 |
| 2 | §0.2 grep #3 대상 = `core/*.py` | `core/has_converter.py` 에는 **매핑표**만. **생성 규칙**은 `remote30_prototype.py:5922-5956` | grep 대상에 `remote30_prototype.py` 추가 |
| 3 | §0.2 grep #4 에 `round_trip_check` 포함 | `routes/r30_combined.py` 에 0건 | 왕복 검증은 별도 호출임을 §13 에 명시 |
| 4 | BUG-3 SOLID 처리 = "closePath + fill" | DXF SOLID/TRACE 는 정점을 **0-1-3-2 나비 순서**로 저장. 그대로 fill 하면 기둥이 모래시계가 되어 지시서 자신의 인수 기준("기둥 = 채워진 사각형 격자")에 불합격 | PR-B 에서 shoelace 부호면적 비교로 순서 판별 후 fill 하도록 구현. 코드 주석·커밋 메시지에 사유 기록 |
