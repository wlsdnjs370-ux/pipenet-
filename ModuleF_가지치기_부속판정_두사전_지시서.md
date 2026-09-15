# 가지치기 · 부속 판정 · 두 사전 지시서 — 헤드 없는 가지는 지우되, 그 자리의 부속은 손질 정본에서 정한다

작성 2026-09-15 · 대상 저장소 `PycharmProjects/JupyterProject` · 앞 지시서 `ModuleF_아이소_평면보존_지시서.md` → `ModuleF_회랑_사슬좌표_지시서.md` 의 **후속** · 다음 지시서 `ModuleF_요소속성_수정카드_지시서.md` 의 **선행** · 그림 `docs/images/iso_plan_preserve/그림16_가지치기와_부속판정.png` · `그림17_부속_판정한벌_표기두벌.png`

문제(오너, 2026-09-15): 입력부터 출력까지는 되는데, 아이소와 표에 **굳이 없어도 되는 잔류 노드·배관**이 같이 달라붙는다. 헤드가 붙어 있지 않은 배관은 전부 지우되, 지워진 자리가 «그냥 꺾인 곳» 인지 «갈래가 있었던 곳» 인지를 가려 엘보 / 직류티 / 분류티를 달아야 한다. 그리고 그 부속을 `.sdf` 와 `.kfp` 가 **각자의 규칙**으로 적어야 한다.

---

## 0. 오너가 답하지 않은 자리의 기본값 — 틀리면 줄을 그어라

| | 기본값 | 근거 · 지우면 어떻게 되나 |
|---|---|---|
| D1 판정 어휘 | `none` · `elbow-45` · `elbow` · `tee` (분류티) · **`tee-run` (직류티, 신설)** · `cross` · **`cross-run`** (십자, 티와 같은 규칙 · 그림 16 표 셋째 줄). 크로스는 판정 결과의 이름만 다르고 표기·등가길이는 티와 같다 | 지금 어휘(`core/fitting_rules.py`)에는 직류티가 없어 «부속 없음» 과 구별이 안 된다(그림 17 「빠진 것 하나」). 크로스가 필요 없으면 두 줄만 지운다 |
| D2 `.kfp` 라벨 | K-Solver **레퍼런스 .kfp 의 어휘**를 쓴다: `elbow`→`"Elbow"` · `elbow-45`→`"Elbow 45"` · `tee`·`cross`→`"Tee(Branch)"` · `tee-run`·`cross-run`→`"Tee(Run)"`. 그림 17 ③ 은 우리 어휘(`"tee"`, `"tee-run"`)로 그렸는데, 저장소 안 `kfp_sdf_converter.py:704~716` 이 레퍼런스 파일(1-1.업무시설_201동_28F)에서 확인한 표기가 저것이라 **그쪽을 따른다** | 우리 어휘 그대로 쓰고 싶으면 사전 파일의 라벨 넉 줄만 바꾼다 — 코드는 안 바뀐다 |
| D3 끝배관 1단 보호 | **회랑에서 끈다.** `prune_dead_pipes(..., keep_head_stub=True)` 키워드를 하나 두고 `expand_worst` 만 `False` 로 부른다. 전체망은 기본값이라 종전 그대로 | 그림 16 ② 의 잔류 스텁이 이 보호가 남긴 것이다. 보호를 켠 채로는 스텁을 지울 수 없다. 단, §2 [하향식] 계측이 0 이 아니면 §6-1 로 멈춘다(앞 지시서 §0 「팔 없는 하향식 그대로」와 부딪히는 자리) |
| D4 45° 근처 | 직진/꺾임 임계는 **지금 값 유지** — 분기점은 `TRUNK_TURN_TOL_DEG = 45°`(≤ 이면 직진 → 직류티, > 이면 분류티), 관통은 종전 엘보 구간(≤22.5° 판정 불가 · ≤67.5° 45° 엘보 · ≤95° 90° 엘보). 임계 근처(22.5~67.5°) 갈래의 개수는 §2 에서 세어 보고만 한다 | 판정 불가로 돌리고 싶으면 §6-3 답으로 정한다 |
| D5 최불리 `.kfp` | **표(design)의 망에서 쓴다.** 지금은 `api_convert.convert_one(worst)` 가 제한 전개를 **한 번 더** 돌려 `.kfp` 를 쓰고(`fittings: []` · `equivalent_length: 0.0` · 관경 전부 `DEFAULT_DN`), `.sdf` 는 표에서 나온다 — 같은 K 인데 두 산출물이 다른 망이다. 이제 최불리 `.kfp` = `sess["design"]["got"]["kfp"]` + 표의 관경 + 부속(kfp 사전) + 등가길이. 표가 없으면 `worst_sdf` 와 같은 코드 `worst_required` 로 막는다 | 종전 경로를 살리고 싶으면 이 줄을 지운다 — 그러면 kfp 사전은 적용할 자리가 없다(최불리 `.kfp` 에 부속이 영영 없다). 관경만 종전(기본값·솔버가 고른다)으로 두고 싶으면 「관경」만 지운다 |
| D6 병합된 분기점 | 노드정리가 지운 직선 분기점(갈래가 지워져 직선이 된 board 노드 · `deg_G ≥ 3`)은 그 자리를 덮는 회랑 배관에 `tee-run` 을 **1건씩** 단다 — 부속표와 `.kfp` 라벨에만 나타나고 등가길이 0, `.sdf` 에는 아무것도 안 실린다 | 라벨이 필요 없으면 지운다. 판정 자체(D1)에는 영향 없다 |
| D7 범위 | **최불리 회랑만.** 전체망 `.kfp` 경로(`convert_to_kfp` 직접 호출)는 바이트 불변. 통합배관망·기계실 계통도는 이 지시서 밖 | — |

---

## 1. 규칙 — 오너 확정 (그림 16 · 17)

### 1-1. 한 줄

**헤드로 물을 나르지 않는 배관은 회랑에서 전부 지운다. 지워진 자리의 부속 «종류» 는 지우기 전의 배관망 — 손질 정본 G — 의 차수와 회랑의 흐름 방향으로 정한다. 판정은 한 벌, 표기는 두 벌.**

가지치기는 계산 범위를 좁히는 일이지 배관을 뜯어내는 일이 아니다 — 지워진 갈래의 티는 건물에 그대로 있다. 지금은 가지치기(회랑 단계)가 차수를 바꿔 「부속 종류」 줄을 다시 정하고 있다(그림 10 의 다섯 번째 ✕: 갈래가 지워진 분기점이 직진이면 «없음», 꺾이면 «엘보» 로 나온다 — 분류티 등가길이 3.0 m 자리에 엘보 1.5 m 가 실린다, 50A 기준).

### 1-2. 판정표 — 두 가지 «자» 로 정한다

| G 에서 그 노드의 물리 차수 `phys` | 회랑에서 흐름이 «직진» | 회랑에서 흐름이 «꺾임» | 세로로 빠짐 (①②③④·가지 상승) |
|---|---|---|---|
| 2 (그냥 지나가거나 꺾인 자리) | 부속 없음 (편향 ≤ 0.5°) | `elbow-45` / `elbow` / 판정 불가 — **종전 규칙 그대로** (`fitting_rules.elbow_fittings`) | `elbow` (종전: 3차원 각 90°) |
| 3 (갈래가 있었던 자리) | **`tee-run`** (직류티) ← 지금은 «없음» | **`tee`** (분류티) ← 지금은 «엘보» ✗ | `tee` (종전) |
| ≥ 4 (십자) | `cross-run` | `cross` | `cross` |
| 상류를 모름 (뿌리·고립) | 판정 불가 — 종전 규칙 (`tee_fittings` 의 `upstream is None`) | | |

직진/꺾임의 자: 분기점(phys ≥ 3)은 유입 방향 대비 편향 ≤ `TRUNK_TURN_TOL_DEG`(45°) 이면 직진. 관통(phys = 2)은 종전대로 `_STRAIGHT_EPS_DEG`(0.5°)·엘보 구간. 자리는 종전대로 **하류 배관(부모가 아닌 쪽) 1개에 1건** — 분기점마다 꺾인 갈래에만 `tee`, 직진 갈래에 `tee-run`.

### 1-3. 정의

| 기호 | 뜻 · 어디서 오나 |
|---|---|
| G | 손질 정본 — `limited["pts"]`·`limited["edges"]` (`restrict_to_worst` 는 헤드만 지우고 **배관은 안 자른다**, `restrict.py:18`). 회랑 전개의 입력 그대로다 |
| `deg_G[v]` | board 노드 v 에 붙은 G 간선 수 |
| `node_ref` | kfp 노드 id → board 노드 (`planar.py:861`, `expand_worst` 반환) |
| `edge_ref` | kfp 배관 id → (board_i, board_j). 없는 배관 = 전개가 만든 세로 토막(①②③④·가지 상승·알람밸브 ①②·상하향식 ③) |
| `n_vert[nid]` | kfp 노드 nid 에 붙은, `edge_ref` 에 없는 배관 수 |
| `deg_kfp[nid]` | 가지치기 뒤 kfp 에서 nid 에 붙은 배관 수 (지금 `build_fittings` 의 `len(links)`) |
| **`phys[nid]`** | `max(deg_cell[cell(node_ref[nid])] + n_vert[nid], deg_kfp[nid])` — `node_ref` 가 없는 노드(전개가 만든 것: 헤드 복제·① 꼭대기 등)는 `deg_kfp` 그대로. `max` 인 이유: `normalize_tee_overlaps` 가 허브에서 관통을 쪼개 만든 접속은 G 에 없어도 물리 접속이다(그 수를 §2 에서 센다) |
| | **★2026-09-15 정정 — `deg_G[node_ref[nid]]` 가 아니라 «칸 차수»다.** `node_ref` 는 격자 한 칸(50mm)당 board 절점을 **하나만** 적어 둔다(`node_ref.setdefault(node_id[tgt], vid)`). 게다가 `used` 가 set 이라 그 칸의 어느 vid 가 적힐지도 **임의**다. 그 한 점의 G 차수는 «칸 전체» 의 차수가 아니다 — 티 둘이 한 칸에 접히면 크로스 하나가 되는 그 자리다. 한 점으로만 세면 차수가 **과소평가**된다: 실측(대명동 K=30) 33개 노드에서 차수가 늘고 `phys ≥ 4` 가 14 → 20 으로 바뀌었다. 구현은 `planar.fold_key` 를 그대로 복제해 칸을 만들고, 그 칸의 vid 들을 **밖으로 나가는 G 간선 수**로 센다(칸 안끼리 이은 간선은 밖에서 보면 한 점이므로 안 센다) — `restrict.corridor_topology` |
| 흐름 | 뿌리(접속점, `require_anchor`)에서의 BFS 부모 — 지금 `bfs_order` 가 주는 `parent` 그대로. **바꾸지 않는다** |
| 덮음 경로 `cover[pid]` | 회랑 평면 배관이 덮는 board 조각의 노드열 (board_i … board_j). 사슬 지시서 path 모드가 만든 것이 있으면 **그것**, 없으면 §3-3 의 한 함수로 만들고 사슬도 그것을 쓴다(두 벌 금지) |

### 1-4. 바뀌는 것 · 안 바뀌는 것

| 바뀐다 | 안 바뀐다 |
|---|---|
| 회랑에서 끝배관 1단 보호 → 꺼짐(잔류 스텁 0) | 전체망 `.kfp` (바이트 불변) · 손질 · 최불리 선정 |
| 갈래가 지워진 분기점의 부속: 없음→직류티 · 엘보→분류티 | 관통 노드의 엘보 판정 · 세로 토막의 티/엘보 · 판정 불가 규칙 · `resolve_eq_len` |
| 부속표 ④ 에 `tee-run`·`cross`·`cross-run` 행이 생긴다 | `.sdf` 어휘(PIPENET: `tee`·`elbow`·`elbow-45`) — 직류티는 실리지 않는다 |
| 최불리 `.kfp` 가 표의 망(관경·부속·등가길이)에서 나온다 (D5) | `engine.py` 의 ①②③④ · `planar.py` 의 스냅·병합·노드정리 (키워드 하나 통과 외 diff 0) |
| 부속 종류의 주인 = 손질 G (그림 10 의 다섯 번째 ✕ 해소) | 안정 키(§18 `spot_key`) · 직접 입력의 문법 |

---

## 2. 먼저 잰다 — 바꾸기 전에, 같은 프로브에 여섯 줄을 더한다

`scripts/_probe_plan_preserve.py` 에 추가. 대명동 K=30 무영역/유영역 · B1F 사진 조건 · 보유 도면 전부:

    [잔류]   끝배관1단보호 __개 (planar 로그 값) · 그 배관 길이 합 __m · 표에 남는 차수 1 비헤드·비급수원 노드 __개 · 아이소 not_terminal __개
    [차수]   회랑 kfp 노드 중 phys ≥ 3 인데 deg_kfp = 2 인 노드 __개 (= 갈래가 지워진 분기점)
             그중 직진(≤45°) __ · 꺾임 __ · 편향각 분포(0/45/90/그 외) · phys ≥ 4 __개
             deg_kfp > deg_G + n_vert 인 노드 __개 (planar 가 만든 접속 — normalize_tee_overlaps)
             노드정리로 사라진 board 분기점(cover 내부 · deg_G ≥ 3) __개
    [부속]   지금 부속표: elbow __ · elbow-45 __ · tee __ · 판정 불가 __ · 등가길이 합(최원 경로) __m
             새 규칙 dry-run: elbow __ · elbow-45 __ · tee __ · tee-run __ · cross __ · 판정 불가 __ · 등가길이 합 __m
             (새 규칙은 §3 를 만들기 전 프로브 안에서 «판정만» 돌려 본다 — 파일은 안 쓴다)
    [경계]   분기점 갈래 중 편향 22.5~67.5° 인 것 __개 (자리 목록)
    [하향식] 스텁을 끄면 «팔 없는 하향식(② 만)» → «팔 있는 하향식(①+②)» 으로 바뀌는 헤드 __개 / 5개 (자리 목록)
    [kfp]    최불리 .kfp (convert_one(worst)) 의 pipe id 집합 == 표 kfp 의 pipe id 집합 인가 · 다르면 몇 개 · 관경: .kfp DN 분포 vs 표 DN 분포

**멈추는 조건**: [하향식] > 0 이면 D3(스텁 끄기)은 §6-1 로 보고하고 **그 부분만** 멈춘다 — 부속 판정(§3-2~3-4)은 진행한다. 나머지는 수치를 보고에 싣고 §3 로 간다.

---

## 3. 만든다

### 3-1. 가지치기 — 키워드 하나

`planar.py:171 prune_dead_pipes(pts, edges, head_vids, seed)` → `prune_dead_pipes(pts, edges, head_vids, seed, *, keep_head_stub=True)`. `False` 면 `planar.py:230~246` 의 «인라인 헤드 막다른 쪽 1개 보호» 블록을 건너뛴다(`kept_stub = 0`). 관말 팔 보호(`kept_sole`)·급수원 관 보호는 그대로.

`build_planar_graph(key, out=None, write=False, **graph)` 는 `graph.get("keep_head_stub", True)` 를 그 호출(`planar.py:719`)에 통과시킨다. `restrict.py:199 expand_worst` 의 호출에만 `keep_head_stub=False` 를 넣는다. `attachable_heads`·`ensure_planar`·`convert_to_kfp` 는 안 넣는다(기본값 → 전체망 불변).

이것이 `planar.py` 에 허용된 **유일한** 변경이다(앞 지시서 §5 금지의 예외 한 줄). 스냅·remap·노드정리·길이 식은 손대지 않는다.

### 3-2. 물리 차수 — `expand_worst` 가 만들어 넘긴다

`restrict.py expand_worst` 안, 사슬 단계 다음 · 반환 직전:

1. `deg_G` = `limited["edges"]` 로 센다 (G 전체 — 가지치기 전).
2. `n_vert`·`deg_kfp` = 세로 처리 뒤 `kfp["pipe_data"]` 와 `edge_ref` 로 센다.
3. `phys[nid]` = 1-3 의 식. `node_ref` 가 없는 노드는 넣지 않는다(`build_fittings` 가 종전대로 `len(links)` 를 쓴다).
4. `interior_junctions[pid]` = `cover[pid]` 의 **내부** 노드(양 끝 제외) 중 `deg_G ≥ 3` 인 것의 수 (D6).
5. 반환에 `"phys": phys, "interior_junctions": interior_junctions` 를 **추가**한다(순수 추가 — 기존 키 불변).

`api_design.py:745 build_design_tables(...)` 에 `phys=got.get("phys"), interior_junctions=got.get("interior_junctions")` 를 넘기고, `tables.py:136` 의 시그니처에 두 키워드(기본 `None`)를 더해 `tables.py:210` 의 `build_fittings(...)` 로 그대로 흘린다. 다른 호출자는 안 넘기므로 종전과 같다.

### 3-3. 덮음 경로 — 한 함수

사슬 단계(path 모드)가 이미 «배관이 덮는 board 조각» 을 갖고 있으면 그것을 `cover` 로 내보낸다. 없으면 `restrict.py` 에 하나 만든다:

    board_cover(pts, adj_G, bi, bj) -> list[vid] | None
      bi 에서 출발해, 현재 노드의 이웃 중 «bj 쪽으로 가장 곧게 이어지는» 이웃(bi→bj 방향과의 편향 최소 · ≤ COLLINEAR 허용각)을 따라 bj 까지 걷는다.
      곧은 이웃이 없거나 bj 에 못 닿으면 None (지어내지 않는다 · 보고에 «cover 실패 __개»).

이 함수가 생기면 사슬 지시서 C2(「그 배관이 덮는 board 조각들의 길이 합」)도 **같은 함수**를 쓴다. 두 벌을 두지 않는다.

### 3-4. 판정 — `build_fittings` 한 곳, 어휘는 `fitting_rules` 한 곳

`core/fitting_rules.py` (두 브랜치가 쓰는 순수 함수 — **기존 함수·상수는 건드리지 않는다**, 순수 추가만):

    TEE_RUN = "tee-run"
    CROSS = "cross"
    CROSS_RUN = "cross-run"

    def tee_split(node, upstream, downstream, *, turn_tol_deg=TRUNK_TURN_TOL_DEG)
        -> tuple[list[str], list[str], int]      # (꺾인 갈래 라벨, 직진 갈래 라벨, 판정 불가 건수)
      tee_fittings 와 같은 자·같은 규칙(상류 없음 → 전부 꺾임 + 판정 불가 수 · 직진 둘 이상 → 전부 꺾임 + 판정 불가 수)이되
      직진 갈래를 버리지 않고 돌려준다. 단 하나 다른 점: 하류가 **1개여도** 판정한다(tee_fittings 는 «분기가 아니다» 로 빈 값을
      돌려주는데, 갈래가 지워진 분기점이 바로 하류 1개짜리 분기점이다 — 차수는 호출자가 phys 로 이미 가렸다).
      tee_fittings 는 그대로 두고(다른 브랜치 호출자), 시험으로 «하류 2개 이상일 때 tee_split 의 꺾인 갈래 == tee_fittings 의 결과» 를 고정한다.

`design/fitting.py build_fittings(net, node_xy, bores, *, parents=None, lib=None, node_z=None, overrides=None, phys=None, interior_junctions=None)`:

1. `phys` 가 `None` 이면 **종전과 동일하게** 동작한다(전체망·다른 호출자 보호). 아래는 `phys` 가 있을 때다.
2. 노드마다 `p = phys.get(nid, len(links))`, `up`·`downs` 는 종전대로.
3. `p == 2` → 종전 관통 블록(`fitting.py:241~283`) 그대로.
4. `p >= 3` → 세로 하류는 종전대로 `TEE`(크로스면 `CROSS`). 평면 하류는 `tee_split(here, up_xy, flat_downs)`: 꺾인 갈래 → `TEE`/`CROSS`, 직진 갈래 → `TEE_RUN`/`CROSS_RUN`, 판정 불가는 종전 방식으로 세고 §18 직접 입력을 종전대로 받는다. **`flat_downs` 가 1개뿐이어도**(갈래가 지워진 분기점) 이 블록을 탄다 — 지금은 `len(flat_downs) < 2` 면 건너뛰는데(`fitting.py:301`) 그 줄이 바로 다섯 번째 ✕ 다. 크로스 여부는 `p >= 4`.
5. `interior_junctions[pid]` 만큼 그 배관에 `TEE_RUN` 을 더 단다(D6).
6. `FITTING_LIB_ID` 에 `"cross": "TEE_BRANCH"`, `"tee-run": None`, `"cross-run": None` 을 더한다. `resolve_eq_len` 은 «`FITTING_LIB_ID` 에 있되 id 가 `None` 인 종류» 에 대해 라이브러리를 보지 않고 `(0.0, "규칙")` 을 돌려준다 — **미해결이 아니다**(미해결은 «값을 모른다» 이고 이것은 «값이 0 이라고 규칙이 정했다» 다). `fitting.py:344` 의 `applied_overrides` 조건(`why != "라이브러리"`)은 `"규칙"` 도 제외하도록 고친다 — 규칙이 정한 0 을 «사람이 채운 값» 으로 세면 안 된다. 어휘에 없는 종류(오타)는 종전대로 `(None, None)` = 미해결.
7. `counts`·`per_pipe`·`unresolved_*` 의 모양은 그대로. 새 어휘가 값으로 들어갈 뿐이다.

`api_design.py:935~947` 의 종류 목록(직접 입력 카드가 고르는 값)에 `{"value": fr.TEE_RUN, "label": "직류티"}` 를 더하고 `fr.TEE` 의 라벨을 「분류티」로 바꾼다. 크로스는 목록에 넣지 않는다(사람이 고를 일이 없다 — 규칙이 차수로 정한다).

### 3-5. 두 사전 — 표기는 파일이 정한다

`cad_project_editor_g/fittings_dict_sdf.json` · `fittings_dict_kfp.json` (라이브러리 `fittings_library_v3.json` 옆). 사전은 **라벨만** 갖는다. 등가길이의 숫자와 라이브러리 id 는 종전대로 `FITTING_LIB_ID`·`resolve_eq_len` 한 곳이다(두 주인 금지).

    fittings_dict_sdf.json
    { "format": "sdf", "target": "PIPENET SDF <Fitting type count/>",
      "vocabulary": ["tee", "elbow", "elbow-45", "gate", "butterfly", "check"],
      "map": { "elbow": "elbow", "elbow-45": "elbow-45", "tee": "tee", "cross": "tee",
               "tee-run": null, "cross-run": null, "none": null },
      "note": "null = 표기 없음. PIPENET 은 직류 흐름 티 항목이 없다(설계팀 수작업 .sdf 도 tee·elbow·elbow-45 뿐). 손실값은 PIPENET 이 제 표로 정한다." }

    fittings_dict_kfp.json
    { "format": "kfp", "schema": "4.0-NFPA13-EQL",
      "map": { "elbow": "Elbow", "elbow-45": "Elbow 45", "tee": "Tee(Branch)", "cross": "Tee(Branch)",
               "tee-run": "Tee(Run)", "cross-run": "Tee(Run)", "none": null },
      "note": "라벨은 K-Solver 레퍼런스 .kfp(1-1.업무시설_201동_28F) 의 표기. equivalent_length 는 배관별 숫자(m) = 표의 부속 등가길이 합 — resolve_eq_len 이 정한 값을 그대로 싣는다." }

`design/fitting.py` 에 `load_fitting_dict(fmt) -> dict` (파일 한 번 읽고 캐시 · 사전에 없는 종류가 오면 **던진다** — 조용히 빠뜨리지 않는다).

- `.sdf`: `design/emit.py:106~109` 의 `fit_by_pipe` 를 만들 때 `dict_sdf["map"][row["type"]]` 로 치환한다. `null` 이면 그 행을 싣지 않는다(직류티·판정 불가). 그 외 `emit_design_sdf`·`sdf_writer.py`·`models.py` 는 손대지 않는다.
- `.kfp`: `design/emit.py` 에 `emit_design_kfp(tables, got, out_path) -> dict` 를 **새로** 둔다(`emit_design_sdf` 와 나란히 — 「표 → .sdf」와 「표 → .kfp」). `copy.deepcopy(got["kfp"])` 에 배관마다: `fittings` = 부속표 ④ 의 그 배관 행을 `dict_kfp["map"]` 으로 치환한 문자열 목록(`null` 은 뺀다 · 같은 종류 n건이면 n번 반복 — `kfp_sdf_converter` 가 읽는 형태), `equivalent_length` = `fittings["per_pipe"][pid]["equivalent_length"]`(같은 숫자 · 반올림 3자리), `nominal_mm` = 배관표 ② 의 `dia`, `diameter` = 그 DN 의 내경(`pipe_specs_for_standard(editor.pipe_library, OPT_PIPE_STD)` 의 `inner_d_mm` — planar 가 쓰는 같은 표). 노드·좌표·길이·`type_id`·`fitting_id`(알람밸브 `VALVE_ALARM`)는 **손대지 않는다**. 파일은 `json.dump(kfp, f, ensure_ascii=False, indent=2)` — `engine.py:929~931` 과 같은 모양.
- `routes/module_f/api_convert.py:174~202` 의 `outputs["worst_kfp"]` 블록: `sess["design"]` 이 없으면 `worst_sdf` 와 같은 `worst_required`(문구: 「표 확정을 먼저」). 있으면 `emit_design_kfp(d["tables"], d["got"], out_path)` 로 쓴다. `convert_one(worst)` 는 이 블록에서 더 이상 부르지 않는다(전체망 블록은 그대로). `summary["worst"]` 의 키(k·nodes·pipes·bytes·filename)는 유지.

`tables.py:340~349` 부속표 ④ 는 판정 어휘 그대로 싣는다(`tee-run` 행이 생긴다). 화면의 부속표는 그 값을 그대로 보이되, 라벨이 있으면 종류 목록의 한글 라벨을 쓴다.

### 3-6. 순서

사슬 단계 **다음**(cover 를 받는다) · `build_design_tables` **전**. `phys` 는 위상만 쓰므로 사슬의 좌표와 무관하다. 속성 수정카드 지시서는 이 작업 **뒤**에 — 카드의 부속 칸이 이 어휘를 쓴다.

---

## 4. 검증

### 4-1. 판정식

| | 판정 | 허용 |
|---|---|---|
| F1 잔류 없음 | 회랑 kfp 에 차수 1 인 노드는 헤드와 급수원뿐. planar 로그 `끝배관1단보호 0` (회랑) · 아이소 `not_terminal` 이 §2 값에서 스텁 수만큼 준다 | 0 개 |
| F2 판정 = f(G, 흐름) | 부속 종류가 **가지치기의 결과에 의존하지 않는다**: 같은 회랑에서 (a) G 에 갈래가 있고 자동으로 지워진 경우와 (b) 손질에서 그 갈래를 처음부터 없앤 G′ 의 경우, 그 노드의 부속이 (a) 티 · (b) 없음/엘보 로 **다르게** 나온다. 그리고 (a) 에서 스텁 보호를 켜든 끄든 그 노드의 종류는 같다 | 합성 시험 |
| F3 표기 두 벌 | `.sdf` 의 `Fitting type` 집합 ⊆ {tee, elbow, elbow-45, gate, butterfly, check}. 최불리 `.kfp` 의 `fittings` 문자열 집합 ⊆ {Elbow, Elbow 45, Tee(Branch), Tee(Run)}. 어느 쪽에도 사전 밖 값이 없다 | 0 건 |
| F4 숫자 하나 | 배관마다 (부속표 ④ 의 종류 목록) = (`.kfp` `fittings` 를 사전 역방향으로 읽은 목록 + null 로 빠진 것) 이고, `.kfp` `equivalent_length` = `per_pipe[pid]["equivalent_length"]` | ≤ 0.001 m |
| F5 전체망 불변 | 전체망 `.kfp` 바이트 동일 · 손질 화면 · 최불리 선정 · 전체망 planar 로그(`끝배관1단보호` 값 포함) 동일 | 바이트 |
| F6 관경 | 최불리 `.kfp` 의 배관별 `nominal_mm` = 배관표 ② `dia` (D5) | 0 건 |
| C1~C4 · W · H1·H3 · I1 | 앞 두 지시서의 판정식 그대로 — 가지치기가 위상을 줄여도 사슬 닫힘·길이·아이소 식은 유지 | 앞 지시서 값 |

### 4-2. 수용 기준

| | 기준 |
|---|---|
| 1 | 대명동 K=30(무영역·유영역) · B1F 사진 조건: F1~F6 통과, 앞 지시서 판정식 통과 |
| 2 | 그림 16 합성망(§4-3)에서 E=`tee-run` · C=`tee` · F=`elbow` · G 스텁 삭제 · H1 관말 — 그리고 `.sdf` 는 C 배관에만 `<Fitting type="tee">`, E 배관에 아무것도 없음, F 배관에 `elbow`; `.kfp` 는 E `["Tee(Run)"]`·0.0, C `["Tee(Branch)"]`·TEE_BRANCH(DN), F `["Elbow"]`·ELBOW_90_STD(DN) |
| 3 | 실도면 부속 개수 전/후 표 — 「엘보→분류티」로 바뀐 자리 수 = §2 [차수]의 «꺾임» 수, 「없음→직류티」 = «직진» 수 + 노드정리로 사라진 분기점 수. 안 맞으면 어느 노드인지 |
| 4 | 최원 경로 등가길이 합 전/후 — 줄어들면 안 된다(엘보 → 티는 항상 커진다; 줄면 버그) |
| 5 | `engine.py` · `worst.py` · `api_edit.py` · `sdf_post.py` · `sdf_writer.py` · `models.py` diff 0. `planar.py` diff = 키워드 통과 두 자리뿐. `fitting_rules.py` 기존 함수·상수 diff 0 |
| 6 | 기존 회귀 전부 통과 · 골든 재생성 없음 · `tests/test_fitting_rules.py` 그대로 통과 |
| 7 | 화면: 표 확정 뒤 아이소에 잔류 스텁이 없고, 부속표에 직류티/분류티가 갈려 보이며, 판정 불가 카드의 종류 목록에 「직류티」가 있다 |

### 4-3. 시험 `tests/test_prune_fittings.py`

- 그림 16 ① 의 합성 board(급수 → E(차수 3, 아래로 H0 K 밖) → H1 자리(인라인, 위로 G 스텁) → C(차수 3, 오른쪽 H3 K 밖) → 위로 F(차수 2 꺾임) → H2). K = {H1, H2}. 회랑 전개 → 표 → `.sdf` · `.kfp` 까지 돌려 수용 기준 2 를 식으로 검사.
- F2: 같은 망에서 E 의 아래 갈래를 board 에서 아예 뺀 G′ → E 는 직진 관통 → 부속 없음(노드정리 병합). 두 결과가 다름을 검사.
- 십자 1개(phys 4 · 하류 셋: 직진 1 + 꺾임 2 → 직진 갈래 `cross-run`, 꺾인 갈래 `cross` 둘). `.sdf` 는 꺾인 두 배관에만 `tee`, `.kfp` 는 `Tee(Run)` 1 · `Tee(Branch)` 2, 등가길이는 꺾인 두 배관만 TEE_BRANCH(DN).
- D6: 갈래가 지워져 직선이 된 분기점이 노드정리로 사라진 경우 → 덮는 배관에 `tee-run` 1건 · `.sdf` 없음.
- `phys=None` 로 부르면 종전 결과와 동일(전체망 보호) — 기존 부속 시험 하나를 `phys` 없이 그대로 돌린다.
- `tee_split` 의 꺾인 갈래 == `tee_fittings` (같은 입력 6벌 — `test_fitting_rules.py` 의 케이스 재사용).
- 사전에 없는 종류를 넣으면 `load_fitting_dict` 경로가 던진다.
- 실도면(대명동 K=30)은 F1~F6 0 건으로 고정 — 골든 아님.

---

## 4-4. 실측 결과 (2026-09-15 · 이 지시서로 작업한 뒤)

| | 대명동 K=30 |
|---|---|
| 가지치기 잔류 | 6 → **0** (남은 차수-1 노드 1개는 알람밸브 세로 스택 끝) |
| 세로 토막 | 62 → **62** — ★스텁을 꺼도 **팔이 안 생긴다**(§6-1 의 우려 미발생) |
| 부속 | elbow 126→95 · elbow-45 16→8 · tee 37→56 · **tee-run 62 · cross 20 · cross-run 14** |
| 등가길이 합 | 224.6 → **269.0 m** (F4 기준: 줄면 버그) |
| 최불리 `.kfp` | 배관 347(옛 두 번째 전개) → **222 = 표** · 관경·등가길이 어긋남 0 |

**D4 정정(오너 2026-09-15)**: 분기의 직진/꺾임은 45° 한 자리에서 자르지 않고
**0/45/90 중 가까운 쪽으로 스냅**한다(경계 22.5·67.5 — 엘보 규칙과 같은 자).
46° 가 89° 와 한 통에 들던 것이 갈렸다. `fitting_rules.snap_turn_deg`.

**D3 정정(오너 2026-09-15)**: [하향식] > 0 이어도 **끈다.** 실측으로 팔이 안
생겼기 때문이다(대명동 62→62 · B1F 33→33 · 노즐 30→30). 관말이 되는 것만으로는
팔이 안 생긴다 — `_pendant_arm_to_tee` 는 **꺾인** 걸음에서만 팔을 만든다.

★**지시서 본문의 줄 번호는 동결본(`cad_project_editor/`, 은퇴한 E 판) 기준이라
  전부 어긋나 있다.** 손댈 자리는 `cad_project_editor_g/` 쪽이다. 예:
  `prune_dead_pipes` 는 171 이 아니라 **188**, 스텁 보호 블록은 230~246 이 아니라
  **247~263**, 호출은 **764**. `build_planar_graph` 는 `**graph` 로 이미
  통과시키므로 **손댈 필요가 없다**(손댈 자리는 `main` 시그니처와 그 호출 둘뿐).

---

## 5. 금지

- `planar.py` 는 `keep_head_stub` 키워드 통과 외 diff 0. 스냅·remap·`normalize_tee_overlaps`·노드정리·길이 식을 바꾸지 말 것.
- `engine.py` 의 ①②③④·재질 복사·알람밸브 `fitting_id` 를 바꾸지 말 것.
- 부속 판정을 `prune_dead_pipes`·planar·`expand_worst` 안에서 하지 말 것 — `expand_worst` 는 **차수를 세어 넘길 뿐**, 종류를 정하는 곳은 `build_fittings` 하나다.
- 사전(json)에 등가길이 숫자나 라이브러리 id 를 적지 말 것 — 사전은 라벨뿐이다.
- 판정 불가를 0·엘보·티로 메우지 말 것(종전 규칙). 직류티의 0 은 「규칙이 정한 0」이라 미해결과 섞지 말 것.
- `.sdf` 에 직류티를 `tee` 로 내지 말 것. `.kfp` 의 `fittings` 에 사전 밖 문자열을 내지 말 것.
- 전체망 `.kfp` 경로(`convert_to_kfp` 직접 호출)에 `phys`·사전·스텁 끄기를 적용하지 말 것.
- 흐름 방향(`bfs_order` 의 부모)·임계 상수(`TRUNK_TURN_TOL_DEG`·엘보 구간)를 바꾸지 말 것 — D4.
- 손질 · 최불리 선정 · 헤드 종류 판정에 되먹이지 말 것. 리팩터링 금지 · 골든 재생성 금지.

## 6. 애매하면 멈추고 묻는다

1. §2 [하향식] > 0 — 스텁을 끄면 «팔 없는 하향식» 이 «팔 있는 하향식» 이 되는 헤드가 있다. 앞 지시서 §0 은 «그대로 둔다» 였다. 자리·개수와 ①·② 길이 변화를 들고 멈춘다(부속 판정은 진행).
2. `phys` 를 만들 `node_ref` 가 없는 **평면** 노드가 있다(전개가 만든 것이 아닌데 역참조가 없다) — 개수·자리. 그 노드는 종전 규칙으로 두고 보고한다.
3. §2 [경계] > 0 — 편향 22.5~67.5° 갈래. D4 는 45° 임계를 유지하는데, 오너가 판정 불가로 돌리고 싶을 수 있다.
4. K-Solver 가 `fittings` 라벨만으로 손실을 **다시** 계산하는가(`equivalent_length` 가 있어도). 저장소의 `kfp_sdf_converter.py:2079` 는 `"tee(run)"` 을 분기 티 L/D 로 «일단» 묶어 두었다 — 오너가 K-Solver 로 확인할 자리. 그때까지 D2 대로 라벨 + 숫자 둘 다 싣는다.
5. 급수원(접속점) 노드의 `phys` — 급수원 관 보호(`planar.py:248~252`)로 남은 배관이 회랑 나무의 뿌리에서 갈라져 있으면 그 자리는 상류가 없어 종전대로 판정 불가다. 개수 보고.
6. cover 실패(§3-3 의 `None`)가 0 이 아니다 — 자리·이유. D6 라벨만 빠지고 판정은 영향 없지만, 사슬 C2 가 같은 함수를 쓴다면 그쪽도 걸린다.
7. §2 [kfp] 에서 최불리 `.kfp` 와 표 kfp 의 배관 집합이 달랐다면 그 차이가 무엇이었는지 — D5 뒤에는 같은 객체라 사라지지만, 왜 달랐는지는 기록한다.

## 7. 보고 형식

    [계측]   §2 여섯 줄 — 도면별
    [가지치기] 회랑 끝배관1단보호 전/후 · 잔류 노드·배관 수 전/후 · 아이소 not_terminal 전/후
    [판정]   부속 개수 전/후 표(종류별) · 바뀐 자리 목록(노드·전→후) · 최원 경로 등가길이 합 전/후
    [표기]   .sdf Fitting type 집합 · .kfp fittings 문자열 집합 · F4 배관별 대조 결과
    [불변]   전체망 .kfp 바이트 비교 · 금지 파일 diff · planar.py diff 두 자리
    [시험]   신규 __건 · 회귀 __건 통과
    [묻는 것] §6 해당 항목과 수치
