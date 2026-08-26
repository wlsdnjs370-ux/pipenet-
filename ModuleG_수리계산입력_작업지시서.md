# Module G 작업 지시서 — 최불리 배관망 확정과 수리계산 입력(SDF) 생성

당신은 FNCADnet 저장소의 유지보수 엔지니어다. 아래를 **항목 순서대로, 항목당 커밋 1개씩**
수행하라. 각 항목은 수용 기준을 통과한 뒤에만 다음으로 진행한다. 판단이 필요한 애매점은
임의로 구현하지 말고 `BLOCKED.md`에 기록하고 다음 항목으로 넘어간다(§7).

작업 대상 트리는 **`cad_project_editor_g/`** 다. `cad_project_editor/`(모듈 E)와
`routes/module_f/`는 **읽기 전용 참조**이며 한 줄도 고치지 않는다.

---

## 0. 이 작업의 정체 (구현 전 필독)

### 0.1 왜 kfp는 되고 sdf는 안 되는가

현재 G의 변환은 `convert_to_kfp` 로 **전체 배관망** `.kfp` 를 만들고, SDF 는 그
`.kfp` 를 `kfp_sdf_converter` 에 태워 **문법만 옮겨 적은 것**이다.

- `.kfp` → K-Fire Solver 가 전체망을 받아 **설계구역·최불리·테이블을 내부에서** 처리한다.
  그래서 전체망을 그대로 넘기는 현재 동작이 옳다.
- `.sdf` → PIPENET 은 **주어진 망을 풀 뿐 설계구역을 스스로 고르지 않는다.**
  지금 나가는 SDF 는 「전체망을 PIPENET 문법으로 재직렬화한 파일」이지 수리계산 입력이 아니다.

따라서 SDF 를 수리계산 입력으로 쓰려면 **최불리 확정 → 위상 정리 → 테이블 채우기**를
프로그램이 미리 해야 한다. 이것이 이번 작업이다. 모듈 A 의 Stage 4~6 과 같은 일이며,
**새로 발명하는 것이 아니라 이미 있는 순수 함수를 옮겨 오는 일**이다(§5).

### 0.2 확정된 설계 결정 (변경 금지)

| # | 항목 | 결정 |
|---|---|---|
| D1 | 최불리 기준 | **모듈 A/F 의 앵커 방식을 그대로 이식.** 급수원에서 배관거리로 최원인 헤드를 앵커로 잡고, 앵커에서 배관거리로 가까운 K개를 설계면적으로 본다. 「먼 순서 K개」가 아니다 |
| D2 | 테이블 값 | **혼합.** 관경 = `max(도면 치수 텍스트, NFPC 103 별표1 최소 호칭경)`. 부속 판정 = `core/fitting_rules.py`, 부속 등가길이 = `cad_project_editor_g/fittings_library_v3.json` |
| D3 | SDF 범위 | **최불리 구간만.** 전체망 SDF 는 내지 않는다. `.kfp` 는 종전대로 전체망 |

### 0.3 두 번 전개 원칙 ★이 작업의 뼈대

같은 편집 그래프(`EditBoard`)에서 **전개를 두 번** 돌린다.

```
EditBoard ─┬─ 전체망 전개 ──────────────→ .kfp        (기존 경로 · 손대지 않는다)
           │
           └─ 최불리 제한 → 제한 전개 → 5표 → .sdf    (이번에 신설)
```

두 산출이 같은 손질 결과·같은 치수 입력(DTO)에서 나오므로 설계 내용이 어긋나지 않는다.
**기존 `.kfp` 경로의 산출물은 이번 작업 전후로 비트 동일해야 한다**(§4 회귀).

---

## 1. 신규 계약

새 패키지 `cad_project_editor_g/services/cad_import/design/` 를 만든다. 화면 없음(Qt import 금지).

```
design/
  __init__.py
  worst.py     최불리 K 선정 (G1)
  restrict.py  corridor 제한 payload (G2)
  bore.py      관경 결정 (G3)
  fitting.py   부속·노즐·기기 (G4)
  tables.py    5개 테이블 조립 (G5)
  emit.py      SDF + SLF 방출 (G6)
  assets/      템플릿 .sdf · 표준 .slf (G6)
```

공개 시그니처(이 형태를 유지할 것):

```python
worst_k_heads(pts, edges, hnodes, sources, k=30, only_heads=None) -> dict
restrict_to_worst(payload, board, worst) -> dict
decide_bores(net, edge_ref, loads, dia_text_pts) -> dict   # pipe_id -> (dia_mm, source)
build_design_tables(net, worst, edge_ref, dia_text_pts, *, project_title) -> PipeTablesG
emit_design_sdf(tables, out_path, *, project_title) -> Path
```

`PipeTablesG` 는 모듈 A 의 `PipeTables` **키 규약을 그대로** 따른다. 자체 dataclass 로
만들되 필드명·단위를 바꾸지 않는다(§3).

---

## 2. 작업 항목 (순서 고정)

### G1. 최불리 K 선정 이식

`routes/module_f/remote30.py:_worst_k_heads` 를 `design/worst.py` 로 옮긴다.
이 함수는 `pts / edges / hnodes / sources` 만 받는 **순수 그래프 함수**이며 Qt·Flask
의존이 없다. **로직을 고치지 말고 그대로 옮긴다.**

호출부는 손질 세션에서 직접 만든다:

```python
b = edit_session.board
w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=K, only_heads=only)
```

`only_heads` 는 도면이 여러 장일 때 한 장으로 좁히는 인자다. 장 나누기는
`routes/module_f/remote30.py:_sheet_frames` 가 모듈 A 의 규칙을 부르는 형태이므로
같은 방식으로 이식한다(장이 하나면 `None`).

이 항목에서는 창을 만들지 않는다. 콘솔에만 찍는다.

**수용 기준**
- 급수 시작 위치가 없으면 계산하지 않고 「급수 시작 위치를 먼저 찍어야 한다」로 막는다.
- 반환 dict 에 `heads / anchor / edges / loads / nodes / far_m / near_m / span_m / total_m / max_load` 가 모두 있다.
- 같은 도면·같은 K 로 모듈 F 의 `/api/module-f/edit/worst` 를 돌린 결과와
  `anchor`, `heads` 집합, `far_m`, `max_load` 가 **완전히 일치**한다. (일치하지 않으면 이식 실패다.)
- 앵커 30개의 배관거리 폭(`span_m`)이 「먼 순서 30개」보다 작다 — B1F 실측 기준값: 앵커 30.3 m vs 먼 순서 95.9 m.

### G2. corridor 제한 전개 + 역참조 노출

`routes/module_f/remote30.py:_restrict_to_worst(payload, board, worst)` 를
`design/restrict.py` 로 옮긴다. 이 함수는 이미 board 를 인자로 받으므로 **시그니처 변경 없이
그대로 쓴다.** 간선을 직접 자르지 말 것 — 남길 헤드만 남기면 `build_planar_graph` 가
제 규칙(물길 필터 → 막다른관 삭제)으로 정리한다. 손으로 자르면 티 겹침·노드정리 불변식이 깨진다.

이어서 **제한 payload 로 두 번째 전개**를 돌리는 진입점을 추가한다. 기존
`convert_to_kfp` 의 저장 경로·반환 규약은 건드리지 않는다. 파일을 쓰지 않고 메모리 상의
망(dict)만 돌려주는 형태여야 한다.

**같은 커밋에서 역참조를 노출한다.** `convert/planar.py:build_planar_graph` 는 내부에
`node_id`(board 노드 인덱스 → kfp 노드 id)와 `remap` 을 이미 갖고 있으나 밖으로 내보내지
않는다. 이것을 결과에 실어 보낸다:

```
edge_ref: {kfp_pipe_id: (board_i, board_j)}   # 전개로 쪼개진 배관도 원 간선을 가리킨다
node_ref: {kfp_node_id: board_i}
```

수직 전개(`convert_to_kfp`)에서 한 배관이 여러 조각으로 나뉘거나 수직 배관이 새로 생기면,
**원 간선을 물려받게** 한다(새로 생긴 수직 배관은 부모 간선 참조를 그대로 상속).
이것이 없으면 G3 의 관경이 엉뚱한 배관에 붙는다.

**수용 기준**
- 기존 전체망 `.kfp` 산출물이 이번 변경 전후로 **비트 동일**(§4).
- 제한 전개 결과의 헤드 수 == `len(worst["heads"])`.
- `edge_ref` 가 제한 망의 **모든** 배관을 덮는다(누락 0). 덮지 못하는 배관이 있으면
  그 배관 종류(수직 드롭·암 등)를 로그로 남기고 `BLOCKED.md` 에 기록.

### G3. 관경 결정 — 혼합 규칙

`design/bore.py`. 모듈 A 의 `remote30_prototype.py:_pipe_diameter` 와
`_nfpc_min_bore_mm` 이 이미 이 규칙 그대로다. **옮겨 온다.**

```
nfpc_min = 별표1('가'칸)[담당 헤드 수]      # 25/32/40/50/65/80/90/100/125/150
text     = 선분에서 수직거리 ≤ 1500 mm 인 가장 가까운 치수 텍스트
dia      = nfpc_min            (text 없음)
         = nfpc_min            (text < nfpc_min — 안전측)
         = text                (그 외)
```

- **담당 헤드 수는 `worst["loads"][(i,j)]` 를 그대로 쓴다.** 이는 corridor 안에서 그 간선이
  책임지는 **선정된 K개 중의 수**다. 전체망 하류 헤드 수를 넣으면 관경이 과대해진다.
- 치수 텍스트는 **DXF 를 다시 읽지 않는다.** `pipeline/handoff.py` 의 캐시에
  `world.texts = [(layer, color, x, y, h, text), ...]` 로 이미 보존돼 있다. 모듈 A 의
  `_extract_dia_text_points` 가 기대하는 `{"t":"T","v":text,"p":[x,y]}` 형태로 어댑터만 쓴다.
- 매칭은 **평면 mm 좌표**에서 한다. `edge_ref` 로 kfp 배관 → 원 board 간선 → `pts` 좌표를
  얻어 그 선분에 매칭한다. 전개된 m 좌표로 매칭하면 안 된다.
- 간선마다 근거를 남긴다: `source ∈ {"text", "nfpc_min", "nfpc_fallback"}`.

**수용 기준**
- 합성 테스트: 담당 헤드 수 1→25, 3→32, 5→40, 10→50, 30→65, 60→80, 100→100, 200→150.
- 합성 테스트: 텍스트 "50A" 가 1200 mm 거리에 있고 별표1 최소가 65 → **65** 채택,
  `source="nfpc_min"`.
- 실도면(B1F) 실행 시 `source` 집계가 로그에 남고, `text` 비율이 0% 가 아니다
  (0% 면 텍스트 어댑터가 죽은 것이다).

### G4. 부속 · 노즐 · 기기

`design/fitting.py`.

- **부속 판정**은 `core/fitting_rules.py` 의 `elbow_fittings(angles)` / `tee_fittings(...)`
  를 이식해 쓴다. 직진해 지나가는 갈래는 **직류티라 계상하지 않는다** — 이 규칙이 이미
  그 모듈 안에 있으니 다시 구현하지 말 것.
- **등가길이**는 D2 에 따라 `cad_project_editor_g/fittings_library_v3.json` 에서 가져온다.
  전개된 배관 dict 에 이미 `equivalent_length` 필드가 있으므로 그 자리에 넣는다.
- **노즐**은 헤드 노드마다 1행. K 값·필요압력은 변환 창의 DTO 값을 그대로 쓴다
  (여기서 새로 정하지 않는다).
- **기기**는 알람밸브(A/V) 1행부터. 손질 단계에서 찍은 알람밸브 위치가 없으면 행을 만들지 않는다.

**수용 기준**
- 합성 테스트: 90° 꺾임 1곳 → 엘보 1, 직진 통과 분기 → 티 0, 갈래 분기 → 분류티 1.
- 부속표의 `pipe` 값이 전부 배관표에 존재하는 라벨이다(고아 참조 0).
- 등가길이가 라이브러리에 없는 부속 종류는 **0 으로 채우지 말고** 미해결로 세어 로그에 남긴다.

### G5. 5개 테이블 조립 ★이번 작업에서 유일하게 «새로 짜는» 것

`design/tables.py`. 제한 전개 망 → `PipeTablesG`.

- 급수원을 뿌리로 **BFS 한 번**을 돌려 노드 라벨 번호·행 순서·배관 방향(`in`→`out`)을
  모두 그 순서에 맞춘다. 표를 위에서 아래로 읽으면 물이 흐르는 순서가 되어야 한다.
  트리에 들어가지 못한 간선(루프 잔여)은 **표 꼬리로 몰아** 배치한다 — 그 꼬리가 곧
  「길이 잘못 트인」 후보 목록이다.
- 단위를 틀리지 말 것(§6 T3): `nodes.x/y` = **mm**, `nodes.elevation` = **m**,
  `pipes.length/elev` = **m**, `pipes.dia` = **호칭경 mm**, 압력 = **Pa**.
  전개 결과는 m 이므로 노드 좌표만 되돌려 곱한다.
- `meta` 에 최소한 다음을 남긴다: 기준개수 K, 앵커 라벨, 최원 유하거리(m),
  설계면적 폭(m), corridor 총연장(m), max_load, 관경 근거 집계.

**수용 기준**
- 배관표 `in`/`out` 라벨이 전부 노드표에 존재한다(고아 참조 0).
- 노즐 수 == `len(worst["heads"])`.
- 배관 길이 합이 `worst["total_m"]` 과 ±1 % 이내(수직 전개분 제외한 평면 성분 기준).
- 표 첫 행이 급수원 인접 배관이고, 마지막 트리 행이 앵커 헤드로 끝난다.

### G6. SDF 방출

`design/emit.py`. 모듈 A 의 `remote30_prototype.py:emit_sdf` 를 이식한다. 템플릿 SDF 를
써야 Graphics 블록(아이소매트릭 표시 메타·schemes·Display-options)이 보존되고,
표준 SLF 가 옆에 있어야 PIPENET 이 호칭경↔내경을 lookup 한다.

- 자산 2종을 `design/assets/` 에 두고, 환경변수 override 를 지원한다
  (`REMOTE30_TEMPLATE_SDF`, `REMOTE30_STANDARD_SLF`).
- **자산이 없으면 경고가 아니라 실패로 처리한다.** 모듈 A 는 경고만 내고 진행하지만,
  G 는 사람이 그 자리에서 보고 있으므로 조용히 이상한 파일이 나가면 안 된다.
- SDF 와 SLF 는 **항상 한 쌍으로** 저장한다(같은 stem).

**수용 기준**
- 산출 SDF 가 PIPENET 에서 열리고, 관경이 "Unset" 으로 뜨지 않는다.
- 자산을 일부러 지우고 실행하면 파일을 만들지 않고 명확한 오류를 낸다.
- 노드·배관·노즐 개수가 G5 테이블과 일치한다.

### G7. 4번째 창 — 「수리계산 입력」

`ui/dialogs/dialog_design_input.py` 신설, `ui/controllers/cad_import_flow.py` 에 편입.
찍기 → 손질 → 변환 **다음**에 온다. G 는 복제본이므로 E 의 「소스 불변」 계약이 없다 —
흐름을 고쳐도 된다.

화면이 보여야 하는 것:
- 기준개수 K 입력(기본 30), 도면 장 선택(여러 장일 때만).
- 앵커 위치와 corridor 미리보기, 간선 굵기 = 담당 헤드 수.
- 요약: 최원 유하거리 · 설계면적 폭 · 총연장 · max_load.
- 관경 근거 집계(텍스트 n건 / 별표1 보강 n건 / 별표1 폴백 n건), 부속 집계, 미해결 건수.
- 저장 버튼 → `.sdf` + `.slf`.

무거운 계산(제한 전개·테이블)은 UI 스레드 밖에서 돌린다 — 기존 `_CadEditBuildThread`
패턴을 따른다.

**수용 기준**
- 급수 시작 위치 미지정, 헤드 종류 미지정 등은 **변환 창과 같은 방식으로** 막고
  사유를 화면에 돌려준다(조용히 실패하지 않는다).
- 창을 닫았다 다시 열어도 직전 K·선정 결과가 유지된다.
- 창이 떠 있는 동안 `.kfp` 저장 동작이 영향받지 않는다.

### G8. 대조 검증

같은 도면을 모듈 A(`/remote30-prototype`)로도 돌려 두 SDF 를 비교한다.

**수용 기준** — 다음 다섯이 맞으면 위상이 같은 것이다.
- 앵커 헤드 좌표(±SNAP), 선정 헤드 집합(K개), 최원 유하거리, max_load, corridor 총연장.
- 관경이 다른 간선은 **전부** 근거(`source`)로 설명된다. 설명되지 않는 차이는 버그다.
- 결과를 `docs/module_g_vs_a.md` 에 표로 남긴다.

---

## 3. 금지·보존 목록

- **`cad_project_editor/`(모듈 E)와 `routes/module_f/` 는 수정 금지.** 참조만 한다.
- 기존 `.kfp` 산출 경로(`convert_to_kfp` 의 저장·반환 규약, DTO 기본값) 변경 금지.
  전체망 `.kfp` 는 이번 작업 전후로 비트 동일해야 한다.
- 이식한 함수의 **로직 변경 금지**: 앵커 선정 3단계, 별표1 매핑표, 직류티 미계상 규칙,
  「헤드만 지우고 배관은 안 자른다」 원칙.
- `PipeTables` 키 이름·단위 변경 금지. 키 하나만 달라도 SDF 가 조용히 비거나 "Unset" 이 된다.
- G 트리 밖(사용자 바탕화면 등)에 중간 산출물을 쓰지 말 것. G 는 제 트리 기준으로 작업
  폴더를 잡으며, 이 분리가 E 와 동시 실행의 전제다.
- 전면 리포맷, 파일 이동/리네이밍, 공개 함수 시그니처 파괴 금지 — keyword 인자 추가만 허용.

---

## 4. 검증 체계

`cad_project_editor_g/tests/test_design_input.py` 신설.

- G1·G3·G4·G5 의 합성 단위 테스트(위 수용 기준 그대로).
- 통합: B1F 실측 도면으로 찍기 저장본 → 손질 → 제한 전개 → 테이블 → SDF 까지 무오류 완주.
- **회귀**: 같은 도면에서 전체망 `.kfp` 를 이번 작업 전/후로 생성해 diff 없음을 증명한다.
  이 회귀가 깨지면 어떤 항목도 완료가 아니다.
- 헤드리스 확인: Qt 없이 `design/` 전 모듈이 import·실행된다
  (`_smoke_headless.py` 와 같은 방식).
- 커밋 메시지 형식: `[G#] 제목 — 수용기준: PASS/FAIL(사유)`

---

## 5. 참조 구현 — 어디서 가져오는가

| 필요한 것 | 위치 | 처리 |
|---|---|---|
| 최불리 K 선정 | `routes/module_f/remote30.py:_worst_k_heads` | 그대로 이식 |
| 도면 장 나누기 | `routes/module_f/remote30.py:_sheet_frames` | 그대로 이식 |
| corridor 제한 | `routes/module_f/remote30.py:_restrict_to_worst` | 그대로 이식(시그니처 유지) |
| 관경 혼합 규칙 | `remote30_prototype.py:_pipe_diameter`, `_nfpc_min_bore_mm` | 그대로 이식 |
| 치수 텍스트 파싱 | `remote30_prototype.py:_extract_dia_text_points`, `_match_diameter_for_segment` | 이식 + 입력 어댑터 |
| 치수 텍스트 원본 | `cad_project_editor_g/services/cad_import/pipeline/handoff.py` → `world.texts` | 이미 있음 |
| 부속 판정 | `core/fitting_rules.py` | 그대로 이식 |
| 부속 등가길이 | `cad_project_editor_g/fittings_library_v3.json` | 이미 있음 |
| 테이블 스키마 | `remote30_prototype.py:PipeTables` | 키 규약 준수 |
| SDF 방출 | `remote30_prototype.py:emit_sdf` | 이식 + 자산 |
| 노드 역참조 | `services/cad_import/convert/planar.py` 내부 `node_id`, `remap` | 밖으로 노출 |

**코드를 옮겨 오되 로직은 고치지 않는다.** 모듈 A 의 `build_input_tables` 는 A 전용
선정 결과 타입(`SelectionResult`)에 묶여 있어 그대로 쓸 수 없다 — G5 에서 **어댑터를
새로 짠다**. 이번 작업에서 실제로 새로 짜는 것은 G5 와 G7 둘뿐이다.

---

## 6. 알려진 함정 (구현 전 숙지)

- **T1 좌표 추적** — 관경 매칭은 평면 mm, 테이블은 전개된 m. 전개로 한 간선이 여러 배관으로
  쪼개지므로 `edge_ref` 역참조가 없으면 관경이 엉뚱한 배관에 붙는다. G2 에서 반드시 해결할 것.
- **T2 두 개의 답** — `.kfp` 는 솔버가, `.sdf` 는 G 가 최불리를 고른다. D1 에 따라 두 결과는
  **다를 수 있다.** 버그가 아니다. 산출물 `meta` 와 화면에 「SDF 의 설계구역은 G 가 앵커
  방식으로 선정」이라고 남겨 혼선을 막는다.
- **T3 단위 혼재** — 같은 dict 안에서 mm 와 m 가 섞인다(§G5). 노드 좌표만 mm 다.
- **T4 담당 헤드 수의 정의** — 별표1 입력은 corridor 안의 K개 기준이다. 전체망 하류 수가 아니다.
- **T5 템플릿 자산** — 없으면 표시 메타가 빠지고 관경이 "Unset" 이 된다. G 는 실패로 처리한다.
- **T6 캐시 분리** — G 는 제 트리에서 작업 폴더를 잡는다. E 와 동시에 띄워도 캐시·찍은 스펙이
  섞이지 않아야 하므로 신규 산출 경로도 G 트리 안에 둔다.

---

## 7. 진행 규칙

- 항목 순서 준수, 항목 병합 금지.
- 애매점은 구현하지 말고 `BLOCKED.md` 에 (항목, 질문, 임시 우회 여부)로 기록 후 진행.
- 아래 두 가지는 아직 미정이다. 임의로 정하지 말고 만나는 시점에 `BLOCKED.md` 에 올릴 것.
  - 기준개수 K 의 기본값 노출 위치 — 매번 입력 vs 프로젝트에 저장.
  - corridor 의 뿌리를 손질 단계의 급수 시작 위치로 볼지, 알람밸브를 별도 기기 행으로 세울지.
- 최종 산출물: 변경 diff 요약, 테스트 결과 전문, `BLOCKED.md`,
  `docs/module_g_design_input.md`(설계 기록), `docs/module_g_vs_a.md`(G8 대조표).
