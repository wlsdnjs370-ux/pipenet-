# FNCADnet 모듈 C — 설계 자동화 워크벤치 개조 지시서 (구현판)

- 대상: 모듈 C (건축 레이어 정리 / Remote 30 워크벤치)
- 기준 문서: 「스프링클러 자동 설계 파이프라인 v6.1」
- 소스 기준: 모듈 C zip (`remote30_workbench.html` / `r30_inspect.py` / `pages.py` / `dxf_parse_progress.py` / `sprinkler_remote30_extractor.py` / `remote30_prototype.py`)

---

# 0. 착수 전 필수 확인

## 0.1 이 작업의 단 하나의 설계 기준

> **모듈 A는 "완성된 소방도면 → 배관망 그래프" 역방향이다.**
> **모듈 C 개조판은 "건축도면 → 배관망 그래프" 순방향이다.**
> **두 파이프라인의 산출 그래프는 동일한 자료구조여야 한다.**

같으면 기존 SDF/KFP/HAS 변환기와 모듈 B 검진기를 **한 줄도 안 고치고** 쓴다. 갈라지면 변환기와 검진기를 영구히 두 벌 유지해야 한다.

**PR-0(모듈 A 그래프 스키마 문서화)의 승인 없이 C5(라우팅)를 구현하지 마라.**

## 0.2 착수 전에 직접 확인할 명령

```bash
# 모듈 A 그래프 노드/엣지 실제 필드명
grep -n "elevation_m\|display_z_m\|edge_len\|_edge_load" core/remote30_graph.py remote30_prototype.py

# 담당 헤드 수 산출 (C570 이 동일 방식을 써야 함)
grep -n "compute_edge_load\|_downstream_heads\|prune_by_load" -A 20 remote30_prototype.py

# 부속류 위상 재구성 (C5 가 같은 규약을 따라야 함)
grep -n "Tee\|Elbow\|fitting" core/*.py | head -40

# 솔버 방출 체인
grep -n "emit_full_sdf\|_emit_kfp\|_emit_has\|round_trip_check" routes/r30_combined.py

# 부하 기준 비성립 조건
sed -n '1,200p' docs/load_extraction.md
```

이 결과를 PR-0 문서에 표로 정리해서 제출하라.

---

# 1. 선행 수정 — 기존 모듈 C의 실제 버그 3건

**설계 자동화 이전에 반드시 고쳐야 한다.** 지금 화면에서 조용히 사라지고 있는 것들이고, 하필 **건축도면에서 가장 중요한 요소들**이다.

## BUG-1. ARC 가 렌더되지 않는다

- 서버(`r30_inspect.py`)는 `{"t": "A", ...}` 를 방출한다.
- 클라이언트(`remote30_workbench.html:drawEntity`)는 `ent.t === "ARC"` 를 검사한다.
- **코드가 달라 조건이 영원히 거짓이다. ARC 는 한 번도 그려진 적이 없다.**

건축도면에서 **문(door)은 개폐 호(ARC)로 그린다.** 문이 안 보이면 검수자가 개구부를 확인할 수 없고, C160 가상 폐합 결과를 눈으로 검증할 수 없다.

```javascript
// remote30_workbench.html drawEntity() 수정
} else if (ent.t === "A") {            // "ARC" → "A"
  const [sx, sy] = worldToScreen(ent.c[0], ent.c[1]);
  const r = ent.r * z;
  if (r < 0.3) return;
  const sa = ent.a[0] * Math.PI / 180;
  const ea = ent.a[1] * Math.PI / 180;
  ctx.beginPath();
  ctx.arc(sx, sy, r, -ea, -sa, false);   // DXF CCW ↔ Canvas CW
  ctx.stroke();
}
```

## BUG-2. HATCH 가 렌더되지 않는다

서버는 `{"t":"H","p":[[x,y],...]}` 를 방출하지만 클라이언트에 분기가 없다. 건축도면에서 **벽 포셰(poché)·기둥 채움·바닥 마감**이 전부 HATCH 다.

```javascript
} else if (ent.t === "H") {
  if (!ent.p || ent.p.length < 3) return;
  ctx.beginPath();
  for (let i = 0; i < ent.p.length; i++) {
    const [sx, sy] = worldToScreen(ent.p[i][0], ent.p[i][1]);
    if (i === 0) ctx.moveTo(sx, sy); else ctx.lineTo(sx, sy);
  }
  ctx.closePath();
  const prev = ctx.globalAlpha;
  ctx.globalAlpha = prev * 0.25;   // 채움은 옅게 — 선 위에 덮이지 않게
  ctx.fill();
  ctx.globalAlpha = prev;
  ctx.stroke();
}
```

## BUG-3. SOLID / 3DFACE / TRACE 가 렌더되지 않는다

서버는 `{"t":"S","p":[[x,y],...]}` 를 방출하지만 분기가 없다. **기둥 채움**이 SOLID 인 도면이 흔하다.

```javascript
} else if (ent.t === "S") {
  if (!ent.p || ent.p.length < 3) return;
  ctx.beginPath();
  for (let i = 0; i < ent.p.length; i++) {
    const [sx, sy] = worldToScreen(ent.p[i][0], ent.p[i][1]);
    if (i === 0) ctx.moveTo(sx, sy); else ctx.lineTo(sx, sy);
  }
  ctx.closePath();
  ctx.fill();
}
```

## BUG 수정 검증 방법

건축 평면도 1장을 올리고 다음을 눈으로 확인하라.

| 확인 | 통과 기준 |
|---|---|
| 문 개폐 호 | 벽 간극마다 1/4원 호가 보인다 |
| 기둥 | 채워진 사각형이 격자로 보인다 |
| 벽 포셰 | 벽 두께 사이가 옅게 채워진다 |

**BUG-1~3 수정은 PR-1 이전에 단독 PR(PR-B)로 먼저 머지하라.** 원본 `remote30_workbench.html` 에도 반영한다(이건 모듈 A 워크벤치의 버그이기도 하므로).

---

# 2. 파일 레이아웃

```
core/design/
  __init__.py
  schema.py                  # dataclass 정의 전부 (Building/Room/Core/Constraints/Design)
  session.py                 # 세션 저장·단계 상태·게이트 플래그
  deterministic/
    __init__.py
    constraints.py           # ★ NFTC 값의 유일한 진실 출처
    nftc_tables.py           # 표 데이터만 (별표1, 표시온도, 보 이격, 기준개수)
    zoning.py                # C3 밸브·방호구역
    head_layout.py           # C4 헤드 배치
    pipe_routing.py          # C5 배관망
    emit_graph.py            # 모듈 A 호환 그래프 방출
  recognize/
    __init__.py
    geom_stats.py            # C130 기하 지문 수집
    arch_category.py         # C140 카테고리 판정
    wall_centerline.py       # C150 벽 중심선화
    opening_close.py         # C160 개구부 간극 폐합
    room_faces.py            # C170 폐합 영역 → 실 폴리곤
    room_label.py            # C180 실명 귀속
    core_detect.py           # C190 코어 판별
  checks/
    __init__.py
    dimensional.py           # 차원 무결성 검사
routes/
  r30_design.py              # 신규 API 전부 (register(app, ...) 패턴)
templates/
  design_workbench.html      # remote30_workbench.html 복제 후 개조
tests/design/
  test_constraints.py
  test_nftc_tables.py
  test_forbidden_patterns.py # 금칙어 회귀
  test_head_layout.py
  test_routing.py
  test_dimensional.py
  fixtures/                  # 실제 도면 기반 골든 파일
```

**기존 파일 수정 허용 범위**

| 파일 | 허용 |
|---|---|
| `templates/remote30_workbench.html` | BUG-1~3 수정만 |
| `routes/pages.py` | `GET /design-workbench` 라우트 1개 추가만 |
| `대조 서버.py` (앱 엔트리) | `r30_design.register(app, ...)` 호출 1줄 추가 |
| 그 외 전부 | **수정 금지** |

---

# 3. C1 — 건축 도면 인식 (인식 셸)

## 3.0 전체 흐름

```
DXF → [C110 파싱] → [C120 가시성] → [C130 기하 지문]
    → [C140 카테고리] → [C150 벽 중심선] → [C160 간극 폐합]
    → [C170 실 폴리곤] → [C180 실명 귀속] → [C190 코어]
    → building_draft.json
```

`C110`·`C120` 은 **기존 `r30_inspect.py` 로직을 그대로 재사용**한다. `_render_entity`, `_insert_matrix`, `_inspect_layer_visibility` 를 함수로 추출해 `core/design/recognize/` 에서 import 하되, **원본 라우트의 동작은 바뀌지 않아야 한다.**

## 3.1 C130 — 기하 지문 수집

레이어 이름 사전에 의존하지 않기 위한 통계 수집. 건축도면은 설계사마다 레이어 규약이 완전히 다르다 (`A-WALL` / `벽체` / `WALL-1` / `건축-벽` / 심지어 `0`).

```python
@dataclass
class LayerFingerprint:
    name: str
    n_entities: int
    type_hist: dict          # {"L": 1240, "A": 88, "H": 12, ...}
    len_median_mm: float     # LINE 길이 중앙값
    len_p90_mm: float
    parallel_pair_ratio: float   # 평행쌍을 이루는 LINE 비율
    offset_peaks_mm: list        # 평행쌍 오프셋 히스토그램 상위 3 peak
    closed_shape_count: int      # 폐합 LWPOLYLINE + 4-LINE 사각형
    closed_repeat_score: float   # 폐합도형 크기 분산의 역수 (반복성)
    arc_attach_ratio: float      # LINE 끝점에 ARC 가 붙은 비율
    text_numeric_ratio: float    # TEXT 중 숫자 전용 비율
    grid_alignment_score: float  # 중심 좌표가 등간격 격자에 정렬되는 정도
```

**평행쌍 판정 파라미터**

| 항목 | 값 | 근거 |
|---|---|---|
| 각도 허용 | ≤ 2.0° | CAD 스냅 오차 흡수 |
| 겹침 비율 | 짧은 쪽 길이의 ≥ 30% | 스쳐 지나가는 선 배제 |
| 오프셋 범위 | 50 ~ 500 mm | 국내 벽 두께 실측 범위 |
| 히스토그램 bin | 10 mm | |
| peak 채택 | 전체 평행쌍의 ≥ 8% 를 차지하는 bin | 노이즈 배제 |

**모든 파라미터는 `core/design/recognize/params.py` 로 빼라. 코드에 하드코딩 금지.**

## 3.2 C140 — 카테고리 판정

**신규 카테고리 12종** (기존 6종과 별개 체계):

`WALL` `DOOR` `WINDOW` `COLUMN` `STAIR` `SHAFT` `ROOM_TEXT` `DIM` `FURNITURE` `GRID` `BEAM` `OTHER`

```python
def arch_category(fp: LayerFingerprint, name_hint: str | None) -> tuple[str, float]:
    """returns (category, confidence 0~1). 지문이 1순위, 이름은 가산점."""
```

**판정 규칙 — 지문 1순위, 이름은 +0.15 가산**

| 카테고리 | 지문 조건 | 기본 신뢰도 |
|---|---|---|
| `WALL` | `parallel_pair_ratio ≥ 0.55` AND `len(offset_peaks) ≥ 1` AND `len_median ≥ 800mm` | 0.80 |
| `DOOR` | `arc_attach_ratio ≥ 0.35` AND ARC 반경 600~1200mm 비율 ≥ 0.5 | 0.75 |
| `WINDOW` | 평행 3~4선 반복 + 벽 중심선 위에 위치 | 0.65 |
| `COLUMN` | `closed_repeat_score ≥ 0.7` AND `grid_alignment_score ≥ 0.6` AND 폐합면적 0.1~4.0㎡ | 0.80 |
| `STAIR` | 등간격 평행 단선 ≥ 6개 다발, 간격 변동계수 ≤ 0.15 | 0.70 |
| `SHAFT` | 소형 폐합(면적 ≤ 6㎡) AND **다층 도면에서 같은 좌표 반복** | 0.85 (단층만이면 0.40) |
| `DIM` | `DIMENSION` 엔티티 존재, 또는 `text_numeric_ratio ≥ 0.8` | 0.85 |
| `ROOM_TEXT` | TEXT 이고 `text_numeric_ratio ≤ 0.3` AND 글자수 1~20 | 0.70 |
| `GRID` | 도면 전체를 가로지르는 초장선(len ≥ bbox 대각의 0.8) + 원형 심볼 | 0.75 |
| `FURNITURE` | 위 어디에도 안 걸리는 INSERT 밀집 | 0.40 |

**이름 힌트 사전** (가산점 전용, 판정 단독 근거 금지):

```python
NAME_HINTS = {
    "WALL":     ["WALL", "벽", "벽체", "A-WALL", "AR-WALL", "W-"],
    "DOOR":     ["DOOR", "문", "출입", "A-DOOR", "DR-"],
    "WINDOW":   ["WINDOW", "창", "창호", "A-GLAZ", "WIN"],
    "COLUMN":   ["COL", "COLUMN", "기둥", "S-COL", "PILLAR"],
    "STAIR":    ["STAIR", "계단", "STR", "A-STRS"],
    "SHAFT":    ["PD", "P.D", "AD", "A.D", "SHAFT", "덕트", "PS", "샤프트", "EPS", "TPS"],
    "ROOM_TEXT":["ROOM", "실명", "NAME", "TEXT", "A-ANNO"],
    "DIM":      ["DIM", "치수", "A-DIMS"],
    "FURNITURE":["FURN", "가구", "A-FURN", "집기"],
    "GRID":     ["GRID", "통심", "A-GRID", "AXIS"],
    "BEAM":     ["BEAM", "보", "GIRDER", "S-BEAM"],
}
```

> **`sprinkler_remote30_extractor.py` 의 `DEFAULT_*_LAYER_KEYWORDS` 를 수정하지 마라.** 모듈 A가 그 사전에 의존한다. 위 사전은 신규 파일에 독립적으로 둔다.

## 3.3 C150 — 벽 중심선화

```
입력: WALL 계열 LINE 집합
출력: centerlines[] = [{p1, p2, thickness_mm, source_pair: (id_a, id_b)}]
```

**알고리즘**

1. LINE 을 각도로 버킷팅 (5° 단위, 0~180°)
2. 같은 버킷 내에서 쌍 후보 생성 — 공간 인덱스(그리드 셀 = 최대 오프셋 500mm)로 근접만
3. 쌍 판정: 각도차 ≤ 2°, 수직 오프셋 `d` 가 `offset_peaks` 중 하나와 ±15mm 이내, 투영 겹침 ≥ 30%
4. 중심선 = 두 선의 중점을 잇는 선분, `thickness_mm = d`
5. **미짝 LINE 처리**: 짝을 못 찾은 WALL LINE 은 버리지 말고 `thickness_mm = None, unpaired = true` 로 보존. 조적벽 단선 표기 도면이 있다
6. 중심선 끝점 스냅: 공차 `snap_tol = max(30, thickness_mm * 0.3)` mm 로 군집

**실패 신호** — `parallel_pair_ratio < 0.3` 이면 이 도면은 벽이 단선 표기일 가능성. `wall_repr = "single_line"` 으로 표시하고 C160/C170 파라미터를 완화 모드로 전환.

## 3.4 C160 — 개구부 간극 가상 폐합 ★ 최대 위험 지점

**여기를 실패하면 두 실이 하나로 합쳐지고, 면적이 두 배가 되고, 용도 판정이 틀리고, 헤드 개수가 틀린다. 오류가 하류로 증폭되는 유일한 지점이다.**

```
입력: centerlines[], DOOR 계열 ARC/LINE
출력: virtual_edges[] = [{p1, p2, kind, evidence, confidence}]
      kind ∈ {"door", "opening", "inferred"}
```

**증거 3종 — 강한 순서대로**

| kind | 판정 | confidence |
|---|---|---|
| `door` | 간극 양끝 300mm 이내에 DOOR ARC 의 끝점이 있고, ARC 반경이 간극 폭의 0.8~1.3배 | 0.90 |
| `opening` | 간극 폭이 700~1800mm 이고 양끝이 **같은 직선(공선, 각도차 ≤2°)** 위 | 0.70 |
| `inferred` | 간극 폭이 200~3000mm 이고 양끝 중심선이 서로 다른 방향 (코너 미접합) | 0.45 |

**알고리즘**

```
for each centerline endpoint e:
    if e 가 다른 중심선과 이미 연결됨 (snap 군집 크기 ≥ 2): continue
    후보 = 반경 3000mm 이내의 미연결 endpoint 들
    for f in 후보:
        gap = dist(e, f)
        if gap < 100: continue          # 스냅으로 처리될 것
        kind, conf = classify_gap(e, f, doors, centerlines)
        if kind: virtual_edges.append(...)
```

**중복 방지** — 한 endpoint 는 가상 간선 1개만 갖는다. 후보가 여럿이면 `confidence` 최대, 동률이면 `gap` 최소.

**절대 규칙** — 가상 간선은 **반드시 `is_virtual: true` 로 표시**하고, 캔버스에서 **점선 + 다른 색**으로 렌더한다. 검수자가 여기를 우선적으로 봐야 한다.

## 3.5 C170 — 폐합 영역 → 실 폴리곤

중심선 + 가상 간선을 **평면 그래프**로 만들고 face 를 추출한다.

**알고리즘 (평면 face 추출)**

1. 모든 간선의 교차점을 계산해 간선을 분할 (`split_edge`) — 모듈 A의 `_split_tee_branches` 와 같은 개념
2. 각 노드에서 나가는 간선을 **각도 순으로 정렬**
3. half-edge 순회: 간선 `(u→v)` 로 들어왔으면, `v` 에서 `(v→u)` 의 **바로 다음 시계방향 간선**을 택한다
4. 순회가 시작 half-edge 로 돌아오면 face 하나 완성
5. 면적이 음수인 face(= 외곽) 1개는 버린다

**face 필터**

| 조건 | 처리 |
|---|---|
| 면적 < 1.0 ㎡ | 버림 (벽 두께 사이 틈 등) |
| 면적 > 전체 bbox 면적의 0.5 | 버림 (외곽 오검출) |
| 변 개수 > 200 | 플래그 (`suspicious_complexity`) 하고 보존 |
| 가상 간선 비율 > 0.5 | 플래그 (`mostly_virtual`) — 검수 우선순위 상위 |

**신뢰도 산출**

```
conf = 0.95
     - 0.25 * (가상 간선 길이 / 전체 둘레)
     - 0.15 * (unpaired 중심선 비율)
     - 0.10 * (변 개수 > 40 이면)
```

## 3.6 C180 — 실명 텍스트 귀속

1. `ROOM_TEXT` 계열 TEXT 중 폴리곤 내부에 있는 것 → 직접 귀속
2. 내부에 없으면 **지시선 추적**: 텍스트 위치에서 반경 2000mm 이내의 LINE 중, 한쪽 끝이 텍스트에 500mm 이내이고 다른 끝이 어떤 폴리곤 내부인 것
3. 한 폴리곤에 텍스트가 여럿이면 **면적 표기(숫자+㎡)를 제외**하고 나머지 중 가장 긴 것
4. 귀속 실패한 폴리곤은 `name: null, needs_input: true`

**용도 추정** — 실명 → NFTC 특정소방대상물 구분 매핑. **추정일 뿐이며 GATE 에서 사람이 확정한다.** 신뢰도 0.95 이상이어도 자동 확정 금지.

```python
USE_HINTS = {
    "업무시설": ["사무실", "사무", "OFFICE", "업무"],
    "공동주택": ["거실", "침실", "안방", "주방", "세대"],
    "판매시설": ["매장", "판매", "SHOP", "STORE", "마트"],
    "숙박시설": ["객실", "ROOM", "숙박"],
    "의료시설": ["병실", "진료", "수술", "처치"],
    "노유자시설": ["보육", "요양", "노인"],
    "주차장":   ["주차", "PARKING", "P.LOT"],
    "창고시설": ["창고", "WAREHOUSE", "저장"],
}
```

## 3.7 C190 — 코어 판별

**단층만으로는 확정하지 마라.** SHAFT 의 결정적 증거는 **여러 층에서 같은 좌표에 같은 크기로 반복**되는 것이다.

```
if 층 도면 ≥ 2:
    for 각 소형 폴리곤 p in 층 f:
        다른 층에서 중심 거리 ≤ 500mm, 면적비 0.7~1.4 인 폴리곤 개수 n
        if n ≥ (층수 - 1) * 0.6:  kind = SHAFT, conf = 0.85
else:
    이름 힌트만으로 후보 제시, conf = 0.40, GATE 확정 필수
```

---

# 4. GATE — 인간 확정 게이트

## 4.1 서버 강제

`/api/design/c2/*` 이후 **모든** 엔드포인트는 진입 시 세션의 `gate.passed` 를 확인한다.

```python
def require_gate(session):
    if not session.gate.passed:
        raise GateNotPassed(unresolved=session.gate.unresolved)

# 응답: HTTP 409
{ "ok": false, "code": "GATE_NOT_PASSED",
  "message": "인간 확정 게이트를 통과해야 진행할 수 있습니다.",
  "unresolved": ["R-1F-012.ceiling.has_finish", "R-1F-013.use"] }
```

UI 잠금은 **시각적 반영일 뿐**이다. 우회해도 서버가 막는다.

## 4.2 확정 항목 — 필수/선택

| 항목 | 필수 | 기본값 허용 | 비고 |
|---|---|---|---|
| 실 폴리곤 분할 | ✔ | ✘ | 병합/분할/삭제 |
| 실별 용도 | ✔ | ✘ | 기준개수·R·헤드종류를 전부 결정하는 단일 실패점 |
| **반자 유무** | ✔ | ✘ | **도면에 없다.** 헤드 종류·신축배관을 결정 |
| 반자고 | 반자 있을 때만 ✔ | ✘ | |
| 천장고(슬래브) | ✔ | 층고 | 기준개수 8m 판정의 대리 지표 |
| 최고 주위온도 | ✘ | 용도별 기본값 | |
| 특수 위험 | ✘ | `null` | 무대부/랙크식/특수가연물/EV충전 |
| 코어 확정 | ✔ | ✘ | 입상관 후보 |
| 장애물 상태 | ✔ | ✘ | `none`/`partial`/`complete` 중 택1 — "미확보" 명시도 확정이다 |
| 건물 층수 | ✔ | ✘ | 수원 20/40/60분 분기 |
| 구조(내화 여부) | ✔ | ✘ | R 2.1 vs 2.3 |

**설계 원칙 — 질의를 한 지점에 모은다.** 파이프라인 중간에 사람을 계속 부르면 자동화가 아니라 방해다. C1 완료 시 `GET /api/design/c1/gate_items/<sid>` 로 **결손 항목 전체를 한 번에 반환**하고, 한 화면에서 전부 채운 뒤 이후는 무인으로 끝까지 돈다.

## 4.3 결손 항목 응답 형식

```jsonc
{
  "ok": true,
  "total": 47,
  "groups": [
    { "field": "ceiling.has_finish", "label": "반자 유무", "required": true,
      "rooms": ["R-1F-001", "R-1F-002", ...], "count": 23,
      "bulk_apply": ["floor", "use"],       // 일괄 적용 가능 축
      "options": [true, false] },
    { "field": "use", "label": "용도", "required": true,
      "rooms": ["R-1F-012"], "count": 1,
      "suggestion": { "R-1F-012": {"value": "업무시설", "confidence": 0.62} },
      "options": ["업무시설", "공동주택", "판매시설", ...] }
  ]
}
```

---

# 5. C2 — 규범 조건 확정 (결정론 코어)

## 5.1 `constraints.py` 스켈레톤

```python
# core/design/deterministic/constraints.py
"""NFTC 수치의 유일한 진실 출처.

다른 어떤 모듈도 NFTC 수치를 하드코딩하지 않는다. 필요하면 여기서 import 한다.
Constraints 는 frozen 이며, 하류 단계가 쓰기를 시도하면 예외가 난다.
"""
from dataclasses import dataclass, field
from typing import Literal

@dataclass(frozen=True)
class Rule:
    code: str            # "RULE-NFTC-211-D"
    article: str         # "2.1.1.1"
    article_text: str    # 조문 원문 전체
    effective_date: str  # "2026-03-01"
    text_hash: str       # sha256 of article_text
    note: str = ""

@dataclass(frozen=True)
class Constraints:
    # ── 기준개수·수원 ──────────────────────────────
    scenario_head_count: int
    water_supply_m3: float
    discharge_minutes: int          # 20 / 40 / 60
    emergency_power_minutes: int    # 20 / 40 / 60

    # ── 헤드 ───────────────────────────────────────
    horizontal_distance_m: float    # R
    head_spacing_square_m: float    # S = √2 R
    wall_clearance_max_m: float     # S / 2
    temp_rating_c: int
    quick_response_required: bool
    k_factor: float
    flow_lpm_min: float             # 80
    pressure_mpa_min: float         # 0.1
    pressure_mpa_max: float         # 1.2
    head_clearance_radius_m: float  # 0.6 (벽은 0.1)
    head_to_wall_clearance_m: float # 0.1
    head_to_ceiling_max_m: float    # 0.3

    # ── 구역 ───────────────────────────────────────
    zone_area_max_m2: float                 # 3000
    zone_grid_relief_m2: float | None       # ★ None 고정 — §9.2 D1
    zone_floors_max: int                    # 1 (예외 3)
    spray_zone_heads_max: int               # 50 (개방형 방수구역)
    spray_zone_heads_min_when_split: int    # 25

    # ── 배관 ───────────────────────────────────────
    branch_heads_per_side_max: int          # 8 — 한쪽 기준. 전체는 16
    cross_main_min_dn: int                  # 40
    tournament_forbidden: bool              # True
    pipe_size_table: dict                   # 별표1
    velocity_limit_mps: dict                # 사내 기준 — 법정 아님

    # ── 표 ─────────────────────────────────────────
    beam_clearance_table: tuple
    head_exempt_places: tuple

    # ── 추적 ───────────────────────────────────────
    nftc_effective_date: str
    trace: tuple                            # tuple[FieldTrace, ...]


def build_constraints(building: "Building") -> Constraints:
    """building.json → Constraints. 이 함수 밖에서 NFTC 값을 만들지 마라."""
    ...
```

## 5.2 기준개수 결정표 (NFTC 103 표 2.1.1.1)

```python
def scenario_head_count(*, use: str, floors_total: int, is_underground_arcade: bool,
                        has_special_combustible: bool, head_mount_height_m: float,
                        is_apartment_unit: bool, is_connected_parking: bool) -> tuple[int, str]:
    """returns (개수, rule_code)"""
    if is_apartment_unit:
        return 10, "RULE-NFTC-211-APT"          # 아파트등의 세대 내
    if is_connected_parking:
        return 30, "RULE-NFTC-211-PARK"         # 각 동이 주차장으로 연결된 구조의 그 주차장 부분
    if floors_total >= 11 or is_underground_arcade:
        return 30, "RULE-NFTC-211-H"

    # 10층 이하
    if use in ("공장", "창고", "랙크식창고"):
        return (30, "RULE-NFTC-211-A") if has_special_combustible else (20, "RULE-NFTC-211-B")

    if use in ("근린생활시설", "판매시설", "운수시설", "복합건축물"):
        # ★ 30 이 되는 것은 판매시설 또는 판매시설이 설치되는 복합건축물뿐이다.
        #    근린생활시설 단독 · 운수시설 단독 → 20.
        if use == "판매시설" or (use == "복합건축물" and building_has_retail):
            return 30, "RULE-NFTC-211-C"
        return 20, "RULE-NFTC-211-D"

    # 그 밖의 것
    if head_mount_height_m >= 8.0:
        return 20, "RULE-NFTC-211-E"
    return 10, "RULE-NFTC-211-F"
```

## 5.3 수원·방사시간

```python
def water_supply(n: int, floors_total: int) -> tuple[float, int]:
    """returns (수원량 m³, 방사시간 분). 20분 고정 금지."""
    if floors_total >= 50:
        return n * 4.8, 60
    if floors_total >= 30:
        return n * 3.2, 40
    return n * 1.6, 20
```

**비상전원 시간도 같은 분기를 쓴다.**

## 5.4 수평거리 R (NFTC 103 2.7.3)

| 설치장소 | R (m) | rule_code |
|---|---|---|
| 무대부·특수가연물 저장취급 | 1.7 | `RULE-NFTC-273-A` |
| 랙크식창고 (특수가연물) | 1.7 | `RULE-NFTC-273-B` |
| 랙크식창고 (그 밖) | 2.5 | `RULE-NFTC-273-C` |
| 공동주택(아파트) 세대 내 거실 | 3.2 | `RULE-NFTC-273-D` |
| 일반 (비내화구조) | 2.1 | `RULE-NFTC-273-E` |
| 내화구조 | 2.3 | `RULE-NFTC-273-F` |

파생값:
```
S_square = R * math.sqrt(2)      # 정방형 헤드 간격
wall_clearance_max = S_square / 2
```

## 5.5 표시온도 (NFTC 103 표 2.7.6)

| 최고 주위온도 | 표시온도 |
|---|---|
| 39℃ 미만 | 79℃ 미만 |
| 39 이상 64 미만 | 79 이상 121 미만 |
| 64 이상 106 미만 | 121 이상 162 미만 |
| 106 이상 | 162 이상 |

## 5.6 별표1 관경표

```python
PIPE_SIZE_TABLE = {
    # 호칭경(mm): 담당 헤드 수 상한
    "가": {25: 2, 32: 3, 40: 5, 50: 10, 65: 30, 80: 60,
           90: 80, 100: 100, 125: 160, 150: 10**9},
    "나": {...},   # 폐쇄형, 무대부·특수가연물
    "다": {...},   # 개방형
}

def min_dn(head_count: int, column: str = "가") -> int:
    for dn in sorted(PIPE_SIZE_TABLE[column]):
        if head_count <= PIPE_SIZE_TABLE[column][dn]:
            return dn
    return 150
```

> **★ 착수 전 필수** — 위 값은 통용값이다. `nftc_tables.py` 에 넣기 전에 **현행 NFTC 원본 별표1과 1:1 대조**하라. 특히 '나' 칸 90mm 행이 '가' 칸과 다르다. 대조 결과를 PR-1 본문에 캡처로 첨부하라.

## 5.7 보 이격표 (NFTC 103 2.7.7.7)

```python
BEAM_CLEARANCE = (
    # (수평거리 상한 m, 수직거리 조건)
    (0.75, "below_beam_bottom"),   # 반사판이 보 하단보다 낮을 것
    (1.00, 0.10),                  # 수직거리 0.1m 미만
    (1.50, 0.15),
    (float("inf"), 0.30),
)
```

## 5.8 조번호 4-tuple 저장

조번호만 저장하면 개정 때마다 전수 재검토가 필요하다. NFTC 103은 최근에도 개정되었다(확인된 최신본 **2026.3.1 시행**).

```python
Rule(
    code="RULE-NFTC-211-D",
    article="2.1.1.1",
    article_text="...조문 원문 전체...",
    effective_date="2026-03-01",
    text_hash=hashlib.sha256(article_text.encode()).hexdigest(),
)
```

**개정 감지 테스트**를 함께 만든다 — `nftc_tables.py` 의 모든 Rule 에 대해 `sha256(article_text) == text_hash` 를 검증. 조문을 고치면 해시가 깨지고, 그때 **영향받는 필드 목록이 테스트 실패 메시지로 출력**된다.

---

# 6. C2B — 화재조기진압용(NFTC 103B) 분기

**활성 판정은 헤드 배치보다 반드시 앞선다.** 103B가 활성되면 수평거리 5종 룰 자체가 무효화되는데, 그 수평거리를 소비하는 단계가 헤드 배치다.

```python
def activate_103b(*, owner_adopts_esfr: bool,
                  site_conditions_ok: bool,
                  commodity_restricted: bool) -> bool:
    """랙크식 창고라고 자동으로 103B 가 되지 않는다.

    랙크식의 기본 트랙은 NFTC 103 2.7.2 (랙 높이 4m/6m마다 헤드).
    103B 는 별도 설비 종류이며 설치장소 구조 조건과 저장물 제한이 있다.
    천장고는 활성 조건이 아니라 활성 후 K값 결정 변수다.
    """
    return owner_adopts_esfr and site_conditions_ok and not commodity_restricted
```

`owner_adopts_esfr` 는 **사람이 입력**한다. 자동 True 금지.

`commodity_restricted` — 제4류 위험물, 타이어, 목재, 종이, 섬유류 등.

---

# 7. C3 — 유수검지장치 위치 + 방호구역 + 헤드 사양

## 7.1 C3.1 유수검지장치 위치 (수동 지정)

사용자 요구사항대로 **밸브는 수동 지정**이다. 시스템은 후보만 제시한다.

```
1. 후보 열거: building.cores[] 중 kind ∈ {SHAFT, STAIR} 이고 confirmed=true
2. 캔버스에서 후보를 하이라이트 (반투명 채움 + 외곽 굵은 선)
3. 사용자 클릭 → 밸브 노드 생성
4. 설치 요건 hard check 즉시 표시
```

**설치 요건 체크리스트**

| 항목 | 값 | 자동 판정 가능 |
|---|---|---|
| 설치 높이 | 0.8 ~ 1.5 m | ✘ — 사용자 입력 |
| 전용실 출입문 | 0.5 m × 1.0 m 이상 | ✘ |
| 실온 | 4℃ 이상 유지 | ✘ |
| 「유수검지장치실」 표지 | 유 | ✘ |
| 담당 구역 면적 | ≤ 3,000 ㎡ | ✔ |
| 담당 층수 | 1개 층 (예외 3) | ✔ |

자동 판정 불가 항목은 **체크박스로 사용자 확인**을 받고 `design.json` 에 기록한다.

## 7.2 C3.2 방호구역 분할

**밸브가 구역을 정의한다.** 밸브 확정 후 구역을 나눈다.

```python
def split_zones(rooms, valves, constraints) -> list[Zone]:
    """
    1. 각 실을 가장 가까운 밸브에 배정 (배관 경로 거리 기준, 직선거리 아님)
    2. 밸브별 담당 면적 합산
    3. 면적 초과 시 분할 — 단, 대부분의 건물은 층당 1구역으로 자명하게 끝난다.
       자명 케이스를 먼저 분기시켜라:
         if 층 면적 ≤ zone_area_max and 밸브 1개: return [단일 구역]
    4. 도달성 검증
    """
```

**도달성 검증** — 각 구역의 최원단 지점이 담당 밸브에서 **배관 경로로 도달 가능**한지. 벽·코어를 통과하지 못하는 경로가 있으면 실패.

```
실패 시: C3 안에서 재분할한다. C5 까지 내려가지 않는다.
UI: 도달 불가 구역을 빨간 해칭 + 원인 표시
```

## 7.3 C3.3 시스템 종류

```python
SystemType = Literal["습식", "건식", "준비작동식", "부압식", "일제살수식"]
```

**부압식을 빠뜨리지 마라.** 파이프라인 v4~v6 문서가 5종에서 부압식을 누락했었다.

| 조건 | 시스템 |
|---|---|
| 동결 우려 없음, 일반 | 습식 |
| 동결 우려, 수손 우려 없음 | 건식 |
| 동결 우려 + 수손 우려 | 준비작동식 |
| 수손 우려 최소화 요구 | 부압식 |
| 무대부·연소확대 우려 | 일제살수식 (개방형) |

## 7.4 C3.4 헤드 사양

`constraints` 를 **읽기만** 한다. R·표시온도를 여기서 재결정하지 마라.

```python
def head_spec(room: Room, c: Constraints) -> HeadSpec:
    orientation = "pendent" if room.ceiling.has_finish else "upright"
    use_flex    = room.ceiling.has_finish          # 신축배관은 반자 구간에만
    return HeadSpec(
        orientation=orientation,
        temp_rating_c=c.temp_rating_c,             # ★ constraints 에서만
        quick_response=c.quick_response_required,
        k_factor=c.k_factor,
        flex_pipe=FlexSpec(
            equivalent_length_m=22.4,              # 한백 표준 F사 유형
            dn=25, inner_dia_mm=28.0,
            c_factor=120.0, physical_length_m=0.7,
        ) if use_flex else None,
    )
```

## 7.5 C3.5 기준개수 순환 의존 절단

기준개수는 헤드 부착높이 8m로 갈리는데, 부착높이는 C4에서 확정된다. **순환이다.**

```
절단 방법: C2 에서 천장고(room.ceiling.slab_height_mm)를 부착높이의 대리 지표로 쓴다.
zone 내 혼재: 한 zone 안에 8m 이상·미만이 섞이면 보수적 값(큰 기준개수) 채택.
```

이 절단을 **코드 주석에 명시**하라. 나중에 "왜 천장고를 쓰지?" 라는 질문이 반드시 나온다.

---

# 8. C4 — 헤드 배치

## 8.1 격자 sweep

```python
def layout_heads(room: Room, c: Constraints, obstacles) -> list[Head]:
    R = c.horizontal_distance_m
    S = c.head_spacing_square_m         # √2 R
    best = None
    axes = principal_axes(room.polygon)  # 실의 주축 (최소 외접 사각형 방향)
    for theta in axes:                   # 보통 2개 (0°, 90°)
        for ox in frange(0, S, S / 8):   # 8 steps
            for oy in frange(0, S, S / 8):
                heads = place_grid(room, theta, ox, oy, S)
                heads = drop_outside(heads, room.polygon)
                if not covers_all(room.polygon, heads, R):
                    heads = add_fill_heads(room.polygon, heads, R)
                score = (len(heads), wall_clearance_variance(heads, room.polygon))
                if best is None or score < best[0]:
                    best = (score, heads)
    return best[1]
```

**sweep 규모** — 2 축 × 8 × 8 = **128 후보**. 실당 128회 피복 검사는 충분히 빠르다.

**피복 검사** — 실 폴리곤을 `R/4` 간격 격자로 샘플링하고, 모든 샘플점이 어떤 헤드로부터 `R` 이내인지 확인.

**목적함수** — `(헤드 수, 벽 이격 분산)`. 사전식 비교. 헤드 수가 같으면 벽 이격이 고른 쪽.

## 8.2 육각(지그재그) 배치 금지

최적 피복은 육각 배치지만 **쓰지 마라.** 헤드가 어긋나 흩어지면 C5의 가지배관이 직선으로 못 간다. 헤드 몇 개 아끼는 이득보다 배관 복잡도 손해가 크다.

**정방형·장방형 격자만 허용.** 장방형은 대각선 ≤ 2R 조건을 만족하는 (S₁, S₂) 조합.

## 8.3 열·행 인덱스 부여

```python
Head(id=..., x=..., y=..., row=int, col=int, branch_axis="x"|"y")
```

이 인덱스가 **C5 가지배관 축 결정의 근거**다. 빠뜨리면 C5가 헤드를 다시 군집화해야 한다.

## 8.4 살수장애 — 보를 60cm 룰에 넣지 마라

```python
# ✘ 이렇게 하지 마라 (기존 파이프라인 문서 v4의 오류)
for obstacle in (ducts, beams, lights, pipes):
    if dist(head, obstacle) < 0.6: FAIL

# ✔ 이렇게 하라
for obs in (ducts, lights, pipes):          # beams 제외
    if dist(head, obs) < c.head_clearance_radius_m:
        if not try_add_head_below(obs):     # 하부 헤드 추가를 먼저 시도
            fail("NFTC 2.7.7.1")

check_beam_clearance(head, beams, c.beam_clearance_table)   # 보는 별도 표
```

**이유** — 보는 수평거리 0.75m 미만에서도 반사판을 보 하단보다 낮추는 조건으로 허용된다. 즉 **60cm 원 안에 보가 들어오는 것이 정상**이다. 보를 60cm 룰에 넣으면 보가 있는 모든 실에서 대량 오탐이 나고 배치가 무한 재실행된다.

## 8.5 장애물 정보가 없을 때

```python
if building.obstacles.status == "none":
    # 살수장애 검사를 건너뛰되, 산출물에 플래그를 남긴다
    design.flags.append(ObstacleUnverified(rooms=[...]))
    # UI 에 경고 배너: "장애물 정보 미확보 — 살수장애 검증 미수행"
```

**조용히 통과시키지 마라.** 검증하지 않은 것과 검증해서 통과한 것은 다르다.

---

# 9. C5 — 배관망 라우팅

## 9.1 단계

| 부호 | 처리 |
|---|---|
| C510 | 헤드 열 군집화 → 가지배관 축 결정 (C4의 `row`/`col` 사용) |
| C520 | 가지배관 생성 + **분기점 기준 한쪽 8개** 상한 분할 |
| C530 | 교차배관 경로 생성 (직교/맨해튼 라우팅) |
| C540 | 주배관·입상관 → 유수검지장치 연결 |
| C550 | 신축배관 부착 (반자 조건부) + 등가길이 |
| C560 | **위상 그래프 확정** — 모듈 A 스키마로 방출 |
| C570 | 담당 헤드 수 누적(후위순회) → 별표1 최소 호칭경 |

## 9.2 라우팅 알고리즘 — C510 ~ C550

배관망 생성은 **구역(zone) 단위로 독립 실행**한다. 각 구역의 유수검지장치가 트리의 루트다. 구역이 여럿이면 구역 수만큼 독립 트리를 만들고, C540 에서 각 트리를 입상관으로 묶는다.

### C510 — 가지배관 축 결정

헤드에는 C4 가 부여한 `row` / `col` 인덱스가 있다. 이걸 그대로 쓴다.

```
1. 구역 내 헤드의 최소 외접 사각형(OBB)을 구한다.
2. 교차배관 축 = OBB 의 장변 방향
   가지배관 축 = OBB 의 단변 방향
   근거: 가지배관은 짧고 많아야 8개 제한에 여유가 생긴다.
3. 단, 유수검지장치가 구역의 단변 쪽에 있으면 축을 뒤집는다.
   교차배관은 밸브에서 곧게 뻗어야 주배관 우회가 없다.
4. 결정된 축을 heads[].branch_axis 에 기록 ("x" | "y").
```

**예외 — 복도형 구역** (OBB 장단변비 ≥ 4.0):
축 판정 없이 **장변 = 교차배관**으로 고정한다. 복도에서 가지배관이 장변을 따라가면 실 안으로 못 들어간다.

### C520 — 가지배관 생성

```
1. 가지배관 축에 수직인 헤드 라인(같은 row 또는 col)을 하나의 가지배관 후보로 묶는다.
   허용 오차: 축 좌표 편차 ≤ head_spacing / 4
2. 각 후보의 헤드를 축 좌표로 정렬한다.
3. 교차배관 예정선(C530 이 정할 직선)과의 교점 = 분기점(tee).
4. 분기점 기준 좌우 헤드 수를 센다.
       left  = 교차배관보다 작은 좌표의 헤드
       right = 큰 좌표의 헤드
5. if max(len(left), len(right)) > branch_heads_per_side_max:
       분할한다. 분할 전략은 아래 우선순위.
```

**분할 전략 (우선순위 순)**

| # | 방법 | 조건 |
|---|---|---|
| 1 | **분기점 이동** | 한쪽만 초과하고 반대쪽에 여유가 있을 때. tee 를 여유 쪽으로 옮겨 균형을 맞춘다 |
| 2 | **교차배관 추가** | 양쪽 다 초과. 교차배관을 하나 더 놓고 헤드 라인을 둘로 나눈다 |
| 3 | **가지배관 분리** | 위 둘로 안 되면 같은 라인을 독립 가지배관 2개로 쪼개고 각각 별도 tee |

**절대 하지 마라** — 8개마다 기계적으로 자르기. 한쪽 8 / 반대쪽 8 = **총 16개가 적법**하다.

### C530 — 교차배관 경로 생성

교차배관은 가지배관 분기점들을 잇는 선이다. 실 경계와 장애물을 피해야 한다.

```
1. 분기점들의 축 좌표 중앙값을 교차배관 기준선으로 잡는다.
2. 기준선이 벽/코어/설치제외 영역을 관통하는지 검사한다.
3. 관통하면 맨해튼 라우팅으로 우회한다 (아래).
4. 우회 후에도 어떤 분기점에 도달 불가면 그 분기점을 인접 교차배관으로 재배정한다.
5. 최소 호칭경 40mm 강제 (constraints.cross_main_min_dn).
```

**맨해튼 라우팅 (직교 격자 A\*)**

```python
def manhattan_route(p_from, p_to, obstacles, grid_mm=500):
    """축 평행 경로만 생성한다. 대각선 금지.

    비용 = 길이 + TURN_PENALTY * 방향전환수
    TURN_PENALTY 를 크게 잡아야 실무 도면처럼 곧은 경로가 나온다.
    """
    TURN_PENALTY = grid_mm * 6      # 방향 1회 전환 = 3m 우회와 등가
    # 4-이웃 A*, 상태 = (cell, incoming_direction)
    # 장애물 = 벽 중심선 buffer(thickness/2 + 50mm) ∪ 코어 ∪ 설치제외 영역
    ...
```

**격자 크기** — `grid_mm = 500` 기본. 실 폭이 2m 미만인 복도가 있으면 `250` 으로 낮춘다. 더 낮추면 탐색 비용이 급증한다.

**벽 통과 처리** — 배관은 실제로 슬리브로 벽을 관통한다. 벽을 절대 장애물로 두면 경로가 안 나온다.

```
벽 통과 비용 = WALL_PIERCE_COST (= grid_mm * 20)
코어·계단실·설치제외 영역 = 절대 장애물 (통과 불가)
```

즉 **벽은 비싸지만 통과 가능**, 코어는 통과 불가다. 이 구분이 없으면 라우팅이 실패하거나 비현실적인 경로가 나온다.

### C540 — 주배관 · 입상관 연결

```
1. 각 구역의 교차배관 시점(밸브에 가장 가까운 끝)을 구한다.
2. 시점 → 유수검지장치 노드까지 맨해튼 라우팅. 이 구간이 주배관.
3. 유수검지장치 → 입상관 접속점: 밸브가 속한 코어의 중심.
4. 입상관은 층간 수직 간선으로 생성한다.
       - 노드 표고 = 층 라벨 → elevation_m (모듈 A 와 동일 규약)
       - 표시 좌표(display_z_m)와 수리 표고(elevation_m)를 분리 보유
       - 배관장 = max(도면상 길이, |표고차|)
5. 최하층 입상관 종점 = 가압송수장치 접속점. 이 노드가 그래프의 최종 급수원.
```

**표시 기하와 수리 기하의 분리를 여기서 반드시 지켜라.** 모듈 A 가 계통도 부분망을 다룰 때 쓰는 규약과 동일해야 한다. 좌표에서 배관장을 역산하는 솔버(K-Fire)를 위해 **좌표 = 실표고인 별도 출력**도 생성해야 한다.

### C550 — 신축배관 부착

```
for head in heads:
    if head.room.ceiling.has_finish:      # 반자 있는 구간만
        가지배관 → 헤드 사이에 신축배관 간선 삽입
        equivalent_length_m = 22.4        # 한백 표준 F사 유형
        dn = 25, inner_dia_mm = 28.0, c_factor = 120.0
        physical_length_m = 0.7
    else:
        가지배관에 상향식 헤드 직결 (FX 없음)
```

**"헤드 1개당 FX 1개"는 조건부다.** 반자 유무에 따라 갈라진다. 지하주차장처럼 반자가 없는 구간은 강관 직결이므로 FX 를 붙이면 안 된다.

### C560 직전 — 부속류 위상 재구성

원본 표기 승계가 아니라 **위상에서 재구성**한다. 모듈 A 와 동일한 규약이다.

```
차수 3 분기점  → 방향벡터 최대각이 최소인 관로를 분기로 판정 → Tee(Branch)
차수 2 이면서 헤드·말단으로 가는 drop pipe → Elbow
```

순방향에서는 우리가 그래프를 직접 만들므로 분기 관계를 이미 알고 있지만, **판정 로직은 모듈 A 와 같은 함수를 써야 한다.** 다르면 같은 형상에 다른 부속류가 붙어 왕복 검증이 깨진다.

### 라우팅 실패 처리

| 실패 | 처리 |
|---|---|
| 분기점 도달 불가 | 인접 교차배관으로 재배정 → 그래도 실패면 `ROUTING_UNREACHABLE` 플래그 + 해당 헤드 목록 |
| 교차배관이 구역을 못 가로지름 | 구역 재분할 필요 → C3 회귀 신호 반환 |
| A\* 탐색 노드 수 초과 (> 200,000) | grid_mm 를 2배로 올려 1회 재시도 → 실패 시 `ROUTING_TIMEOUT` |

**실패를 조용히 삼키지 마라.** 도달 못한 헤드를 빼고 그래프를 완성하면, 수리계산은 통과하는데 실제로는 미방호 구역이 생긴다. 반드시 플래그를 남기고 UI 에 빨간 해칭으로 표시한다.

## 9.3 C520 — 토너먼트 방식 금지를 hard constraint 로

NFTC는 가지배관 배열이 토너먼트 방식이 아닐 것을 요구한다. **MST/Steiner 로 총연장을 최소화하면 대칭 분기(토너먼트형)로 수렴할 수 있다. 즉 최적화가 위법 형상을 만든다.**

```python
def check_tournament(graph, source) -> list[str]:
    """급수원에서 각 헤드까지의 경로에서 차수 3 이상 분기가
    연속 2회 이상 나타나면 토너먼트 의심."""
    violations = []
    for head in heads(graph):
        path = shortest_path(graph, source, head)
        streak = 0
        for node in path:
            if degree(graph, node) >= 3:
                streak += 1
                if streak >= 2:
                    violations.append(f"{head}: {node} 연속 분기")
                    break
            else:
                streak = 0
    return violations
```

**허용 형상은 빗살 구조뿐이다.** 교차배관 1개 + 그로부터 갈라지는 가지배관들.

**부수 효과가 좋다** — 빗살이 강제되면 배관망이 **트리로 고정**되어 담당 헤드 수 누적과 해석적 수리계산이 성립한다.

## 9.4 C520 — 8개는 한쪽 기준

```python
# ✘ 금지
if len(heads_on_branch) > 8: split()

# ✔ 올바름
BRANCH_HEADS_PER_SIDE_MAX = c.branch_heads_per_side_max   # 8
left  = heads_left_of(tee)
right = heads_right_of(tee)
if len(left) > BRANCH_HEADS_PER_SIDE_MAX or len(right) > BRANCH_HEADS_PER_SIDE_MAX:
    split()
# 가지배관 전체로는 최대 16개가 적법하다. 8 로 자르면 과분할된다.
```

변수명을 `branch_heads_max` 로 짓지 마라. **반드시 `per_side` 를 포함**하라. 금칙어 테스트가 이걸 검사한다.

## 9.5 C570 — 관경은 법정 최소부터

```
법정 최소(별표1) → 유속 검증 → 필요 시에만 상향
```

**"별표 + 1단계"를 기본값으로 넣지 마라.** 실무 관행이지 규정이 아니고, 무조건 상향하면 자재비가 과다해진다.

```python
def size_pipes(graph, c: Constraints):
    # 1. 후위순회로 담당 헤드 수 누적 (모듈 A 의 compute_edge_load 와 동일 방식)
    load = compute_edge_load(graph, source=valve_node)
    for edge in graph.edges:
        dn = min_dn(load[edge], column="가")           # 법정 최소
        v  = velocity(flow(load[edge]), dn)
        while v > velocity_limit(dn, c) and dn < 150:  # 유속 초과 시에만 상향
            dn = next_dn(dn)
            v = velocity(flow(load[edge]), dn)
        edge.dn = dn
    # 2. 상류 ≥ 하류 단조성 강제
    enforce_monotonic(graph, source=valve_node)
```

유속 상한(≤50A 6 m/s / ≥65A 10 m/s)은 **사내 기준이며 법정이 아니다.** `settings/velocity.yml` 로 분리하고 기본값에 `# 사내 기준, 법정 아님` 주석을 달아라.

## 9.6 C560 — 모듈 A 스키마 방출

**PR-0 문서가 정의한 스키마 그대로.** 특히 다음을 확인하고 맞춰라.

- 노드가 **표시 좌표**(`display_z_m`)와 **수리 표고**(`elevation_m`)를 분리 보유하는지
- 엣지 배관장이 `max(도면상 길이, |표고차|)` 로 보정되는지
- 부속류가 원본 표기 승계가 아니라 **위상에서 재구성**되는지 (차수3 분기점 → Tee(Branch), 차수2 drop pipe → Elbow)

**모듈 A/B 코드를 수정하지 마라.** 수정이 필요하다면 C560 의 스키마가 틀렸다는 뜻이다.

---

# 10. 데이터 스키마

## 10.1 building.json (GATE 산출물)

```jsonc
{
  "schema": "fncadnet.building/1",
  "session_id": "d3f1...",
  "source": {
    "dxf": "1F_평면도.dxf",
    "content_hash": "sha256:...",
    "floors": [{"label": "1F", "dxf": "1F_평면도.dxf", "origin": [0, 0]}]
  },
  "scale": { "unit_to_mm": 1.0, "wall_thickness_peaks_mm": [100, 150, 200],
             "wall_repr": "double_line" },
  "building": {
    "floors_total": 12,
    "floors_underground": 2,
    "structure": "내화구조",
    "is_underground_arcade": false
  },
  "rooms": [{
    "id": "R-1F-012",
    "floor": "1F",
    "polygon": [[x,y], ...],
    "area_m2": 82.4,
    "perimeter_m": 38.2,
    "virtual_edge_ratio": 0.08,
    "name": "사무실",
    "use": "업무시설",
    "ceiling": { "has_finish": true, "finish_height_mm": 2700, "slab_height_mm": 3200 },
    "ambient_temp_max_c": 30,
    "special_hazard": null,
    "head_exempt": false,
    "confidence": { "polygon": 0.91, "name": 0.88, "use": 0.62 },
    "provenance": { "polygon": "C170", "name": "C180", "use": "GATE" },
    "flags": []
  }],
  "virtual_edges": [
    { "p1": [x,y], "p2": [x,y], "kind": "door", "confidence": 0.90,
      "evidence": "ARC r=900 at 120mm", "rooms": ["R-1F-012", "R-1F-013"] }
  ],
  "cores": [{
    "id": "SH-01", "kind": "SHAFT",
    "polygon": [[x,y], ...], "area_m2": 2.1,
    "vertical_aligned_floors": ["B1","1F","2F","3F"],
    "confidence": 0.85, "confirmed": true
  }],
  "obstacles": {
    "status": "partial",
    "beams":  [{ "polyline": [...], "depth_mm": 600, "source": "구조도면" }],
    "ducts":  [],
    "lights": []
  },
  "gate": {
    "passed": true,
    "passed_at": "2026-08-03T14:20:00+09:00",
    "operator": "jinwon",
    "unresolved": [],
    "edits": [ { "op": "split", "room": "R-1F-007", "into": ["R-1F-007a","R-1F-007b"] } ]
  }
}
```

`gate.edits` 를 남기는 이유 — 사람이 어떤 실을 고쳤는지가 **C1 인식 셸의 학습·개선 데이터**가 된다. 반드시 기록하라.

## 10.2 constraints.json (C2 산출물, 읽기 전용)

```jsonc
{
  "schema": "fncadnet.constraints/1",
  "nftc_effective_date": "2026-03-01",
  "scenario_head_count": 20,
  "water_supply_m3": 32.0,
  "discharge_minutes": 20,
  "emergency_power_minutes": 20,
  "horizontal_distance_m": 2.3,
  "head_spacing_square_m": 3.253,
  "wall_clearance_max_m": 1.626,
  "temp_rating_c": 72,
  "quick_response_required": false,
  "k_factor": 80.0,
  "flow_lpm_min": 80,
  "pressure_mpa": { "min": 0.1, "max": 1.2 },
  "head_clearance_radius_m": 0.6,
  "head_to_wall_clearance_m": 0.1,
  "head_to_ceiling_max_m": 0.3,
  "zone_area_max_m2": 3000,
  "zone_grid_relief_m2": null,
  "zone_floors_max": 1,
  "spray_zone_heads_max": 50,
  "spray_zone_heads_min_when_split": 25,
  "branch_heads_per_side_max": 8,
  "cross_main_min_dn": 40,
  "tournament_forbidden": true,
  "pipe_size_table": { "가": {"25":2,"32":3,"40":5,"50":10,"65":30,
                              "80":60,"90":80,"100":100,"125":160,"150":1000000000} },
  "velocity_limit_mps": { "_note": "사내 기준 — 법정 아님",
                          "le_50a": 6.0, "ge_65a": 10.0 },
  "beam_clearance_table": [
    { "horizontal_lt_m": 0.75, "vertical": "below_beam_bottom" },
    { "horizontal_lt_m": 1.00, "vertical_lt_m": 0.10 },
    { "horizontal_lt_m": 1.50, "vertical_lt_m": 0.15 },
    { "horizontal_gte_m": 1.50, "vertical_lt_m": 0.30 }
  ],
  "trace": [
    { "field": "scenario_head_count", "code": "RULE-NFTC-211-E",
      "article": "2.1.1.1", "text_hash": "sha256:...",
      "effective_date": "2026-03-01" }
  ]
}
```

## 10.3 design.json (C3~C5 산출물)

```jsonc
{
  "schema": "fncadnet.design/1",
  "valves": [{
    "id": "AV-1F-01", "core_id": "SH-01", "floor": "1F",
    "point": [x, y], "system_type": "습식",
    "manual": true,
    "requirements_confirmed": {
      "mount_height_ok": true, "door_size_ok": true,
      "room_temp_ok": true, "signage_ok": true
    }
  }],
  "zones": [{
    "id": "Z-1F-01", "valve_id": "AV-1F-01", "floor": "1F",
    "rooms": ["R-1F-001", "..."], "area_m2": 1240.5,
    "reachability": "ok",
    "scenario_head_count": 20,
    "scenario_head_count_source": "conservative_max"
  }],
  "heads": [{
    "id": "H-1F-0001", "zone_id": "Z-1F-01", "room_id": "R-1F-012",
    "x": 12340.0, "y": 8820.0,
    "row": 2, "col": 5, "branch_axis": "x",
    "orientation": "pendent", "temp_rating_c": 72,
    "k_factor": 80.0, "quick_response": false,
    "flex": { "equivalent_length_m": 22.4, "dn": 25 }
  }],
  "graph": { "...": "모듈 A 스키마 — PR-0 문서 참조" },
  "flags": [
    { "code": "OBSTACLE_UNVERIFIED", "rooms": ["R-1F-012"],
      "message": "장애물 정보 미확보 — 살수장애 검증 미수행" }
  ]
}
```

---

# 11. API 계약

전부 신규. 기존 `/api/remote30/*` 는 건드리지 않는다. `routes/r30_design.py` 에 `register(app, *, UPLOAD_DIR, _save_upload, DESIGN_SESSION_DIR)` 패턴으로 구현.

```
GET  /design-workbench                        신규 페이지

POST /api/design/session                      → {"session_id": "..."}
GET  /api/design/session/<sid>                현재 단계·산출물 상태

POST /api/design/c1/recognize                 DXF 업로드 → NDJSON
GET  /api/design/c1/gate_items/<sid>          결손 항목 (§4.3)
POST /api/design/gate/confirm                 GATE 통과. unresolved 있으면 422

POST /api/design/c2/constraints               → constraints.json
POST /api/design/c2b/esfr                     103B 활성 (인간 승인 필수)

POST /api/design/c3/valves                    밸브 좌표 배열 + 요건 확인
POST /api/design/c3/zones                     구역 분할 + 도달성 검증
POST /api/design/c4/heads                     헤드 배치 → NDJSON (실별 진행률)
POST /api/design/c5/route                     라우팅 → 그래프

POST /api/design/emit                         모듈 A 호환 그래프 방출
                                              → 기존 SDF/KFP/HAS 변환기 체이닝
GET  /api/design/checks/<sid>                 차원 무결성 검사 결과
```

## 11.1 `/api/design/c1/recognize` NDJSON 스펙

기존 inspect 스트림 형식을 **그대로 유지**하고 메시지 타입만 추가한다. 클라이언트 파서를 재사용할 수 있다.

```jsonc
{"type":"parse","stage":"read","done":12345678,"total":141000000}
{"type":"phase","phase":"render","total":284000}
{"type":"progress","phase":"foreground","entities":[...],"done":2000,"total":284000,"bbox":{...}}
{"type":"phase","phase":"recognize"}
{"type":"fingerprint","layers":[{"name":"A-WALL","category":"WALL","confidence":0.86,...}]}
{"type":"centerlines","edges":[{"p1":[x,y],"p2":[x,y],"thickness_mm":150}]}
{"type":"virtual_edges","edges":[{"p1":[x,y],"p2":[x,y],"kind":"door","confidence":0.90}]}
{"type":"rooms","rooms":[{"id":"R-1F-012","polygon":[[x,y],...],"area_m2":82.4,...}]}
{"type":"cores","cores":[...]}
{"type":"result","ok":true,"dxf_token":"...","bbox":{...},"counts":{...},
 "gate_items_url":"/api/design/c1/gate_items/<sid>"}
```

## 11.2 캐시 키

기존 inspect 는 `f"{INSPECT_CACHE_VERSION}_{content_hash}"` 를 쓴다. 인식 결과는 **파라미터에 따라 달라지므로** 키에 포함한다.

```python
cache_key = f"{DESIGN_CACHE_VERSION}_{content_hash}_{sha256(json.dumps(recognize_params, sort_keys=True))[:12]}"
```

`params.py` 를 고치면 캐시가 자동으로 미스된다. 이게 없으면 파라미터를 튜닝해도 옛 결과가 나온다.

---

# 12. UI 명세

`templates/remote30_workbench.html` 을 `templates/design_workbench.html` 로 **복제**한 뒤 개조. 원본은 BUG-1~3 수정 외에 손대지 않는다.

## 12.1 재사용 (수정 금지)

`render()` / `drawEntity()` / `worldToScreen()` / `screenToWorld()` / `zoomAtPoint()` / `zoomCenter()` / `fitToBBox()` / `resizeCanvas()` / NDJSON 스트림 소비 루프 / 휠·키보드 바인딩.

## 12.2 DOM ID 명명 규칙

기존은 `wb-*` 접두사. 신규는 **`dw-*`** 를 쓴다.

| ID | 역할 |
|---|---|
| `dw-canvas` | 캔버스 |
| `dw-dxf` | DXF 파일 입력 |
| `dw-stepper` | 좌측 단계 스테퍼 |
| `dw-layer-list` | 건축 12종 토글 |
| `dw-design-list` | 설계 산출물 토글 |
| `dw-gate-panel` | 결손 항목 폼 |
| `dw-room-props` | 실 속성 편집 패널 |
| `dw-tool-select` / `dw-tool-merge` / `dw-tool-split` / `dw-tool-delete` | 실 편집 도구 |
| `dw-tool-valve` | 밸브 지정 모드 |
| `dw-run-c2` ~ `dw-run-c5` | 단계 실행 버튼 |
| `dw-emit-btn` | 솔버 파일 방출 |
| `dw-checks` | 차원 무결성 결과 |

## 12.3 state 확장

```javascript
const state = {
  // ── 기존 유지 ──
  entities: [], bbox: null, layers: [], layerState: {},
  view: {zoom:1, panX:0, panY:0}, dpr: 1,
  dxfFile: null, dxfToken: null, drag: null, fitZoom: null,

  // ── 신규 ──
  sessionId: null,
  stage: "c1",                    // c1 | gate | c2 | c3 | c4 | c5 | emit
  gatePassed: false,
  tool: "pan",                    // pan | select | merge | split | delete | valve
  centerlines: [],
  virtualEdges: [],
  rooms: [],                      // {id, polygon, area_m2, name, use, ceiling, confidence, ...}
  cores: [],
  selection: new Set(),           // 선택된 room id
  design: { valves: [], zones: [], heads: [], graph: null },
  overlayVisible: {
    rooms: true, roomLabels: true, virtualEdges: true, cores: true,
    obstacles: true, valves: true, zones: true, heads: true, pipes: true,
  },
};
```

## 12.4 렌더 순서 (뒤 → 앞)

```javascript
const RENDER_ORDER = [
  "GRID", "DIM", "FURNITURE",          // 건축 배경 — alpha 0.20
  "WALL", "COLUMN", "STAIR", "WINDOW", // 건축 주요 — alpha 0.45
  "DOOR",                              // alpha 0.6
  "rooms",                             // 실 폴리곤 반투명 채움 — alpha 0.12
  "virtualEdges",                      // ★ 점선 + 경고색
  "cores",                             // 반투명 채움 + 굵은 외곽
  "zones",                             // 구역 경계 굵은 선
  "obstacles",                         // 보/덕트 — 해칭
  "pipes",                             // 배관
  "heads",                             // 헤드 심볼
  "valves",                            // 밸브 심볼
  "roomLabels",                        // 텍스트는 항상 최상단
];
```

## 12.5 색상

```javascript
const ARCH_COLORS = {
  WALL: "#94a3b8", DOOR: "#fbbf24", WINDOW: "#67e8f9", COLUMN: "#a1a1aa",
  STAIR: "#c4b5fd", SHAFT: "#f472b6", ROOM_TEXT: "#22c55e", DIM: "#52525b",
  FURNITURE: "#3f3f46", GRID: "#3f3f46", BEAM: "#fb923c", OTHER: "#71717a",
};
const DESIGN_COLORS = {
  room: "#38bdf8", virtualEdge: "#f87171",   // ★ 경고색 — 검수 우선
  core: "#f472b6", zone: "#a78bfa", obstacle: "#fb923c",
  pipeBranch: "#3b82f6", pipeCross: "#2563eb", pipeMain: "#1d4ed8",
  head: "#ef4444", valve: "#facc15",
};
```

**가상 폐합선은 반드시 경고색 + 점선.** 여기가 오류의 최대 발생원이므로 검수자가 즉시 알아봐야 한다.

```javascript
ctx.setLineDash([6, 4]);
ctx.strokeStyle = DESIGN_COLORS.virtualEdge;
ctx.lineWidth = 2.0;
// ... 그린 뒤
ctx.setLineDash([]);
```

## 12.6 실 편집 도구

| 도구 | 조작 | 서버 반영 |
|---|---|---|
| `select` | 클릭 → 우측 속성 패널 | 없음 (로컬) |
| `merge` | 2개 이상 선택 → 버튼 | `POST /api/design/gate/confirm` 시 `edits` 로 |
| `split` | 실 위에 선 긋기 (2점) | 동일 |
| `delete` | 선택 → Delete 키 | 동일 |

**히트 테스트** — point-in-polygon (ray casting). 겹치면 면적 작은 쪽 우선.

```javascript
canvas.addEventListener("click", (e) => {
  if (state.tool === "pan") return;
  const rect = canvas.getBoundingClientRect();
  const [wx, wy] = screenToWorld(e.clientX - rect.left, e.clientY - rect.top);
  const hit = state.rooms
    .filter(r => pointInPolygon([wx, wy], r.polygon))
    .sort((a, b) => a.area_m2 - b.area_m2)[0];
  if (!hit) return;
  if (state.tool === "select") { showRoomProps(hit); }
  else if (e.shiftKey) { state.selection.add(hit.id); }
  else { state.selection.clear(); state.selection.add(hit.id); }
  render();
});
```

## 12.7 밸브 지정 모드

Remote 30 의 알람밸브 클릭과 같은 조작.

```
1. dw-tool-valve 클릭 → state.tool = "valve", 커서 crosshair
2. cores[] 를 하이라이트 (외곽 3px + 채움 alpha 0.3)
3. 클릭 → 가장 가까운 core 에 스냅 (반경 2000mm 이내면 core 중심으로)
4. POST /api/design/c3/valves → 요건 체크리스트 모달
5. 확인 후 방호구역 자동 분할 → 도달 불가 구역은 빨간 해칭
```

## 12.8 결손 항목 폼

표 형식. 행 = 실, 열 = 용도 / 반자 / 반자고 / 천장고 / 주위온도 / 특수위험.

**일괄 적용** — "같은 층 전체", "같은 용도 전체", "선택한 실 전체" 3종. 이게 없으면 실이 200개인 도면에서 사용자가 포기한다.

**정렬** — 신뢰도 낮은 순. 검수 우선순위가 곧 정렬 순서다.

**진행 표시** — `47건 중 12건 남음`. 0이 되면 `dw-run-c2` 활성.

---

# 13. 테스트 명세

## 13.1 금칙어 회귀 (`test_forbidden_patterns.py`)

**코드와 문서 양쪽**을 검사한다. CI 필수.

```python
FORBIDDEN = [
    (r"×\s*20\s*분",                        "G1 방사시간 20분 고정"),
    (r"discharge_minutes\s*=\s*20\b",       "G1 하드코딩"),
    (r"water_supply.*\*\s*1\.6(?!\s*if)",   "G1 층수 분기 없는 1.6 고정"),
    (r"activate_103[bB]\s*=\s*True",        "G2 103B 자동 활성"),
    (r'rack_storage["\']?\s*:\s*.*103B',    "G2 랙크식 자동 분기"),
    (r"branch_heads_max",                   "G3 per_side 누락 변수명"),
    (r"beams.*0\.6|0\.6.*beams",            "G4 보를 60cm 룰에 포함"),
    (r"별표\s*\+\s*1|min_dn\(.*\)\s*\+\s*1","G5 관경 자동 상향"),
    (r"zone_grid_relief_m2\s*[:=]\s*3700",  "D1 결정 전 구현"),
    (r"tau_water.*(FAIL|재설계|redesign)",   "D2 결정 전 구현"),
]
SCAN_PATHS = ["core/design/", "routes/r30_design.py", "templates/design_workbench.html"]
```

## 13.2 `constraints` 단위 테스트 — 구체 케이스

| # | 입력 | 기대 | 검증 포인트 |
|---|---|---|---|
| T1 | 근린생활시설 단독, 8층, 부착 4m | **20** | ★ 30 이 아니다 |
| T2 | 운수시설 단독, 5층 | **20** | ★ |
| T3 | 판매시설, 6층 | 30 | |
| T4 | 복합건축물(판매시설 포함), 9층 | 30 | |
| T5 | 복합건축물(판매시설 없음), 9층 | **20** | ★ |
| T6 | 업무시설, 7층, 부착 7.5m | 10 | 8m 미만 |
| T7 | 업무시설, 7층, 부착 8.0m | 20 | 경계 |
| T8 | 업무시설, 11층 | 30 | 11층 이상 |
| T9 | 아파트 세대 내 | 10, R=3.2 | |
| T10 | 공장, 특수가연물, 5층 | 30 | |
| T11 | 공장, 일반, 5층 | 20 | |

**수원 경계**

| # | 기준개수 | 층수 | 기대 수원 | 방사시간 |
|---|---|---|---|---|
| W1 | 20 | 29 | 32.0 ㎥ | 20분 |
| W2 | 20 | 30 | 64.0 ㎥ | 40분 |
| W3 | 20 | 49 | 64.0 ㎥ | 40분 |
| W4 | 20 | 50 | 96.0 ㎥ | 60분 |
| W5 | 30 | 55 | 144.0 ㎥ | 60분 |

**관경 경계 (별표1 '가')**

| 담당 헤드 수 | 기대 호칭경 |
|---|---|
| 2 | 25 |
| 3 | 32 |
| 4 | 40 |
| 5 | 40 |
| 6 | 50 |
| 10 | 50 |
| 11 | 65 |
| 30 | 65 |
| 31 | 80 |
| 161 | 150 |

**보 이격 경계**

| 수평거리 | 기대 |
|---|---|
| 0.74 m | `below_beam_bottom` |
| 0.75 m | 0.10 |
| 0.99 m | 0.10 |
| 1.00 m | 0.15 |
| 1.49 m | 0.15 |
| 1.50 m | 0.30 |

## 13.3 라우팅 테스트

| # | 시나리오 | 기대 |
|---|---|---|
| R1 | 한쪽 8개 / 반대쪽 8개 = 총 16 | **통과** (분할 없음) |
| R2 | 한쪽 9개 | 분할 발생 |
| R3 | 대칭 이분 분기 그래프 | `check_tournament` FAIL |
| R4 | 빗살 구조 | 통과 |
| R5 | 관경 상류 65A / 하류 80A | `enforce_monotonic` 이 하류를 65A 로 |
| R6 | 담당 4개 구간 | 40A (별표 최소, +1단계 아님) |

## 13.4 헤드 배치 테스트

| # | 시나리오 | 기대 |
|---|---|---|
| H1 | 10m × 10m 정사각 실, R=2.3 | 피복 100%, 헤드 수 ≤ 12 |
| H2 | 폭 2m 복도 20m, R=2.3 | 벽 이격 ≤ S/2 |
| H3 | 3㎡ 구획실 | 헤드 **1개** (면적 무관 최소 1) |
| H4 | 보가 헤드에서 0.5m | `below_beam_bottom` 요구, 60cm FAIL 아님 |
| H5 | 덕트가 헤드에서 0.4m | 하부 헤드 추가 시도 → 실패 시에만 FAIL |
| H6 | 장애물 status=none | FAIL 없음 + `OBSTACLE_UNVERIFIED` 플래그 |

## 13.5 차원 무결성 검사 (`checks/dimensional.py`)

**게이트만으로는 부족하다.** 모델이 건물과 대응하는지 보는 검사. 전부 값싸다.

| 검사 | 판정 | 우선순위 |
|---|---|---|
| 헤드 총수 vs 방호면적 ÷ (S×L) 이론값 | ±10% 이탈 시 플래그 | 중 |
| 실 폴리곤 면적 합 vs 건축개요 연면적 | 불일치율 > 5% 플래그 | 상 |
| 마스킹(설치제외) 면적 비율 | 연면적의 30% 초과 시 플래그 | 중 |
| 배관 총연장 ÷ 헤드 수 | 과거 도면 회귀 범위 이탈 시 플래그 | 하 |
| **기준층 교차검증** | 같은 평면인데 층별 헤드 수가 다르면 플래그 | **최상** |

**기준층 교차검증을 먼저 구현하라.** 정답 없이 **자기모순만으로 오류를 검출**한다. 아파트·오피스텔은 기준층이 반복되므로 공짜 검증기다.

```python
def check_typical_floor_consistency(design, building):
    groups = group_floors_by_polygon_similarity(building)   # IoU ≥ 0.95
    for g in groups:
        counts = [head_count(design, f) for f in g]
        if len(set(counts)) > 1:
            yield Flag("TYPICAL_FLOOR_MISMATCH", floors=g, counts=counts)
```

## 13.6 왕복 검증

```
C560 그래프 → SDF 방출 → 재파싱 → 그래프 비교
```
노드 수 · 엣지 수 · 총 배관장 · 담당 헤드 수가 일치해야 한다. 기존 `round_trip_check` 를 재사용하라.

---

# 14. 금지 사항

## 14.1 절대 금지 (위반 시 PR 반려)

| # | 금지 | 이유 |
|---|---|---|
| G1 | 수원·방사시간·비상전원 **20분 고정** | 층수 분기 필수. 초고층에서 수원 최대 1/3 과소 |
| G2 | 랙크식 창고 → **103B 자동 활성** | 기본 트랙은 NFTC 103 2.7.2. 저장물 제한 존재 |
| G3 | 가지배관 8개를 **전체 기준**으로 자르기 | 분기점 기준 한쪽. 전체는 16개 |
| G4 | 60cm 살수공간 검사에 **보(BEAM) 포함** | 2.7.7.7 별도 표. 대량 오탐 |
| G5 | 관경 **"별표 + 1단계" 자동 상향** | 관행이지 규정 아님. 자재비 과다 |
| G6 | 근린생활·운수 단독에 **기준개수 30** | 판매시설·판매시설 있는 복합건축물만 30 |
| G7 | `sprinkler_remote30_extractor.py` **키워드 수정** | 모듈 A 회귀 |
| G8 | **모듈 A/B 코드 수정** | 필요하면 C560 스키마가 틀린 것 |
| G9 | 신뢰도 미달 시 **기본값으로 채워 진행** | 조용한 실패 |
| G10 | 결정론 코어에 **임계값·확률·휴리스틱** | 인식 셸 소관 |
| G11 | NFTC 수치를 `constraints.py` **밖에** 작성 | 단일 진실 출처 파괴 |
| G12 | 육각(지그재그) 헤드 배치 | 배관 복잡도 폭증 |

## 14.2 결정 대기 — 결론 전까지 구현 금지

**D1. 격자형 배관방식 3,700㎡ 완화**
`zone_grid_relief_m2` 를 `None` 으로 고정. C5가 트리 라우팅만 하므로 격자형이 성립하지 않고, 따라서 면적 완화의 근거가 없다. **루프 라우팅 모드를 신설할지, 완화를 쓰지 않을지 결정된 뒤에 구현한다.**

**D2. τ_water(수원고갈시간)를 재설계 강제 트리거로 사용**
법정 수원은 **최소 방수량 80 LPM 기준**으로 산정되므로, 실제 방수량이 크면 τ가 짧아지는 것은 규정 구조상 당연하며 위법이 아니다. hard 판정으로 쓰면 **법정보다 과도한 수원을 시스템이 강제**한다. **참고 지표로만 계산하고 재설계 트리거에서 제외.**

## 14.3 착수 전 확인

| # | 항목 | 조치 |
|---|---|---|
| V1 | 별표1 관경표 | 현행 NFTC 원본과 1:1 대조. 캡처 첨부 |
| V2 | 조번호 전수 | 2026.3.1 시행본 대조. 4-tuple 저장 |
| V3 | 신축배관 등가길이 | 한백 표준 F사 유형 확인. 호칭경별 표로 |
| V4 | Hazen-Williams C값 | 관종·설비종류별 표 (습식 흑관 / 건식·준비작동 / CPVC) |
| V5 | 유속 상한 6/10 m/s | 법정 아님. 설정 파일 분리 + 주석 |

---

# 15. PR 분할

**순서가 중요하다. 앞의 것 없이 뒤를 만들면 만든 것이 맞는지 알 방법이 없다.**

| PR | 내용 | 파일 | 완료 기준 |
|---|---|---|---|
| **PR-B** | BUG-1~3 렌더 수정 | `remote30_workbench.html` | 문 호·기둥·해칭이 화면에 보임 (캡처 첨부) |
| **PR-0** | 모듈 A 그래프 스키마 문서화 | `docs/module_a_graph_schema.md` | 노드/엣지 필드 표 + `elevation_m`↔`display_z_m` 분리 확인. **승인 필수** |
| **PR-1** | `constraints` + 표 + 테스트 | `nftc_tables.py`, `constraints.py`, `test_constraints.py`, `test_forbidden_patterns.py` | T1~T11·W1~W5·관경/보 경계 전부 통과, 금칙어 0건, V1 캡처 첨부 |
| **PR-2** | 세션·게이트 골격 | `session.py`, `r30_design.py` (세션·게이트만) | 게이트 미통과 시 C2 이후 전부 409 |
| **PR-3** | 페이지 + 캔버스 골격 | `design_workbench.html` | 기존 렌더러 재사용 확인, 건축 12종 토글 동작 |
| **PR-4** | C1 인식 셸 | `recognize/*` | 실 도면 N장 벤치마크 리포트 (실 폴리곤 정답률) |
| **PR-5** | GATE UI | `design_workbench.html` (편집 도구·결손 폼) | 결손 0건으로 통과 가능. `gate.edits` 기록 |
| **PR-6** | C2/C2B 연결 | `r30_design.py` | `constraints.json` 생성, 103B 인간 승인 |
| **PR-7** | C3 밸브·구역 | `zoning.py` | 밸브 수동 지정, 도달성 검증 동작 |
| **PR-8** | C4 헤드 배치 | `head_layout.py` | H1~H6 통과, 차원 무결성 검사 첨부 |
| **PR-9** | C5 라우팅 + 방출 | `pipe_routing.py`, `emit_graph.py` | R1~R6 통과, 왕복 검증 통과 |
| **PR-10** | 하류 체이닝 | `r30_design.py` (emit) | 모듈 A/B **무수정** 확인. SDF/KFP/HAS 3종 생성 |

**PR-1 을 앞에 두는 이유** — `constraints` 가 모든 하류의 입력이다. 이게 틀리면 뒤의 모든 테스트가 틀린 기준으로 통과한다.

**PR-4(C1)를 뒤로 미룬 이유** — 가장 어렵고 가장 나중이어도 된다. 그전까지는 사람이 실 경계를 지정하는 반자동으로 버틸 수 있다. **대체 불가능한 것부터 만든다.**

---

# 16. 에이전트 작업 규칙

1. **PR-0 승인 전에 C5(PR-9)를 구현하지 마라.**
2. 기존 파일 수정 시 수정 이유를 커밋 메시지에 명시. `sprinkler_remote30_extractor.py` / 모듈 A / 모듈 B 는 수정 금지.
3. **구현과 명세가 어긋나면 명세를 고치지 말고 보고하라.** 코드 주석에 `[문서정합]` 태그로 불일치를 표기하는 기존 관행을 따른다.
4. NFTC 수치는 **`constraints.py` / `nftc_tables.py` 에만** 쓴다. 다른 파일에서 필요하면 import.
5. 임계값을 새로 도입할 때는 **근거를 함께 남겨라.** 근거가 없으면 `params.py` 로 빼고 `# 미검증` 주석.
6. §14.2 결정 대기 2건은 **구현하지 말고 자리만 비운다.** `None` 또는 `NotImplementedError`.
7. 각 PR에 **차원 무결성 검사 결과**를 첨부.
8. 벤치마크는 **실제 프로젝트 도면**으로. 합성 도면 통과 보고는 받지 않는다.
9. 인식 셸 결과에는 항상 `confidence` 와 `provenance` 를 붙인다. 근거 없는 값 금지.
10. **에러를 삼키지 마라.** `except Exception: pass` 는 기존 코드의 렌더 경로에만 허용된다(부분 렌더가 전체 실패보다 낫기 때문). 설계 로직에서는 금지.

---

# 부록 A. 달성하는 것과 달성하지 못하는 것

**달성한다** — 규범 조건 확정, 헤드 사양 선정, 격자 배치, 제약 검사, 담당 헤드 수 누적, 법정 관경 산정, 트리 라우팅, 솔버 입력 생성, 왕복 검증, trace 기록.

이것들은 가능한 정도가 아니라 **사람보다 낫다.** 신축배관 등가길이 누락률 0%, 표 조회 오류 0%, 기준구역 누락 0%. 전부 사람이 지루해서 빠뜨리는 항목이다.

**달성하지 못한다** — 도면에서 실과 용도를 신뢰성 있게 읽는 일, 반자·장애물을 아는 일, **시공 가능성을 판단하는 일**.

앞의 둘은 정보가 도면 밖에 있어서이고, 마지막은 **판정기 자체가 없어서**다. 배관 경로를 하나 생성했을 때 수리적으로 되는지는 계산으로 알 수 있지만, 그 경로에 덕트가 지나가는지·행거를 걸 구조체가 있는지·시공팀이 실제로 그렇게 시공하는지는 **어떤 계산으로도 답이 나오지 않는다.**

그래서 목표는 완전 무인이 아니라 **초안 자동 생성 + 자동 검증 + 사람은 시공성 판단**이다. 이 역할 분담은 누가 정한 것이 아니라 **정보의 소재를 따라 저절로 갈린다.**

# 부록 B. 인식 셸과 결정론 코어의 테스트 전략 차이

| | 인식 셸 (`recognize/`) | 결정론 코어 (`deterministic/`) |
|---|---|---|
| 검증 방법 | 벤치마크 도면셋, 통계 | 단위 테스트, 100% 커버리지 |
| 성능 표현 | "도면 N장에서 정답률 X%" | "통과 / 실패" |
| 도면이 바뀌면 | **깨진다** | 안 깨진다 |
| 투자 회수 | 계속 재투자 | **한 번 맞으면 영구** |
| 실패 양상 | 조용히 틀린다 | 시끄럽게 틀린다 |

**두 구역을 한 파이프라인으로 뭉치면 전체 신뢰도가 인식 셸 수준으로 하향 평준화된다.** 디렉터리를 나눈 이유가 이것이고, 결정론 코어에 임계값·확률·휴리스틱을 넣지 말라는 규칙(G10)의 이유도 이것이다.

---

# 부록 C. 운영 명세

## C.1 성능 예산

기존 모듈 C 는 141MB 도면에서 파싱만 36초가 걸린다(주석에 실측 기록). 인식 단계가 그 위에 얹히므로 예산을 미리 정한다.

| 도면 규모 | 단계 | 목표 (cold) | 목표 (캐시 hit) |
|---|---|---|---|
| ~5 MB (단위세대·소규모) | C1 전체 | ≤ 8 초 | ≤ 1 초 |
| ~30 MB (일반 층 평면) | C1 전체 | ≤ 35 초 | ≤ 3 초 |
| ~140 MB (대형 지하층) | C1 전체 | ≤ 180 초 | ≤ 10 초 |
| 실 200개 구역 | C4 헤드 배치 | ≤ 20 초 | — |
| 헤드 600개 구역 | C5 라우팅 | ≤ 30 초 | — |

**첫 페인트 목표는 3초다.** 기존 `FIRST_FLUSH_N = 2000` 을 유지하고, 인식 결과(`rooms`/`virtual_edges`)는 **엔티티 렌더가 끝난 뒤 별도 메시지로** 흘린다. 인식을 기다리느라 도면이 안 보이면 안 된다.

**예산 초과 시 대응**

| 단계 | 대응 |
|---|---|
| C130 지문 수집 | 레이어당 샘플링 — 엔티티 20,000개 초과 시 무작위 20,000개만 |
| C170 face 추출 | 간선 50,000개 초과 시 bbox 사분면으로 분할 처리 후 병합 |
| C4 배치 sweep | 실 면적 ≥ 500㎡ 이면 sweep step 을 8 → 4 로 |
| C5 A\* | §9.2 라우팅 실패 처리 참조 |

**측정을 코드에 심어라.** 각 단계 종료 시 소요 시간을 `design.timings` 에 기록하고 `/api/design/session/<sid>` 응답에 포함한다. 예산 초과가 나면 어디서 났는지 즉시 알 수 있어야 한다.

## C.2 세션 저장과 동시성

```
DESIGN_SESSION_DIR/
  <session_id>/
    meta.json           # 단계 상태, 타임스탬프, operator
    building.json       # GATE 산출물
    constraints.json    # C2 산출물 (읽기 전용)
    design.json         # C3~C5 산출물
    graph.json          # C560 방출 그래프
    checks.json         # 차원 무결성 결과
    timings.json
    audit.ndjson        # 감사 로그 (append-only)
```

**규칙**

1. `session_id` 는 UUID4. 추측 가능한 값 금지.
2. 각 파일은 **원자적 쓰기** — `.tmp` 로 쓰고 `os.replace`. 기존 inspect 캐시가 쓰는 방식과 동일하다.
3. `constraints.json` 은 **한 번 쓰면 불변**이다. 재생성이 필요하면 새 파일(`constraints.v2.json`)을 만들고 `meta.json` 이 현행을 가리킨다. 덮어쓰지 마라 — 어떤 제약으로 설계했는지가 감사 대상이다.
4. **낙관적 잠금** — 각 산출물에 `version` 필드를 두고, 갱신 요청은 `if_version` 을 보낸다. 불일치면 `409` + 현재 내용 반환.
5. 세션 만료 — 30일. 만료 전 `meta.json` 에 `expires_at` 기록. 자동 삭제는 하지 말고 목록에서만 숨긴다.

**동시 편집** — 같은 세션을 두 탭에서 열 수 있다. GATE 편집은 낙관적 잠금으로 충돌을 감지하고, C2 이후 실행은 **단계별 뮤텍스**로 직렬화한다(같은 세션에서 C4 가 두 번 동시에 돌면 안 된다).

## C.3 감사 로그 (`audit.ndjson`)

**설계 도서는 날인 문서다.** 어떤 값이 어디서 왔는지 재구성할 수 없으면 감리 대응이 불가능하다.

```jsonc
{"ts":"2026-08-03T14:02:11+09:00","actor":"system","stage":"C1","event":"recognize_done",
 "detail":{"rooms":47,"virtual_edges":31,"low_conf_rooms":8}}
{"ts":"...","actor":"jinwon","stage":"GATE","event":"room_split",
 "detail":{"room":"R-1F-007","into":["R-1F-007a","R-1F-007b"],"reason":"문 간극 미폐합"}}
{"ts":"...","actor":"jinwon","stage":"GATE","event":"use_override",
 "detail":{"room":"R-1F-012","from":"업무시설","to":"판매시설","suggested_conf":0.62}}
{"ts":"...","actor":"system","stage":"C2","event":"constraints_built",
 "detail":{"scenario_head_count":30,"rule":"RULE-NFTC-211-C","water_supply_m3":48.0}}
{"ts":"...","actor":"jinwon","stage":"C3","event":"valve_placed",
 "detail":{"id":"AV-1F-01","core":"SH-01","point":[12340,8820],"manual":true}}
{"ts":"...","actor":"system","stage":"C5","event":"routing_flag",
 "detail":{"code":"ROUTING_UNREACHABLE","heads":["H-1F-0231"]}}
```

**필수 기록 이벤트**

| 이벤트 | 이유 |
|---|---|
| 사람이 인식 결과를 고친 것 전부 | **C1 개선의 학습 데이터**. 어떤 실을 왜 고쳤는지가 다음 버전의 파라미터를 정한다 |
| 용도 override (제안값 ≠ 확정값) | 기준개수가 여기서 갈린다. 감리 질의 1순위 |
| 밸브 배치와 요건 확인 결과 | 자동 판정 불가 항목을 사람이 확인했다는 증거 |
| `constraints` 생성 시 적용된 rule code 전부 | trace 의 실체 |
| 모든 플래그 발생 | 검증하지 않은 것과 통과한 것의 구분 |

**append-only.** 수정·삭제 API 를 만들지 마라.

## C.4 오류 처리 규약

기존 코드의 `except Exception: pass` 는 **렌더 경로에만** 허용된다 — 엔티티 하나가 깨져도 나머지는 그려야 하기 때문이다. 그 관행을 설계 로직으로 가져오지 마라.

| 구역 | 규약 |
|---|---|
| 렌더·파싱 | 개별 엔티티 실패는 `dropped_types` 에 집계하고 계속. 기존 동작 유지 |
| 인식 셸 | 실패 시 해당 산출물을 `null` + `confidence: 0` + `error` 필드로. 조용히 빈 배열 반환 금지 |
| 결정론 코어 | **예외를 던진다.** 기본값으로 대체 금지 |
| API | 4xx/5xx + `{"ok":false,"code":...,"message":...}`. 스트리밍 중 실패는 `{"type":"error",...}` 라인 |

**금지 패턴**

```python
# ✘ 조용한 실패
try:
    n = scenario_head_count(...)
except Exception:
    n = 20            # ← 절대 금지. 기준개수를 추측하면 안 된다

# ✔
n = scenario_head_count(...)   # 실패하면 던진다
```

## C.5 배포·롤백

1. 신규 라우트는 **feature flag** 뒤에 둔다 — `DESIGN_WORKBENCH_ENABLED` 환경변수. 기본 off.
2. `/remote30-workbench` 는 **끝까지 그대로 동작해야 한다.** 모듈 A 사용자가 영향받으면 안 된다.
3. PR-B(렌더 버그 수정)만 예외적으로 기존 화면에 즉시 반영한다. 이건 **버그 수정이지 기능 변경이 아니다.**
4. 롤백 기준 — 기존 워크벤치의 inspect 응답 시간이 20% 이상 느려지면 즉시 되돌린다.

---
