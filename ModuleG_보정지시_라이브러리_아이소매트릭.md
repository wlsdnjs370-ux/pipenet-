# Module G 보정 지시서 — 배관 규격 라이브러리 바인딩과 아이소매트릭 좌표

당신은 FNCADnet 저장소의 유지보수 엔지니어다. `ModuleG_수리계산입력_작업지시서.md` 의
G1~G8 로 만든 SDF 방출 경로에 **두 가지 결함**이 실측으로 확인되었다. 아래를
**항목 순서대로, 항목당 커밋 1개씩** 수행하라. 각 항목은 수용 기준을 통과한 뒤에만
다음으로 진행한다. 애매점은 구현하지 말고 `BLOCKED.md` 에 기록하고 넘어간다(§5).

작업 대상은 **`cad_project_editor_g/`** 다. `cad_project_editor/`(모듈 E)와
`routes/module_f/`, `remote30_prototype.py`, `pipenet_converter/` 는 **읽기 전용 참조**다.

---

## 0. 증상과 원인 (구현 전 필독 — 이미 조사 완료)

산출물 `B1F 현장조사 소화설비 평면도_수리계산입력.sdf` 를 PIPENET Spray Module 에서
연 결과다. Node 92 · Pipe 61 · Nozzle 30 으로 **위상 자체는 정상**이다. 잘못된 것은
직렬화 두 곳뿐이다.

### 증상 1 — Pipe 표의 `Type` 이 전부 "None defined"

호칭경(`Diameter`)과 길이(`Length`)는 정상인데 배관 규격만 비어 있다.

- **원인**: `pipenet_converter/src/pipenet_converter/sdf_writer.py:93-100` 의
  `_build_links_element` 는 `<Pipe-set>` 을 **두 개만** 만든다 — 빈 placeholder 하나와
  파이프를 담는 것 하나. **`<Pipe-type>` 을 전혀 쓰지 않는다.** `Pipe.material` 필드는
  받아만 두고 직렬화에서 버려진다.
- 테이블에는 값이 있다: `design/tables.py:161` 이 `"type": "KSD 3507"` 을 채우고,
  `design/emit.py` 가 `material=` 로 넘긴다. **표에서 죽는 게 아니라 XML 에서 죽는다.**
- **모듈 A 는 어떻게 하나**: `write_sdf` 를 부른 **뒤에 XML 후처리**로 6종
  `<Pipe-type>` 을 주입하고 Pipe-set 을 재구성한다
  (`remote30_prototype.py:7083-7175`). G 의 `emit.py` 는 `write_sdf` 만 부르고
  이 후처리를 하지 않는다. **이것이 유일한 차이다.**
- **자산은 정상이다.** `assets/2. Pipenet_hand_FX28.slf` 의 `Item-name` 에
  `KSD 3507 · KSD 3562 · KSD 3576 · DP · CPVC2 · FX` 6종이 모두 있다. SLF 를 고칠
  일은 없다 — SDF 쪽 바인딩만 만들면 된다.

### 증상 2 — 망이 캔버스 한 점에 뭉쳐서 그려짐

- **원인**: `design/emit.py:tables_to_network` 는 노드 좌표를 `x/1000.0` **(mm→m)**
  로만 바꾼다. B1F 실측 도면은 한 변이 수백 m 라 좌표가 ±100 대 실수가 되고,
  PIPENET 스키매틱 캔버스 단위에서는 사실상 한 점이다.
- **모듈 A 는 어떻게 하나**: `remote30_prototype.py:6935-6948` 에서
  **bbox 중심 → (0,0)**, **가장 긴 축 → 약 3000 unit** 으로 정규화한 뒤 넘긴다.
- 또한 A 는 표시 전용 z(`display_z`)를 **같은 배율로** 정규화해
  `<Position z=..>` 로 실어 보낸다. G 는 이 값을 아예 넘기지 않는다.

### 증상 3 — 함께 고칠 것: `io-node` 값이 규약 밖

`sdf_writer._build_node_element` 는 `metadata["io_node"]` 가 없으면 `node_type` 을
그대로 쓴다. G 는 `"base"`/`"input"` 을 넣는데 PIPENET 규약은 **`"No"` / `"Input"`**
이다(레퍼런스 SDF 전부 이 표기). 지금은 관대하게 열리지만 규약 위반이다.

### 이 보정이 건드리지 않는 것

좌표·표시 z 는 **표시 전용**이다. 수리계산은 `length` · `elevation` · `rise` 로
하므로 **아래 작업 후에도 계산 결과는 비트 단위로 같아야 한다**(§4 회귀).

---

## 1. 신규 계약

`cad_project_editor_g/services/cad_import/design/sdf_post.py` 를 신설한다. 화면 없음.

```python
SCHEDULE_DEFS: list[tuple[str, str, list[tuple[float, int]]]]   # (name, c_factor, [(size_m, max_vel)])

normalize_node_coords(tables, *, canvas_units=3000.0) -> float
    # bbox 중심 → (0,0), 가장 긴 축 → canvas_units. 반환값은 적용된 배율.
    # tables.nodes 의 x,y 를 in-place 로 바꾸고 display_z 가 있으면 같은 배율을 곱한다.

bake_isometric(tables, *, iso_z_scale=1.0, ref_label=None, no_lift_labels=None) -> None
    # 30° 등각투영을 x,y 에 in-place 로 굽는다. elevation 은 건드리지 않는다.

inject_pipe_types(sdf_path, sched_by_pipe: dict[str, str]) -> None
    # 방출된 SDF 를 다시 읽어 Pipe-set 을 schedule 별로 재구성하고 <Pipe-type> 을 주입.
```

`emit_design_sdf` 의 시그니처에 `iso: bool = False`, `iso_z_scale: float = 1.0`,
`canvas_units: float = 3000.0` 를 **keyword 인자로만** 추가한다. 기본값으로 부르면
지금과 같은 호출부가 그대로 동작해야 한다.

---

## 2. 작업 항목 (순서 고정)

### G9. 배관 규격(Pipe-type) 주입 ★증상 1의 해결

`sdf_post.inject_pipe_types` 를 구현하고 `emit_design_sdf` 의 `write_sdf` **직후**에
호출한다. 모듈 A 의 `remote30_prototype.py:7083-7175` 를 그대로 옮긴다.

1. **6종 스케줄 정의**를 상수로 둔다. 이름은 SLF 의 `Item-name` 과
   **철자·공백까지 동일**해야 PIPENET 이 Pipe-type ↔ Schedule(내경)을 바인딩한다.
   `KSD 3507`(공백 있음) · `KSD 3562` · `KSD 3576` · `DP` · `CPVC2`(숫자 2 포함) · `FX`.
   호칭경 집합과 velocity 컨벤션(≤50 mm = 6, ≥65 mm = 10)은 A 의 정의를 그대로 쓴다.
2. **Pipe-set 재구성**. `tables.pipes[].type` 값별로 파이프를 나눠, 그 schedule 의
   `<Pipe-type>` 을 첫 자식으로 가진 Pipe-set 에 담는다.
3. **빈 placeholder Pipe-set 을 맨 앞에 유지한다.** PIPENET 은 `<Links>` 의 첫
   Pipe-set 을 blank/default 슬롯으로 예약한다 — 이게 없으면 우리 Pipe-type 이 그
   슬롯에 흡수돼 관경이 "Unset" 이 된다(`sdf_writer.py:96-98` 주석, 레퍼런스 SDF 3종에서 확인).
4. **쓰이지 않는 나머지 스케줄**도 `<Pipe-type>` 만 있고 `<Pipe>` 는 없는 Pipe-set 으로
   정의해 둔다 → PIPENET UI 의 schedule 드롭다운에 노출되어, 사용자가 표에서 배관을
   골라 관종을 바꿀 수 있다. **이것이 «라이브러리를 가져오게 한다»의 실제 내용이다.**
5. 최종 `<Links>` 순서: `[빈 placeholder] + [쓰인 schedule Pipe-set 들] + [빈 정의 Pipe-set 들]`
   그 다음 `<Nozzle>` · `<Valve>`.

**수용 기준**
- 산출 SDF 를 PIPENET 에서 열면 Pipe 표의 `Type` 열이 전부 `KSD 3507` 로 채워지고
  "None defined" 가 **0건**이다.
- `Diameter` 열이 종전 값 그대로다(65/25/150/40 …). 관경이 "Unset" 으로 뜨지 않는다.
- 배관을 하나 골라 드롭다운을 열면 6종이 모두 선택지로 보인다.
- 합성 테스트: `tables.pipes` 에 `type` 이 두 종류(예: KSD 3507 · CPVC2) 섞여 있으면
  Pipe-set 이 그만큼 분리되고, 각 파이프가 제 schedule 쪽에 들어간다.

### G10. 관종 선택을 표까지 잇기

지금 `tables.py:161` 은 전량 `KSD 3507` 하드코딩이다. 이것을 **인자로 받는다.**

- `build_design_tables(..., default_schedule="KSD 3507", schedule_by_pipe=None)`
  형태로 확장한다. `schedule_by_pipe` 는 `{pipe_id: schedule_name}` 이며 없으면 전량 기본값.
- 유효성 검사: `SCHEDULE_DEFS` 에 없는 이름이 오면 **조용히 기본값으로 떨어지지 말고**
  오류로 세우고 로그에 남긴다. 오타 하나로 다시 "None defined" 가 되는 것을 막는다.
- 배관별 관종을 어디서 정할지는 아직 미정이다(§5). 이 항목에서는 **전체 기본값 선택**
  까지만 화면에 노출한다.

**수용 기준**
- 기본값을 `CPVC2` 로 바꿔 방출하면 PIPENET 의 `Type` 열이 전부 CPVC2 로 뜨고,
  같은 호칭경에 대해 내경이 KSD 3507 과 다르게 잡힌다(SLF 바인딩이 실제로 걸렸다는 증거).
- 없는 스케줄 이름을 주면 파일을 만들지 않고 오류를 낸다.

### G11. 좌표 정규화 ★증상 2의 해결

`sdf_post.normalize_node_coords` 를 구현하고 `tables_to_network` **앞에서** 부른다.
모듈 A(`remote30_prototype.py:6935-6948`)와 같은 규칙이다.

```
cx, cy = bbox 중심
scale  = canvas_units / max(폭, 높이)        # canvas_units 기본 3000
x' = (x - cx) * scale ,  y' = (y - cy) * scale
display_z 가 있으면 display_z' = display_z * scale   # ★x,y 와 같은 배율이어야 비례가 맞는다
```

- `tables_to_network` 의 `/1000.0` **(mm→m)은 제거한다.** 정규화가 이미 스케일을
  흡수하므로 두 번 나누면 다시 뭉친다. `elevation`(m)·`length`(m)·`rise`(m)는
  **절대 건드리지 않는다.**
- `io_node` 를 함께 고친다(증상 3): `metadata={"io_node": "Input" if … else "No"}` 로
  넘긴다. `node_type` 문자열을 그대로 흘려보내지 않는다.

**수용 기준**
- 산출 SDF 의 `<Position>` 좌표 폭이 약 3000 unit 이고, PIPENET 캔버스에서
  망이 화면을 채운다(한 점 뭉침 해소).
- `<Node io-node=…>` 값이 `Input` 또는 `No` 뿐이다. `base` 가 **0건**.
- **회귀**: 정규화 전/후 SDF 의 `length` · `rise` · `elevation` · `bore` 값이
  **완전히 동일**하다. 좌표만 달라야 한다.

### G12. 아이소매트릭 베이크와 조절

`sdf_post.bake_isometric` 을 구현한다. 공식은 `routes/r30_combined.py:_bake_isometric_node_coords`
와 **동일해야 한다** — 다른 공식을 쓰면 같은 망이 SDF·KFP·HAS 에서 다르게 보인다.

```
COS30, SIN30 = 0.8660254037844387, 0.5
lift = (평면대각선 * 0.5 * iso_z_scale) / (표고범위)      # 표고범위 0 이면 lift = 0
x' = (x - y) * COS30
y' = (x + y) * SIN30 + (elev - e_ref) * lift
```

- `e_ref`(lift 영점)는 기본 bbox 중앙, `ref_label` 이 주어지면 그 노드의 표고.
  **알람밸브를 영점으로 잡는 것을 권장한다** — 안 그러면 이음매에서 두 망이 찢어진다
  (모듈 A 실측: 대명동에서 약 11.6 m 벌어짐).
- `no_lift_labels` 노드(라이저·기계실 계통도)는 lift 를 건너뛴다. schematic y 가 이미
  수직을 인코딩하고 있어 표고 lift 를 또 더하면 이중부호로 계통도가 구부러진다.
- **호출 순서는 정규화 → 베이크다.** 순서를 바꾸면 lift 배율이 어긋난다.
- 헤드 z-돌출은 여기서 적용하지 않는다(평면 Y 를 기울여 가지배관을 꼬이게 만든다).

4번째 창(`dialog_design_input.py`)에 조절 칸을 붙인다:

| 칸 | 기본값 | 설명 |
|---|---|---|
| 아이소매트릭 보기 | 꺼짐 | 켜면 30° 등각으로 굽는다 |
| 고도 펼침 배율 | 1.0 | `iso_z_scale`. 0.5~3.0 범위 |
| 캔버스 크기 | 3000 | `canvas_units`. 망이 커/작게 보일 때 조절 |
| lift 영점 | 알람밸브 | 없으면 표고 중앙 |

이 네 칸은 **표시 전용**임을 화면에 한 줄로 명시한다 — 수리계산 결과는 바뀌지 않는다.

**수용 기준**
- 아이소매트릭을 켜고 방출한 SDF 를 PIPENET 에서 열면 층·입상관이 세로로 분리돼 보인다.
- **회귀**: 켬/끔 두 SDF 의 `length` · `rise` · `elevation` · `bore` · 부속 · 노즐 유량이
  **완전히 동일**하다. `<Position>` 값만 다르다.
- `iso_z_scale` 을 1.0 → 2.0 으로 올리면 세로 분리 폭만 커지고 평면 배치는 그대로다.
- 표고가 모두 같은 평면 전용 망에서 베이크해도 예외 없이 동작한다(`lift = 0`).

### G13. 대조 검증

같은 도면을 모듈 A 로도 돌려 두 SDF 를 비교하고 `docs/module_g_vs_a.md` 에 추가한다.

**수용 기준**
- Pipe-set 구성(개수·순서·각 Pipe-type 의 Name/Schedule)이 A 의 산출과 동일하다.
- 좌표 폭(bbox)이 A 의 산출과 ±5 % 이내다.
- 노드·배관·노즐 개수, `length` 합계가 종전 G 산출과 동일하다(위상 불변 증명).

---

## 3. 금지·보존 목록

- **`pipenet_converter/` 를 고치지 말 것.** Pipe-type 주입은 방출 **후처리**로 한다.
  writer 를 고치면 모듈 A·F 의 산출물까지 함께 흔들린다.
- `remote30_prototype.py` · `routes/` · `cad_project_editor/`(모듈 E) 수정 금지.
- 좌표·표시 z 이외의 값(길이·표고·낙차·관경·부속·유량)은 이 보정에서 **한 개도**
  바뀌면 안 된다. 회귀로 증명할 것.
- SLF 자산을 편집하지 말 것. 6종 스케줄은 이미 정의돼 있다.
- 공개 함수 시그니처 파괴 금지 — keyword 인자 추가만 허용. 기본값으로 부르면
  종전 동작이어야 한다.

---

## 4. 검증 체계

`cad_project_editor_g/tests/test_sdf_post.py` 신설.

- G9: 합성 테이블(스케줄 2종 혼합) → Pipe-set 분리·Pipe-type 주입·placeholder 유지.
- G11: 합성 좌표 → bbox 폭 3000, 중심 (0,0), `display_z` 동일 배율.
- G12: 표고 0 인 평면망에서 `lift=0`, 표고 있는 망에서 켬/끔 산출의 계산값 동일.
- **회귀(가장 중요)**: 실도면 B1F 로 종전 G 산출과 새 산출을 비교해
  `length` · `rise` · `elevation` · `bore` · `<Fitting>` · `<Nozzle>` 이 전부 동일하고
  `<Position>` 과 `<Pipe-type>` 만 달라진 것을 diff 로 보인다.
- 커밋 메시지 형식: `[G#] 제목 — 수용기준: PASS/FAIL(사유)`

---

## 5. 진행 규칙 · 미정 사항

- 항목 순서 준수, 항목 병합 금지.
- 아래는 아직 정하지 않았다. 임의로 구현하지 말고 만나는 시점에 `BLOCKED.md` 에 올릴 것.
  - **배관별 관종 지정 방법** — 손질 단계에서 구간을 찍어 지정할지, 변환 창에서
    레이어·구간별로 매핑할지, 전체 기본값 하나로 둘지.
  - **아이소매트릭 설정의 저장 위치** — 매번 입력할지, 프로젝트에 저장해 다음에도 따라올지.
- 최종 산출물: 변경 diff 요약, 테스트 결과 전문, `BLOCKED.md`,
  `docs/module_g_vs_a.md` 갱신, PIPENET 에서 연 화면 캡처(Type 열이 채워진 것).
