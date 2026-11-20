# Module F 통합 작업 지시서 — 평면도 전용에서 특허 전 국면으로

당신은 FNCADnet 저장소의 유지보수 엔지니어다. 아래를 **항목 순서대로, 항목당 커밋 1개씩**
수행하라. 각 항목은 수용 기준을 통과한 뒤에만 다음으로 진행한다. 판단이 필요한 애매점은
임의로 구현하지 말고 `BLOCKED.md`(저장소 루트)에 기록하고 다음 항목으로 넘어간다(§5).

이 지시서의 규범 문서는 **`scripts/특허도면(한백수정본V1).pptx`** 다. 아래의 모든 항목은
특허 부호(S···)에 대응하며, 부호와 어긋나는 구현은 오구현으로 본다. 덤프본은
`data/_patent_dump.txt`(`scripts/_dump_patent.py` 로 재생성).

작업 대상: **`routes/module_f/`** · **`templates/module_f.html`** · **`cad_project_editor_g/`**(엔진).

**동결 해제(이 지시서 한정)**: `routes/r30_combined.py` — 단, **H-1 의 승격 리팩터에 한한다.**
동작을 바꾸는 수정은 금지하며, A 의 기존 산출이 바이트 동일해야 한다(H-1 수용기준).

읽기 전용 참조(수정 금지): `remote30_prototype.py`, `cad_project_editor/`(모듈 E — 동결),
`pipenet_converter/`, `routes/r30_system.py`, `routes/r30_machineroom.py`, 그 밖의 `routes/r30_*.py`.

---

## 0. 이 작업의 정체 (구현 전 필독)

### 0.1 F 의 결손 — 특허 대비

F 는 **평면도 한 장**에 대한 제1~4국면(S100~S560)을 이미 구현하고 있다. 없는 것은
**S650(추가도면 회귀)과 제5국면 S700(병합·결과 출력) 통째**다.

| 부호 | 단계 | F | A 의 구현 |
|---|---|---|---|
| S650 | 추가도면 여부 → 평면도·계통도·기계실 반복 | 없음 | 도면 종류별 라우트 |
| S710 | 급수방식 선택 (펌프 / 자연낙차 / 1차 감압 / 2차 감압) | 없음 | `r30_combined` |
| S720 | 급수방식별 입상관 구성 | 없음 | `extract_system_path` · `_system_path_to_riser_dict` |
| S730 | 기계실 배관 전단 접속 (수원 이동 · 낙차 부여) | 없음 | `extract_machine_room_path` |
| S740 | 입상관–헤드배관 결합 (기준점 번호 10 공통 절점) | 없음 | `_remap_riser_to_head_av` |
| S750 | 입력파일 생성 | **부분** (평면도만) | `_emit_format_bundle` |
| S760 | 타 형식 변환 | 있음 | 동일 |
| S770 | 산출물 일괄 압축 | 없음 | `_emit_format_bundle` |

F 의 33개 엔드포인트는 전부 도면 1장짜리다. `common.py` 의 `DIAGRAMS` 는 헤드 종류
그림이지 도면 종류가 아니다 — **도면 종류 개념 자체가 F 에 없다.**

### 0.2 결정적 사실 — A 는 이미 import 되고 있다

`remote30_prototype.py` 는 `PySide6`/`PyQt` **0건**으로 완전 헤드리스다. F 는 이미
세 곳에서 A 를 직접 물고 있다:

```
routes/module_f/api_pick.py:146    import remote30_prototype as A   # detect_heads (D2)
routes/module_f/common.py:181      from remote30_prototype import _categorize_layer
routes/module_f/remote30.py:82     from remote30_prototype import detect_sheet_frames
```

**이 작업에서 새 추출 로직은 쓰지 않는다.** A 의 계통도·기계실 추출기를 같은 방식으로
붙이고, 화면을 만들고, S700 으로 흐름을 잇는 것이 전부다.

붙일 A 함수 (전부 `remote30_prototype.py`, 수정 금지):

```
build_system_graph            (1330)   계통도 그래프
extract_system_path           (1490)   계통도 급수경로 → 입상관 (S720)
_system_path_to_riser_dict    (1639)
extract_clean_system_network  (1812)   조각난 계통도 대응
_network_to_riser_dict        (1841)
extract_machine_room_path     (2007)   기계실 (S730)
_machine_room_path_to_dict    (2125)
```

### 0.3 확정된 설계 결정 (변경 금지)

| # | 항목 | 결정 |
|---|---|---|
| H-D1 | A 엔진 소유권 | **import 재사용.** A 를 고치지도, F 안에 베끼지도 않는다. `api_pick.py` 의 D2 선례와 같은 구도 |
| H-D2 | S700 오케스트레이션 | ~~라우트 본문 승격~~ → **접합만 한다.** 실측 결과 S700 원시함수가 이미 전부 모듈 레벨이다(`core/remote30_full_network.py`: `build_riser`·`prepend_machine_room_to_riser`·`stitch_riser_and_heads`·`emit_full_sdf`). 승격할 이유가 없어졌고 `r30_combined.py` 동결도 **그대로 유지**한다. 아래 H-D7 참조 |
| H-D7 | 평면 쪽이 다르다 | A 의 590줄 라우트는 대부분 «A 의 평면 경로»(`_PROTOTYPE_JOBS` → `build_input_tables`)다. F 의 평면은 사람이 손질한 board 위 G 의 `select_and_expand`→`build_design_tables` 다. **다른 것은 평면 쪽뿐이고 S700 은 공유돼 있다** — 그래서 F 는 제 오케스트레이션을 갖되 원시함수는 공유한다. 사본은 생기지 않는다 |
| H-D8 | 기준점 번호 | 특허 S550 «기준점 번호 = 10» · S740 «10 을 공통 절점으로 결합». A 의 헤드망도 `{10,11,12,…}`, 라이저 빌더 4종 모두 `av_node_label="10"`. 그런데 G 는 BFS 로 1 부터 매긴다 → **+9 오프셋**으로 정확히 일치. 라벨이 박힌 자리(배관·노즐·부속·기기의 in/out)를 빠짐없이 옮기되 **배관 이름(label·pipe)은 옮기지 않는다** |
| H-D3 | 도면 슬롯 | 세션이 **평면도 · 계통도 · 기계실 3슬롯**을 갖는다(S650). 평면도는 필수, 나머지 둘은 선택. 슬롯마다 제1~4국면을 독립 수행 |
| H-D4 | 급수방식 | **사람이 고른다**(S710, 4종). 자동 추정 금지 — E 의 「표시가 없으면 추측하지 않는다」 승계 |
| H-D5 | 기준점 규약 | 결합 절점은 **기준점 번호 10**(S550·S740). F-1 의 `source_selection_required` 규약과 통일 |
| H-D6 | 좌표 처리 | 계통도는 **수직 막대로 재배치**, 기계실은 **평면 좌표 보존**(S730 주석 그대로). 임의 변경 금지 |

### 0.4 순서의 원칙

수치를 바꾸는 작업은 **H-5(결합)** 하나뿐이다. H-1~H-4 는 계약·승격·입력이라 평면도
단독 산출이 **바이트 동일**해야 한다. H-5 이후로만 새 기준선을 뜬다.
**H-5 이전의 어떤 산출물도 이후 회귀 기준선으로 쓰지 말 것.**

---

## 1. 신규 계약

```
routes/module_f/
  slots.py          [신설] 도면 슬롯 3종 · S650 회귀 상태기계
  api_slot.py       [신설] 슬롯 열기/전환/삭제
  api_merge.py      [신설] S700 — 급수방식·입상관·결합·산출
  common.py         [수정] SLOT_KINDS 추가
templates/module_f.html   [수정] 슬롯 탭 + 급수방식 패널 + 병합 결과
routes/r30_combined.py    [승격] 라우트 본문 → build_combined_network()
```

새 엔드포인트(`/api/module-f/*` 규약 유지):

```
POST /api/module-f/slot/open        도면 종류 지정 열기 (kind=plan|system|machineroom)
POST /api/module-f/slot/switch      활성 슬롯 전환
GET  /api/module-f/slot/state       3슬롯 진행 상태 (S650 판단 근거)
POST /api/module-f/merge/mode       급수방식 선택 (S710)
POST /api/module-f/merge/build      S720~S740 실행
GET  /api/module-f/merge/preview    결합망 미리보기
POST /api/module-f/merge/emit       S750·S760·S770 (zip)
```

---

## 2. 작업 항목 (순서 고정)

### H-0. 도면 슬롯 계약 (S650 · H-D3)

`slots.py` 신설. 세션 하나가 `{plan, system, machineroom}` 3슬롯을 들고, 각 슬롯이
기존 F 세션의 도면 상태(찍기·손질·설계)를 독립으로 갖는다. 활성 슬롯 개념을 넣어
기존 33개 엔드포인트가 **활성 슬롯에 대해** 동작하게 한다.

- 기존 엔드포인트의 요청/응답 스키마는 **바꾸지 않는다**. `sid` 만으로 부르면
  활성 슬롯(기본 `plan`)에 붙는다 — 구 클라이언트가 그대로 돌아야 한다.
- 슬롯 미개설 상태에서 `system`/`machineroom` 을 부르면 `slot_not_opened` 로 실패.

**수용 기준**: `tests/test_module_f_*.py` 전 항목 PASS(슬롯 도입 전과 동일).
새 `tests/test_module_f_slots.py` 가 3슬롯 독립성과 활성 전환을 확인.

### H-1. S700 오케스트레이션 승격 (H-D2)

`routes/r30_combined.py` 의 `remote30_combined_build()` 본문을 모듈 레벨
`build_combined_network(...) -> dict` 로 들어올린다. `_remap_riser_to_head_av` 도
모듈 레벨로 승격. 라우트는 **요청 파싱 + 승격 함수 호출 + 응답 직렬화**만 남는다.

- Flask `request`/`current_app` 접근은 전부 라우트 쪽에 남기고, 승격 함수는
  **순수 인자만** 받는다. 승격 함수 안에 HTTP 개념이 남으면 실패로 본다.
- `register()` 가 주입받던 `_RISER_HEIGHT_FRAC` 등은 승격 함수의 기본값 인자로.

**수용 기준**: 승격 전후로 `/api/remote30/combined/build` 산출이 **바이트 동일**.
기준 도면 1건으로 승격 전 산출을 먼저 떠 두고 `fc /b` 로 대조한 로그를 커밋 메시지에 남긴다.
A 의 기존 테스트 전 항목 PASS.

### H-2. 계통도 슬롯 (S100~S370 on 계통도 · S720 준비)

`system` 슬롯에서 A 의 `build_system_graph` · `extract_system_path` 를 부른다.
조각난 계통도는 `extract_clean_system_network` 로 폴백(메모리
`lh306-system-dxf-fragmented` 의 실측 — 풀 도면 PIPE 레이어는 단일망 추출 불가).

- 추정으로 이은 edge 는 **점선 + 다른 색**으로 실측과 구분해 그린다
  (메모리 `verification-distinguish-estimated-edges`. 통합 렌더 금지).
- 계통도 좌표는 H-D6 대로 수직 막대 재배치. 막대 길이는 S740 규칙
  (`헤드 분포 세로폭 × 0.8`).

**수용 기준**: 계통도 DXF 1건에서 입상관 dict 가 나오고, 절점 수·연장이
A 의 `/api/remote30/system/extract` 와 일치. 불일치 시 그 차이를 `BLOCKED.md` 에 기록.

### H-3. 기계실 슬롯 (S730 준비)

`machineroom` 슬롯에서 A 의 `extract_machine_room_path` 를 부른다.
좌표는 **평면 그대로 보존**(H-D6).

- 초대형 기계실 도면은 메모리 `machineroom-parse-budget-guard` 의 예산 가드를
  반드시 태운다 — 비용은 `OTHER` 레이어에 있으므로 `ARCH` 만 스킵하면 이득이 0이다.

**수용 기준**: 기계실 DXF 1건에서 경로 dict 가 나오고 A 의
`/api/remote30/machineroom/extract` 와 절점 수 일치. 파싱 예산 초과 시 부분 결과 + 경고.

### H-4. 급수방식 선택 (S710 · S720 · H-D4)

`/merge/mode` 로 4종(펌프 가압 / 자연낙차 / 1차 감압 / 2차 감압) 중 하나를 사람이 고른다.
고른 방식이 H-2 의 입상관 구성에 반영된다(감압밸브·펌프 접속 구조).

- **자동 추정 금지.** 미선택 상태로 `/merge/build` 를 부르면 `supply_mode_required` 로 실패.
- 화면은 라디오 4개 + 각 방식의 입상관 모식도.

**수용 기준**: 4종 각각에 대해 입상관 구성이 달라지는 것을 테스트가 확인.
미선택 시 실패하는 것도 확인.

### H-5. 병합 (S730 · S740) — **수치가 바뀌는 유일한 항목**

H-1 의 `build_combined_network` 를 F 의 3슬롯 결과로 부른다.
기계실 → 입상관 전단 접속(S730), 입상관 끝단 유수검지장치 절점 ↔ 헤드배관 같은 절점
결합(S740, 기준점 번호 10).

- 계통도·기계실 슬롯이 비어 있으면 **평면도 단독**으로 지나가야 한다(둘 다 선택).
  이 경우 산출은 H-4 이전과 바이트 동일해야 한다.
- 결합 실패(절점 불일치 등)는 임의로 메우지 말고 **미도달로 보고**한다(S340 원칙 승계).

**수용 기준**: 3장 결합 산출이 A 의 `/api/remote30/combined/build` 와 일치.
평면도 단독 경로가 회귀 없음. 여기서 새 골든을 뜬다.

### H-6. 산출 (S750 · S760 · S770)

결합망을 5종 입력표(절점·관로·헤드·관이음쇠·기기)로 기록하고(S550), 형식 변환(S760)은
**S750 결과 파일을 원본으로** 수행한다 — 별도 산출 금지. 전 형식이 항상 같은 배관망을
가리켜야 한다. 마지막에 형식별 파일 + 입력표 + 호칭경 대조 자료를 하나로 압축(S770).

**수용 기준**: zip 안의 `.sdf`/`.kfp`/`.has` 가 같은 절점 수·연장을 갖는다(교차 검증).
호칭경 대조 자료가 동봉된다.

### H-7. 골든 · 문서

`docs/module_f_integrated.md` 에 특허 부호 ↔ 구현 위치 대조표를 남긴다.
`tests/module_f_integrated_golden.json` 으로 3장 결합 골든 고정.

**수용 기준**: 전 테스트 PASS. 문서의 부호 대조표가 실제 코드 위치와 일치.

---

## 3. 회귀 기준선

- H-1 이전: A 의 `/api/remote30/combined/build` 산출 (바이트 동일 대조용)
- H-5 이전: F 의 평면도 단독 산출 (`tests/module_f_complete_golden.json`)
- H-5 이후: 3장 결합 골든 (`tests/module_f_integrated_golden.json`)

## 4. 하지 말 것

- A(`remote30_prototype.py`, `routes/r30_system.py`, `routes/r30_machineroom.py`) 수정
- `r30_combined.py` 의 **동작** 변경 (승격 리팩터만 허용)
- 새 추출·계산 로직 작성 (§0.2)
- 급수방식·기준점의 자동 추정 (H-D4 · H-D5)
- 특허 부호와 어긋나는 단계 순서 (특히 S200 → S300 선후, S342 → S343 선후)

## 5. BLOCKED 규약

판단이 필요한 애매점은 구현하지 말고 `BLOCKED.md` 에 다음 형식으로 기록한다:

```
### H-B<n> <한 줄 제목>
- 항목: H-<n>
- 특허 부호: S<···>
- 무엇이 애매한가:
- 실측 근거:
- 임의 구현하지 않은 이유:
```
