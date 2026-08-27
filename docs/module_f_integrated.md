# 모듈 F 통합 — 특허 부호 ↔ 구현 위치

규범 문서는 `scripts/특허도면(한백수정본V1).pptx` 다(덤프: `data/_patent_dump.txt`,
재생성 `scripts/_dump_patent.py`). 아래 표는 그 부호가 **실제로 어느 코드에**
있는지의 대조다. 부호와 어긋나는 구현은 오구현으로 본다.

이 작업 이전의 F 는 **평면도 한 장**에 대한 제1~4국면(S100~S560)만 있었다.
없던 것은 S650(추가도면 회귀)과 제5국면 S700(병합·결과 출력) 통째다.

---

## 어디의 무엇을 합쳤나

| | 가져온 것 | 왜 |
|---|---|---|
| **모듈 A** | 계통도·기계실 추출, 헤드 후보 인식(`detect_heads`), S700 전 단계 원시함수, 레이어 사전 | 실도면에서 검증된 추출기. 헤드리스(PySide6 0건)라 그대로 import 된다 |
| **모듈 E** | 「표시가 없으면 추측하지 않는다」 판정 철학, 사람이 찍고 손질하는 게이트 | 지저분한 실측 도면 앞에서 판단을 사람에게 물을 수 있는 유일한 계통 |
| **모듈 G** | 최불리 선정·관경·부속·표고·SDF 방출 (`design/`) | G1~G18 이 사실상 F 를 위한 엔진 R&D 였다 |
| **모듈 F** | 위 셋을 **브라우저 하나**에 태운 것 | 로그인 게이트 뒤 원격 사용이 가능한 유일한 계통 |

**A 를 고치지 않는다.** `remote30_prototype.py` · `routes/r30_*.py` 는 읽기 전용이다.
F 는 그 함수를 import 해서 부를 뿐이다. 사본을 만들지 않는다.

---

## 평면도의 두 길 — 자동(A) / 수동(E)

같은 평면도에서 같은 것(최불리 헤드군)을 뽑는 길이 둘이고, **업로드할 때
사람이 고른다**(고르기 전에는 단계바가 「도면 열기」 하나다).

| | 사람이 정하는 것 | 나머지 | 흐름 |
|---|---|---|---|
| **자동 (A)** | 알람밸브 한 점 + 헤드 영역 | 헤드 검출·그래프 복원·앵커·최불리 K 전부 자동 | 도면 열기 · 자동 추출 · 수리계산 · 통합 |
| **수동 (E)** | 색으로 배관·헤드·급수원을 직접 찍음 | 손질·자동이음은 도구로 보조 | 도면 열기 · 찍기 · 손질 · 변환 · 수리계산 · 통합 |

구현: `routes/module_f/auto.py` · `api_auto.py` (A 의
`select_worst30_heads_anchored` + `build_input_tables` 를 그대로 부른다).

**라벨 규약이 갈린다** — G(수동)는 BFS 로 1부터라 통합 앞에서 +9, A(자동)는
처음부터 10 이라 옮기지 않는다(`merge.label_offset_for`). 잘못 먹이면 결합이
성립하지 않는 것을 테스트가 강제한다.

**자동 경로가 비우는 것** — A 의 표에는 배관별 관경 근거(`dia_src`)와 meta
「앵커 노드」가 없다. 그래서 화면은 관경 색 나누기를 잠그고(「자동 경로는
배관별 근거를 남기지 않습니다」), 앵커 겹원·최원 유하거리 점선도 그리지
않는다 — 없는 것을 있는 척 그리지 않는다. 최원 거리 «수치»(far_m)는
`selection.distances` 에서 나오므로 요약에는 남는다.

**자동 표 + 실계통도 통합 실측** — `scripts/_probe_auto_merge.py`.
LH306 자동 표(절점 118 · 기준점 10 · 오프셋 0) + 대명동 계통도 라이저(절점 56)
→ S740 절점 173 = 56+118−1 · 노즐 27 보존 · 3형식 연장 273.487 m 일치.
노드 라벨은 기준점 10 만 겹치고 배관 라벨은 안 겹친다(주석의 약속을 실측으로).

## 최불리 선정의 보조 (수동 경로)

- **기준개수 표** — NFTC 103 표 2.1.1.1 을 서버(`core/nftc_rules.py`
  `reference_count_options`)가 그대로 내려보낸다. 화면에 옮겨 적지 않는다 —
  법정 수치의 출처는 하나다. 30 고정이 아니라 10 · 20 · 30 을 표에서 고른다.
- **영역 지정(zones)** — A 의 zones 를 수동 경로에도. 사각형 합집합으로
  후보를 가두되, 가둔 것은 «후보» 지 «망» 이 아니다 — 경로는 영역 밖을
  지나서라도 급수원까지 간다. 상한 64(`MAX_ZONES`).
- **최원 유하거리 경로(anchor_path)** — far_m 이 «어느 줄» 인지. 급수원→앵커
  절점 열을 손질·수리계산 두 캔버스에 같은 빨간 점선으로 그린다
  (`design/worst.py` 가 ①의 Dijkstra `prev` 를 되짚는다 — 다시 풀지 않는다).

## 제1국면 — 도면 입력·인식 (S100)

| 부호 | 단계 | 구현 |
|---|---|---|
| S110 | 도면 해독 | 평면도: E `PickSession.open` / 계통도·기계실: A `parse_dxf_for_view` → `routes/module_f/subdrawing.py:parse_subdrawing` |
| S120 | 도면층 용도 분류 | A `_categorize_layer` ← `routes/module_f/common.py:_layer_category` |
| S130 | 요소 선별·재분류 | E 의 찍기(사람 클릭) — `routes/module_f/api_pick.py` |
| S140 | 헤드 후보 추출 | A `detect_heads`(R1~R5·신뢰도) ← `api_pick.py:module_f_pick_suggest` · **제안만, 확정은 사람** |
| S210 | 유수검지장치 탐지 | 계통도: 사람이 찍는다(아래 참조) |
| S220 | 기준점 확정 | 평면도: `/edit/flow` 급수원 지정 · 계통도: 두 점 찍기 |

> **S220 에서 F 는 「사용자 지정」만 쓴다.** 특허의 우선순위는 사용자 지정 →
> 자동 탐지 → 최다 접속 절점이지만, 계통도의 펌프·알람밸브는 도면마다 기호가
> 달라 자동 탐지가 조용히 틀리면 경로가 통째로 다른 곳으로 간다.

## 제2국면 — 배관망 그래프 구축 (S300)

평면도는 E/G 의 board 위에서, 계통도·기계실은 A 의 그래프에서 돈다.

| 부호 | 단계 | 구현 |
|---|---|---|
| S310 | 정합 허용오차 산정 | F 의 자동 이음 사다리 — `routes/module_f/graph.py` (`AUTOJOIN_LADDER_MM`) |
| S320 | 절점 정합·관로 등록 | E `board` · A `build_system_graph` |
| S330 | 복선 배관 단선화 | E 의 손질 |
| S340 | 단절 접속 복원 | F 의 자동 이음 — **A 처럼 재고 E 처럼 붙인다**(`common.py` 머리말) |
| S350 | 헤드 절점 결합 | G `select_and_expand` (G19 세로 처리) |
| S360 | 기준점 배관망 접속 | `/edit/flow` |
| S370 | 가지식 정형화 | G `restrict.py` |

## 제3국면 — 기준개수 헤드군 선정 (S400)

| 부호 | 단계 | 구현 |
|---|---|---|
| S410~S450 | 배관 연장·정렬·급수경로·직선 통합·직각 정형 | G `design/worst.py` · `design/restrict.py` |

## 제4국면 — 배관 정보 입력 (S500)

| 부호 | 단계 | 구현 |
|---|---|---|
| S510 | 담당 헤드 수 | G `restrict.tree_loads` · `worst["loads"]` |
| S520 | 호칭경 결정 | G `design/bore.py:decide_bores` — **안전측** `max(별표1 최소, 도면 표기)` |
| S530 | 표고 규칙 | G `design/emit.py` |
| S540 | 사용자 확인·보완 | 수리계산 패널 |
| S550 | 입력표 5종 | G `design/tables.py:build_design_tables` |
| S560 | 등각투상 | `api_design.py` `iso` 옵션 (표시 전용) |

> **관경 근거는 화면에서 갈라 보인다.** 도면 텍스트(실선 하늘색) / 별표1 보강
> (실선 노랑) / 별표1 폴백(**점선** 회색). 규약으로만 정한 구간을 실측과 같은
> 모양으로 그리지 않는다.

## S650 — 추가도면 회귀

| 구현 | 위치 |
|---|---|
| 도면 슬롯 3종 (평면도·계통도·기계실) | `routes/module_f/slots.py` |
| 슬롯 열기/전환/상태 | `routes/module_f/api_slot.py` |

세션 dict 는 평면 그대로다 — 그 내용이 «지금 활성인 슬롯» 의 도면 상태다.
도면별 키를 열거하지 않고 **세션 전역만 적고 나머지 전부**를 슬롯 상태로 본다.

## 제5국면 — 병합·결과 출력 (S700)

| 부호 | 단계 | 구현 |
|---|---|---|
| S710 | 급수방식 선택 | `merge.py:SUPPLY_MODES` — 엔진 `ZoneType` 과 1:1 · **자동 추정 없음** |
| S720 | 급수방식별 입상관 | 계통도 추출(`subdrawing.extract_system`) 또는 `build_riser` |
| S730 | 기계실 전단 접속 | `prepend_machine_room_to_riser` |
| S740 | 입상관–헤드배관 결합 | `stitch_riser_and_heads` · **기준점 10** |
| S750 | 입력파일 생성 | `emit.py:emit_merged` → `emit_full_sdf` (+SLF 동봉) |
| S760 | 형식 변환 | **그 SDF 파일을 원본으로** KFP·HAS |
| S770 | 일괄 압축 | `emit_merged` 의 zip |

### 기준점 번호 — 접합의 핵심

특허 S550 은 «기준점 번호 = 10», S740 은 «10 을 공통 절점으로 결합» 이다.
A 의 헤드망도 `{10, 11, 12, …}` 이고 라이저 빌더 4종 모두 `av_node_label="10"`.
그런데 **G 는 BFS 로 1 부터** 매긴다.

`merge.py:to_head_tables` 가 **+9 오프셋**으로 옮긴다. 라벨은 노드표에만 있지
않다 — 배관·노즐·부속·기기의 `in`/`out` 에도 박혀 있어 전부 같이 옮기고,
반대로 배관 이름(`pipes.label`·`fittings.pipe`·`equipment.pipe`)과 노즐 참조
(`@/n`)는 노드 라벨이 아니므로 옮기지 않는다.

### S760 의 불변량은 «수» 가 아니다

SDF→KFP/HAS 변환은 직선 위 통과절점을 통합한다(S440·S443 ·
`kfp_sdf_converter.simplify_passthrough_nodes`). 그래서 절점 수는 **정당하게**
줄어든다 — 실측으로 라이저 체인 때문에 59 → 11 이 됐다.

통합이 보존해야 하는 것은 **총 연장과 노즐**이다(S443: 통합된 관로의 길이는
합산하여 보존된다). `emit.py:cross_check` 가 그 둘로 견준다.

---

## 검증

| 무엇 | 어디 |
|---|---|
| 슬롯 상태기계 | `tests/test_module_f_slots.py` |
| S700 접합(라벨 오프셋·고아 참조) | `tests/test_module_f_merge.py` |
| 계통도·기계실 어댑터 | `tests/test_module_f_subdrawing.py` |
| 통합 골든(실도면) | `tests/test_module_f_integrated.py` + `module_f_integrated_golden.json` |
| 라우트 등록 | `tests/smoke/test_route_inventory.py` |
| 슬롯 라우트 실측 | `scripts/_verify_module_f_slots.py` |
| 계통도·기계실 실측 | `scripts/_verify_module_f_sub.py` |
| 제5국면 실측 | `scripts/_verify_module_f_merge.py` |
| 화면 JS 구문·스코프 | `scripts/_verify_module_f_js.py` |

`node --check` 만으로 화면의 성공을 선언하지 않는다 — 구문이 맞아도 다른
스코프의 헬퍼를 부르면 클릭하는 순간 ReferenceError 로 죽는다.

## 하지 말 것

- `remote30_prototype.py` · `routes/r30_*.py` · `cad_project_editor/`(E) 수정
- 새 추출·계산 로직 작성 (전부 A·G 것을 부른다)
- 급수방식·기준점의 자동 추정
- 특허 부호와 어긋나는 단계 순서 — 특히 **S200 → S300**, **S342 → S343**
- 산출 형식을 결합망 객체에서 각자 뽑기 (S760 은 SDF **파일**이 원본이다)
