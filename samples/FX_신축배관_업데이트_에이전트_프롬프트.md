# Remote 30 — 신축배관(FX) 표준화 + 웹 테이블 편집기 구현 지시서

## 0. 컨텍스트

이 저장소는 FNCADnet의 Module A(Remote 30) 프로토타입이다: DXF 도면 → 배관망 추출 → 최불리 30헤드 선정 → 5테이블(Nodes/Pipes/Nozzles/Fittings/Equipment) → PIPENET SDF + SLF 동봉 + KFP 출력.

핵심 파일:
- `remote30_prototype.py` — 파이프라인 본체 (stage 0~6, 약 4,350줄)
- `core/remote30_constants.py` — 튜닝 상수 (STEEL_PIPE_TYPE, CPVC_C_FACTOR 등)
- Flask 서버와 웹 프론트는 별도 파일에 있음 (이 지시서의 게이트 패턴 설명 참조)

현재 신축배관(FX)은 `build_input_tables()` 안에서 두 경로로 Equipment에 추가된다:
1. **도면 추출**: `SP 후렉시블` 레이어 LWPOLYLINE의 endpoint를 헤드 500mm 이내로 매칭 → 해당 헤드의 파이프에 `desc="FX"` Equipment 부착
2. **자동 보충** (`# 1.5) FX 보충` 주석 블록): 도면에서 못 찾은 헤드 포함, 선정된 30헤드 전원에 FX 1개씩 보장

두 경로 모두 `"eq_len": 15.62` 가 **하드코딩**되어 있다 (파일 내 2회 등장 — `grep -n '15\.62' remote30_prototype.py` 로 확인). 이 값은 구(舊) 레퍼런스(2. Pipenet_hand) 기준이며, **한백 표준이 F사 유형(22.4m)으로 확정되어 교체 대상**이다.

파이프라인은 이미 2단 분리 + 사람 검토 게이트 패턴을 갖고 있다:
- `run_stages_0_2()` — 파싱/배관망/헤드 인식 → 마지막 `stage2_complete` 이벤트로 job state 저장
- (웹에서 사용자가 헤드 추가/삭제 편집)
- `run_stages_3_5()` — 편집 결과(`user_added_heads`/`user_deleted_indices`)를 받아 선정→테이블→**SDF emit까지 한 번에** 실행

이번 작업의 본질: **이 게이트 패턴을 한 번 더 복제**하여 "테이블 생성"과 "emit" 사이에 신축배관 검토 게이트를 끼워 넣고, FX 값을 표준 상수(표 1)와 인스턴스 테이블(표 2)로 분리하는 것.

---

## 1. 확정 표준값 (한백 표준 = F사 유형, 2026-07 확정)

| 항목 | 값 | 비고 |
|---|---|---|
| 등가길이 (eq_len) | **22.4 m** | Equipment item에 입력. 기존 15.62 전면 교체 |
| 관경 | 25A (내경 28mm) | 말단 파이프는 기존 NFPC 별표1 추론 유지 (헤드≤2개 → 25A 최소, 이미 정합) |
| C값 (조도) | 120 | 기존 유지. **C값으로 손실을 흉내 내는 방식 금지** |
| 물리 길이 | 0.7 m (참고) | 파이프 기하(도면 거리)에 이미 포함 — **eq_len이나 파이프 길이에 별도 가산 금지 (이중 계상 방지)** |
| 유속 | FX 스케줄 velocity=10 (기존 `_SCHEDULE_DEFS` 정의 유지) | |

---

## 2. Task 목록 (순서대로 실행)

### Task 1 — 표 1: 규격 프로파일 상수화 (`core/remote30_constants.py`)

`remote30_constants.py` 끝에 추가:

```python
# ── 신축배관(FX) 규격 프로파일 — "표 1" (원본 규격표) ──────────────
# 값은 여기에만 존재한다. build_input_tables 등 사용처는 반드시 이 dict를 참조.
# 프로파일 추가 시(예: 상가형) 여기에만 항목을 늘린다.
FX_SPEC_PROFILES: dict[str, dict] = {
    "HANBAEK_STD": {           # F사 유형 — 한백 표준 (2026-07 확정)
        "eq_len_m": 22.4,      # Equipment 등가길이
        "nominal_dn": 25,      # 25A — 말단 파이프 최소 호칭경과 정합
        "inner_dia_mm": 28.0,  # SLF Size-definition 검증 대상 (Task 6)
        "c_factor": 120,
        "phys_len_m": 0.7,     # 참고용 — 파이프 기하에 포함됨. 가산 금지.
    },
}
FX_DEFAULT_PROFILE = "HANBAEK_STD"
AV_EQ_LEN_M = 12.9             # 알람밸브 등가길이 (기존값 상수화만, 값 변경 없음)
```

그리고 `remote30_prototype.py`에서:
- `"eq_len": 15.62` **2곳 모두** → `FX_SPEC_PROFILES[FX_DEFAULT_PROFILE]["eq_len_m"]` 참조로 교체
- A/V 블록의 `"eq_len": 12.9` → `AV_EQ_LEN_M` 참조로 교체
- 15.62 관련 주석("권위 레퍼런스 … 15.62m 채택")은 "한백 표준 F사 유형 22.4m (FX_SPEC_PROFILES 참조)"로 갱신

**완료 기준**: `grep -rn '15\.62' --include='*.py' .` 결과가 0건 (또는 이력 주석만 잔존).

### Task 2 — 표 2: equipment 항목 스키마 확장

`tables.equipment`에 append되는 dict에 다음 필드 추가 (FX·A/V 공통, 기존 필드 유지):

```python
{
    ...,  # pipe, in, out, label, desc, eq_len, rel_pos (기존)
    "spec_ref": FX_DEFAULT_PROFILE,   # 표 1의 어느 프로파일을 참조하는지. A/V는 "AV_STD"
    "source": "extracted",            # "extracted"(도면 추출) | "supplemented"(자동 보충) | "manual"(사용자 수정)
    "override_flag": False,           # 사용자가 값을 직접 덮어썼는지
    "override_note": "",              # 수동 수정 사유 (선택)
    "drawing_len_mm": fx_len_mm,      # 도면상 후렉시블 물리 길이 — QA/대사용. 보충 경로는 None
}
```

주의: 도면 추출 경로에서 이미 계산하는 `fx_len_mm` 변수를 버리지 말고 `drawing_len_mm`으로 기록할 것. `write_csv_tables` / `write_xlsx_tables`의 equipment 헤더·행에도 새 컬럼 반영.

### Task 3 — 파이프라인 게이트 분리: `run_stages_3_5` → 테이블까지 + `run_stage_6_emit` 신설

**기존 헤드 편집 게이트(`run_stages_0_2` → job state → `run_stages_3_5`) 패턴을 그대로 복제한다.**

1. `run_stages_3_5()`에서 stage 6(SDF emit) 블록을 **분리 제거**하고, 마지막에 다음 이벤트를 yield하며 종료:
   ```python
   yield evt({"type": "stage5_complete",
              "tables": tables.as_dict(),      # 직렬화 헬퍼 필요 시 PipeTables에 추가
              "fx_review": {
                  "equipment": tables.equipment,   # 전량 — [:8] 캡 금지
                  "profiles": FX_SPEC_PROFILES,    # 편집기 드롭다운용
                  "default_profile": FX_DEFAULT_PROFILE,
              }})
   ```
   서버는 `stage2_complete`와 동일한 방식으로 이 데이터를 job state에 저장한다.

2. 신규 함수:
   ```python
   def run_stage_6_emit(
       out_dir: Path, job_id: str,
       tables: PipeTables,                      # job state에서 복원
       edited_equipment: list[dict] | None = None,  # 웹 편집 결과. None이면 원본 그대로
       *, project_title: str = "Remote 30 Prototype",
   ) -> Iterator[dict]:
   ```
   동작:
   - `edited_equipment`가 오면 검증 후 `tables.equipment`를 교체
   - 검증 규칙: `eq_len`은 float > 0; `spec_ref`는 `FX_SPEC_PROFILES`에 존재하거나 `override_flag=True`; 사용자가 값을 바꾼 행은 `source="manual"`, `override_flag=True` 강제; 프로파일 기준 ±50% 초과 편차는 경고 이벤트(`{"type":"warning", ...}`)로 통지하되 차단하지 않음
   - 이후 기존 stage 6 로직(emit_sdf → SLF 동봉 → emit_kfp → zip → `done` 이벤트) 그대로 실행

3. 기존 원샷 호출자가 있으면 하위호환 래퍼 유지:
   ```python
   def run_stages_3_6(...):  # 기존 시그니처 유지 — 3_5 실행 후 곧바로 6 실행
   ```
   저장소 전체에서 `run_stages_3_5` 호출부를 grep하여 영향 범위 확인 후 결정할 것.

### Task 4 — 서버 라우트 (Flask, 기존 게이트 라우트 패턴 준수)

기존 헤드 편집 finalize 라우트를 찾아 **네이밍·상태관리 컨벤션을 동일하게** 따라 추가:
- `run_stages_3_5` 완료 시 `stage5_complete` 데이터를 job state에 저장 (기존 stage2 저장 방식과 동일)
- 신규 엔드포인트 (예): `POST /jobs/<job_id>/fx/finalize` — body: `{"equipment": [...]}` → `run_stage_6_emit` SSE 스트림 반환
- 편집 없이 바로 진행하는 경로도 유지: body 생략 시 원본 equipment로 emit

### Task 5 — 웹 테이블 편집기 (프론트)

stage 5 완료 후 "신축배관 검토" 패널을 표시. 기존 헤드 편집 UI의 스타일·상호작용 컨벤션을 따른다.

테이블 컬럼 구성:

| 컬럼 | 편집 | 내용 |
|---|---|---|
| FX # / Desc | 불가 | equipment label, FX/AV 구분 배지 |
| 헤드 노드 / 파이프 | 불가 | in·out 노드, pipe label |
| 출처 | 불가 | 배지: 도면추출 / 자동보충 / 수동 (source 필드) |
| 규격 (spec_ref) | **드롭다운** | FX_SPEC_PROFILES 키 목록 + "직접 입력" |
| 등가길이 (eq_len) | **숫자 입력** | 기본은 프로파일 값 표시(읽기전용 느낌). 직접 수정 시 override_flag=True + 행 하이라이트 |
| 도면 물리길이 | 불가 | drawing_len_mm (보충 행은 "—"). QA 참고용 |
| 비고 (override_note) | 텍스트 | 수동 수정 시 사유 입력 권장 |

동작 요구:
- 상단 일괄 버튼: "전체 표준 적용 (HANBAEK_STD)" / 행별 "초기화(프로파일 값으로)"
- override된 행은 시각적으로 구분 (색/아이콘) — "사람이 의도적으로 바꿈"이 한눈에 보여야 함
- 유효성: eq_len 비어있음/0 이하 → 확정 차단; ±50% 편차 → 경고 표시하되 진행 허용
- "확정 후 SDF 생성" 버튼 → Task 4 엔드포인트 호출 → 기존 진행 이벤트 UI로 stage 6 스트림 표시

### Task 6 — SLF 내경 검증 (자산 확인 — 코드 수정 아님)

`resolve_standard_slf()`가 가리키는 표준 SLF에서 다음을 파싱·확인하고 결과를 보고할 것:
- KSD 3507 스케줄의 25A(0.025) Size-definition 내경이 F사 기준 **28mm(±0.5)** 인지
- FX 스케줄 25A 내경도 동일 기준인지

**불일치 시 SLF를 수정하지 말 것** — 공유 자산이므로 수치·경로를 리포트만 하고 사람 결정을 기다린다. Hazen-Williams에서 내경은 4.87제곱으로 작용하므로 이 검증이 등가길이보다 결과에 더 크게 영향한다.

---

## 3. 가드레일 (하지 말 것)

1. **이중 계상 금지**: phys_len 0.7m를 eq_len이나 파이프 길이에 더하지 말 것. 파이프 길이는 도면 기하에서 이미 나온다.
2. **C값 조작 금지**: 등가길이 대신 C값을 낮춰 손실을 흉내 내는 방식(G사 유형)은 어떤 경로로도 도입하지 말 것 — 다른 관로 계산을 왜곡한다.
3. **emit/후처리 단계에서 FX 값 임의 변경 금지**: 신축배관은 형식승인 고정 제품이다. emit 단계나 향후 재최적화 로직이 FX의 eq_len·관경을 자동 조정해서는 안 된다. FX는 항상 "주어진 제약"이다.
4. `_SCHEDULE_DEFS`의 스케줄 이름·velocity 값 변경 금지 (SLF Item-name 바인딩이 깨진다). 상수 참조로의 치환만 허용.
5. Template SDF / 표준 SLF 자산 파일 수정 금지 (Task 6은 검증·보고만).
6. 기존 SSE 이벤트 타입·형식은 유지하고 **새 타입 추가**로 확장할 것 (프론트 하위호환).

## 4. 완료 기준 (acceptance)

1. `grep -rn '15\.62' --include='*.py' .` → 코드 0건
2. 샘플 DXF 1건으로 전체 플로우 실행 시: 선정 헤드 30개 전원에 FX 1개씩, 모든 FX `eq_len=22.4`, `spec_ref="HANBAEK_STD"`, source가 extracted/supplemented로 올바르게 구분됨
3. 편집기에서 특정 행 eq_len을 수정 → 확정 → 생성된 SDF의 해당 `<Equipment equivalent-length>`에만 수정값 반영, `override_flag=True` + `source="manual"` 기록
4. 편집 없이 확정한 경우 기존 원샷 결과와 SDF가 동일 (회귀 없음 — diff로 확인)
5. xlsx/csv equipment 시트에 신규 컬럼(spec_ref/source/override_flag/override_note/drawing_len_mm) 포함
6. KFP 변환·zip 번들 기존대로 동작
7. Task 6 SLF 검증 결과 리포트 제출 (일치/불일치 수치 명시)

## 5. 참고 앵커 (라인 번호는 드리프트 가능 — 검색어 기준)

- FX 도면 추출: `en.get("l") != "SP 후렉시블"` 블록 (~L3437)
- FX 자동 보충: `# 1.5) FX 보충` (~L3494)
- 하드코딩 값: `"eq_len": 15.62` ×2 (~L3490, L3510), `"eq_len": 12.9` (~L3519)
- 스케줄 정의: `_SCHEDULE_DEFS` 내 `("FX", "120", [(0.025, 10)])` (~L3869)
- 미리보기 캡: `tables.equipment[:8]` (~L4304)
- 게이트 패턴 원형: `run_stages_0_2` docstring (~L4006) — "호출자(서버)는 … job state에 저장해두고 … finalize 호출 시 run_stages_3_5()에 전달"
- NFPC 최소관경: `_nfpc_min_bore_mm` (헤드≤2 → 25)
