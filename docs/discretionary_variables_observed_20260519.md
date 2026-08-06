# 공백변수 5종 — 다이소 세종 SDF 9종 실측 분포

**분석일**: 2026-05-19
**대상**: `data/tmp_inspect_260415/` 의 다이소 세종허브센터 SDF 9개
**원시 데이터**: `tmp_analysis_20260519/discretionary_raw.json`
**추출 스크립트**: `tmp_analysis_20260519/extract_discretionary.py`

이 문서는 박사논문이 정의한 5종 공백변수가 **실제로 어떻게 분포하는지** 다이소 세종허브센터 9개 SDF 파일에서 추출한 결과를 정리한다. 목적은:

1. "자의 판단" 변수와 "결정적/회사 표준" 변수를 데이터로 분리
2. 우리 코드(`phd_rules.py:115-211`)의 디폴트값이 실측과 일치하는지 검증
3. ④ 배관 라우팅 / ⑦ 재설계 루프 설계에 반영할 우선순위 도출

---

## 0. 분석 대상 SDF 인벤토리

| # | 파일명 | Title (시스템 식별자) | elev_span | 추정 압력존 |
|---|---|---|---|---|
| 1 | 1-1. 지상4층 창고 | `Daiso_SP_4F_WAREHOUSE` | 49.15 m | HSP 펌프 가압 |
| 2 | 1-2. 지상4층 랙(OSR) | `Daiso_SP_4F_RACK(OSR) Upper part` | 48.95 m | HSP 펌프 가압 |
| 3 | 1-5. 지상1층(Side Picking) | `Daiso_SP_1F (Side Picking)` | 29.15 m | MSP/펌프 |
| 4 | 2-1. 지하1층 펌프가압 (근거리) | `Daiso_SP_B1F` | 7.25 m | 펌프 직결 |
| 5 | 2-2. 지하1층 펌프가압 (원거리) | `Daiso_SP_B1F_Parking Area` | 7.25 m | 펌프 직결 |
| 6 | 3-1. 지원동 4층 | `Daiso_SP_Ancillary_4F` | 22.90 m | 별동 HSP |
| 7 | 3-2. 지원동 지하2층 | `Daiso_SP_Ancillary_B2F` | 9.85 m | 별동 LLSP |
| 8 | 4-1. 지상4층 랙(BUILDING-2) | `Daiso_SP_4F_RACK(BUILDING) Upper part` | 49.15 m | HSP 펌프 |
| 9 | 4-2. 지상1층 인랙 유량맞춤 | `Daiso_SP_1F_IN-RACK` | 45.55 m | MAX-Q 시나리오 |

9개 SDF는 한 건물의 **압력존 × 위치 × 특수 시나리오** 조합으로 박사논문 ①번 공백변수 구조와 일치.

---

## 1. 변수별 실측 분포

### ① 기준구역 (reference_zones) — **결정적, 자의 아님**

- 9개 모두 `<Design-options specification-type="remote-nozzle">`
- 9개 모두 title이 `zone + position + 특수태그` 패턴 (`Upper part`, `가장 가까운/먼`, `유량 맞춤`)
- 한 건물에서 9개 SDF로 분리 — 박사논문 "최대 12개 기준구역" 구조와 일치
- **결론**: ①번은 `floors[]` 입력만 주면 알고리즘으로 도출되는 결정적 변수.
  현재 코드(`generate_reference_zones`)가 정합. **수정 불필요**.

### ② 자연낙차 시작점 (natural_drop_start_floor) — **SDF 외부 변수**

- 9개 모두 `<Node io-node="Input" elevation="0.0">` (상대좌표)
- 즉 자연낙차 시작점 자체는 SDF에 안 박혀 있음
- 단 `elev_span` 패턴으로 **펌프 직결 vs 자연낙차 자동 분류 가능**:
  - span < 10 m → 펌프 직결 (지하층, 단일 층 SDF)
  - span 20~50 m → 다층 펌프 가압 또는 자연낙차
- **결론**: ②번 결정은 SDF 외부 메타(`floors[]`, `hb_case`)에 의존하는 게 맞음.
  현재 코드(`max(MSP floors)`)는 유지하되, **±1층 perturbation 후보**를 ⑦ 재설계 루프에 노출.

### ③ FX 신축배관 등가길이 — **회사 표준, 자의 아님**

```
n=28, 범위 22.40 ~ 22.40 m, 표준편차 0.00
등장 SDF: 3-1. 지원동 4층 (1개 SDF)
```

- 다이소 실측 단일값 = **22.4 m**
- 현재 코드 디폴트 `fx_equivalent_length_m = 0.6` → **30배 이상 차이, 즉시 교체 필요**
- 표준편차 0 = 다이소가 "회사 표준"으로 박아둔 값
- 22.4 m 해석: FX 본체 길이만이 아니라 **본체 + 양끝 피팅 + 곡률 등가길이 합산값**으로 추정
  - 일반적 PIPENET 검증값 `0.6 m`는 KSD 3507 20A 직선 호스만의 본체 길이
  - 22.4 = 본체(약 0.6~1.0) + 양쪽 어댑터·엘보·티 피팅 등가길이 약 20 m
- **사용 패턴**: FX는 지원동 4층 SDF에만 등장. 본동 본관/랙은 직결.
  → "FX 사용 여부"는 zone/용도 단위 결정. 사용 시 등가길이는 22.4 단일값.
- **결론**: 회사 프로파일(`designer_profile.fx_eq_length_m = 22.4`)로 분리.
  자의 변수 아님. ⑦ 루프에 안 들어감.

### ④ AV / PV 등가길이 — **진짜 자의 변수, 데이터로 확인됨** ✅

```
AV (n=7) : 12.0 또는 24.0 m, σ=6.41 (정확히 2값으로 분리)
PV (n=2) : 12.0 또는 24.0 m (각 1회씩)
PRV      : 9개 SDF 모두 등장 없음
```

| 등가길이 | SDF | 추정 원인 |
|---|---|---|
| **24 m** | 1-1, 1-2 (4F 창고/OSR), 1-5 (1F Side Picking), 4-1 (랙 BUILDING-2 PV) | 본관 + 상류 직류티·곡관 피팅 합산 |
| **12 m** | 2-1, 2-2 (지하1F), 3-1, 3-2 (지원동), 4-2 (인랙 PV) | 밸브 본체 단독 (2-1 SDF에 `"A/V 150A"` 명시) |

**관찰**:
- 정확히 2배 차이 → 같은 부품 다른 표기 방식
- 본동 랙·창고 = 24, 지하·지원동·인랙 = 12
- → 다이소 엔지니어가 SDF 작성 시 **상류 피팅을 AV 등가길이에 포함하느냐 분리하느냐**의 자의 판단
- 우리 코드 디폴트 `av_equivalent_length_m = 12.9` / `pv_equivalent_length_m = 10.1`은
  PIPENET 검증문서 값으로, **실제 다이소 분포(12 또는 24) 어느 쪽과도 불일치**

**결론**:
- AV/PV는 진짜 자유도. ⑦ 재설계 루프에 perturbation 후보 (12 ↔ 24) 노출 필요.
- 디폴트는 보수적인 24 채택 권장 (계산상 마찰손실 ↑ → 안전측).
- ChangeLog에 선택 이유 기록 (회사 관례·검토자 합의 등).

**PRV**: 9개 SDF에 등장 없음. 다이소 세종은 PRV 없는 설계.
→ 박사논문이 명시한 LSP존이 다이소엔 없거나, 다른 방식(자연낙차+오리피스)으로 해결.
별도 LSP 포함 건물 데이터 수집 필요.

### ⑤ 펌프 운전점 — **SDF 외부, 우리 설계 정합**

- 9개 모두 `<Pump>` 태그 **0개**
- 펌프 출력은 input 노드의 압력 (별도 메모로 부여)
- **결론**: ⑤번은 SDF 외부 메타. 펌프 사양서/카탈로그 별도 입력 채널 필요.
  현재 코드의 `decide_discretionary_variables(pump_rated_q_lpm=..., pump_rated_h_m=...)` 외부 주입 구조가 맞음.

---

## 2. 보너스 — 헤드 설계유량 자유도

```
9개 SDF 정확히 2값으로 분리: 80.0 또는 160.0 L/min
```

| 값 | SDF | 의미 |
|---|---|---|
| **80 L/min** | 2-1, 2-2 (지하1F), 3-1, 3-2 (지원동) | NFTC 2.2.1 최소값, P≥1 bar |
| **160 L/min** | 1-1, 1-2 (4F 창고/OSR), 1-5 (1F SP), 4-1, 4-2 (랙) | K=80 헤드, P=4 bar → Q=80·√4=160 |

**해석**: 랙·창고(고위험 용도)는 K=80 헤드 기준 압력 4 bar 설계. 일반 용도는 1 bar.
이건 자의가 아닌 **용도별 NFTC 룰**. NFTC 2.7.6 온도등급 + NFTC 103B ESFR 조합으로 결정.

→ 현재 코드는 헤드 K-factor를 `HeadSpec.k_factor_lpm_bar05`로 받지만,
**용도×K×P → 설계유량**의 자동 결정 함수가 추가되어야 함.

---

## 3. ⑦ 재설계 루프 설계 의사결정 매트릭스

| 변수 | 자유도 등급 | 자동화 전략 | ⑦ 루프 perturbation? |
|---|---|---|---|
| ① 기준구역 | L2 (결정적) | `generate_reference_zones` 유지 | ❌ 안 함 |
| ② 자연낙차 시작점 | L3 (정책) | 디폴트 = 최상 MSP, override 가능 | ✅ ±1층 스윕 (⑦E) |
| ③ FX 등가길이 | L1 (회사표준) | 다이소 프로파일 = 22.4 m (현재 0.6 → 교체) | ❌ 안 함 |
| ④ AV/PV 등가길이 | L4 (탐색) | 디폴트 24 (보수적), 12↔24 perturbation | ✅ (신규 ⑦F 추가 검토) |
| ⑤ 펌프 운전점 | L4 (탐색) | 외부 카탈로그 입력, 펌프 모델 후보 스윕 | ✅ (⑦C) |

박사논문의 5가지 대안 중:
- ⑦A 지름 / ⑦B 라우팅 / ⑦D 루프 = 규칙 내 자유도 (공백변수 아님)
- ⑦C 펌프 = 공백변수 ⑤
- ⑦E 고도 = 공백변수 ②
- **신규 ⑦F (AV/PV 등가길이) 검토 권장** = 공백변수 ④

---

## 4. 코드 액션 아이템

### P1 (디폴트값 교정 — 즉시)
- `phd_rules.py:136` `fx_equivalent_length_m: float = 0.6` → **22.4** (다이소 표준)
  - 단, 회사별 다를 수 있으므로 `designer_profile`에서 주입하는 게 더 안전
- `phd_rules.py:140-141` `av_equivalent_length_m=12.9` / `pv_equivalent_length_m=10.1`
  → **24.0** (보수적 디폴트) + perturbation 후보 `[12.0, 24.0]` 명시

### P2 (데이터 모델 확장)
- `DiscretionaryVariables`에 다음 필드 추가:
  ```python
  confidence: dict[str, float]   # 변수별 자동결정 신뢰도 0~1
  alternatives_considered: dict[str, list[Any]]   # ⑦ 루프 후보군
  source: dict[str, str]   # "designer_profile" | "auto_algorithm" | "user_override"
  ```

### P3 (designer profile 인프라)
- `data/cad_sdf_learning_profile.json`에 회사별 프로파일 섹션 추가:
  ```json
  {
    "daiso_sejong": {
      "fx_eq_length_m": 22.4,
      "av_eq_length_default_m": 24.0,
      "av_eq_length_alternatives": [12.0, 24.0],
      "head_design_flow_lpm": {"warehouse": 160, "office": 80},
      "natural_drop_offset_from_top_msp": 0
    }
  }
  ```

### P4 (⑦ 재설계 루프 재설계)
- 5가지 대안 중 **⑦C, ⑦E, (신규 ⑦F)**는 공백변수 perturbation
- **⑦A, ⑦B, ⑦D**는 규칙 내 자유도 — 알고리즘 분리
- 각 후보 평가 시 ChangeLog에 `(value, source, alternatives_considered, kpi_delta)` 기록

### P5 (데이터 보강 필요)
- **LSP존이 있는 건물 SDF 추가 수집** — PRV 등가길이 자유도 확인
- **다른 회사 SDF 수집** — FX 22.4가 다이소 표준인지 업계 표준인지 확인
- **펌프 사양서 PDF 수집** — ⑤번 자동화의 input 채널 확보

---

## 5. 수리계산 학습 노트 (사용자용)

이 분석을 진행하며 노출된 수리계산 핵심:

1. **Hazen-Williams 식**이 PIPENET의 마찰손실 기본 공식. C값 100/120이 다이소 SDF에 보임 (스틸/CPVC 차이).
2. **Q = K · √P** 공식이 헤드 유량 설계의 기본. K=80 헤드에서 P=4 bar이면 Q=160 L/min. SDF의 `Flow-define`은 이 Q를 직접 고정한 형태.
3. **등가길이**는 피팅의 마찰손실을 동일 직경 직관 길이로 환산한 값. AV/PV/PRV 등가길이가 자의 변수가 되는 건 "어디까지를 밸브로 보고 어디부터를 배관으로 볼 것인가"의 경계 판단 때문.
4. **remote-nozzle 사양**은 "가장 먼 헤드 기준으로 압력·유량을 만족시키는 설계". 박사논문 ①번 공백변수의 본질.
5. **자연낙차 vs 펌프 가압** = 고가수조에서 떨어지는 압력으로만 헤드 압력을 만들 수 있는 층(MSP)과 펌프 도움이 필요한 층(HSP)의 경계. ②번 공백변수.

---

**다음 단계 권장 순서**:
1. P1 디폴트값 교정 (1시간) — 즉시 안전 향상
2. P3 designer profile 인프라 (반나절) — 회사별 확장성
3. P2 데이터 모델 확장 + P4 ⑦ 루프 재설계 (1~2일)
4. P5 데이터 보강은 병행 진행
