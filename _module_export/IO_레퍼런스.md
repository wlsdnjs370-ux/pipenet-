# 도메인 서버 — 배관망 추출 / 검진 모듈 I/O 레퍼런스

> 생성 기준: `JupyterProject-domain` 저장소. 소스 전체는 같은 폴더의
> `배관망_추출_모듈_전체소스.py.txt`, `검진_모듈_전체소스.py.txt` 참조.
> 모든 경로는 도메인 저장소 기준 `파일:행` 표기.

---

## 1) 배관망 추출 모듈

### 구성 파일
| 파일 | 역할 |
|------|------|
| `remote30_prototype.py` | 파이프라인 진입·오케스트레이션, SDF/KFP/HAS emit 래퍼 |
| `remote30_full_network.py` | 전체망 그래프 구성(노드/파이프 dict 생성) |
| `sprinkler_remote30_extractor.py` | 헤드 선정·수리계산(하젠-윌리엄스)·테이블 빌드·아이소 PNG |
| `remote30_ml.py` | 타일 기반 ML 헤드 검출(보조) |
| `head_detector.py` | 템플릿/YOLO 헤드 검출(보조) |

### 엔드포인트 (`대조 서버.py`)
- `POST /api/remote30/extract` — 추출 실행 (`대조 서버.py:6132`)
- `POST /api/remote30/auto_process` — DXF→최종 산출까지 원샷 (`대조 서버.py:6550`)

### 입력 (INPUT)
- **DXF 도면 파일** 업로드 또는 토큰 참조 (`대조 서버.py:6135, 6145`)
- 주요 파라미터:
  - `auto_detect_alarm` / `alarm_x,alarm_y` — 알람밸브(현 "밸브") 자동/수동 좌표
  - 레이어 키워드 필터(PIPE/HEAD/TEXT), `snap_tol`, `head_to_pipe_tol`, `diameter_text_search_radius`
  - `elevation_alarm_m`, `elevation_head_m` — 층고/헤드 표고(m)
  - `k_factor` — 노즐 K계수, `design_flow_per_head_lpm` — 헤드당 설계유량(L/min)
  - `remote_head_count`(기본 30), `remote_mode` = `"length"` | `"hydraulic"`
  - `natural_fall_height_m` — 자연낙차 높이
  - `fallback_dia_mm`, `cad_unit_to_m`(0.001=mm→m)
  - `emit_sdf`, `emit_csv`, `zone_bbox`(관심영역), override heads/pipes(JSON)

### 처리 단계 (PROCESSING)
1. **DXF 파싱** — ezdxf 모델스페이스 원시 엔티티 추출
2. **Pipenet 전용 필터** — PIPE/HEAD/TEXT 카테고리만, CAD 숨김레이어 컷
3. **헤드 검출 + 그래프 구성** — epsilon-cluster 노드화, 무향 그래프, 알람밸브 자동식별, Dijkstra 로 최악단 K개 헤드 선정 (`sprinkler_remote30_extractor.py:1300+`)
4. **관경 추정** — TEXT 정규식 + NFPC 공칭경 룩업 (`remote30_prototype.py:3137`)
5. **수리손실 계산**(mode=hydraulic) — 관별 하젠-윌리엄스 마찰 + 헤드별 정수두
   - `hw_friction_loss_kgcm2()` (`sprinkler_remote30_extractor.py:623`)
   - 공식: `dp_mpa = 6.174e4 · Q(lpm)^1.85 · L(m) / (C^1.85 · d(mm)^4.87)` → kgf/cm²
6. **테이블 빌드** — Nodes/Pipes/Nozzles/Fittings/Equipment 5종 (`sprinkler_remote30_extractor.py:805`)
7. **Export** — SDF/KFP/HAS/CSV/XLSX + 아이소 PNG

### 출력 (OUTPUT / 산출물)
| 산출물 | 함수 | 비고 |
|--------|------|------|
| `.sdf` | `emit_sdf()` (`remote30_prototype.py:3538`) | PIPENET 스프링클러 XML |
| `.kfp` | `emit_kfp()` (`remote30_prototype.py:3824`) → `kfp_sdf_converter` | **4.0-NFPA13-EQL** (방금 수정, 솔버 4.0.0 호환) |
| `.has` | `emit_has()` (`remote30_prototype.py:3858`) | HASS 통합형 |
| `.slf` | `resolve_standard_slf()` (`remote30_prototype.py:103`) | 표준 라이브러리 동봉 |
| `.csv ×4` | emit_csv 플래그 | Nodes/Pipes/Nozzles/Fittings |
| `.xlsx` | `build_pipenet_tables()` | 4시트 워크북 |
| **아이소 PNG** | `_plot_extracted_isometric()` (`sprinkler_remote30_extractor.py:768`) | 선정 헤드 + 경로 엣지 등각 다이어그램 |

**수리 산출 수치(관/헤드별, 브라우저 프리뷰 JSON `대조 서버.py:6238`):**
- 관별: `friction_loss`, `base_length_m`, `velocity_mps`, `nominal_bore_mm`
- 헤드별: `Friction Loss(kgcm2)`, `Static Head(kgcm2)`, `Total Loss(kgcm2)`, `Path Length(m)`
- 프리뷰: `selected_heads_xy`, `path_edges_xy`, `sdf_tables`, 각 다운로드 URL

> **주의:** 추출 모듈 자체는 마찰손실을 **수치로만** 산출하고 그래프는 그리지 않음.
> 아이소 PNG 에도 마찰 데이터는 표시되지 않음. "마찰손실 그래프"는 아래 **검진 모듈** 산출물.

---

## 2) 검진 모듈

### 구성 파일
| 파일 | 역할 |
|------|------|
| `pipenet_validator.py` | 메인 검증 클래스 `PipenetGuideValidator` (`:215`, `.validate()` `:246`) |
| `pipenet_validator_v4.py` | v4 래퍼 — 12시나리오 집계·3소스 트레이스(NFTC+HB+PhD), PIPE.001~006 불변 |
| `nftc_rules.py` | 국가화재안전기준(NFTC) 결정트리 정의 |
| `phd_rules.py` | 박사논문 규칙 — 압력존(HSP/MSP/LSP/LLSP), 공백변수 5종, 12시나리오, 불균형지표(ΔP·CV·τ) |
| `hb_rules.py` | 핸드북 규칙 — 시스템 유형·유속한계·처언압·재질선정 |

### 엔드포인트 (`대조 서버.py`)
- `POST /api/validate` — 검진 실행 (`대조 서버.py:2283`)

### 입력 (INPUT)
- **설계계산서**: DOCX 또는 PDF (`report_file`, 필수)
- **SDF**: 계통도/토폴로지 (`sdf_file`, 선택)
- 파싱 추출 항목(`pipenet_validator.py:246`): 설계정보(공식·재질·가용경), 관구성행(라벨·입출력노드·구경·길이·표고·C계수·피팅등가장), 설계관행, 노즐구성(K·필요유량·압력범위), 유입유량, 펌프유량, 관유동(입출압·마찰손실·유속·유량), 노즐유동(입구압·필요vs실제유량·편차), 장비(FX/AV/PV), 탄성밸브

### 검사 규칙 (CHECKS)
- **PIPE.001~006 하드룰** (`pipenet_validator.py:816-975`):
  - 001 계통 토폴로지 일치 / 002 고압재질(≥1.2MPa→KSD3562) / 003 C계수(CPVC 150, 강관 120) / 004 세대내부 CPVC / 005 세대인입 ≤65A / 006 헤드수 기준 최소구경
- **하젠-윌리엄스 마찰 재검산** (`:267`) — 계산 vs 보고값, `hw_ok` 판정
- **유속검사** (`:269`) — 가지배관 ≤6.0 m/s, 그외 ≤10.0 m/s
- **노즐 유량·압력** — 최소유량 ≥80 L/min, 압력 0.1~12.0 MPa
- **장비 등가장** — FX 13~21m, AV 12.9±0.1, PV 10.1±0.1
- **탄성밸브 압력강하** — ±0.05 kgf/cm² 허용
- NFTC/HB/PhD 규칙은 트레이스·시나리오 집계에 활용

### 출력 (OUTPUT) — `/api/validate` JSON (`대조 서버.py:2303`)
- `summary`: pass/fail/warning 카운트
- `results`: PASS/FAIL/WARNING 메시지 분류
- `stats`(= "검진 통계"): 관수·노즐수·최소노즐유량·최소노즐압·최대유속(가지/주)·HW 검사수·HW 실패수·HW 최대오차 등
- `tables`: pipes/nozzles/equipment/valves 상세행 (관별 마찰손실·HW기대값·유속·PIPE.00x 결과 포함)
- `rules.pipe`: PIPE.001~006 + NFTC/HB/PhD 트레이스
- `insights`: 엔지니어링/경제성 조언
- `visualizations`: **그래프 3종** (아래)

### 그래프 (CHARTS) — matplotlib PNG(base64) 임베드
1. **관 유속 vs 한계** (`대조 서버.py:836`)
   - X=배관 라벨, Y=유속(m/s), 파란선=실제·빨간점선=한계, 토폴로지별(가지/주) 위반 검출
2. **노즐 압력-유량 산점도** (`대조 서버.py:859`)
   - X=입구압(kgf/cm²G), Y=실제유량(L/min), 녹색=≥80·빨강=<80, 기준선 80 L/min·1.0 kgf/cm²
3. **★ 마찰손실 비율 막대그래프** "Friction Loss Ratio by Pipe" (`대조 서버.py:1000`)
   - **X = 배관 라벨, Y = 마찰손실 ÷ 길이 (kg/cm²/m)**
   - 파란막대=엔지니어링 플래그 관, 회색=그외
   - 빨간 점선 = Threshold 1.00
   - 빨간 ▽ 마커 = "Sharp Increase"(급증) 이상점
   - 급증 판정: `ratio > 1.0 AND delta > max(1.0, prev_ratio × 0.75)` (`대조 서버.py:927-998`)
   - 각 급증점마다 원인분석 카드(장관·고유속·피팅집중 등) + 권고 생성

---

## 요약: "인풋 → 아웃풋" 한눈에

| 모듈 | 인풋 | 핵심 아웃풋 | 그래프 |
|------|------|-------------|--------|
| **추출** | DXF + 층고/K/유량 파라미터 | SDF·KFP·HAS·CSV·XLSX + 아이소PNG + 관/헤드별 마찰손실·유속·경로길이 수치 | 아이소 PNG(마찰 미표시) |
| **검진** | 설계계산서(DOCX/PDF) + SDF | pass/fail·통계·상세테이블·PIPE.001~006 판정·조언 | 유속 vs 한계 · 노즐 압력-유량 · **마찰손실 비율(급증검출)** |
