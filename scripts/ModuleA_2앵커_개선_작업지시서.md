# Module A 개선 작업 지시서 — 2앵커(수동 밸브 + 헤드 영역) 기반 추출 재구조화

당신은 이 저장소의 유지보수 엔지니어다. 아래 지시를 **항목 순서대로, 항목당 커밋 1개씩** 수행하라.
각 항목은 수용 기준 테스트를 통과한 뒤에만 다음으로 진행한다. 애매한 지점은 임의로 구현하지
말고 BLOCKED로 기록한 뒤 다음 항목으로 넘어간다(§5).

---

## 0. 컨텍스트 (수정 전 필독)

- 대상 코드: `ModuleA/remote30_prototype.py`(주 파이프라인), `ModuleA/routes/r30_combined.py`(라우트),
  `ModuleA/sprinkler_remote30_extractor.py`(구형 경로), `ModuleA/core/remote30_constants.py`,
  `ModuleA/core/remote30_graph.py`.
- 검증 도면(fixture): `1__입력도면_대명동_단위세대_평면도.dxf` — 단위세대 5개 + 우측 범례 표.
  건축 전체가 `L4` 레이어(SPLINE 11,770개 포함), 소방은 `-소화(SP…)`/`SP 후렉시블` 계열.
- 현행 평면도 추출 순서(remote30_prototype.py):
  `filter_pipenet_only` → `detect_heads` → `_build_graph` → `collapse_parallel_ladders`
  → `_weld_dangling_endpoints` → `_bridge_components` 계단식(`for tol in (200.0, 500.0, 1000.0,
  2000.0, 5000.0, 10000.0)` 로 grep) → 헤드 drop line(`HEAD_BRIDGE_MAX_MM`)
  → `spt_source` 결정(`alarm_xy` canonical → `_nearest_graph_node` 폴백)
  → `_restrict_to_branch_region` → `force_spanning_tree` → `select_worst30_heads`.
- 확정된 진단: 신뢰 입력 두 개(`alarm_xy`, `branch_zones`)가 파이프라인 **끝**에서 사후 필터로
  소비된다. 그 앞 단계(전역 브릿지·영역 무관 헤드 검출·blind nearest 소스 결합)에서 오염이
  먼저 일어난다. 실측 근거:
  - 범례 표본 헤드 2개(블록 `A$C60792707` @ (288201,−233417), `A$C3F157AFD` @ (288201,−234617))가
    본망 최근접 배관 정점에서 각 3,858mm / 4,491mm — `HEAD_BRIDGE_MAX_MM`(5,000) 이내라
    팬텀 헤드로 부착 가능.
  - 인접 세대망 간 거리가 10,000mm 계단식 브릿지 사거리 이내 — 세대 간 가짜 봉합 가능.
  - 소스 좌표의 blind nearest가 606mm 거리의 실제 2차 배관 대신 102mm 거리의 고립
    노이즈 조각(layer 0 지시선)에 붙을 수 있음.
- 설계 원칙(이번 작업의 헌법): **앵커가 수집 범위(W), 봉합 방향, 승인 기준, 실패 정의를 유도한다.**

## 1. 신규 계약 — anchored 모드

- `anchored=True` 모드 신설: `alarm_xy` **필수**, `head_region` **필수**(다각형; 사각형 union
  입력도 다각형으로 승격). 둘 중 하나라도 없으면 `anchored` 진입 불가 — 명시적 에러.
- **anchored 미사용 시 기존 동작을 비트 단위로 보존한다(골든 불변).** 모든 신규 로직은
  anchored 분기 안에서만 발동한다.
- 작업창 자동 유도: `W = convex_hull(head_region ∪ {alarm_xy})` 를 `ANCHOR_W_MARGIN_MM`
  (신규 상수, 기본 3000.0, `remote30_constants.py`에 정의)만큼 팽창(dilate)한 다각형.
  사용자에게 추가로 묻지 않는다.

## 2. 작업 항목 (순서 고정)

### W1. `detect_heads` 영역 게이트 + 미도달 헤드 보고
1. `detect_heads(..., region: HeadRegion | None = None)` 파라미터 추가. `region` 지정 시
   **최종 승인 후보에 한해** point-in-region 판정을 적용한다. R1~R5 신호 계산·클러스터링
   로직 자체는 수정 금지. `region=None`이면 현행과 동일.
2. 파이프라인 후단에서 "region 안이지만 source에서 도달 불가"인 승인 헤드 목록을 산출해
   audit(→W7)에 기록한다. 조용한 drop 금지 — anchored 모드에서 미도달 헤드는 추출 결함 신호다.

**수용 기준**
- fixture + 서쪽 세대 1곳을 덮는 region 다각형(테스트 안에 하드코딩) 실행 시,
  (288201,−233417)·(288201,−234617)의 두 INSERT가 승인 헤드 목록에 **없음**.
- `region=None` 실행 결과가 변경 전과 동일(회귀).

### W2. 소스 결합 규칙 교체
1. 신규 함수 `attach_source(alarm_xy, graph, comp_of, accepted_heads, edge_len, audit)`:
   - 1순위 후보 = region 내 승인 헤드를 1개 이상 보유한 컴포넌트의 노드들. 그중 최근접에 부착.
   - 1순위가 `SOURCE_BRIDGE_MAX_MM` 내에 없으면 거리 상한 안에서 순차 완화하되,
     선택 근거(거리, 컴포넌트 헤드 수, 완화 단계)를 audit에 기록.
2. anchored 모드에서 기존 `_nearest_graph_node` 단독 폴백 사용 금지.

**수용 기준(합성 단위 테스트)**
- 소스 좌표에서 102mm 거리에 헤드 없는 2노드 고립 조각, 606mm 거리에 승인 헤드 보유
  본망이 있는 인공 그래프에서 본망에 부착됨.

### W3. 표적 브릿지 (전역 계단식 대체)
1. anchored 모드에서 전역 `_bridge_components` 계단식 호출을 다음으로 대체:
   봉합 후보 쌍 = `comp(source)` ↔ `{region 내 승인 헤드 보유 컴포넌트}` 만.
   기존과 동일한 tol 계단(200→10,000), 병합 후 재평가 루프, `bridge_edges_out` 유지.
2. `force_connect`(무제한 봉합)는 anchored 평면도 경로에서 호출 금지 — 기계실 경로 전용임을
   호출부 주석으로 명시.

**수용 기준**
- fixture + 서쪽 세대 region 실행 시 최종망(SPT 후)에 다른 세대의 노드가 0개.
- 사용된 모든 bridge edge의 양단이 comp(source) 성장 이력에 속함(audit로 검증).

### W4. 영역 표현 통일 — `HeadRegion`
1. `HeadRegion` 신설(shapely `Polygon` 기반 — extractor가 이미 shapely 의존):
   `from_rects(list[tuple])` / `from_polygon(list[xy])` / `contains(pt)` / `dilate(mm)`.
2. `branch_zones`(rect list)와 extractor의 `zone_bbox` 소비처를 `HeadRegion` 참조로 통일.
   라우트(`r30_combined.py`의 `branch_zones` 파싱부)는 기존 rect 입력을 `from_rects`로 승격.
3. `_restrict_to_branch_region`의 **내부 의미론 변경 금지**(밖=corridor 1개, 안=루프 보존,
   도달 불가 시 no-op). 입력 타입 교체와 in_region 판정 위임만 수행.

**수용 기준**
- 기존 rect 입력 경로 회귀 통과. L자형 다각형 region 단위 테스트에서 사각형 union이
  물었을 이웃 노드가 제외됨.

### W5. 공간한정 조건부 재선별 (플래그, 기본 off)
1. 발동 조건: anchored 모드 **그리고** 1차 명목 수집(`filter_pipenet_only`) 결과로
   region 내 승인 헤드에 도달하는 망을 구성하지 못한 경우에만. (S140 조건부 재선별의
   공간 한정 실시예 — 코드 주석에 이 문구로 근거를 남길 것.)
2. 재선별 범위: **W 내부**의 `OTHER` 카테고리 레이어에서 LINE/ARC/LWPOLYLINE만
   저-prior 후보로 승인. SPLINE/ELLIPSE/3DFACE/DIMENSION/HATCH/닫힌 PL(`CLOSED_PL_TOL_MM`)은
   음성 유형으로 승인 금지.
3. 승인된 비명목 edge는 태깅되어 audit의 비명목 점유율 집계에 들어간다.

**수용 기준**
- 플래그 off 시 fixture 결과 불변. 레이어명을 무의미 문자열로 전부 치환한 변조 fixture
  (테스트에서 ezdxf로 생성)에서 플래그 on 시 추출이 성공하고, SPLINE은 후보에 0개.

### W6. 관경 텍스트 위생
1. `sprinkler_remote30_extractor.py`의 naive 관경 정규식(`\b(20|25|…)\b`)을 제거하고
   `remote30_constants`의 `_DIA_TEXT_PATTERNS` + `_DIA_TEXT_NOISE_KW` + `_VALID_DIA_MM`
   참조로 교체(구현 이원화 청산).
2. anchored 모드에서 관경 텍스트 후보를 W 내부로 제한.

**수용 기준**
- 단위 테스트: `'NO.20'` 이 관경 후보로 잡히지 않음. `'25A'`, `'Ø65'`, 순수 `'50'`은 잡힘.
- 범례 좌표대(x≈287,000~291,000)의 관경 문자가 W 밖일 때 후보에서 제외됨.

### W7. `ExtractionAudit` 리포트
1. dataclass 신설:
   `heads {detected_in_region, attached, unreachable: list[xy]}`,
   `bridges [{p1, p2, len_mm, layers}]`, `welds [...]`, `head_drops [...]`,
   `nonnominal {edge_count, len_mm, ratio}`, `corridor {node_count, len_mm}`,
   `source_attach {dist_mm, method, escalation}`.
2. 파이프라인 반환값에 포함하고 `r30_combined.py` 응답 JSON에 직렬화한다.
   프론트 렌더링은 이번 범위 외 — JSON 계약까지만.
3. 기존 `weld_edges`/`bridge_edges`/`head_drop_edges` 자료구조를 재사용해 채운다(중복 계산 금지).

**수용 기준**
- fixture anchored 실행 시 audit JSON이 스키마대로 채워지고, W3의 bridge 검증이 이 audit로 수행됨.

## 3. 금지·보존 목록

- 값·로직 변경 금지: `SNAP_TOL_MM`(50), `MIN_PIPE_EDGE_MM`, `HEAD_BRIDGE_MAX_MM`,
  `SOURCE_BRIDGE_MAX_MM`, R1~R5 헤드 시그니처 규칙(게이트 추가만 허용), `KNOWN_HEAD_BLOCKS`,
  `collapse_parallel_ladders`, `_weld_dangling_endpoints`, `force_spanning_tree`,
  `_restrict_to_branch_region` 내부 로직.
- 전면 리포맷, 파일 이동/리네이밍, 공개 함수 시그니처 파괴 금지 — keyword 인자 추가만 허용.
- anchored 미사용 경로의 산출물은 **비트 동일**해야 한다. 회귀 스크립트로 증명할 것.

## 4. 검증 체계

- `tests/test_anchored_extraction.py` 신설. fixture는 위 DXF. 최소 assert 목록:
  - 레이어 분류: `-소화(SP가지관)`→PIPE, `소화기CO2`→EXCLUDE, `L4`→OTHER(ARCH 아님),
    `SHEET-TEXT`→TEXT, `1.68℃하향식`→HEAD.
  - 헤드층 원시 CIRCLE(r=52.6mm) 9개가 R2로 승인됨.
  - W1·W3·W5·W6의 각 수용 기준.
- 회귀: anchored 미지정 전체 실행의 주요 산출(JSON/SDF/KFP) diff 없음.
- 커밋 메시지 형식: `[W#] 제목 — 수용기준: PASS/FAIL(사유)`.

## 5. 진행 규칙

- 항목 순서 준수, 항목 병합 금지. 판단이 필요한 애매점은 구현하지 말고 `BLOCKED.md`에
  (항목, 질문, 임시 우회 여부)로 기록 후 다음 항목 진행.
- 이 변경으로 특허 상세서(S100~S600 문서)의 문구와 코드가 어긋나게 되는 지점이 생기면
  해당 코드에 `# [문서정합]` 주석을 남기고 최종 보고서에 목록화한다(문서 수정은 범위 외).
- 최종 산출물: 변경 diff 요약, 테스트 결과 전문, `BLOCKED.md`, `[문서정합]` 목록.
