# BLOCKED — ModuleA 2앵커 개선 작업 (작업지시서 §5 기록)

형식: (항목, 질문/애매점, 임시 우회 여부)

## 1. W4 — shapely 전제 오류
- **항목**: W4 (HeadRegion)
- **질문**: 작업지시서는 "shapely `Polygon` 기반 — extractor가 이미 shapely 의존"이라 명시하나,
  실제로는 shapely 미설치 + 저장소 전체에 shapely import 0건 (전제가 사실과 다름).
  신규 외부 의존성(shapely)을 추가할지는 배포 환경(anaconda system site-packages) 정책 판단 필요.
- **임시 우회**: 예 — 동일 API(`from_rects`/`from_polygon`/`contains`/`dilate`)의
  순수 파이썬 HeadRegion으로 구현. 추후 shapely 도입 결정 시 내부 구현만 교체 가능.

## 2. 회귀 기준선 — golden 1건이 작업 시작 전부터 불일치
- **항목**: §3 비트동일 회귀 증명 (전 항목 공통)
- **질문**: `tests/characterization` 실행 결과, 작업 시작 시점(HEAD=ae05799)에 이미
  `combined_build__plane_daemyeong` 이 golden과 불일치 (직전 커밋 ae05799 적응형 tol/
  라우팅 penalty의 의도된 변경 + golden 재생성은 "사용자 육안 확인 후"로 보류 중인 상태).
  stale golden 을 기준으로는 비트동일을 증명할 수 없음.
- **임시 우회**: 예 — 작업 시작 시점 HEAD의 실제 산출을 `data/_w_baseline/`(비추적)에
  동결하고, 각 W 항목 후 전 케이스를 이 baseline과 diff 하여 비트동일을 증명.
  golden 파일 자체는 수정하지 않음(사용자 확인 게이트 유지).

## 3. W1.2 — audit 기록 시점
- **항목**: W1 (미도달 헤드 보고)
- **질문**: "파이프라인 후단에서 audit(→W7)에 기록" — audit 객체(ExtractionAudit)는 W7에서
  신설되고, anchored 파이프라인 경로 자체도 W2·W3에서 형성되므로 W1 시점에는 기록할
  장소가 없음.
- **임시 우회**: 예 — W1에서는 계산 함수(`find_unreachable_region_heads`)까지 신설·테스트하고,
  anchored 경로/audit 배선은 W2~W7 진행에 맞춰 완성.

## 4. W3 — 도면 전제 불일치 → 브릿지 양단 W-창 한정
- **항목**: W3 (표적 브릿지)
- **질문**: 지시서 전제("세대별 배관망이 분리돼 있고 가짜 봉합은 브릿지에서만 발생")가
  실측과 부분 불일치 — 대명동 fixture 는 weld 후에도 세대 경계를 넘는 컴포넌트가 존재
  (최대 comp x=245k→284k, comp12 x=252k→262k). 헤드 보유 컴포넌트만 대상으로 좁혀도
  그 컴포넌트의 *동측* 지점에서 봉합되면 동측 우회 경로가 최종망에 유입됨
  (수용기준 FAIL 재현: 노드 (259436, -229271) 포함).
- **임시 우회**: 예 — §0 설계원리("앵커가 봉합 방향을 유도")에 따라 브릿지 양단 후보를
  작업창 W(=convex_hull(head_region ∪ {alarm_xy}) + ANCHOR_W_MARGIN_MM 팽창) 내부로
  한정하는 `within` 파라미터를 `bridge_targeted` 에 추가. 짝짓기 대상(comp(source) ↔
  헤드 보유 컴포넌트)·tol 계단·merge-and-reevaluate 는 지시서 그대로.
  head_region 이 `.pts` 를 노출하지 않으면 W 제한 없이 동작(W4 HeadRegion 이 정식화).

## 5. W5 — 음성 유형(SPLINE 등) 판정 전제와 파서 현실 불일치
- **항목**: W5 (공간한정 조건부 재선별)
- **질문**: 지시서는 재선별 범위에서 "SPLINE/ELLIPSE/… 음성 유형 승인 금지"를
  요구하나, 파서(parse_dxf_bundle)는 SPLINE/ELLIPSE 를 flattening 해 "PL" 타입
  dict 로 담으므로 파이프라인 입력(entity dict) 수준에서는 원 dxftype 을 구분할
  수 없음. entity dict 에 원천 태그를 추가하면 골든 characterization 의
  entities_sig(raw dict 해시)가 바뀌어 비-anchored 비트동일 요구를 위반.
- **임시 우회**: 예 — 재선별 헬퍼(collect_spatial_reselect_segments)가 DXF
  원본을 ezdxf 로 직접 스캔해 실제 dxftype 으로 승인/음성 판정. 이 때문에
  anchored 함수에 dxf_path 인자가 추가됨(재선별 발동 시에만 사용). 블록(INSERT)
  내부의 OTHER 선분은 이 스캔 범위 밖 — 저-prior 후보 특성상 수용.

## 6-1. W7 — bridges "layers" 필드는 그래프 수준에서 복원 불가
- **항목**: W7 (ExtractionAudit)
- **질문**: 스키마의 `bridges [{p1, p2, len_mm, layers}]` 중 `layers` — 브릿지는
  그래프(노드 좌표 + edge 길이만 보유) 수준에서 생성되므로 양단이 유래한 원
  entity 의 layer 정보가 이미 소실됨. 노드→layer 역매핑을 만들려면 _build_graph
  변경(골든 위험) 또는 중복 자료구조 신설이 필요.
- **임시 우회**: 예 — `layers: null` 로 스키마 자리만 채워 기록. 추후 필요 시
  _build_graph 에 선택적 node→layer 출력 kwarg 를 추가해 채울 수 있음.

## 6. fixture 파일명 불일치 (해소됨 — 기록용)
- **항목**: §0 검증 도면
- **질문**: 지시서의 `1__입력도면_대명동_단위세대_평면도.dxf` 는 저장소에 없음.
- **임시 우회**: 불필요 — 증거로 확정. `samples/dxf/대명동201동 단위세대_layer정리.dxf` 가
  지시서 실측 근거와 정확히 일치 (L4 SPLINE 11,770개, 범례 블록 `A$C60792707`
  @ (288201,−233417) / `A$C3F157AFD` @ (288201,−234617) 동일).
