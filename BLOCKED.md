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
- **후속 (2026-07-29, P1)**: 가짜 봉합의 원인이 브릿지 하나가 아니었음이 확인됨.
  `_weld_dangling_endpoints` 는 노드↔노드만 이을 수 있어, 가지관 끝점이 주배관
  *중간* 에 닿는 실제 T분기를 구조적으로 복원할 수 없었다. 그 끝점들은 전방
  콘 안의 엉뚱한 노드(주배관 반대편 끝, 다른 조각)에 붙어 오접합이 됐다.
  대명동 서측 세대 실측: 그런 후보 47건 중 32건이 오접합, 허위 배관 10.7m.
  해소: `_split_tee_branches`(P1) 를 weld *이전* 에 두어 edge 를 제자리에서
  쪼갠다. 배출망 79.25m → 73.54m, weld 364 → 299, 헤드 25/25 유지.
  잡힌 148건의 갭 중앙값 0.0mm / 최대 0.1mm — 전부 정확한 CAD T분기였다.

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

## 7. 유속 수렴 사이징 — 골든 1건 재생성 보류 + 피팅 등가길이 미모델링
- **항목**: 통합 유속 수렴 (`core/hydraulic_solver.py`, V1~V6)
- **질문 1**: 내경 산정 근거가 "수요유량 1패스"에서 "과토출 유량 수렴"으로 바뀌어
  `combined_build__plane_daemyeong` 골든의 `dia`/`velocity_mps` 가 필연적으로 달라짐.
  골든 재생성은 §2 와 같은 이유(사용자 육안 확인 게이트)로 아직 못 함.
- **임시 우회**: 예 — 골든 파일 미수정. 대명동 201동 실측으로 무해함을 확인:
  `changed=0` (기존 내경 그대로 통과), 위반 0건, 과토출 1.27배, 소스압 2.10 bar.
  즉 이 도면에서는 내경이 하나도 안 바뀌고 stamp 값(`velocity_mps`)만 과토출 기준으로
  갱신됨 (max 1.92 → 1.97 m/s).
- **질문 2**: `<Fittings>` 등가길이를 해석에 넣지 않음(equipment 의 명시 `eq_len` 만 반영).
  손실 과소 → 헤드 간 압력차 과소 → 과토출 과소 평가 = **비보수측**.
- **임시 우회**: 예 — 사이징 판정에만 `safety=1.1` 유량 할증으로 상쇄. 피팅 등가길이
  테이블을 해석에 직접 넣는 것은 별도 과제.
- **질문 3**: 극단적으로 얇은 초기망(전 구간 25A 등)은 최원단 0.1 MPa 를 물리적으로
  만족시킬 수 없어 소스압이 폭주(실측 1772 bar)하고 고정점 반복이 발산함.
- **임시 우회**: 예 — 발산 회차에는 유량의 **크기**를 믿지 않고 방향("너무 얇다")만
  써서 한 치수씩만 승급. 대명동 백지설계 기준 최종 내경이 150/200A 일색 → 32~100A
  등급 분포로 개선됨 (승급 1279회 → 316회).
- **질문 4 (domain-slim 포팅)**: domain-slim 에만 있는 `size_combined_bores`
  (역할별 내경: 평면도=규약 유지 / 입상관=단일 균일경 / 기계실=한 단계 굵게)와
  유속 수렴이 같은 값을 두 번 정한다. 두 정책의 우선순위가 지시서에 없음.
- **임시 우회**: 예 — 수렴 루프를 `size_combined_bores` **뒤**에 두고
  `keep_existing=True` 로 실행. never-shrink 라 역할별 정책은 보존되고, 그 결과가
  유속을 못 견딜 때만 승급된다. 대명동 실측 `changed=0` (역할별 배정이 이미
  유속을 만족) — 두 정책이 충돌하지 않음을 확인.

## 6. fixture 파일명 불일치 (해소됨 — 기록용)
- **항목**: §0 검증 도면
- **질문**: 지시서의 `1__입력도면_대명동_단위세대_평면도.dxf` 는 저장소에 없음.
- **임시 우회**: 불필요 — 증거로 확정. `samples/dxf/대명동201동 단위세대_layer정리.dxf` 가
  지시서 실측 근거와 정확히 일치 (L4 SPLINE 11,770개, 범례 블록 `A$C60792707`
  @ (288201,−233417) / `A$C3F157AFD` @ (288201,−234617) 동일).
