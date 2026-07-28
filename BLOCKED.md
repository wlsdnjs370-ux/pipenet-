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

## 4. fixture 파일명 불일치 (해소됨 — 기록용)
- **항목**: §0 검증 도면
- **질문**: 지시서의 `1__입력도면_대명동_단위세대_평면도.dxf` 는 저장소에 없음.
- **임시 우회**: 불필요 — 증거로 확정. `samples/dxf/대명동201동 단위세대_layer정리.dxf` 가
  지시서 실측 근거와 정확히 일치 (L4 SPLINE 11,770개, 범례 블록 `A$C60792707`
  @ (288201,−233417) / `A$C3F157AFD` @ (288201,−234617) 동일).
