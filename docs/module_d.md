# 모듈 D — PIPENET 결과를 PIPENET 없이 출력한다

대상 코드: `core/d_display_model.py`(D1), `d_result_binder.py`(D2),
`d_iso_renderer.py`(D3), `d_label_layout.py`(D4), `d_report_typeset.py`(D5),
`d_batch.py`(D6), `routes/d_output.py` + `templates/d_output.html`(웹 UI).
근거 지시서: `samples/모듈D_작업지시서.md`.

입력은 `.sdf`(도형·표시 스킴)와 결과 `.xml`(XDSET) 두 장이다. `.slf` 는 쓰지 않는다.
출력은 ISO 도면 PDF 와 수리계산서 PDF, 그리고 둘을 묶은 합본이다.

## 1. 왜 PIPENET 을 띄우지 않는가

지시서 §7-1 이 GUI 자동화·COM·프로세스 기동을 전부 막았다. 라이선스 좌석과
사람 클릭에 묶이면 일괄 처리가 성립하지 않는다. 그래서 SDF 와 XML 만으로
완결한다 — 이 저장소 어디에도 PIPENET 실행 코드는 없다.

같은 이유로 표제란에 `PIPENET Schematic` 을 찍지 않는다(§7-2). 도면 아래에
생성 출처를 명시한다: "PIPENET 이 계산한 결과를 옮겨 그린 도면이다. 값은 결과
XML 원문이며 이 프로그램이 계산하거나 보간하지 않는다."

## 2. 값은 어디서 오는가 — 계산하지 않는다는 말의 뜻

- **수치는 전부 결과 XML 원문이다.** 추정·보간이 없다(§7-4). XML 에 없는 값은
  빈칸으로 두고 그 라벨 이름을 리포트에 **전량** 싣는다. 개수만 세지 않는다.
- **단위 환산 계수는 `<Units>` 블록에서 읽는다**(§7-5). 하드코딩이 없다.
  선언된 계수와 실측이 어긋나면 **실측을 쓰고 그 사실을 경고로 남긴다** —
  대명동 샘플에서 `flowrate` 선언 `lit/min` 60000 대 실측 1 같은 경우가 나온다.
- **원본 파일을 고치지 않는다**(§7-3). 표시 스킴을 바꿔도 사본에만 쓴다.
- **워드를 거치지 않는다**(§7-6). 계산서는 matplotlib PdfPages 로 직접 조판한다.

## 3. SDF 와 XML 을 어떻게 맞추는가 — 결손 처리 정책

라벨로 맞춘다. 한쪽에만 있는 라벨은 **교집합에서 빼고 이름을 전량 나열한다**
(사용자 확정 (B)안). `JoinReport` 가 9 갈래로 나눠 담는다:

```
sdf_only_nodes / xml_only_nodes        # 도형에만 / 결과에만 있는 노드
sdf_only_pipes / xml_only_pipes
sdf_only_nozzles / xml_only_nozzles
divergent_pipes                        # 양쪽에 있으나 양끝 노드가 어긋난 관로
duplicate_labels                       # 한 파일 안에서 중복된 라벨
matched (bool) / summary() (사람이 읽는 한 줄)
```

`matched` 는 **개수가 아니라 bool** 이다 — 전량 일치면 `True`. 화면에는
`summary()` 를 띄운다.

## 4. 표시 항목은 우리가 정하지 않는다

지시서 §3.4: "SDF 의 `Link-schemes` / `Node-schemes` 자식 요소가 곧 선택지다."
그래서 `LINK_ITEMS` 15종 · `NODE_ITEMS` 4종을 SDF 스키마에서 그대로 옮겼고,
UI 드롭다운·배치 프리셋·라우트 검증이 **이 한 벌만** 본다.

프리셋 3종(`d_iso_renderer.PRESETS`):

| 이름 | 관로 표시 | 노드 표시 |
|---|---|---|
| 유량본 | Pipe volumetric flow | None |
| 압력본 | Pipe pressure difference | Node pressure |
| 압력본_옥내소화전 | Pipe bore | Node pressure |

**어느 프리셋을 쓸지는 자동 판별하지 않는다.** 코퍼스 260 세트 실측에서 SDF 의
`link_scheme` 은 마지막 편집 상태일 뿐 설비 종류와 무관했고(옥내소화전 안에서만
`Pipe type` 80 · `None` 55 · `Pipe volumetric flow` 22), 노즐 유무로도 갈리지
않았다(260 세트 전량 노즐 보유). 호출자가 지정한다.

## 5. `show_labels` 의 실제 의미

**이름표(태그)만 끈다.** 고른 표시 항목의 값 글자는 그대로 남는다
(`d_iso_renderer.py` 의 `if show_labels: parts.append(label)`).
글자를 하나도 남기지 않으려면 이름표를 끄고 표시 항목도 둘 다 `None` 으로 둔다.
UI 문구를 "라벨"이 아니라 "이름표"로 적은 이유다.

## 6. 수치 일치 — 실측 (§2.2 수용 기준)

대명동 단위세대 자동화 세트(관로 136 · 노드 137 · 노즐 30)를 PIPENET 이
워드로 내보낸 계산서와 셀 단위로 대조했다:

```
PIPE CONFIGURATION               136행 / 136행   불일치 0
FLOW IN PIPES                    136행 / 136행   불일치 35   ← 전부 자릿수
NOZZLE CONFIGURATION              30행 /  30행   불일치 0
FLOW THROUGH NOZZLES              30행 /  30행   불일치 1    ← 자릿수
DESIGNED DIAMETERS & FLOWRATES   136행 / 136행   불일치 0
WARNINGS  유속 초과 2 / 최소압 미달 18            판정 불일치 0

자릿수 불일치 36 건 / 설명 못 한 불일치 0 건
```

36 건은 전부 **표시 자릿수 경계**다. 어긋난 폭이 XML 자체 정밀도 안에 있음을
같이 재 두었다 — Friction loss 최대 0.009687(XML 정밀도 0.01), Inlet-Outlet
pressure 최대 0.009687(정밀도 1), 노즐 Calculated flow 최대 1e-11(정밀도 1e-11).
**값이 다른 것이 아니라 반올림 경계에서 마지막 자리가 갈린다.**

## 7. 웹 UI — 한 번 읽고 여러 번 그린다

§2.2 는 "표시 항목 변경이 **재계산 없이** 즉시 반영된다"를 수용 기준으로 둔다.
그래서 `/api/module-d/load` 가 한 번만 파싱해 모델을 잡에 붙잡아 두고,
`/api/module-d/draw` 는 `job["bound"] or job["model"]` 을 재사용한다. 실측
재도시 0.5~1.6초(관로 136 규모), 계산서 18쪽·649행·본문 8.5pt, 합본 3조각
20쪽 4~5초.

| 경로 | 하는 일 |
|---|---|
| `GET /module-d-output` | 워크벤치 화면 |
| `POST /api/module-d/load` | SDF(필수) + 결과 XML(선택) → 잡 개설. **여기서만 파싱한다** |
| `POST /api/module-d/draw` | 표시 항목을 여러 개 골라 관로×노드 조합마다 ISO 한 장씩 (낱장 PDF + 미리보기 PNG, 두 장 이상이면 합본도). 한 번에 24장까지 |
| `POST /api/module-d/report` | 수리계산서 한 부. 결과 XML 이 없으면 400 |
| `POST /api/module-d/book` | 계산서 + 프리셋 도면들을 한 권으로 |
| `GET /api/module-d/result/<job>/<file>` | 산출물 내려받기 |

**세트 하나씩이다.** 폴더째 던지는 일괄 처리는 D6 의 CLI(`python -m core.d_batch`)
가 맡는다 — 지시서가 "웹 UI 는 마지막, 그 전까지는 CLI 로 충분하다"고 못박았고,
브라우저로 수백 세트를 올리는 것보다 폴더를 걸어 두는 편이 실제로 빠르다.

## 8. 정직하게 실패한다

- **결과 XML 없이도 도면은 나온다.** 단 SDF 만으로 되는 항목(관경·길이·라벨)에만
  값이 붙고, 그 사실을 화면에 적는다. 계산서는 거절한다(400).
- **겹치는 라벨은 겹친 채 두지 않고 그리지 않는다.** 떨어뜨린 라벨 이름을 리포트에
  싣는다(D4).
- **합본 파일이 없으면 조각이 몇 장 나왔든 "묶었다"고 하지 않는다.** 일부 조각만
  성공한 경우와 합본 자체가 실패한 경우를 구분해 보고한다.
- **일괄 처리는 세트 하나가 터져도 나머지를 계속 간다.** 실패 사유는 세트별로
  리포트에 남는다(§2.2).

## 9. 적용 한계

| 경우 | 왜 안 되는가 | 어떻게 하는가 |
|---|---|---|
| 프리셋 자동 선택 | SDF 의 저장된 스킴이 설비 종류와 무관하다(§4 실측) | 호출자가 지정한다. 자동 추론을 넣지 않는다 |
| 결과에만 있고 도형이 없는 라벨 | 그릴 좌표가 없다 | 도면에서 빼고 `undrawn_result_*` 로 이름을 남긴다 |
| `.slf` 기반 항목 | 입력에서 뺐다(사용자 확정: SLF 는 선택사항) | 해당 항목은 XML 에 있는 값만 쓴다 |
| 자릿수 경계 불일치 | XML 정밀도 자체의 한계 | §6 처럼 폭을 재서 정밀도 안임을 보인다. 값을 손보지 않는다 |

## 10. 검증 자산

- `tests/test_module_d/` — 97 passed.
- D1 은 코퍼스 341개 SDF 회귀로 파싱 안정성을 확인했다.
- 웹 UI 는 `node --check` 가 아니라 **playwright 실브라우저 런타임**으로 통과시킨다
  (콘솔 오류/경고 0 건, 내려받기 링크 전량 `%PDF-` 매직바이트 확인). 함수-지역
  헬퍼를 다른 스코프에서 참조해 `ReferenceError` 가 났던 전례 때문이다.
