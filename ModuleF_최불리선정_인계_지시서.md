# 최불리 선정 인계 지시서 — 손질이 고른 K개를 수리계산이 그대로 받는다

작성 2026-09-09 · 대상 저장소 `PycharmProjects/JupyterProject`

---

## 0. 증상과 원인 — 한 줄

**손질에서 영역을 지정해 K개를 골랐는데, 수리계산은 그 선정을 버리고 도면 전체에서 K개를 다시 뽑는다.** 그래서 손질 화면의 corridor 와 수리계산 화면의 corridor 가 다른 헤드를 잇고, 기준 헤드도 최원 경로도 달라진다.

실측 (B1F · 사용자 화면): 손질 영역 박스 안 헤드 **20개** · 설계 SDF 헤드 **30개**. 밑그림을 켜면 두 망이 **정확히 겹친다** — 좌표는 무결하고 **어느 헤드인가만** 다르다.

---

## 1. 소스 — 어디서 끊기나

### 손질 `/edit/worst` — 영역을 받아 K개를 고르고 세션에 남긴다

    routes/module_f/api_edit.py  (≈304~416)

    zones = body.get("zones")
    ...
    in_zone = {hi for hi, d in enumerate(b.disks) if any(... 사각형 안 ...)}
    ...
    w = _worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=k,
                       only_heads=only, source_index=src_index, head_xy=b.disks)
    ...
    sess["worst"] = w                 # w["heads"] = K개 헤드의 disk 번호
    sess["worst_zones"] = w["zones"]
    sess["worst_k"] = k               # [F-10b] 수리계산 K 도 여기서 같이 맞춘다

`w["heads"]` 는 **`b.disks` 의 인덱스** 목록이다. `w["worst_head"]` · `w["worst_path"]` · `w["loads"]`(corridor) 도 함께 들어 있다.

### 수리계산 `/design/build` — 세션의 선정을 안 읽는다

    routes/module_f/api_design.py:539

    got = select_and_expand(payload, es.board, k=cfg["k"],
                            selected_source=sel)        # ★only_heads 없음

`select_and_expand` 시그니처에는 `only_heads=None` 자리가 **이미 있다**:

    cad_project_editor_g/services/cad_import/design/restrict.py:265
    def select_and_expand(payload, board, *, k=None, only_heads=None, ...):
        ...
        wet = probe["wet"]
        cand = wet if only_heads is None else (set(only_heads) & wet)
        worst = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=k, only_heads=cand)

`only_heads` 가 `None` 이면 `cand = wet` = **도면 전체의 물닿는 헤드**. 거기서 K개를 다시 뽑는다.

`api_design.py` 에서 `sess["worst"]` · `sess["worst_zones"]` 를 읽는 곳: **0건** (grep 확인).

손질 코드의 주석이 이미 이렇게 말한다 — 「★수리계산 설정의 K 도 여기서 같이 맞춘다. 기준개수는 손질에서 정하는 값 하나뿐이고」. **K 는 맞췄는데 선정 결과는 빠뜨린 것이다.**

### 손질 corridor 미리보기는 **이미 있다** — 손대지 말 것

    routes/module_f/remote30.py  _worst_view()
        "corridor": [[x1,y1,x2,y2,load] for (a,c),load in w["loads"].items()]
        "worst_path": ...
    static/module_f.js:594~  Remote 30 — 최불리 «배관망(corridor)» 굵기 비례 + 펄스

손질 화면은 corridor 를 굵은 흰 선으로 이미 그린다. 사용자 화면(사진 1)의 박스 안 흰 격자선이 그것이다. **이 작업은 손질 화면을 건드리지 않는다.** 수리계산이 그 선정을 받으면 두 화면의 corridor 가 저절로 같아진다.

---

## 2. 조치 — `/design/build` 한 곳

### 2-1. 손질 선정을 `only_heads` 로 넘긴다

    w = sess.get("worst")
    only = None
    if w and w.get("heads"):
        only = {int(i) for i in w["heads"]}
    got = select_and_expand(payload, es.board, k=cfg["k"],
                            selected_source=sel, only_heads=only)

### 2-2. 인덱스 공간이 같은지 **먼저 확인**한다

`w["heads"]` 는 `b.disks` 인덱스다. `select_and_expand` 안의 `wet` 은 `build_planar_graph` 의 `wet_head_idx` 로, `payload["hcov"]` 인덱스다. **두 인덱스가 같은 순서인지**는 `es.convert_payload()` 가 `hcov` 를 `board.disks` 순서 그대로 싣는가에 달렸다.

`restrict_to_worst` 가 이미 `worst["heads"]` 를 `board.disks` 인덱스로 쓰고 거기서 `hcov` 를 다시 만들므로(`disks = list(board.disks); kept = [disks[i] for i in keep_idx]; out["hcov"] = kept`) 같은 공간일 가능성이 높다. **그래도 재서 확인한다** — 손질 `w["heads"]` 의 disk 좌표와 설계 표 헤드 노드 좌표가 K개 전부 6~19 mm 안에서 1:1 맞는지(BLOCKED §30 이 쓴 그 대응).

맞지 않으면 **멈추고 보고**한다. 인덱스가 어긋난 채 넘기면 엉뚱한 헤드 K개가 조용히 선정된다 — 지금보다 나쁘다.

### 2-3. 손질이 고른 K개를 전개가 못 붙이면 **말한다**

`cand = set(only_heads) & wet` 에서 빠지는 헤드가 있을 수 있다 — 선정은 board 도달로 세고 전개는 더 엄격하다(BLOCKED B4: 868 대 619). `select_and_expand` 는 지금 `excluded_heads` 로 개수만 돌려주고 `print` 한다.

    · `len(cand) < K` 이면 조용히 K 를 깎지 않는다 — 손질 `/edit/worst` 가
      「기준개수를 «조용히» 줄이지 않는다」고 막는 것과 **같은 규칙**을 여기서도 쓴다
    · 응답과 화면에 「손질에서 고른 K개 중 n개는 전개가 붙이지 못했습니다 —
      손질에서 그 헤드의 배관을 이어 주세요」 를 싣는다. 헤드 목록(disk 번호·좌표)도
    · **다른 헤드로 채우지 않는다.** 채우면 사람이 고른 것이 아닌 것이 산출에 들어간다(S340 · D-F10-3)

### 2-4. `sess["worst"]` 가 없으면 — 지금처럼 하되 **말한다**

손질에서 최불리를 안 눌렀으면 지금처럼 전체에서 뽑는다. 다만 응답과 화면에 「손질에서 최불리를 정하지 않아 도면 전체에서 K개를 뽑았습니다」 를 싣는다. 조용히 다르게 동작하는 갈래를 두지 않는다.

### 2-5. K 가 어긋나면 — 손질 것을 믿는다

`len(w["heads"]) != cfg["k"]` 이면 세션이 낡은 것이다. 손질의 `sess["worst_k"]` 를 쓰고 그 사실을 `steps` 에 남긴다. 기준개수의 입력칸은 손질에 하나뿐이라는 것이 이 저장소의 결정이다.

---

## 3. 검증

### 3-1. 계측 — 조치 전/후 (B1F · 영역 지정 · K=손질 값)

`scripts/_probe_worst_handoff.py` 로 새로 만든다.

| 항목 | 조치 전 (기대) | 조치 후 (기대) |
|---|---|---|
| 손질 `w["heads"]` 수 | K | K |
| 설계 표 헤드(노즐) 수 | **K 와 다름** (전체에서 뽑음) | **K** |
| 손질 헤드 ↔ 설계 헤드 좌표 1:1 대응 | 일부만 | **K / K** (6~19 mm 안) |
| 손질 `worst_head` == 설계 「기준 헤드 노드」 | 다름 | **같음** |
| 손질 `far_m` == 설계 `worst_path_m` | 다름 | **같음** (±0.05 m) |
| 손질 corridor 간선 수 == 설계 `edge_ref` 로 되짚은 board 간선 수 | 다름 | 같음 (세로 구간 제외) |
| `excluded_heads` | — | 0 (0 이 아니면 그 목록) |

### 3-2. 화면 — 사용자가 본 그 장면

    · B1F 를 열어 손질에서 영역 박스를 그리고 최불리를 누른다
    · 「표 확정」 → 수리계산 → 「배관 밑그림 깔기」 켜기
    · 손질 박스 안 헤드가 수리계산 삼각형과 **하나하나 겹치는지**
    · 빨간 점선(최원 경로)이 두 화면에서 **같은 줄**인지
    · 아이소를 켰을 때도 같은 헤드·같은 줄인지

### 3-3. 시험

`tests/test_module_f_worst_handoff.py`

    · 손질 선정이 있으면 그 K개가 그대로 설계 표에 온다 (전체에서 다시 안 뽑는다)
    · 손질 선정이 없으면 전체에서 뽑되 응답에 그 사실이 실린다
    · 손질이 고른 헤드 중 전개가 못 붙이는 것이 있으면 다른 헤드로 채우지 않고 보고한다
    · 인덱스 공간 — 손질 disk 번호 → 설계 헤드 좌표가 1:1 이다
    · 기존 회귀 전부 통과

---

## 4. 수용 기준

| | 기준 |
|---|---|
| 1 | 사진 1(손질)의 헤드 K개가 사진 2(수리계산)의 헤드 K개와 **좌표로 1:1** |
| 2 | 두 화면의 빨간 점선(최원 경로)이 **같은 절점 열** |
| 3 | 손질 `far_m` 과 설계 `worst_path_m` 차이 **0.05 m 이내** |
| 4 | 손질에서 최불리를 안 누른 세션 — 산출 **바이트 불변** + 안내 문구 |
| 5 | 전개가 못 붙인 헤드가 있을 때 — 산출에 **다른 헤드가 대신 들어가지 않는다** |
| 6 | 손질 화면 코드 **무변경** |

---

## 5. 금지 사항

- 손질 화면(`api_edit.py` 의 `/edit/worst`, `_worst_view`, `module_f.js` 의 corridor 그리기)을 건드리지 말 것 — 이미 옳게 동작한다
- `select_and_expand` · `worst_k_heads` 의 선정 규칙을 바꾸지 말 것 — 넘기는 인자 하나만 추가한다
- 전개가 못 붙인 헤드를 **다른 헤드로 대체하지 말 것**
- K 를 조용히 줄이지 말 것 — 손질과 같은 규칙으로 막고 말한다
- `bake_isometric` · 좌표 변환 일체 손대지 말 것 — 밑그림이 정확히 겹친다는 것이 실측으로 확인됐다
- 골든 재생성 금지 · 리팩터링 금지

## 6. 보고 형식

    [인덱스 확인]  손질 disk 번호 ↔ 설계 헤드 좌표 __/K 대응 (최대 오차 __ mm)
    [계측표]       §3-1 조치 전 / 후
    [화면]         손질 K개 ↔ 수리계산 K개 겹침 · 최원 경로 일치 · 아이소에서도 일치
    [excluded]     전개가 못 붙인 헤드 __개 (목록)
    [시험]         신규 __건 · 회귀 __건 통과
    [막힌 것]      못 고친 것 · 이유 · 필요한 판단

**인덱스 공간이 다르다고 확인되면 §2-1 을 적용하지 말고 그 자리에서 멈춰 보고한다.**
