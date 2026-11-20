# 경로 상수 함수화 지시서 — import 시점에 굳는 경로 없애기

작성 2026-09-08 · 대상 저장소 `PycharmProjects/JupyterProject`
목적: **모듈 A·F 통합 전에, 「import 순서에 따라 경로가 갈리는」 구조를 먼저 없앤다**

---

## 0. 무엇을 고치는가 — 그리고 무엇은 안 고치는가

### 고치는 것

경로가 **모듈 레벨 상수**로 import 시점에 굳고, 그 뒤 `_boot()` 가 값을 덮어써도 **이미 복사해 간 모듈은 따라오지 않는** 구조.

    services/cad_import/pipeline/handoff.py:42
        OUT_DIR = pick_out_dir()                    ← import 때 굳는다

    services/cad_import/pick/io.py:13
        NEW_DIR = handoff.OUT_DIR                   ← ★값을 복사한다

    routes/module_f/common.py  _boot()
        handoff.OUT_DIR = handoff.pick_out_dir()    ← 여기서 덮어도
                                                       io.NEW_DIR 은 안 바뀐다

`pick/__init__.py:10` 이 `NEW_DIR` 을 다시 재수출하므로 복사가 한 번 더 일어난다.

**증상**: 오류가 안 난다. 파일이 조용히 엉뚱한 폴더에 생기거나 안 보인다. `_boot()` 전에 import 된 모듈과 후에 import 된 모듈이 **한 프로세스 안에서 서로 다른 경로**를 붙들고 있다.

### 안 고치는 것 (별건 — 손대지 말 것)

- **`sys.path` 조작**. 모듈 첫머리 import 가 `_boot()` 전에 안 되는 것은 경로 등록 문제이지 이 지시서의 대상이 아니다. 이 작업으로 그게 풀리지 않는다.
- **`_boot()` 의 ②③ 방어** — 은퇴한 E 루트 차단, `services` 가 G 엔진인지 검증. **절대 지우지 말 것.** 두 트리의 최상위 패키지 이름이 둘 다 `services` 라 이 방어가 유일한 안전장치다.
- **경로 값 자체.** 함수화만 한다. 같은 실행 조건에서 **같은 경로가 나와야 한다.**

---

## 1. 대상 목록 — 전수

### A. 값이 굳고, 복사되는 것 (핵심)

| # | 자리 | 현재 | 문제 |
|---|---|---|---|
| 1 | `pipeline/handoff.py:42` | `OUT_DIR = pick_out_dir()` | import 때 굳음. `_boot()` 가 덮음 |
| 2 | `pick/io.py:13` | `NEW_DIR = handoff.OUT_DIR` | **값 복사 — 패치가 안 따라온다** |
| 3 | `pick/__init__.py:10` | `NEW_DIR` 재수출 | 복사가 한 번 더 |
| 4 | `pipeline/disp_cache.py:15` | `_DISP_CACHE_DIR = import_write_root()` | import 때 굳음. `_boot()` 가 덮음 |
| 5 | `pipeline/expand.py:15` | `DWG = s1.DWG_DIR` | 값 복사 (원본이 절대경로라 지금은 무해하나 같은 패턴) |
| 6 | `pipeline/flow.py:68` | `DWG = s1.DWG_DIR` | 위와 같음 |

읽는 자리: `handoff.py:102` · `handoff.py:107` · `disp_cache.py:19` · `pick/io.py:31` · `pick/io.py:36`

### B. 패치 대상에서 아예 빠진 것

| # | 자리 | 현재 | 문제 |
|---|---|---|---|
| 7 | `pick/io.py:14` | `STD_DIR = os.path.join("docs", "import", "0단계_표준샘플")` | **cwd 상대경로.** `_boot()` 가 안 건드린다 |

`common.py` 주석이 이 문제를 이미 알고 있다 — 「데스크톱 G 는 cwd 가 편집기 폴더라 상대경로 "docs/import" 로 여기를 가리킨다. 웹서버는 cwd 가 프로젝트 루트라 같은 상대경로가 엉뚱한 곳을 가리키므로, 부팅 때 절대경로로 고정한다」. **그런데 `STD_DIR` 은 고정 목록에 없다.**

### C. 범위 밖 — 보고만 할 것

| # | 자리 | 현재 |
|---|---|---|
| 8 | `convert/planar.py:46` | `OUT = os.path.join(os.path.expanduser("~"), "Desktop", f"{KEY}_유저정리5.kfp")` |

**사용자 바탕화면에 파일을 쓴다.** 서버에서 돌면 서버 계정 바탕화면에 쓴다. 이 지시서에서 **고치지 말고**, 어디서 쓰이는지·실제로 파일이 생기는지만 확인해 보고할 것. 조치는 사용자 판단.

---

## 2. 조치

### 2-1. 쓰기 루트에 명시적 setter 를 둔다

지금은 `_boot()` 가 함수 자체를 람다로 갈아끼운다.

    handoff.import_write_root = lambda: work

영리하지만 **누가 언제 갈아끼웠는지 추적이 안 되고**, 이미 값을 복사해 간 모듈은 못 잡는다. 주입점을 하나로 명시한다.

    # services/cad_import/pipeline/handoff.py
    _WRITE_ROOT_OVERRIDE = None

    def set_write_root(path):
        """쓰기 루트를 바깥에서 못박는다. 서버가 부팅 때 한 번 부른다.

        데스크톱 실행은 부르지 않는다 — 그때는 아래 기본 규칙이 맞다.
        """
        global _WRITE_ROOT_OVERRIDE
        _WRITE_ROOT_OVERRIDE = str(path) if path else None

    def import_write_root():
        if _WRITE_ROOT_OVERRIDE:
            return _WRITE_ROOT_OVERRIDE
        if getattr(sys, "frozen", False):
            base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
            return os.path.join(base, "K-Fire", "cad_import")
        return os.path.join("docs", "import")

**기존 세 분기(override / frozen / 소스)의 값은 그대로다.** 분기 하나가 앞에 붙을 뿐이다.

### 2-2. 상수를 전부 함수로 바꾼다

| # | 바꾸는 것 | 바꾼 뒤 |
|---|---|---|
| 1 | `handoff.OUT_DIR` | **삭제.** 읽는 자리(`:102`, `:107`)를 `pick_out_dir()` 호출로 |
| 2 | `pick/io.NEW_DIR` | `def new_dir(): return handoff.pick_out_dir()` |
| 3 | `pick/__init__.py` | `NEW_DIR` 대신 `new_dir` 을 재수출. `__all__` 도 함께 |
| 4 | `disp_cache._DISP_CACHE_DIR` | **삭제.** `_disp_cache_dir()` 이 `import_write_root()` 를 직접 호출 |
| 5·6 | `expand.DWG` · `flow.DWG` | **삭제.** 쓰는 자리에서 `s1.DWG_DIR` 을 직접 참조 |
| 7 | `pick/io.STD_DIR` | `def std_dir(): return os.path.join(handoff.import_write_root(), "0단계_표준샘플")` |

**⑦ 주의** — 지금 `STD_DIR` 은 `"docs/import/0단계_표준샘플"` 이고, `import_write_root()` 의 소스 실행 분기가 `"docs/import"` 다. 즉 **소스 실행에서는 값이 같다.** 서버(override 적용) 에서만 달라지며, 그게 원래 의도한 동작이다. 이 점을 커밋 메시지에 남길 것.

`NEW_DIR` 을 밖에서 쓰는 곳이 저장소 안에 더 있는지 **반드시 전수 확인**하고(웹 라우트·시험·스크립트 포함) 같이 고칠 것. 하나라도 남으면 이 작업의 의미가 없다.

### 2-3. `_boot()` 를 정리한다

    # routes/module_f/common.py  _boot()
    #  전
    handoff.import_write_root = lambda: work
    handoff.OUT_DIR = handoff.pick_out_dir()
    disp_cache._DISP_CACHE_DIR = work

    #  후
    handoff.set_write_root(work)

세 줄이 한 줄이 된다. `disp_cache` import 가 이 목적으로만 있었다면 그것도 함께 정리한다(다른 용도가 있으면 남긴다).

**②③ 방어와 `os.makedirs` 두 줄은 그대로 둔다.**

---

## 3. 검증

### 3-1. 이 작업의 성패를 가르는 시험 하나

`tests/test_g_path_late_binding.py` 로 새로 만든다.

    · G 모듈을 «_boot() 전에» 먼저 import 한다
        import services.cad_import.pick.io as io
        import services.cad_import.pipeline.disp_cache as dc
    · 그다음 set_write_root(임시폴더) 를 부른다
    · io.new_dir() · dc._disp_cache_dir() 가 **임시폴더 기준으로 나오는지** 확인

**지금 코드로 돌리면 실패해야 한다.** 실패하는 것을 먼저 확인하고(빨간 줄을 눈으로 본 뒤) 고칠 것 — 안 그러면 이 시험이 무엇을 지키는지 아무도 모르게 된다.

### 3-2. 모듈 레벨 경로 상수가 0인지

    grep -rn "^[A-Za-z_][A-Za-z0-9_]*\s*=.*\(os\.path\.join\|import_write_root\|pick_out_dir\|default_edits_dir\|OUT_DIR\|\"docs\"\|'docs'\)" \
      cad_project_editor_g/services --include=*.py

조치 후 남는 것은 `stage1.py:105-106`(`_APP_ROOT` · `DWG_DIR`) 과 `planar.py:46`(범위 밖) 뿐이어야 한다.
`stage1.py` 의 둘은 `__file__` 기준 **절대경로**라 cwd·부팅 순서에 안 걸린다 — 그대로 둔다.

### 3-3. 실제 파일이 같은 자리에 가는지 (제일 중요)

시험 통과만으로는 부족하다. **실제로 한 번 돌려서 확인한다.**

    · 웹서버를 띄우고 대명동 평면도를 열어 찍은스펙을 저장한다
    ·   → 조치 전과 **같은 폴더·같은 파일명**에 생기는지
    · 기존에 저장된 찍은스펙을 연다
    ·   → 조치 전과 같이 열리는지
    · 표시캐시(_edit_disp_cache_*.json)가 같은 폴더에 생기는지
    · 데스크톱 G 를 그 폴더에서 실행해 같은 도면이 그대로 이어지는지
    ·   ← 두 실행이 같은 폴더를 쓴다는 것이 이 구조의 전제다(common.py 주석)

---

## 4. 수용 기준

| | 기준 |
|---|---|
| 1 | `test_g_path_late_binding.py` 통과. **조치 전에는 실패했음을 로그로 남길 것** |
| 2 | 모듈 레벨 경로 상수 = `stage1.py` 2건 + `planar.py` 1건(범위 밖) 뿐 |
| 3 | `_boot()` 의 몽키패치 3줄 → `set_write_root()` 1줄 |
| 4 | 찍은스펙 저장·로드 경로가 조치 전과 **동일** (실제 파일로 확인) |
| 5 | 표시캐시 경로가 조치 전과 동일 |
| 6 | 데스크톱 G 실행 경로가 조치 전과 동일 |
| 7 | 기존 시험 전부 통과 |
| 8 | `NEW_DIR` 을 밖에서 쓰던 자리가 저장소에 하나도 안 남음 |

---

## 5. 금지 사항

- **`sys.path` 조작을 건드리지 말 것** — 별건이다. 이 작업으로 모듈 첫머리 import 가 되지는 않는다
- **`_boot()` 의 E 루트 차단·`services` 검증을 지우지 말 것** — 두 트리 패키지 이름이 같아 이것이 유일한 방어다
- **경로 값을 바꾸지 말 것** — 함수화만. 같은 조건에 같은 경로
- `planar.py:46` 의 바탕화면 경로를 **고치지 말 것** — 확인·보고만
- 리팩터링·파일 이동·이름 변경 금지
- 골든 재생성 금지

## 6. 보고 형식

    [시험]      test_g_path_late_binding — 조치 전 실패 / 조치 후 통과 (로그 첨부)
    [상수]      모듈 레벨 경로 상수 7건 → __건 (남은 것과 이유)
    [_boot]     3줄 → 1줄
    [경로 동일] 찍은스펙 저장 __ / 로드 __ / 표시캐시 __ / 데스크톱 G __
    [NEW_DIR]   밖에서 쓰던 자리 __곳 전부 이관 (목록)
    [범위 밖]   planar.py:46 — 어디서 불리는지 · 실제로 파일이 생기는지
    [막힌 것]   못 고친 것 · 이유 · 필요한 판단
