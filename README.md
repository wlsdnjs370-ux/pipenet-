# CAD → PIPENET 자동 도식 변환 (Remote 30)

스프링클러 도면(`.dxf`)에서 가장 먼/가까운 헤드 30개를 자동 선정하고, 추출된 배관망을 PIPENET 의 SDF/SLF 포맷으로 직렬화하는 모듈입니다.

## 사전 요구사항

- **Python 3.10+** (dataclass / `from __future__ import annotations` / `|` union type 사용)
- 운영체제: Windows / Linux / macOS 모두 가능 (Docker 도 지원)
- PIPENET (별도 상용 도구, 산출 SDF 를 열어 검증)

## 설치

```bash
git clone <repo-url>
cd JupyterProject

python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux / macOS
source .venv/bin/activate

pip install -r requirements.txt
```

## 핵심 자산 파일

`remote30_prototype.py` 의 `emit_sdf` 는 두 자산 파일을 필요로 합니다:

| 파일 | 역할 |
|---|---|
| `3-1형_자연낙차_LSP_4F_OA_지하층포함_120m~200m미만_6.6K로 감압_알람밸브.sdf` | Template SDF — PIPENET 의 Graphics 블록(아이소매트릭 표시 메타·schemes·Display-options) 보존 |
| `OA_3-1형_지하층포함_120~200m미만_35F.slf` | Standard SLF — 6 schedule (KSD 3507/3562/3576/DP/CPVC/FX) + 표준 노즐(SP-HEAD / INDOOR HYDRANT) + 표준 펌프 정의. 결과 폴더에 동봉 |

두 파일은 git 에 포함되어 있어 clone 직후 자동 인식됩니다.

### 자산 경로 override (선택)

다른 위치에 자산을 두고 싶거나, Docker/CI 환경에서 별도 경로로 마운트하는 경우 환경변수로 지정합니다:

```bash
export REMOTE30_TEMPLATE_SDF=/path/to/template.sdf
export REMOTE30_STANDARD_SLF=/path/to/standard.slf
```

해석 우선순위:
1. 환경변수 (절대·상대 둘 다 허용, 상대는 cwd 기준)
2. `remote30_prototype.py` 와 같은 디렉토리의 표준 파일명

둘 다 실패하면 `RuntimeWarning` 이 발행되고, 결과 SDF 의 아이소매트릭 표시가 누락되거나 PIPENET 의 diameter UI 가 "Unset" 으로 표시됩니다.

## 사용

(진입점은 프로젝트 상태에 따라 다를 수 있습니다 — Flask 서버 `대조 서버.py` / 직접 Python 호출 등.)

직접 호출 예:

```python
from pathlib import Path
from remote30_prototype import run_stages_0_2, emit_sdf

# 단계 0~2 (파싱·배관망·헤드 인식)
for evt in run_stages_0_2(Path("도면.dxf"), job_id="test"):
    print(evt)

# stage2_complete 이벤트의 데이터로 tables 만든 뒤
emit_sdf(tables, Path("out.sdf"), project_title="My Project")
```

## Production 운영 — 본체 PC 를 24/7 서버로

개발용 `app.run` 대신 **waitress** (Windows production WSGI) 로 띄우고, 부팅 시 자동 실행 + Cloudflare Tunnel 로 외부 도메인 노출.

### 1. 초기 설정 (한 번만)

```bash
pip install -r requirements.txt
cp .env.example .env
# .env 의 FLASK_SECRET_KEY 를 새 값으로 교체 (이미 자동 생성됐다면 그대로 둠):
python -c "import secrets; print(secrets.token_hex(32))"
```

### 2. 서버 실행

```bash
# 수동 (한 번 띄우기)
python serve.py

# 또는 더블 클릭 / cmd
start_server.bat
```

`start_server.bat` 은 크래시 시 5초 후 자동 재시작 루프가 포함됩니다.

### 3. 부팅 시 자동 시작 (Windows 작업 스케줄러)

```cmd
schtasks /create /tn "Remote30 Server" ^
  /tr "\"C:\Users\admin\PycharmProjects\JupyterProject\start_server.bat\"" ^
  /sc onstart /ru SYSTEM /rl HIGHEST
```

또는 GUI: `taskschd.msc` → 작업 만들기 → 트리거 "시작할 때" → 동작 `start_server.bat`.

추가 권장 설정 (제어판 → 전원 옵션):
- 절대 끄지 않음 (디스플레이/슬립 모두)
- 빠른 시작 비활성화

### 4. 외부 접속 — Cloudflare Tunnel + 자체 도메인

#### 4-1. 도메인 구매 (~$10/년)

- **Cloudflare Registrar** (https://dash.cloudflare.com → Domains) 권장 — 원가 마진 없이 도메인 등록
- 또는 Namecheap / Gabia 등에서 산 뒤 Cloudflare 에 DNS 위탁

#### 4-2. cloudflared 설치 + 터널 생성

```cmd
winget install Cloudflare.cloudflared

cloudflared tunnel login                                  # 브라우저 인증
cloudflared tunnel create remote30                        # 터널 생성
cloudflared tunnel route dns remote30 cad.yourdomain.com  # DNS 매핑
```

`%USERPROFILE%\.cloudflared\config.yml` 작성:

```yaml
tunnel: remote30
credentials-file: C:\Users\<user>\.cloudflared\<tunnel-id>.json

ingress:
  - hostname: cad.yourdomain.com
    service: http://localhost:5051
  - service: http_status:404
```

#### 4-3. 터널 실행 (자동 시작 권장)

```cmd
# 서비스로 등록 (관리자 cmd)
cloudflared service install
```

이후 `cad.yourdomain.com` 으로 외부 접속 → 로그인 게이트(비밀번호) 통과 → 메인 화면. HTTPS 자동 (Let's Encrypt 인증서 Cloudflare 가 발급).

### 5. 운영 후 점검

| 항목 | 확인 방법 |
|---|---|
| 서버 살아있나 | `curl http://localhost:5051/login` → 200 |
| 외부 도메인 살아있나 | `curl https://cad.yourdomain.com/login` → 200 |
| 자동 시작 등록됨 | `schtasks /query /tn "Remote30 Server"` |
| cloudflared 서비스 | `Get-Service cloudflared` (PowerShell) |
| 디스크 차오름 방지 | `data/uploads/`, `data/overall_runs/` 주기 정리 |

## 문제 해결

| 증상 | 원인 / 조치 |
|---|---|
| PIPENET 에서 diameter 가 "Unset" | Standard SLF 가 결과 폴더에 동봉되지 않음. `.slf` 파일을 `.sdf` 와 같은 폴더에 두거나, `REMOTE30_STANDARD_SLF` 환경변수로 경로 지정 |
| 아이소매트릭 도식 메타 누락 | Template SDF 를 찾지 못함. `REMOTE30_TEMPLATE_SDF` 환경변수 또는 모듈 디렉토리 확인 |
| `ModuleNotFoundError: pipenet_converter` | `pip install -e pipenet_converter` 또는 `pipenet_converter/src` 를 PYTHONPATH 에 추가 |
