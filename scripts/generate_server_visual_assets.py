from __future__ import annotations

import subprocess
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "generated_reports" / "server_screenshots"
OUT.mkdir(parents=True, exist_ok=True)
CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")


BASE_CSS = """
html,body{margin:0;background:#e8edf3;font-family:'Malgun Gothic','Noto Sans KR',sans-serif;color:#172033}
.shell{width:1280px;height:820px;padding:24px;box-sizing:border-box;background:linear-gradient(135deg,#eef3f8,#dbe5ef)}
.layout{display:grid;grid-template-columns:210px 1fr;gap:18px;height:100%}
.side{background:#fff;border:1px solid #cfd7e3;border-radius:0;padding:16px;box-shadow:0 8px 22px rgba(15,23,42,.08)}
.side h3{margin:0 0 14px;font-size:15px;color:#0f172a}
.menu{display:block;width:100%;text-align:left;margin:7px 0;padding:12px;border:1px solid #d7dee8;background:#f8fafc;border-radius:0;font-weight:700;color:#334155}
.menu.active{background:#12345a;color:white;border-color:#12345a}
.main{display:flex;flex-direction:column;gap:14px;min-width:0}
.header{background:#fff;border:1px solid #cfd7e3;border-radius:0;padding:18px 22px;box-shadow:0 8px 22px rgba(15,23,42,.08)}
.kicker{font-size:12px;letter-spacing:.16em;color:#2563eb;font-weight:900;margin:0 0 4px}
h1{margin:0;font-size:28px}
.card{background:#fff;border:1px solid #cfd7e3;border-radius:0;padding:18px;box-shadow:0 8px 22px rgba(15,23,42,.08)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.pill{display:inline-block;padding:5px 10px;border-radius:999px;font-size:12px;font-weight:900}
.pass{background:#dcfce7;color:#166534}.fail{background:#fee2e2;color:#991b1b}.eng{background:#dbeafe;color:#1d4ed8}.econ{background:#dcfce7;color:#15803d}
table{width:100%;border-collapse:collapse;font-size:13px;background:#fff}th{background:#18233a;color:white;padding:9px;border:1px solid #334155}td{padding:8px;border:1px solid #cbd5e1}
.network{height:330px;background:#f8fafc;border:1px solid #cbd5e1;position:relative;overflow:hidden}
.pipe{position:absolute;height:3px;background:#64748b;transform-origin:left center}.pipe.red{background:#ef4444;height:5px}.pipe.green{background:#16a34a;height:5px}
.node{position:absolute;width:8px;height:8px;border-radius:50%;background:#0f172a}
.maptitle{position:absolute;left:14px;top:12px;background:white;border:1px solid #cbd5e1;padding:6px 10px;font-weight:800}
.chart{height:210px;background:linear-gradient(180deg,#f8fafc,#fff);border:1px solid #cbd5e1;position:relative}
.bar{position:absolute;bottom:28px;width:22px;background:#2563eb}.bar.red{background:#ef4444}.axis{position:absolute;left:38px;right:20px;bottom:28px;height:1px;background:#94a3b8}
.note{font-size:13px;color:#475569;line-height:1.55}
"""


def write_html(name: str, body: str) -> Path:
    path = OUT / f"{name}.html"
    path.write_text(f"<!doctype html><meta charset='utf-8'><style>{BASE_CSS}</style>{body}", encoding="utf-8")
    return path


def chrome_shot(url: str, out: Path) -> None:
    subprocess.run(
        [
            str(CHROME),
            "--headless=new",
            "--disable-gpu",
            "--window-size=1280,820",
            f"--screenshot={out}",
            url,
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def page_shell(active: str, content: str) -> str:
    menus = ["검증결과", "설계 최적화 가이드", "결과 데이터 테이블", "검진 통계", "상세리포트", "CAD-아이소매트릭 대조"]
    buttons = "".join(f"<button class='menu {'active' if m==active else ''}'>{i+1}. {m}</button>" for i, m in enumerate(menus))
    return f"""
    <div class='shell'><div class='layout'>
      <aside class='side'><h3>PIPENET 메뉴</h3>{buttons}</aside>
      <main class='main'>
        <section class='header'><p class='kicker'>PIPE CALC VALIDATION</p><h1>PIPENET 수리계산 검증</h1></section>
        {content}
      </main>
    </div></div>
    """


def create_assets() -> None:
    # Actual running server landing screen.
    chrome_shot("http://127.0.0.1:5050/", OUT / "server_actual_main.png")

    validation = page_shell(
        "검증결과",
        """
        <section class='grid2'>
          <div class='card'><h2>검증 결과</h2>
            <table><tr><th>구분</th><th>내용</th></tr>
            <tr><td><span class='pill pass'>적합</span></td><td>하겐윌리엄 선언 및 마찰손실 재계산 PASS</td></tr>
            <tr><td><span class='pill fail'>부적합</span></td><td>Pipe 114: 가지배관 유속 6.044 &gt; 6.000 m/s</td></tr>
            <tr><td><span class='pill pass'>적합</span></td><td>C-Factor: KSD=120, CPVC=150 일치</td></tr></table>
          </div>
          <div class='card'><h2>SDF 아이소매트릭 배관망</h2><div class='network'>
            <div class='maptitle'>FAIL/PASS 레이어 표시</div>
            <div class='pipe green' style='left:120px;top:190px;width:260px;transform:rotate(-12deg)'></div>
            <div class='pipe green' style='left:370px;top:135px;width:210px;transform:rotate(18deg)'></div>
            <div class='pipe red' style='left:525px;top:200px;width:115px;transform:rotate(38deg)'></div>
            <div class='pipe' style='left:250px;top:235px;width:170px;transform:rotate(35deg)'></div>
            <div class='node' style='left:115px;top:186px'></div><div class='node' style='left:372px;top:130px'></div>
            <div class='node' style='left:526px;top:196px'></div><div class='node' style='left:635px;top:282px'></div>
          </div></div>
        </section>
        """,
    )
    chrome_shot(write_html("server_validation", validation).as_uri(), OUT / "server_validation.png")

    optimization = page_shell(
        "설계 최적화 가이드",
        """
        <section class='grid2'>
          <div class='card'><h2>공학적 마찰손실 최적화 조치</h2><div class='network'>
            <div class='maptitle'>Friction Loss Spike Map</div>
            <div class='pipe' style='left:90px;top:210px;width:210px;transform:rotate(-8deg)'></div>
            <div class='pipe red' style='left:298px;top:181px;width:160px;transform:rotate(22deg)'></div>
            <div class='pipe' style='left:445px;top:240px;width:180px;transform:rotate(-18deg)'></div>
          </div><p class='note'>m당 마찰손실 &gt; 1.0 kg/cm²/m 구간만 빨간색으로 표시합니다.</p></div>
          <div class='card'><h2>시공사 경제성 확보 방안</h2><div class='network'>
            <div class='maptitle'>Economy Optimization Candidate Map</div>
            <div class='pipe green' style='left:100px;top:190px;width:180px;transform:rotate(8deg)'></div>
            <div class='pipe green' style='left:275px;top:220px;width:210px;transform:rotate(-20deg)'></div>
            <div class='pipe' style='left:480px;top:145px;width:150px;transform:rotate(30deg)'></div>
          </div><p class='note'>저유속, 압력여유 과다, 대구경 밸브, CPVC 대구경 후보를 구분합니다.</p></div>
        </section>
        """,
    )
    chrome_shot(write_html("server_optimization", optimization).as_uri(), OUT / "server_optimization.png")

    table_view = page_shell(
        "결과 데이터 테이블",
        """
        <section class='card'><h2>결과 데이터 테이블</h2>
          <div style='margin-bottom:10px'><span class='pill fail'>빨강: 기준 위반</span> <span class='pill eng'>파랑: 공학 후보</span> <span class='pill econ'>초록: 경제성 후보</span></div>
          <table><tr><th>Pipe</th><th>배관 역할</th><th>유속</th><th>HW 재계산</th><th>공학 후보</th><th>경제 후보</th><th>판정 사유</th></tr>
          <tr><td>5</td><td>other</td><td>9.574 / 10.0</td><td>PASS</td><td>유속여유 부족</td><td>-</td><td>그 밖의 배관 기준 10 m/s 이내</td></tr>
          <tr><td>97</td><td>branch</td><td>5.948 / 6.0</td><td>PASS</td><td>유속여유 부족</td><td>-</td><td>하류 교차분기 없음</td></tr>
          <tr><td>114</td><td>branch</td><td>6.044 / 6.0</td><td>PASS</td><td>-</td><td>-</td><td>기준 초과로 수정 대상</td></tr>
          </table>
        </section>
        """,
    )
    chrome_shot(write_html("server_table", table_view).as_uri(), OUT / "server_table.png")

    stats = page_shell(
        "검진 통계",
        """
        <section class='grid2'>
          <div class='card'><h2>Pipe Velocity Check</h2><div class='chart'><div class='axis'></div>
            <div class='bar' style='left:70px;height:80px'></div><div class='bar' style='left:120px;height:130px'></div>
            <div class='bar red' style='left:170px;height:168px'></div><div class='bar' style='left:220px;height:100px'></div></div></div>
          <div class='card'><h2>Nozzle Pressure-Flow</h2><div class='chart'><div class='axis'></div>
            <div class='bar' style='left:70px;height:95px;background:#16a34a'></div><div class='bar' style='left:120px;height:115px;background:#16a34a'></div>
            <div class='bar' style='left:170px;height:150px;background:#16a34a'></div><div class='bar' style='left:220px;height:110px;background:#16a34a'></div></div></div>
        </section>
        <section class='grid2'><div class='card'><h3>결과서 통계</h3><table><tr><td>배관 수</td><td>107</td></tr><tr><td>노즐 수</td><td>30</td></tr></table></div>
        <div class='card'><h3>SDF 통계</h3><table><tr><td>배관 수</td><td>107</td></tr><tr><td>특수설비</td><td>31</td></tr></table></div></section>
        """,
    )
    chrome_shot(write_html("server_stats", stats).as_uri(), OUT / "server_stats.png")


if __name__ == "__main__":
    create_assets()
    for p in sorted(OUT.glob("server_*.png")):
        print(p, p.stat().st_size)
