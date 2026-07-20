"""템플릿(HTML)의 외부 CDN 링크를 로컬(static/vendor) 경로로 교체.

localize_cdn_assets.py 로 자산을 내려받은 뒤 실행한다.
멱등(idempotent): 이미 로컬 경로면 변화 없음.

실행:
    python scripts/rewrite_cdn_links.py [--base <worktree_root>]
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

# 정확 일치 URL → 로컬 경로
_EXACT = {
    "https://cdn.tailwindcss.com": "/static/vendor/js/tailwind.js",
    "https://cdn.jsdelivr.net/npm/dxf-parser@1.1.2/dist/dxf-parser.js": "/static/vendor/js/dxf-parser.js",
    "https://cdn.jsdelivr.net/npm/three@0.128.0/build/three.min.js": "/static/vendor/js/three-0.128.0.min.js",
    "https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js":
        "/static/vendor/js/OrbitControls-0.128.0.js",
    "https://cdn.jsdelivr.net/npm/three@0.160.0/build/three.min.js": "/static/vendor/js/three-0.160.0.min.js",
    "https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.css":
        "/static/vendor/css/pretendard.css",
}

# 제거할 preconnect 라인 (외부 prefetch 힌트 — 로컬화 후 불필요)
_PRECONNECT_RE = re.compile(
    r'[ \t]*<link\s+rel="preconnect"\s+href="https://(?:fonts\.googleapis\.com|fonts\.gstatic\.com|cdn\.jsdelivr\.net)"[^>]*>\s*\n',
    re.IGNORECASE,
)

# Google Fonts CSS 링크 → 로컬 통합 fonts.css
_GFONTS_CSS_RE = re.compile(
    r'<link[^>]*href="https://fonts\.googleapis\.com/css2[^"]*"[^>]*>',
    re.IGNORECASE,
)
_LOCAL_FONTS_LINK = '<link rel="stylesheet" href="/static/vendor/fonts/fonts.css">'


def _rewrite(html: str) -> tuple[str, int]:
    n = 0
    # 1) preconnect 제거
    html, c = _PRECONNECT_RE.subn("", html)
    n += c
    # 2) google fonts css → 로컬 (첫 번째만 로컬 링크로, 나머지 중복은 제거)
    matches = _GFONTS_CSS_RE.findall(html)
    if matches:
        html = _GFONTS_CSS_RE.sub("", html)
        # head 안의 적당한 위치 대신, 원래 첫 링크 자리를 보존하기 어려우므로 </head> 앞 또는 첫 위치에 삽입
        # 간단·안전하게: <title> 다음 줄 또는 <head> 다음에 삽입
        if "/static/vendor/fonts/fonts.css" not in html:
            if re.search(r"</title>", html, re.IGNORECASE):
                html = re.sub(r"(</title>)", r"\1\n  " + _LOCAL_FONTS_LINK, html, count=1, flags=re.IGNORECASE)
            else:
                html = re.sub(r"(<head[^>]*>)", r"\1\n  " + _LOCAL_FONTS_LINK, html, count=1, flags=re.IGNORECASE)
        n += len(matches)
    # 3) 정확 일치 URL 치환
    for src, dst in _EXACT.items():
        if src in html:
            html = html.replace(src, dst)
            n += 1
    return html, n


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=str(Path(__file__).resolve().parent.parent))
    args = ap.parse_args()
    tdir = Path(args.base).resolve() / "templates"
    total = 0
    for html_path in sorted(tdir.glob("*.html")):
        original = html_path.read_text(encoding="utf-8")
        new, n = _rewrite(original)
        if n and new != original:
            html_path.write_text(new, encoding="utf-8")
            print(f"  [rewrite] {html_path.name}: {n} 교체")
            total += n
    print(f"[rewrite] base={tdir.parent.name}  총 {total} 교체")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
