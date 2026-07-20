"""외부 CDN 자산을 로컬(static/vendor)로 내려받아 self-host 화.

목적: 브라우저가 fonts.googleapis.com / cdn.jsdelivr.net / cdn.tailwindcss.com
      등 외부 호스트로 요청을 보내지 않도록, 모든 정적 자산을 로컬에 고정한다.
      (LH 네트워크 격리 증빙 — 외부 연결 0)

실행:
    python scripts/localize_cdn_assets.py [--base <worktree_root>]

산출물:
    static/vendor/js/*.js
    static/vendor/fonts/fonts.css        (Inter + JetBrains Mono, @font-face 로컬화)
    static/vendor/fonts/files/*.woff2
    static/vendor/css/pretendard.css + files

멱등(idempotent): 이미 받은 파일은 건너뛰지 않고 덮어쓴다(재현 가능).
"""
from __future__ import annotations

import argparse
import re
import sys
import urllib.request
from pathlib import Path

# 모던 브라우저 UA — Google Fonts 가 woff2(@font-face) 를 반환하도록 유도
_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# 통합 폰트 — 전체 템플릿에서 쓰는 모든 weight 를 한 번에 (Inter 400-800, JetBrains Mono 400-600)
_FONTS_CSS_URL = (
    "https://fonts.googleapis.com/css2"
    "?family=Inter:wght@400;500;600;700;800"
    "&family=JetBrains+Mono:wght@400;500;600"
    "&display=swap"
)

# 단순 JS/CSS 자산 — (URL, 로컬 상대경로)
_SIMPLE_ASSETS = [
    ("https://cdn.tailwindcss.com", "js/tailwind.js"),
    ("https://cdn.jsdelivr.net/npm/dxf-parser@1.1.2/dist/dxf-parser.js", "js/dxf-parser.js"),
    ("https://cdn.jsdelivr.net/npm/three@0.128.0/build/three.min.js", "js/three-0.128.0.min.js"),
    ("https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js",
     "js/OrbitControls-0.128.0.js"),
    ("https://cdn.jsdelivr.net/npm/three@0.160.0/build/three.min.js", "js/three-0.160.0.min.js"),
    ("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.css",
     "css/pretendard.css"),
]


def _fetch(url: str, *, binary: bool) -> bytes | str:
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    with urllib.request.urlopen(req, timeout=60) as resp:  # noqa: S310  (고정된 신뢰 URL)
        data = resp.read()
    return data if binary else data.decode("utf-8")


def _download_simple(vendor: Path) -> list[str]:
    log = []
    for url, rel in _SIMPLE_ASSETS:
        dst = vendor / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        data = _fetch(url, binary=True)
        dst.write_bytes(data)
        log.append(f"  [js/css] {url}\n           -> {rel} ({len(data):,} bytes)")
    return log


def _download_fonts(vendor: Path) -> list[str]:
    """Google Fonts CSS 를 받아 woff2 를 로컬화하고 CSS 의 URL 을 로컬 경로로 치환."""
    log = []
    fonts_dir = vendor / "fonts"
    files_dir = fonts_dir / "files"
    files_dir.mkdir(parents=True, exist_ok=True)

    css = _fetch(_FONTS_CSS_URL, binary=False)
    # @font-face 안의 https://fonts.gstatic.com/....woff2 전부 추출
    urls = re.findall(r"https://fonts\.gstatic\.com/[^)]+\.woff2", css)
    seen: dict[str, str] = {}
    for i, u in enumerate(dict.fromkeys(urls)):
        # 파일명: 원본 마지막 경로 토큰 사용 (충돌 시 인덱스)
        name = u.rsplit("/", 1)[-1]
        if name in seen.values():
            name = f"{i}_{name}"
        data = _fetch(u, binary=True)
        (files_dir / name).write_bytes(data)
        seen[u] = name
        log.append(f"  [font]   {name} ({len(data):,} bytes)")

    # CSS 내 gstatic URL 을 로컬 상대경로로 치환
    for u, name in seen.items():
        css = css.replace(u, f"files/{name}")
    (fonts_dir / "fonts.css").write_text(css, encoding="utf-8")
    log.append(f"  [css]    fonts.css ({len(seen)} woff2 로컬화)")
    return log


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=str(Path(__file__).resolve().parent.parent),
                    help="worktree 루트 (기본: 이 스크립트의 상위)")
    args = ap.parse_args()

    base = Path(args.base).resolve()
    vendor = base / "static" / "vendor"
    print(f"[localize] base   = {base}")
    print(f"[localize] vendor = {vendor}")

    log: list[str] = []
    log += _download_simple(vendor)
    log += _download_fonts(vendor)

    print("\n".join(log))
    print(f"\n[localize] done - {vendor} 에 자산 고정됨. 템플릿의 CDN 링크를 로컬 경로로 교체하세요.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
