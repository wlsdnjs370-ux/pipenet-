"""
ThemeManager — QSS 파일 기반 테마 로드/전환 관리자.

지원 테마 (사용자 설정값): 'light', 'dark', 'system'
  - 'system': OS 색상 설정을 감지하여 'light' 또는 'dark' 중 하나를 자동 적용
QSS 파일 위치: ui/themes/<name>.qss  ('light.qss', 'dark.qss' 두 파일만 필요)
"""
from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtWidgets import QApplication

log = logging.getLogger(__name__)

_THEMES_DIR = Path(__file__).parent / "themes"

# 사용자가 선택한 모드 ('light' | 'dark' | 'system')
_CURRENT_MODE: str = "system"
# 실제로 적용된 테마 ('light' | 'dark') — 'system'은 감지 후 해소된 값
_CURRENT_THEME: str = "light"


# ---------------------------------------------------------------------------
# 내부 헬퍼
# ---------------------------------------------------------------------------

def _resolve_system_theme() -> str:
    """OS 다크/라이트 설정을 감지하여 'dark' 또는 'light' 반환.

    탐지 우선순위:
      1. Qt 6.5+ QGuiApplication.styleHints().colorScheme()
      2. Windows 레지스트리 폴백 (winreg)
      3. 감지 실패 시 'light' 기본값
    """
    # --- 방법 1: Qt 내장 API (Qt 6.5+, 이벤트 기반 감지와 일관성 보장) ---
    try:
        from PySide6.QtCore import Qt
        from PySide6.QtGui import QGuiApplication

        app = QGuiApplication.instance()
        if app is not None:
            scheme = app.styleHints().colorScheme()
            if scheme == Qt.ColorScheme.Dark:
                return "dark"
            if scheme == Qt.ColorScheme.Light:
                return "light"
    except AttributeError:
        pass  # Qt < 6.5 폴백으로 진행

    # --- 방법 2: Windows 레지스트리 ---
    try:
        import winreg  # Windows 전용 표준 라이브러리

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if value == 1 else "dark"
    except (OSError, ImportError):
        pass

    return "light"


# ---------------------------------------------------------------------------
# 공개 API
# ---------------------------------------------------------------------------

def available_themes() -> list[str]:
    """사용 가능한 테마 이름 목록 반환 ('light', 'dark')."""
    return sorted(p.stem for p in _THEMES_DIR.glob("*.qss"))


def apply(app: QApplication, theme_name: str = "system") -> bool:
    """
    지정한 테마를 앱 전역 스타일시트로 적용한다.

    Args:
        theme_name: 'light' | 'dark' | 'system'
            - 'system'은 OS 설정을 감지하여 'light' / 'dark' 중 하나로 해소

    Returns:
        True  — 적용 성공
        False — QSS 파일을 찾을 수 없음 (에러 로그 출력 후 기존 스타일 유지)
    """
    global _CURRENT_MODE, _CURRENT_THEME

    _CURRENT_MODE = theme_name

    resolved = _resolve_system_theme() if theme_name == "system" else theme_name
    qss_path = _THEMES_DIR / f"{resolved}.qss"

    if not qss_path.exists():
        log.error("테마 파일을 찾을 수 없습니다: %s", qss_path)
        return False

    try:
        qss = qss_path.read_text(encoding="utf-8")
        # url() 경로를 절대경로로 치환 — CWD 의존 문제 방지
        themes_abs = _THEMES_DIR.as_posix()
        qss = qss.replace("url(ui/themes/", f"url({themes_abs}/")
        app.setStyleSheet(qss)
        _CURRENT_THEME = resolved
        log.debug("테마 적용 완료: mode=%s → resolved=%s", theme_name, resolved)
        return True
    except OSError as e:
        log.error("테마 파일 읽기 실패: %s", e)
        return False


def current_theme() -> str:
    """현재 실제로 적용 중인 테마 이름 반환 ('light' 또는 'dark').

    'system' 모드일 때는 OS 감지 결과로 해소된 값을 반환한다.
    """
    return _CURRENT_THEME


def current_mode() -> str:
    """사용자가 선택한 테마 모드 반환 ('light' | 'dark' | 'system')."""
    return _CURRENT_MODE


def toggle(app: QApplication) -> str:
    """
    라이트 ↔ 다크 테마를 토글한다. ('system' 모드는 해제됨)

    Returns:
        토글 후 적용된 테마 이름
    """
    next_theme = "dark" if _CURRENT_THEME == "light" else "light"
    apply(app, next_theme)
    return _CURRENT_THEME
