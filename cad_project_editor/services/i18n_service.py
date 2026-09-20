"""
K-Fire 다국어 번역 서비스 (Application 계층)
PySide6 의존성 없음 — 순수 Python.
"""
import json
import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)


class I18nService:
    """다국어 번역 서비스 — 싱글턴"""

    _instance: "I18nService | None" = None

    def __init__(self) -> None:
        self._lang: str = "ko"
        self._strings: dict[str, str] = {}
        self._fallback: dict[str, str] = {}

    @classmethod
    def instance(cls) -> "I18nService":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def load(self, lang: str = "ko") -> None:
        """언어 사전 로드. 시도 순서: {lang}.json → {lang}.json.bak → ko fallback"""
        self._lang = lang
        self._strings = self._try_load(lang)
        if lang != "ko":
            self._fallback = self._try_load("ko")
        else:
            self._fallback = {}
        logger.info("i18n loaded: lang=%s, keys=%d", lang, len(self._strings))

    def t(self, key: str, **kwargs) -> str:
        """
        번역 함수.
        조회 순서: 현재 언어 사전 → 폴백(한국어) → 키 원문 반환
        kwargs로 동적 치환: t("노드 '{name}' 삭제", name="N1")
        """
        text = self._strings.get(key) or self._fallback.get(key) or key
        if kwargs:
            try:
                text = text.format(**kwargs)
            except (KeyError, IndexError):
                pass  # 치환 실패 시 원문 반환
        return text

    @property
    def lang(self) -> str:
        return self._lang

    @staticmethod
    def available_languages() -> list[tuple[str, str]]:
        """사용 가능한 언어 목록 반환: [(코드, 표시명), ...]"""
        return [("ko", "한국어"), ("en", "English"), ("ja", "日本語"), ("vi", "Tiếng Việt"), ("ch", "简体中文"), ("es", "Español")]

    def _try_load(self, lang: str) -> dict[str, str]:
        """로드 시도 순서: {lang}.json → {lang}.json.bak → ko fallback → 빈 dict"""
        for suffix in (f"{lang}.json", f"{lang}.json.bak"):
            data = self._load_json_file(suffix)
            if data is not None:  # 빈 dict({})도 유효 — is not None 체크 필수
                if suffix.endswith(".bak"):
                    logger.warning("i18n: %s.json 로드 실패, .bak에서 복구 완료", lang)
                return data
        if lang != "ko":
            logger.warning("i18n: %s .bak도 없음, ko fallback 사용", lang)
            return self._try_load("ko")
        return {}

    def _load_json_file(self, filename: str) -> "dict[str, str] | None":
        """locales/{filename} 로드. 없거나 파싱 오류이면 None 반환."""
        base = Path(os.path.dirname(os.path.abspath(__file__))).parent
        path = base / "locales" / filename
        if not path.exists():
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else None
        except Exception as e:
            logger.error("i18n load error (%s): %s", path, e)
            return None

    def _load_json(self, lang: str) -> dict[str, str]:
        """하위 호환 래퍼 — _try_load 로 위임."""
        return self._try_load(lang)


def _t(key: str, **kwargs) -> str:
    """모듈 레벨 편의 래퍼. I18nService.instance().t()를 호출한다."""
    return I18nService.instance().t(key, **kwargs)


def user_msg(key: str, **kwargs) -> str:
    """사용자 노출 메시지 — locale 키(한국어 원문)로 호출하고 현재 언어로 변환한다.

    UseCase/Service에서 error_message 등을 만들 때 이 함수를 쓰면
    UI가 _t()를 다시 호출하지 않아도 된다(이미 번역된 문자열).
    """
    return _t(key, **kwargs)


def format_pump_error(code: str, params: dict) -> str:
    """
    PumpCurveValidationError의 (code, params)를 현재 언어 문자열로 변환.

    params 중 "name" 키는 자체가 번역 키(pump.param.*)이므로 이중 번역 수행.
    domain 타입을 import하지 않으므로 계층 원칙 유지.
    """
    translated_params = {
        k: _t(str(v)) if k == "name" else v
        for k, v in params.items()
    }
    return _t(code).format(**translated_params)