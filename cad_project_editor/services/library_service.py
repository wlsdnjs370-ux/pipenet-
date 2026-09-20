"""
LibraryService — 배관/노즐/피팅 라이브러리 조회/저장 담당

Phase 2-2 리팩토링: editor_core.py에서 라이브러리 메서드 추출.
PySide6 의존성 없음 (순수 서비스 계층).
"""
import json
import logging
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from domain.fitting_models import (
    FittingDefinition,
    normalize_equivalent_lengths,
    validate_library,
    FittingSchemaError,
)
from domain.pipe_sizing_presets import (
    HeadRangePreset,
    builtin_head_range_presets,
    builtin_library_dict,
    presets_from_library,
)

logger = logging.getLogger(__name__)

# 전역 디폴트 라이브러리 파일 (user ↔ factory 시드 공통)
FACTORY_LIBRARY_FILES: tuple[str, ...] = (
    "pipe_library.json",
    "nozzle_library.json",
    "fittings_library_v3.json",
    "pump_library.json",
    "pipe_sizing_library.json",
)


# =========================================================
# 시스템 보호 카테고리 표시명 → 현행 category_id 정규화 테이블 (SSOT)
#
# 과거 버전에서 category_id 명명 규칙이 바뀐 카테고리(예: "nozzle" → "water_spray")를
# display_name 기준으로 현행 ID로 되돌린다.
# 사용자가 새로 만든 카테고리는 이 테이블에 없으므로 자동으로 보호된다.
# =========================================================
_SYSTEM_DISPLAY_TO_CATEGORY: dict[str, str] = {
    # Nozzle library
    "head":            "head",
    "water spray":     "water_spray",
    "hose nozzle":     "hose_nozzle",
    # Fittings(Valve) library
    "alarm valve":     "alarm_valve",
    "preaction valve": "preaction_valve",
    "dry pipe valve":  "dry_pipe_valve",
    "swing check":     "swing_check",
    "smolensky check": "smolensky_check",
    "gate valve":      "gate_valve",
    "butterfly valve": "butterfly_valve",
    "globe valve":     "globe_valve",
    "angle valve":     "angle_valve",
    "ball valve":      "ball_valve",
    "vane type flow switch": "flow_switch_vane",  # M2-C: M5.9에서 선이관
    "strainer":        "strainer",           # D3 추가
}


def canonicalize_category_id(raw_category_id: str, display_name: str) -> str:
    """시스템 카테고리는 display_name 기준으로 현행 category_id로 정규화한다.

    _SYSTEM_DISPLAY_TO_CATEGORY 테이블에 없는 값(사용자 정의 카테고리)은
    raw_category_id를 그대로 반환하여 임의 변조를 방지한다.
    """
    canonical = _SYSTEM_DISPLAY_TO_CATEGORY.get((display_name or "").strip().lower())
    if canonical:
        return canonical
    return (raw_category_id or "").strip()


class LibraryService:
    """
    배관·노즐·피팅·펌프 라이브러리 데이터를 보유하고 조회/저장하는 서비스.
    생성자에서 모든 라이브러리를 None으로 초기화하며,
    외부에서 직접 속성에 데이터를 주입합니다.
    """

    def __init__(self) -> None:
        self.pipe_library: dict | None = None
        self.nozzle_library: dict | None = None
        self.fittings_library_v3: list[FittingDefinition] | None = None  # v3.0 NFPA 등가길이
        self.pump_library: dict | None = None
        self.pipe_sizing_library: dict | None = None  # 헤드갯수기준 카탈로그 (전역 전용, .kfp 비내장)

    # =========================================================
    # 배관 라이브러리
    # =========================================================

    def _robust_get_library_item(self, standard, dn_mm) -> dict | None:
        if not self.pipe_library:
            return None
        pipe_list = self.pipe_library.get("pipe_types", [])

        tgt_std = str(standard).strip().upper() if standard else None
        tgt_dn = int(dn_mm) if dn_mm is not None else None

        for row in pipe_list:
            if tgt_std:
                row_std = str(row.get("standard", "")).strip().upper()
                if row_std != tgt_std:
                    continue
            if tgt_dn is not None:
                row_dn_val = row.get("nominal_d_mm") or row.get("nominal_mm")
                try:
                    if int(row_dn_val) != tgt_dn:
                        continue
                except Exception:
                    continue
            return row
        return None

    def get_pipe_spec(self, standard, dn_mm) -> dict | None:
        row = self._robust_get_library_item(standard, dn_mm)
        if not row:
            return None

        raw_id = (
            row.get("inner_d_mm")
            or row.get("inner_mm")
            or row.get("ID_mm")
            or row.get("diameter")
            or 0.0
        )
        raw_rough = row.get("roughness_mm") or 0.045
        raw_c = (
            row.get("C_HW_default")
            or row.get("C")
            or row.get("hw_c")
            or row.get("HW_C")
            or 120.0
        )
        raw_name = row.get("name") or row.get("id") or f"{standard} DN{dn_mm}"

        return {
            "name": raw_name,
            "inner_d_mm": float(raw_id),
            "roughness_mm": float(raw_rough),
            "C": float(raw_c),
        }

    def get_pipe_inner_diameter_mm(self, dn_mm, current_standard=None) -> float:
        spec = self.get_pipe_spec(current_standard, dn_mm)
        if spec:
            return spec["inner_d_mm"]
        return float(dn_mm)

    # ---- 규격/DN 목록 게이트웨이 (staticmethod) -------------------------

    @staticmethod
    def get_pipe_standards(pipe_library: dict | None) -> list[str]:
        """pipe_library에서 규격(standard) 목록을 중복 없이 정렬하여 반환."""
        if not pipe_library:
            return []
        standards: set[str] = set()
        for row in pipe_library.get("pipe_types", []):
            std = row.get("standard")
            if std:
                standards.add(str(std).strip())
        return sorted(standards)

    @staticmethod
    def get_dn_list_for_standard(
        pipe_library: dict | None, standard: str
    ) -> list[dict]:
        """지정 규격에 해당하는 pipe_types 행 목록을 반환 (순서 유지)."""
        if not pipe_library:
            return []
        return [
            row
            for row in pipe_library.get("pipe_types", [])
            if str(row.get("standard", "")).strip() == standard
        ]

    @staticmethod
    def get_pipe_type_name(
        pipe_library: dict | None, standard: str, dn: str
    ) -> str:
        """standard와 dn이 일치하는 행의 type_name을 반환. 없으면 빈 문자열."""
        if not pipe_library:
            return ""
        for row in pipe_library.get("pipe_types", []):
            if (
                str(row.get("standard", "")).strip() == standard
                and str(row.get("dn", "")).strip() == str(dn).strip()
            ):
                return row.get("type_name", "")
        return ""

    # ----------------------------------------------------------------------

    def save_pipe_library(self, to_factory: bool = False) -> bool:
        """배관 라이브러리를 JSON 파일로 저장.

        to_factory=False (기본): 사용자 폴더(`%LOCALAPPDATA%\\K-Fire\\`)에 저장 — 개인 작업용
        to_factory=True: 워크스페이스 루트(설치 디렉토리)에 저장 — 공장 초기값 갱신용
        """
        if not self.pipe_library:
            return False
        base = self._install_path() if to_factory else self._base_path()
        filepath = base / "pipe_library.json"
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(self.pipe_library, f, indent=2, ensure_ascii=False)
            logger.info("[LIB] 배관 라이브러리 저장 완료: %s", filepath)
            return True
        except Exception as exc:
            logger.warning("[LIB] 배관 라이브러리 저장 실패: %s", exc)
            return False

    # =========================================================
    # 배관구경 산정(헤드갯수기준) 라이브러리
    # =========================================================

    def get_head_range_presets(self) -> list[HeadRangePreset]:
        """로드된 배관구경 산정 라이브러리의 기준 목록 (없거나 깨지면 내장 폴백)."""
        presets = presets_from_library(self.pipe_sizing_library)
        if presets:
            return presets
        return list(builtin_head_range_presets())

    def load_pipe_sizing_library_from_disk(self) -> dict:
        """전역(사용자 폴더) 파일 → 메모리 로드. 없거나 깨지면 내장 폴백.

        앱 시작·새 문서 리셋·구버전 .kfp 폴백이 전부 이 로더 하나를 쓴다
        (editor.__init__() 리셋이 메모리를 비우므로, 재로드 경로가 갈라지면
        화면에서 사라지는 사고가 재발한다).
        """
        path = self._base_path() / "pipe_sizing_library.json"
        data = None
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            logger.warning("[LIB] 배관구경 산정 라이브러리 로드 실패(%s): %s", path, exc)
        if not (isinstance(data, dict) and data.get("presets")):
            data = builtin_library_dict()
        self.pipe_sizing_library = data
        return data

    def save_pipe_sizing_library(self, to_factory: bool = False) -> bool:
        """배관구경 산정 라이브러리를 JSON 파일로 저장 (경로 규칙은 배관과 동일).

        메모리가 비어 있으면 내장 기본 카탈로그를 기록해 파일 부재를 복구한다.
        """
        data = self.pipe_sizing_library or builtin_library_dict()
        base = self._install_path() if to_factory else self._base_path()
        filepath = base / "pipe_sizing_library.json"
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            logger.info("[LIB] 배관구경 산정 라이브러리 저장 완료: %s", filepath)
            return True
        except Exception as exc:
            logger.warning("[LIB] 배관구경 산정 라이브러리 저장 실패: %s", exc)
            return False

    # =========================================================
    # 노즐 라이브러리
    # =========================================================

    def get_nozzle_spec(self, nozzle_name: str) -> dict | None:
        """노즐 이름(display_name)으로 항목 조회 — categories.items에서만 조회한다."""
        if not self.nozzle_library:
            return None
        # 새 스키마: categories 순회
        for cat in self.nozzle_library.get("categories", []):
            for item in cat.get("items", []):
                if item.get("display_name") == nozzle_name or item.get("name") == nozzle_name:
                    return self._nozzle_item_with_legacy_keys(item)
        return None

    def get_nozzle_by_id(self, item_id: str) -> dict | None:
        """노즐 ID로 항목 조회."""
        if not self.nozzle_library:
            return None
        for cat in self.nozzle_library.get("categories", []):
            for item in cat.get("items", []):
                if item.get("id") == item_id:
                    return self._nozzle_item_with_legacy_keys(item)
        return None

    def get_nozzle_categories(self) -> list[dict]:
        """노즐 라이브러리의 전체 카테고리 목록 반환."""
        if not self.nozzle_library:
            return []
        return list(self.nozzle_library.get("categories", []))

    def get_nozzle_items_by_category(self, category_id: str) -> list[dict]:
        """특정 카테고리의 노즐 항목 목록 반환."""
        if not self.nozzle_library:
            return []
        for cat in self.nozzle_library.get("categories", []):
            if cat.get("category_id") == category_id:
                return [self._nozzle_item_with_legacy_keys(item) for item in cat.get("items", [])]
        return []

    def get_all_nozzle_items_flat(self) -> list[dict]:
        """전체 카테고리의 노즐 항목을 flat list로 반환 (하위 호환용)."""
        if not self.nozzle_library:
            return []
        result = []
        for cat in self.nozzle_library.get("categories", []):
            for item in cat.get("items", []):
                result.append(self._nozzle_item_with_legacy_keys(item))
        return result

    def save_nozzle_library(self, to_factory: bool = False) -> bool:
        """노즐 라이브러리를 JSON 파일로 저장. 매개변수 의미는 save_pipe_library 참조."""
        if not self.nozzle_library:
            return False
        base = self._install_path() if to_factory else self._base_path()
        filepath = base / "nozzle_library.json"
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(self.nozzle_library, f, indent=2, ensure_ascii=False)
            logger.info("[LIB] 노즐 라이브러리 저장 완료: %s", filepath)
            return True
        except Exception as exc:
            logger.warning("[LIB] 노즐 라이브러리 저장 실패: %s", exc)
            return False

    @staticmethod
    def _nozzle_item_with_legacy_keys(item: dict) -> dict:
        """새 스키마 노즐 항목에 레거시 키 별칭을 추가한 얕은 복사를 반환.

        소비처(node_controller, dialog_node_properties 등)가
        item["name"], item["K_val"] 등으로 접근하는 것을 보호한다.
        """
        out = dict(item)
        # display_name → name 별칭
        if "name" not in out and "display_name" in out:
            out["name"] = out["display_name"]
        # Head 항목: K_SI → K_val 별칭 (노즐과 동일한 키로 접근 가능하게)
        if "K_val" not in out and "K_SI" in out:
            out["K_val"] = out["K_SI"]
        return out

    # =========================================================
    # 헤드(스프링클러) 라이브러리
    # =========================================================

    def get_head_spec(self, head_name: str) -> dict | None:
        """head_name(display_name 형식, 예: "80(5.6)")으로 Head 스펙을 반환."""
        if self.nozzle_library:
            for cat in self.nozzle_library.get("categories", []):
                if cat.get("category_id") == "head":
                    for item in cat.get("items", []):
                        if item.get("display_name") == head_name:
                            return {
                                "K_SI": float(item.get("K_SI", 0.0)),
                                "K_US": float(item.get("K_US", 0.0)),
                                "required_pressure_bar": float(item.get("min_p", 0.1)),
                            }
                    break
        return None

    def get_head_items(self) -> list[dict]:
        """Head 항목 전체 목록 반환 — 콤보박스/다이얼로그용.

        반환 형식: [{"display_name": "80(5.6)", "K_SI": 80.0, "K_US": 5.6, ...}, ...]
        """
        if self.nozzle_library:
            for cat in self.nozzle_library.get("categories", []):
                if cat.get("category_id") == "head":
                    return list(cat.get("items", []))
        return []

    # =========================================================
    # 피팅 라이브러리 v3.0 (NFPA 13 등가길이 기반)
    # =========================================================

    def load_fittings_v3(self, path: Path) -> list[FittingDefinition]:
        """fittings_library_v3.json → FittingDefinition 목록으로 로딩.

        M2.3 구현 — 스키마 검증(validate_library) 강제.
        DN 키 정규화(normalize_equivalent_lengths) 자동 적용.
        로딩 성공 시 self.fittings_library_v3 에 저장하고 목록을 반환.
        실패 시 FittingSchemaError 또는 기타 예외가 전파된다.
        """
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)

        raw_items = raw.get("items", [])
        result: list[FittingDefinition] = []
        for d in raw_items:
            result.append(FittingDefinition.from_dict(d))

        validate_library(result)
        self.fittings_library_v3 = result
        logger.info("[LIB] v3 부속 라이브러리 로딩 완료: %d개 항목 (%s)", len(result), path)
        return result

    def get_fitting_data_v3(self, fitting_id: str) -> FittingDefinition | None:
        """v3 라이브러리에서 ID로 FittingDefinition 조회."""
        if not self.fittings_library_v3:
            return None
        for item in self.fittings_library_v3:
            if item.id == str(fitting_id):
                return item
        return None

    def get_fittings_v3_flat(self) -> list[FittingDefinition]:
        """v3 라이브러리 전체 항목을 반환한다."""
        return list(self.fittings_library_v3) if self.fittings_library_v3 else []

    def save_fittings_library_v3(self, to_factory: bool = False) -> bool:
        """v3 부속 라이브러리를 JSON 파일로 저장.

        to_factory=False (기본): 사용자 폴더(%LOCALAPPDATA%\\K-Fire\\)에 저장.
        to_factory=True: 설치 폴더에 저장 (공장 초기값 갱신용).
        """
        if not self.fittings_library_v3:
            return False
        base = self._install_path() if to_factory else self._base_path()
        filepath = base / "fittings_library_v3.json"
        try:
            payload = {
                "schema_version": "3.0",
                "description": "K-Fire v3.0 부속·밸브 등가길이 라이브러리. NFPA 13 2025판 Table 28.2.4.4.1 기반.",
                "reference": {"pipe_standard": "ASTM A135 #40", "hazen_williams_C": 120, "unit": "m"},
                "items": [item.to_dict() for item in self.fittings_library_v3],
            }
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2, ensure_ascii=False)
            logger.info("[LIB] v3 부속 라이브러리 저장 완료: %s", filepath)
            return True
        except Exception as exc:
            logger.warning("[LIB] v3 부속 라이브러리 저장 실패: %s", exc)
            return False

    def find_nozzle_category_by_display(self, display_name: str) -> dict | None:
        """노즐 라이브러리에서 display_name이 일치하는 카테고리를 반환."""
        if not self.nozzle_library:
            return None
        for cat in self.nozzle_library.get("categories", []):
            if cat.get("display_name") == display_name:
                return cat
        return None

    def find_category_type_id(self, display_name: str) -> str | None:
        """display_name으로 노즐 라이브러리를 순회하여 type_id를 반환.
        못 찾으면 None."""
        cat = self.find_nozzle_category_by_display(display_name)
        if cat:
            return cat.get("type_id")
        return None

    # =========================================================
    # 펌프 라이브러리
    # =========================================================

    @staticmethod
    def next_pump_id(pump_library: dict) -> str:
        """펌프 라이브러리에서 다음 PUMP_NNN ID를 반환한다.

        pump_library: {"pumps": [...]} 형태의 dict.
        dialog_node_properties 와 dialog_pump_library 양쪽에서 공유하는 SSOT.
        """
        max_num = 0
        for item in (pump_library or {}).get("pumps", []):
            pid = item.get("id", "")
            if pid.startswith("PUMP_"):
                try:
                    num = int(pid[5:])
                    if num > max_num:
                        max_num = num
                except ValueError:
                    pass
        return f"PUMP_{max_num + 1:03d}"

    def get_pump_list(self) -> list[dict]:
        """펌프 라이브러리의 전체 항목 목록 반환 (콤보박스/다이얼로그용)."""
        if not self.pump_library:
            return []
        return list(self.pump_library.get("pumps", []))

    def get_pump_spec(self, pump_id: str) -> dict | None:
        """펌프 ID로 항목 조회."""
        if not self.pump_library:
            return None
        for item in self.pump_library.get("pumps", []):
            if item.get("id") == pump_id:
                return dict(item)
        return None

    def save_pump_library(self, to_factory: bool = False) -> bool:
        """펌프 라이브러리를 JSON 파일로 저장. 매개변수 의미는 save_pipe_library 참조."""
        if not self.pump_library:
            return False
        base = self._install_path() if to_factory else self._base_path()
        filepath = base / "pump_library.json"
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(self.pump_library, f, indent=2, ensure_ascii=False)
            logger.info("[LIB] 펌프 라이브러리 저장 완료: %s", filepath)
            return True
        except Exception as exc:
            logger.warning("[LIB] 펌프 라이브러리 저장 실패: %s", exc)
            return False

    # =========================================================
    # 내부 헬퍼
    # =========================================================

    def _install_path(self) -> Path:
        """설치(읽기 전용) 경로 — 번들된 기본 JSON이 위치한 곳"""
        try:
            return Path(__file__).resolve().parent.parent
        except Exception:
            return Path(".")

    def _factory_seed_bases(self) -> list[Path]:
        """공장 시드 JSON 후보 루트 (유저 폴더는 제외)."""
        bases: list[Path] = []
        if getattr(sys, "frozen", False):
            try:
                bases.append(Path(sys.executable).resolve().parent)
            except Exception:
                pass
            meipass = getattr(sys, "_MEIPASS", None)
            if meipass:
                bases.append(Path(meipass))
        bases.append(self._install_path())
        return bases

    def resolve_factory_seed(
        self,
        filename: str,
        *,
        seed_dir: Path | None = None,
    ) -> Path | None:
        """설치/번들 쪽 공장 시드 파일 경로를 찾는다. 유저 폴더는 보지 않는다."""
        if seed_dir is not None:
            candidate = Path(seed_dir) / filename
            return candidate if candidate.exists() else None

        seen: set[str] = set()
        for base in self._factory_seed_bases():
            for candidate in (
                base / filename,
                base / "data" / filename,
                base / "assets" / filename,
            ):
                key = str(candidate)
                if key in seen:
                    continue
                seen.add(key)
                if candidate.exists():
                    return candidate
        return None

    def reset_user_libraries_to_factory(
        self,
        *,
        seed_dir: Path | None = None,
    ) -> dict:
        """유저 전역 DB를 공장 시드로 백업·교체한다.

        Returns:
            {
              "ok": bool,              # 4종 모두 교체 성공
              "restored": [str, ...],  # 교체된 파일명
              "backed_up": [str, ...], # 백업 경로
              "errors": [str, ...],    # 실패 사유
            }
        """
        user_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "K-Fire"
        user_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        restored: list[str] = []
        backed_up: list[str] = []
        errors: list[str] = []

        for name in FACTORY_LIBRARY_FILES:
            dest = user_dir / name
            src = self.resolve_factory_seed(name, seed_dir=seed_dir)
            if src is None:
                errors.append(f"seed_missing:{name}")
                logger.warning("[LIB] 공장 시드 없음 — 초기화 스킵: %s", name)
                continue
            try:
                if dest.exists():
                    bak = user_dir / f"{name}.bak.{ts}"
                    # 동일 초 재실행 시 충돌 방지
                    n = 1
                    while bak.exists():
                        bak = user_dir / f"{name}.bak.{ts}_{n}"
                        n += 1
                    dest.replace(bak)
                    backed_up.append(str(bak))
                shutil.copy2(src, dest)
                restored.append(name)
                logger.info("[LIB] 공장 시드로 초기화: %s -> %s", src, dest)
            except Exception as exc:
                errors.append(f"copy_failed:{name}:{exc}")
                logger.warning("[LIB] 공장 시드 초기화 실패 (%s): %s", name, exc)

        return {
            "ok": len(restored) == len(FACTORY_LIBRARY_FILES) and not errors,
            "restored": restored,
            "backed_up": backed_up,
            "errors": errors,
        }

    def _base_path(self) -> Path:
        """JSON 파일 저장 경로 — %LOCALAPPDATA%/K-Fire (쓰기 가능)
        Program Files 쓰기 금지 우회. 최초 호출 시 기본 JSON을 복사한다."""
        user_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "K-Fire"
        user_dir.mkdir(parents=True, exist_ok=True)

        for name in FACTORY_LIBRARY_FILES:
            dest = user_dir / name
            src = self.resolve_factory_seed(name)
            if not dest.exists() and src is not None:
                try:
                    shutil.copy2(src, dest)
                    logger.info("[LIB] 기본 라이브러리 복사: %s -> %s", src, dest)
                except Exception as exc:
                    logger.warning("[LIB] 기본 라이브러리 복사 실패: %s", exc)

        return user_dir

    # =========================================================
    # 사용자 노즐 카테고리 색상 관리
    # =========================================================

    def get_nozzle_category_color(self, category_id: str) -> tuple[int, int, int] | None:
        """카테고리에 저장된 color_rgb를 (R, G, B) 튜플로 반환.

        is_system_protected 카테고리나 color_rgb 필드가 없는 카테고리는 None 반환.
        None을 받은 호출부는 MD5 해시 기반 폴백으로 처리한다.
        """
        if not self.nozzle_library:
            return None
        for cat in self.nozzle_library.get("categories", []):
            if cat.get("category_id") != category_id:
                continue
            raw = cat.get("color_rgb")
            if isinstance(raw, (list, tuple)) and len(raw) == 3:
                try:
                    return (int(raw[0]), int(raw[1]), int(raw[2]))
                except Exception:
                    pass
        return None

    def assign_colors_to_user_categories(self) -> int:
        """color_rgb 없는 사용자 노즐 카테고리에 팔레트 색을 순차 배정한다.

        - is_system_protected=True 카테고리는 건너뜀
        - 이미 다른 사용자 카테고리가 쓰는 색을 제외한 미사용 팔레트 색 중 첫 번째 배정
        - 팔레트가 고갈되면 MD5 해시 폴백 사용
        - 반환: 새로 색이 배정된 카테고리 수
        """
        from domain.node_visual import _USER_NOZZLE_PALETTE, _auto_color_for_user_category

        if not self.nozzle_library:
            return 0

        # 이미 쓰이고 있는 색 목록 수집 (사용자 카테고리만)
        used_colors: set[tuple[int, int, int]] = set()
        for cat in self.nozzle_library.get("categories", []):
            if cat.get("is_system_protected", False):
                continue
            raw = cat.get("color_rgb")
            if isinstance(raw, (list, tuple)) and len(raw) == 3:
                try:
                    used_colors.add((int(raw[0]), int(raw[1]), int(raw[2])))
                except Exception:
                    pass

        assigned = 0
        for cat in self.nozzle_library.get("categories", []):
            if cat.get("is_system_protected", False):
                continue
            raw = cat.get("color_rgb")
            if isinstance(raw, (list, tuple)) and len(raw) == 3:
                continue  # 이미 색 있음

            cat_id = str(cat.get("category_id", ""))
            available = [c for c in _USER_NOZZLE_PALETTE if c not in used_colors]
            if available:
                chosen = available[0]
            else:
                logger.warning("[LIB] 팔레트 고갈 — 해시 폴백: category_id='%s'", cat_id)
                chosen = _auto_color_for_user_category(cat_id)

            cat["color_rgb"] = list(chosen)
            used_colors.add(chosen)
            logger.debug("[LIB] 카테고리 색상 배정: '%s' → %s", cat_id, chosen)
            assigned += 1

        return assigned

    # =========================================================
    # 내장 라이브러리 정규화
    # =========================================================

    def normalize_embedded_library_category_ids(self) -> int:
        """내장 노즐/밸브 라이브러리의 시스템 카테고리 ID를 현행 규칙으로 정규화한다.

        구버전 .kfp 파일에 내장된 라이브러리는 category_id 명명 규칙이
        현재와 다를 수 있다(예: Water Spray의 "nozzle" → "water_spray").
        is_system_protected=True 인 카테고리만 대상으로 하여
        사용자가 직접 추가한 카테고리는 보호한다.

        반환: 교정된 카테고리 수 (0이면 변경 없음)
        """
        changed = 0
        for lib in (self.nozzle_library,):
            if not lib:
                continue
            for cat in lib.get("categories", []):
                if not cat.get("is_system_protected", False):
                    continue
                old_cid = str(cat.get("category_id", "") or "")
                new_cid = canonicalize_category_id(old_cid, cat.get("display_name", ""))
                if new_cid and new_cid != old_cid:
                    cat["category_id"] = new_cid
                    logger.debug(
                        "[LIB] 카테고리 ID 정규화: '%s' → '%s' (display='%s')",
                        old_cid, new_cid, cat.get("display_name", ""),
                    )
                    changed += 1
        return changed
