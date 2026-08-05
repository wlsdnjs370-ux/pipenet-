# -*- coding: utf-8 -*-
"""C2B — 화재조기진압용 스프링클러(NFTC 103B) 활성 판정 — 지시서 §6.

**랙크식 창고라고 자동으로 103B 가 되지 않는다.** 랙크식의 기본 트랙은 NFTC 103
2.7.2 이고, 103B 는 별도 설비 종류다. 설치장소 구조 조건과 저장물 제한이 붙으며
그 판단은 사람이 한다. 천장고는 활성 조건이 아니라 활성 뒤 K값을 정하는 변수다.

판정이 헤드 배치보다 앞서는 이유 — 103B 가 켜지면 NFTC 103 의 수평거리 5종 룰이
통째로 무효가 되는데, 그 수평거리를 실제로 쓰는 단계가 C4 다. 뒤늦게 켜면 이미
깔린 헤드가 근거를 잃는다.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

# 103B 가 켜지면 근거를 잃는 constraints 필드. 값을 그대로 두면 하류가 NFTC 103
# 기준으로 계산하고도 통과한다 — 지우고 왜 없는지를 남긴다.
VOID_FIELDS = (
    "horizontal_distance_m",
    "head_spacing_square_m",
    "wall_clearance_max_m",
    "k_factor",
)

# 사람이 답해야 하는 세 조건. 화면·라우트가 각자 나열하면 한쪽만 늘었을 때
# 묻지 않은 조건이 거짓으로 확정된다.
INPUTS = ("owner_adopts_esfr", "site_conditions_ok", "commodity_restricted")


def activate_103b(*, owner_adopts_esfr: bool, site_conditions_ok: bool,
                  commodity_restricted: bool) -> bool:
    """세 조건이 모두 사람 손으로 확정됐을 때만 참."""
    return owner_adopts_esfr and site_conditions_ok and not commodity_restricted


def _reason(*, owner_adopts_esfr: bool, site_conditions_ok: bool,
            commodity_restricted: bool) -> str:
    if not owner_adopts_esfr:
        return "발주자가 103B 를 채택하지 않았다 — 기본 트랙은 NFTC 103 이다."
    if not site_conditions_ok:
        return "설치장소 구조 조건을 만족하지 않는다."
    if commodity_restricted:
        return "저장물이 103B 제한 품목이다(제4류 위험물·타이어·목재·종이·섬유류 등)."
    return "발주자 채택 + 설치장소 조건 충족 + 저장물 제한 없음."


@dataclass(frozen=True)
class EsfrDecision:
    """사람이 내린 103B 판단. 세 입력과 판단자를 함께 남긴다."""

    active: bool
    owner_adopts_esfr: bool
    site_conditions_ok: bool
    commodity_restricted: bool
    operator: str
    decided_at: str
    reason: str
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema": "fncadnet.esfr/1",
            "active": self.active,
            "owner_adopts_esfr": self.owner_adopts_esfr,
            "site_conditions_ok": self.site_conditions_ok,
            "commodity_restricted": self.commodity_restricted,
            "operator": self.operator,
            "decided_at": self.decided_at,
            "reason": self.reason,
            "note": self.note,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "EsfrDecision":
        """`active` 는 파일 값을 믿지 않고 세 조건에서 다시 계산한다.

        `apply()` 가 `active` 하나만 보고 NFTC 103 근거를 지우므로, 손으로 고친
        `esfr.json` 이 조건과 어긋난 채 들어오면 설계 근거가 소리 없이 사라진다.
        """
        conditions = {key: bool(d.get(key)) for key in INPUTS}
        return cls(
            active=activate_103b(**conditions),
            operator=str(d.get("operator") or ""),
            decided_at=str(d.get("decided_at") or ""),
            reason=str(d.get("reason") or ""),
            note=str(d.get("note") or ""),
            **conditions,
        )


def decide(*, operator: str, owner_adopts_esfr: bool, site_conditions_ok: bool,
           commodity_restricted: bool, decided_at: str, note: str = "") -> EsfrDecision:
    """세 입력 + 판단자 → 결정. 판단자 없이는 만들지 않는다.

    `owner_adopts_esfr` 는 사람이 입력한다. 도면·용도에서 유도하지 마라.
    """
    if not (operator or "").strip():
        raise ValueError("103B 활성은 사람이 판단한다 — 확정자를 입력하세요.")
    conditions = {
        "owner_adopts_esfr": bool(owner_adopts_esfr),
        "site_conditions_ok": bool(site_conditions_ok),
        "commodity_restricted": bool(commodity_restricted),
    }
    return EsfrDecision(
        active=activate_103b(**conditions),
        operator=operator.strip(),
        decided_at=decided_at,
        reason=_reason(**conditions),
        note=(note or "").strip(),
        **conditions,
    )


def apply(constraints: dict, decision: EsfrDecision) -> dict:
    """constraints.json 에 판단을 싣는다. 활성이면 무효 필드를 비운다.

    비운 자리에 103B 표를 채워 넣지 않는다 — 103B 의 K값·간격표는 아직 구현되지
    않았고, 없는 표를 NFTC 103 값으로 대신하면 그 설계는 근거 없이 통과한다.
    """
    out = dict(constraints)
    out["esfr"] = decision.to_dict()
    if not decision.active:
        return out
    for field in VOID_FIELDS:
        out[field] = None
    out["esfr"]["voided_fields"] = list(VOID_FIELDS)
    out["trace"] = [t for t in out.get("trace", []) if t.get("field") not in VOID_FIELDS]
    return out
