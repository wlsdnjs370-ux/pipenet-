from __future__ import annotations

from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any

from head_detector import TriangleHeadDetector


@dataclass(slots=True)
class AIVisionResult:
    heads: list[dict[str, Any]] = field(default_factory=list)
    labels: list[dict[str, Any]] = field(default_factory=list)
    regions: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    model_versions: dict[str, str] = field(default_factory=dict)
    stats: dict[str, Any] = field(default_factory=dict)


class AIVisionExtractor:
    """AI-assisted DXF object extraction.

    Current production scope is intentionally narrow:
    - trained YOLO model for triangle sprinkler-head detection
    - template fallback from the existing head detector

    OCR and segmentation stay as extension points until corresponding
    datasets and runtime dependencies are prepared.
    """

    def __init__(
        self,
        *,
        base_dir: Path,
        model_path: Path | None = None,
        template_dir: Path | None = None,
    ) -> None:
        self.base_dir = base_dir
        self.model_path = model_path or self._resolve_model_path()
        self.template_dir = template_dir or (base_dir / "data" / "head_templates")
        self.detector = TriangleHeadDetector(self.template_dir, self.model_path)

    def enhance_from_payload(
        self,
        *,
        dxf_path: Path,
        cad_payload: dict[str, Any],
        visible_layers: set[str] | None = None,
    ) -> AIVisionResult:
        entities = list(cad_payload.get("entities") or [])
        bounds = cad_payload.get("bounds") or {}
        network_layers = set(visible_layers or cad_payload.get("networkLayers") or [])

        warnings: list[str] = []
        if not self.detector.yolo_detector.available:
            warnings.append("Trained YOLO weights unavailable; template fallback only.")
        if not bounds:
            warnings.append("CAD payload bounds missing; AI extraction skipped.")
            return AIVisionResult(
                warnings=warnings,
                model_versions={
                    "triangle_head_detector": self.model_path.name if self.model_path.exists() else "missing",
                },
                stats={"head_count": 0, "detector_mode": "unavailable"},
            )

        head_boxes = self.detector.detect(entities, bounds, network_layers)
        return AIVisionResult(
            heads=[
                {
                    "id": f"ai-head-{idx:04d}",
                    "bbox": box,
                    "source": "triangle_head_yolo" if self.detector.yolo_detector.available else "template",
                }
                for idx, box in enumerate(head_boxes, start=1)
            ],
            warnings=warnings,
            model_versions={
                "triangle_head_detector": self.model_path.name if self.model_path.exists() else "missing",
            },
            stats={
                "head_count": len(head_boxes),
                "detector_mode": "yolo+template" if self.detector.yolo_detector.available else "template",
                "dxf_path": str(dxf_path),
                "network_layer_count": len(network_layers),
            },
        )

    def _resolve_model_path(self) -> Path:
        candidates = [
            self.base_dir / "models" / "triangle_head_yolo_ai" / "weights" / "best.pt",
            self.base_dir / "runs" / "detect" / "models" / "triangle_head_yolo_ai" / "weights" / "best.pt",
            self.base_dir / "models" / "triangle_head_yolo" / "weights" / "best.pt",
            self.base_dir / "runs" / "detect" / "models" / "triangle_head_yolo" / "weights" / "best.pt",
            self.base_dir / "yolo26n.pt",
            self.base_dir / "yolo11n.pt",
        ]
        for path in candidates:
            if path.exists():
                return path
        return candidates[0]


@lru_cache(maxsize=1)
def get_cached_ai_vision_extractor(base_dir: str) -> AIVisionExtractor:
    return AIVisionExtractor(base_dir=Path(base_dir))
