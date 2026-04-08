from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import cv2
import numpy as np

try:
    import torch
except Exception:  # pragma: no cover - optional at runtime
    torch = None

try:
    from ultralytics import YOLO
except Exception:  # pragma: no cover - optional at runtime
    YOLO = None


@dataclass
class RenderTransform:
    min_x: float
    min_y: float
    scale: float
    width: int
    height: int
    pad_x: float
    pad_y: float

    def world_to_image(self, x: float, y: float) -> tuple[int, int]:
        ix = int(round((x - self.min_x) * self.scale + self.pad_x))
        iy = int(round(self.height - ((y - self.min_y) * self.scale + self.pad_y)))
        return ix, iy

    def image_to_world_box(self, x1: int, y1: int, x2: int, y2: int) -> dict[str, float]:
        min_x = (x1 - self.pad_x) / self.scale + self.min_x
        max_x = (x2 - self.pad_x) / self.scale + self.min_x
        max_y = self.min_y + ((self.height - y1) - self.pad_y) / self.scale
        min_y = self.min_y + ((self.height - y2) - self.pad_y) / self.scale
        return {"minX": min_x, "minY": min_y, "maxX": max_x, "maxY": max_y}


def render_entities(
    entities: list[dict],
    rect: dict[str, float],
    visible_layers: set[str],
    width: int = 1280,
    height: int = 1280,
) -> tuple[np.ndarray, RenderTransform] | None:
    min_x = float(rect["minX"])
    min_y = float(rect["minY"])
    max_x = float(rect["maxX"])
    max_y = float(rect["maxY"])
    world_w = max(max_x - min_x, 1.0)
    world_h = max(max_y - min_y, 1.0)
    scale = min((width - 32) / world_w, (height - 32) / world_h)
    if scale <= 0:
        return None
    pad_x = (width - world_w * scale) / 2
    pad_y = (height - world_h * scale) / 2
    transform = RenderTransform(min_x=min_x, min_y=min_y, scale=scale, width=width, height=height, pad_x=pad_x, pad_y=pad_y)
    image = np.zeros((height, width), dtype=np.uint8)

    for entity in entities:
        if visible_layers and entity.get("layer") not in visible_layers:
            continue
        entity_type = entity.get("type")
        if entity_type == "LINE":
            p1 = transform.world_to_image(entity["start"]["x"], entity["start"]["y"])
            p2 = transform.world_to_image(entity["end"]["x"], entity["end"]["y"])
            cv2.line(image, p1, p2, 255, 2, cv2.LINE_AA)
        elif entity_type == "LWPOLYLINE":
            pts = np.array([transform.world_to_image(point["x"], point["y"]) for point in entity.get("points", [])], dtype=np.int32)
            if len(pts) >= 2:
                cv2.polylines(image, [pts], bool(entity.get("closed")), 255, 2, cv2.LINE_AA)
        elif entity_type == "CIRCLE":
            center = transform.world_to_image(entity["center"]["x"], entity["center"]["y"])
            radius = max(int(round(float(entity["radius"]) * scale)), 1)
            cv2.circle(image, center, radius, 255, 2, cv2.LINE_AA)
        elif entity_type == "ARC":
            center = transform.world_to_image(entity["center"]["x"], entity["center"]["y"])
            radius = max(int(round(float(entity["radius"]) * scale)), 1)
            cv2.ellipse(
                image,
                center,
                (radius, radius),
                0,
                -float(entity["endAngle"]),
                -float(entity["startAngle"]),
                255,
                2,
                cv2.LINE_AA,
            )
    return image, transform


def merge_boxes(boxes: list[dict[str, float]], center_distance: float = 220.0) -> list[dict[str, float]]:
    merged: list[dict[str, float]] = []
    for box in boxes:
        found = None
        for other in merged:
            if world_overlap(box, other) or world_center_distance(box, other) < center_distance:
                found = other
                break
        if found is None:
            merged.append(dict(box))
        else:
            found["minX"] = min(found["minX"], box["minX"])
            found["minY"] = min(found["minY"], box["minY"])
            found["maxX"] = max(found["maxX"], box["maxX"])
            found["maxY"] = max(found["maxY"], box["maxY"])
    return merged


def world_overlap(left: dict[str, float], right: dict[str, float]) -> bool:
    return not (
        left["maxX"] < right["minX"]
        or right["maxX"] < left["minX"]
        or left["maxY"] < right["minY"]
        or right["maxY"] < left["minY"]
    )


def world_center_distance(left: dict[str, float], right: dict[str, float]) -> float:
    lx = (left["minX"] + left["maxX"]) / 2
    ly = (left["minY"] + left["maxY"]) / 2
    rx = (right["minX"] + right["maxX"]) / 2
    ry = (right["minY"] + right["maxY"]) / 2
    return float(np.hypot(lx - rx, ly - ry))


class TriangleHeadTemplateDetector:
    def __init__(self, template_dir: Path) -> None:
        self.templates = self._load_templates(template_dir)

    def _load_templates(self, template_dir: Path) -> list[np.ndarray]:
        templates: list[np.ndarray] = []
        for path in sorted(template_dir.glob("*.png")):
            image = cv2.imread(str(path), cv2.IMREAD_GRAYSCALE)
            if image is None:
                continue
            _, thresh = cv2.threshold(image, 16, 255, cv2.THRESH_BINARY)
            coords = cv2.findNonZero(thresh)
            if coords is None:
                continue
            x, y, w, h = cv2.boundingRect(coords)
            cropped = thresh[y:y + h, x:x + w]
            edges = cv2.Canny(cropped, 40, 120)
            if edges.size > 0:
                templates.append(edges)
        return templates

    def detect(self, entities: list[dict], rect: dict[str, float], visible_layers: set[str]) -> list[dict[str, float]]:
        if not self.templates:
            return []
        render = render_entities(entities, rect, visible_layers, width=1024, height=1024)
        if render is None:
            return []
        image, transform = render
        edges = cv2.Canny(image, 40, 120)
        detections = self._match_templates(edges)
        world_boxes = [transform.image_to_world_box(*box) for box in detections]
        return merge_boxes(world_boxes, center_distance=350.0)

    def _match_templates(self, edges: np.ndarray) -> list[tuple[int, int, int, int]]:
        candidates: list[tuple[int, int, int, int, float]] = []
        scales = (0.6, 0.8, 1.0, 1.2, 1.4, 1.8)
        for template in self.templates:
            for scale in scales:
                resized = cv2.resize(template, None, fx=scale, fy=scale, interpolation=cv2.INTER_LINEAR)
                th, tw = resized.shape[:2]
                if th < 12 or tw < 12 or th >= edges.shape[0] or tw >= edges.shape[1]:
                    continue
                result = cv2.matchTemplate(edges, resized, cv2.TM_CCOEFF_NORMED)
                ys, xs = np.where(result >= 0.42)
                for x, y in zip(xs.tolist(), ys.tolist()):
                    score = float(result[y, x])
                    candidates.append((x, y, x + tw, y + th, score))
        candidates.sort(key=lambda item: item[4], reverse=True)
        selected: list[tuple[int, int, int, int]] = []
        for x1, y1, x2, y2, _score in candidates:
            candidate = (x1, y1, x2, y2)
            if any(box_iou(candidate, other) > 0.2 or box_center_distance(candidate, other) < 28 for other in selected):
                continue
            selected.append(candidate)
        return selected


class TriangleHeadYoloDetector:
    def __init__(self, model_path: Path) -> None:
        self.model_path = model_path
        self.model = None
        if YOLO is not None and model_path.exists():
            self.model = YOLO(str(model_path))

    @property
    def available(self) -> bool:
        return self.model is not None

    def detect(self, entities: list[dict], rect: dict[str, float], visible_layers: set[str]) -> list[dict[str, float]]:
        if self.model is None:
            return []
        render = render_entities(entities, rect, visible_layers, width=1280, height=1280)
        if render is None:
            return []
        image, transform = render
        rgb = cv2.cvtColor(image, cv2.COLOR_GRAY2BGR)
        device = 0 if torch is not None and torch.cuda.is_available() else "cpu"
        result = self.model.predict(
            source=rgb,
            imgsz=1280,
            conf=0.18,
            iou=0.35,
            device=device,
            verbose=False,
        )[0]
        boxes: list[dict[str, float]] = []
        if result.boxes is not None:
            for xyxy in result.boxes.xyxy.detach().cpu().numpy().tolist():
                x1, y1, x2, y2 = [int(round(v)) for v in xyxy]
                boxes.append(transform.image_to_world_box(x1, y1, x2, y2))
        return merge_boxes(boxes, center_distance=240.0)


class TriangleHeadDetector:
    def __init__(self, template_dir: Path, model_path: Path) -> None:
        self.template_detector = TriangleHeadTemplateDetector(template_dir)
        self.yolo_detector = TriangleHeadYoloDetector(model_path)

    def detect(self, entities: list[dict], rect: dict[str, float], visible_layers: set[str]) -> list[dict[str, float]]:
        detections: list[dict[str, float]] = []
        if self.yolo_detector.available:
            detections.extend(self.yolo_detector.detect(entities, rect, visible_layers))
        detections.extend(self.template_detector.detect(entities, rect, visible_layers))
        return merge_boxes(detections, center_distance=260.0)


def box_iou(left: tuple[int, int, int, int], right: tuple[int, int, int, int]) -> float:
    x1 = max(left[0], right[0])
    y1 = max(left[1], right[1])
    x2 = min(left[2], right[2])
    y2 = min(left[3], right[3])
    inter = max(0, x2 - x1) * max(0, y2 - y1)
    if inter == 0:
        return 0.0
    area_l = (left[2] - left[0]) * (left[3] - left[1])
    area_r = (right[2] - right[0]) * (right[3] - right[1])
    return inter / max(area_l + area_r - inter, 1)


def box_center_distance(left: tuple[int, int, int, int], right: tuple[int, int, int, int]) -> float:
    lx = (left[0] + left[2]) / 2
    ly = (left[1] + left[3]) / 2
    rx = (right[0] + right[2]) / 2
    ry = (right[1] + right[3]) / 2
    return float(np.hypot(lx - rx, ly - ry))
