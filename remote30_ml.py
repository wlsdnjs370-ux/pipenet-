"""Remote 30 ML 헤드 검출 - 타일 분할 + 고해상도 추론.

head_detector.TriangleHeadYoloDetector 는 PNG 1280×1280 으로 도면 전체를 압축하는데,
한국 sprinkler 도면(예: 대명동201동 = 800m × 4.7Mm 단위)에서는 헤드 심볼이
픽셀 미만으로 깨져 거의 검출되지 않는다.

이 모듈은 도면을 NxN 타일로 분할하고 각 타일을 독립적으로 1280px PNG 로 렌더해서
YOLO 추론한 뒤, world-coord 결과를 병합한다. 헤드 심볼이 충분한 픽셀을 갖게 됨.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import numpy as np

from head_detector import (
    render_entities,
    merge_boxes,
    TriangleHeadYoloDetector,
)

try:
    import cv2
except ImportError:
    cv2 = None

HEAD_TEMPLATES_DIR = Path("data/head_templates")
ALARM_TEMPLATES_DIR = Path("data/alarm_templates")

# 학습된 sprinkler 다중 클래스 모델 우선. 없으면 fallback to triangle_head_yolo.
SPRINKLER_YOLO_MODEL_CANDIDATES = [
    Path("models/sprinkler_yolo/weights/best.pt"),
    Path("models/triangle_head_yolo_ai/weights/best.pt"),
    Path("models/triangle_head_yolo/weights/best.pt"),
]


def resolve_sprinkler_model_path() -> Path | None:
    """sprinkler_yolo 모델이 학습돼 있으면 우선 사용, 없으면 triangle_head_yolo fallback."""
    for p in SPRINKLER_YOLO_MODEL_CANDIDATES:
        if p.exists():
            return p
    return None


def _entity_world_bbox(entity: dict) -> tuple[float, float, float, float] | None:
    """Entity 의 world bbox (minX, minY, maxX, maxY) 반환. INSERT 등 점 entity 는 작은 사각."""
    t = entity.get("type")
    try:
        if t == "LINE":
            x1 = float(entity["start"]["x"]); y1 = float(entity["start"]["y"])
            x2 = float(entity["end"]["x"]); y2 = float(entity["end"]["y"])
            return (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        if t == "LWPOLYLINE":
            pts = entity.get("points") or []
            if not pts:
                return None
            xs = [float(p["x"]) for p in pts]
            ys = [float(p["y"]) for p in pts]
            return (min(xs), min(ys), max(xs), max(ys))
        if t in ("CIRCLE", "ARC"):
            cx = float(entity["center"]["x"]); cy = float(entity["center"]["y"])
            r = float(entity["radius"])
            return (cx - r, cy - r, cx + r, cy + r)
    except (KeyError, TypeError, ValueError):
        return None
    return None


def _rect_overlaps(rect: tuple[float, float, float, float], box: tuple[float, float, float, float]) -> bool:
    return not (rect[2] < box[0] or box[2] < rect[0] or rect[3] < box[1] or box[3] < rect[1])


def detect_heads_with_tiles(
    *,
    entities: list[dict[str, Any]],
    rect: dict[str, float],
    visible_layers: set[str],
    model_path: Path,
    tile_grid: int = 2,
    tile_px: int = 1280,
    overlap: float = 0.10,
    conf: float = 0.18,
    iou: float = 0.35,
    class_names: list[str] | None = None,
) -> dict[str, Any]:
    """타일 분할 YOLO 헤드 검출.

    Returns:
        {
          'boxes': [{'minX','minY','maxX','maxY'}, ...],  # world-coord bbox 병합 결과
          'tiles': N,           # 처리한 타일 수
          'tile_grid': tile_grid,
          'tile_px': tile_px,
          'image_count': 실제 추론한 PNG 수
        }
    """
    if cv2 is None:
        raise RuntimeError("OpenCV(cv2) 미설치 - 타일 추론 불가")
    detector = TriangleHeadYoloDetector(Path(model_path))
    if not detector.available:
        raise RuntimeError("YOLO 모델 가중치를 로드할 수 없습니다 (head_detector.TriangleHeadYoloDetector.available=False)")

    if tile_grid < 1:
        tile_grid = 1
    tile_grid = min(tile_grid, 8)  # 안전 상한 (8x8 = 64 타일)
    tile_px = max(640, min(int(tile_px), 4096))

    minX = float(rect["minX"]); minY = float(rect["minY"])
    maxX = float(rect["maxX"]); maxY = float(rect["maxY"])
    dx = (maxX - minX) / tile_grid
    dy = (maxY - minY) / tile_grid

    all_boxes: list[dict[str, float]] = []
    image_count = 0

    for i in range(tile_grid):
        for j in range(tile_grid):
            tile_minX = minX + i * dx - overlap * dx
            tile_minY = minY + j * dy - overlap * dy
            tile_maxX = minX + (i + 1) * dx + overlap * dx
            tile_maxY = minY + (j + 1) * dy + overlap * dy
            tile_rect = {"minX": tile_minX, "minY": tile_minY, "maxX": tile_maxX, "maxY": tile_maxY}
            tile_bounds = (tile_minX, tile_minY, tile_maxX, tile_maxY)

            tile_entities = []
            for e in entities:
                ebbox = _entity_world_bbox(e)
                if ebbox is None:
                    continue
                if _rect_overlaps(tile_bounds, ebbox):
                    tile_entities.append(e)
            if len(tile_entities) < 5:
                continue  # 거의 빈 타일은 스킵 (배경뿐)

            render = render_entities(tile_entities, tile_rect, visible_layers, width=tile_px, height=tile_px)
            if render is None:
                continue
            image, transform = render
            rgb = cv2.cvtColor(image, cv2.COLOR_GRAY2BGR)
            result = detector.model.predict(
                source=rgb, imgsz=tile_px, conf=conf, iou=iou, device="cpu", verbose=False,
            )[0]
            image_count += 1
            if result.boxes is None:
                continue
            xyxy_arr = result.boxes.xyxy.detach().cpu().numpy().tolist()
            cls_arr = result.boxes.cls.detach().cpu().numpy().tolist() if result.boxes.cls is not None else [0] * len(xyxy_arr)
            conf_arr = result.boxes.conf.detach().cpu().numpy().tolist() if result.boxes.conf is not None else [1.0] * len(xyxy_arr)
            for xyxy, c, cf in zip(xyxy_arr, cls_arr, conf_arr):
                x1, y1, x2, y2 = [int(round(v)) for v in xyxy]
                box_world = transform.image_to_world_box(x1, y1, x2, y2)
                box_world["cls"] = int(c)
                box_world["conf"] = float(cf)
                all_boxes.append(box_world)

    # 클래스별 분리해서 merge (다른 클래스끼리 합치지 않음)
    classes_seen = sorted({b.get("cls", 0) for b in all_boxes})
    merged_all: list[dict] = []
    for cls in classes_seen:
        cls_boxes = [b for b in all_boxes if b.get("cls", 0) == cls]
        merged = merge_boxes(cls_boxes, center_distance=200.0)
        for m in merged:
            m["cls"] = cls
        merged_all.extend(merged)

    return {
        "boxes": merged_all,
        "tiles": tile_grid * tile_grid,
        "tile_grid": tile_grid,
        "tile_px": tile_px,
        "image_count": image_count,
        "raw_detections": len(all_boxes),
        "class_names": class_names or [],
    }


# ============================================================
# Phase 1: 색상 기반 검출 (직접 도면 DXF → 색상 마스크)
# ============================================================
def detect_by_color_on_dxf(
    *,
    entities: list[dict[str, Any]],
    rect: dict[str, float],
    visible_layers: set[str],
    tile_grid: int = 3,
    tile_px: int = 1600,
    overlap: float = 0.10,
    color_specs: list[dict] | None = None,
    min_area_px: int = 8,
    max_area_px: int = 800,
) -> dict[str, Any]:
    """DXF entity 를 컬러로 PNG 렌더 후 HSV 마스크로 빨강/노랑 헤드 심볼 검출.

    Red triangle/dot, Yellow circle 같은 sprinkler 표준 심볼은 매우 일관된 색상이라
    HSV 임계값 + contour detection 만으로도 정확히 잡을 수 있다.

    color_specs: [{"name", "hsv_low", "hsv_high"}] — 기본은 빨강 + 노랑
    """
    if cv2 is None:
        raise RuntimeError("OpenCV(cv2) 미설치")
    import numpy as _np

    if color_specs is None:
        color_specs = [
            {"name": "red", "hsv_low": (0, 120, 80),   "hsv_high": (10, 255, 255)},
            {"name": "red2", "hsv_low": (170, 120, 80), "hsv_high": (180, 255, 255)},
            {"name": "yellow", "hsv_low": (20, 120, 120), "hsv_high": (35, 255, 255)},
        ]

    minX = float(rect["minX"]); minY = float(rect["minY"])
    maxX = float(rect["maxX"]); maxY = float(rect["maxY"])
    dx = (maxX - minX) / tile_grid
    dy = (maxY - minY) / tile_grid

    all_detections: list[dict] = []
    image_count = 0

    for i in range(tile_grid):
        for j in range(tile_grid):
            tile_rect = {
                "minX": minX + i * dx - overlap * dx,
                "minY": minY + j * dy - overlap * dy,
                "maxX": minX + (i + 1) * dx + overlap * dx,
                "maxY": minY + (j + 1) * dy + overlap * dy,
            }
            tile_bounds = (tile_rect["minX"], tile_rect["minY"], tile_rect["maxX"], tile_rect["maxY"])

            tile_entities = []
            for e in entities:
                ebbox = _entity_world_bbox(e)
                if ebbox is None: continue
                if _rect_overlaps(tile_bounds, ebbox):
                    tile_entities.append(e)
            if len(tile_entities) < 5:
                continue

            # 컬러 PNG 렌더 — entity 의 color 정보 사용
            color_img = _render_entities_color(tile_entities, tile_rect, visible_layers, tile_px)
            if color_img is None:
                continue
            image_count += 1
            hsv = cv2.cvtColor(color_img, cv2.COLOR_BGR2HSV)

            # 각 색상 마스크 → contour → 검출
            combined_mask = _np.zeros(hsv.shape[:2], dtype=_np.uint8)
            for spec in color_specs:
                low = _np.array(spec["hsv_low"], dtype=_np.uint8)
                high = _np.array(spec["hsv_high"], dtype=_np.uint8)
                m = cv2.inRange(hsv, low, high)
                combined_mask = cv2.bitwise_or(combined_mask, m)

            # morph: 작은 노이즈 제거 + 가까운 픽셀 묶기
            kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
            combined_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)
            combined_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_OPEN, kernel)

            contours, _ = cv2.findContours(combined_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            for c in contours:
                area = cv2.contourArea(c)
                if area < min_area_px or area > max_area_px:
                    continue
                x, y, w, h = cv2.boundingRect(c)
                # 검출된 픽셀 좌표 → world 좌표 (tile 의 inverse transform 사용)
                from head_detector import render_entities as _render
                world_w = max(tile_rect["maxX"] - tile_rect["minX"], 1.0)
                world_h = max(tile_rect["maxY"] - tile_rect["minY"], 1.0)
                scale_px = min((tile_px - 32) / world_w, (tile_px - 32) / world_h)
                pad_x = (tile_px - world_w * scale_px) / 2
                pad_y = (tile_px - world_h * scale_px) / 2
                # render_entities 는 y를 위→아래로 뒤집어 그림. inverse:
                wx_min = tile_rect["minX"] + (x - pad_x) / scale_px
                wx_max = tile_rect["minX"] + (x + w - pad_x) / scale_px
                wy_max = tile_rect["maxY"] - (y - pad_y) / scale_px
                wy_min = tile_rect["maxY"] - (y + h - pad_y) / scale_px
                all_detections.append({
                    "minX": min(wx_min, wx_max), "minY": min(wy_min, wy_max),
                    "maxX": max(wx_min, wx_max), "maxY": max(wy_min, wy_max),
                })

    merged = merge_boxes(all_detections, center_distance=200.0)
    return {
        "boxes": merged,
        "tiles": tile_grid * tile_grid,
        "tile_grid": tile_grid,
        "tile_px": tile_px,
        "image_count": image_count,
        "raw_detections": len(all_detections),
    }


def _render_entities_color(entities, rect, visible_layers, tile_px):
    """entity 의 DXF color index 를 BGR 색상으로 변환해 컬러 PNG 렌더.
    sprinkler 도면의 빨간 헤드/노란 헤드를 색상 마스크로 잡기 위해 필요.
    DXF ACI(Auto CAD Color Index) → BGR 매핑.
    """
    if cv2 is None:
        return None
    import numpy as _np
    minX = float(rect["minX"]); minY = float(rect["minY"])
    maxX = float(rect["maxX"]); maxY = float(rect["maxY"])
    world_w = max(maxX - minX, 1.0); world_h = max(maxY - minY, 1.0)
    scale = min((tile_px - 32) / world_w, (tile_px - 32) / world_h)
    if scale <= 0:
        return None
    pad_x = (tile_px - world_w * scale) / 2
    pad_y = (tile_px - world_h * scale) / 2
    # 어두운 배경 (DXF AutoCAD 기본 배경과 비슷)
    image = _np.full((tile_px, tile_px, 3), 30, dtype=_np.uint8)

    def w2i(x, y):
        return (int(round(pad_x + (x - minX) * scale)), int(round(tile_px - pad_y - (y - minY) * scale)))

    def color_for(layer_name):
        """layer 이름에서 색상 추론 — sprinkler 도면 관례."""
        s = str(layer_name)
        if "헤드" in s or "HEAD" in s.upper() or "SP-H" in s.upper():
            return (0, 0, 255)  # red BGR
        if "하향식" in s or "1.68" in s or "1.93" in s:
            return (0, 0, 255)  # red - downward heads
        if "헤드반경" in s or "반경" in s:
            return (0, 255, 255)  # yellow
        if "알람" in s or "ALARM" in s.upper() or "ALV" in s.upper():
            return (0, 0, 255)
        return (200, 200, 200)  # gray for other

    for entity in entities:
        if visible_layers and entity.get("layer") not in visible_layers:
            continue
        bgr = color_for(entity.get("layer", ""))
        etype = entity.get("type")
        try:
            if etype == "LINE":
                p1 = w2i(entity["start"]["x"], entity["start"]["y"])
                p2 = w2i(entity["end"]["x"], entity["end"]["y"])
                cv2.line(image, p1, p2, bgr, 2, cv2.LINE_AA)
            elif etype == "LWPOLYLINE":
                pts = _np.array([w2i(p["x"], p["y"]) for p in entity.get("points", [])], dtype=_np.int32)
                if len(pts) >= 2:
                    cv2.polylines(image, [pts], bool(entity.get("closed")), bgr, 2, cv2.LINE_AA)
            elif etype == "CIRCLE":
                c = w2i(entity["center"]["x"], entity["center"]["y"])
                r = max(int(round(float(entity["radius"]) * scale)), 1)
                # 헤드 layer 의 원은 색칠. 일반 layer 의 원은 윤곽선만
                if "헤드" in str(entity.get("layer", "")) or "반경" in str(entity.get("layer", "")):
                    cv2.circle(image, c, r, bgr, -1)  # filled
                else:
                    cv2.circle(image, c, r, bgr, 2, cv2.LINE_AA)
            elif etype == "ARC":
                c = w2i(entity["center"]["x"], entity["center"]["y"])
                r = max(int(round(float(entity["radius"]) * scale)), 1)
                cv2.ellipse(image, c, (r, r), 0, -float(entity["endAngle"]), -float(entity["startAngle"]), bgr, 2, cv2.LINE_AA)
        except Exception:
            continue
    return image


def detect_heads_by_layer_insert(
    *,
    msp,
    settings,
    layer_match_fn,
) -> list[dict]:
    """가장 신뢰성 높은 헤드 검출:
    HEAD layer 의 INSERT 또는 CIRCLE entity 위치를 직접 사용.
    DXF 의 ground truth — 모델/매칭 없이도 100% 정확.

    이 함수는 layer 분류에 의존하므로 워크벤치에서 사용자가 카테고리 수정한 후 호출하면 정확함.
    """
    heads = []
    for e in msp:
        layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
        if layer_match_fn(layer, settings.exclude_layer_keywords):
            continue
        if not layer_match_fn(layer, settings.head_layer_keywords):
            continue
        etype = e.dxftype()
        if etype == "INSERT":
            heads.append({
                "x": float(e.dxf.insert.x), "y": float(e.dxf.insert.y),
                "source": f"INSERT[{e.dxf.name}]@{layer}",
            })
        elif etype == "CIRCLE":
            heads.append({
                "x": float(e.dxf.center.x), "y": float(e.dxf.center.y),
                "source": f"CIRCLE@{layer}",
            })
    return heads
