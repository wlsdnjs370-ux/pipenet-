from __future__ import annotations

import argparse
import random
import shutil
from pathlib import Path

import cv2
import numpy as np
import yaml


def load_templates(template_dir: Path) -> list[np.ndarray]:
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
        templates.append(cropped)
    if not templates:
        raise RuntimeError(f"No templates found in {template_dir}")
    return templates


def random_background(size: int) -> np.ndarray:
    canvas = np.zeros((size, size, 3), dtype=np.uint8)
    for _ in range(random.randint(20, 60)):
        color = random.randint(30, 90)
        x1, y1 = random.randint(0, size - 1), random.randint(0, size - 1)
        x2, y2 = random.randint(0, size - 1), random.randint(0, size - 1)
        cv2.line(canvas, (x1, y1), (x2, y2), (color, color, color), random.randint(1, 2), cv2.LINE_AA)
    for _ in range(random.randint(4, 10)):
        color = random.randint(30, 80)
        x1, y1 = random.randint(0, size - 80), random.randint(0, size - 80)
        x2, y2 = random.randint(x1 + 20, min(size - 1, x1 + 220)), random.randint(y1 + 20, min(size - 1, y1 + 220))
        cv2.rectangle(canvas, (x1, y1), (x2, y2), (color, color, color), 1, cv2.LINE_AA)
    for _ in range(random.randint(4, 12)):
        color = random.randint(35, 85)
        center = (random.randint(0, size - 1), random.randint(0, size - 1))
        radius = random.randint(8, 40)
        cv2.circle(canvas, center, radius, (color, color, color), 1, cv2.LINE_AA)
    noise = np.random.normal(0, 6, (size, size, 3)).astype(np.int16)
    return np.clip(canvas.astype(np.int16) + noise, 0, 255).astype(np.uint8)


def augment_template(template: np.ndarray) -> np.ndarray:
    scale = random.uniform(0.75, 1.6)
    angle = random.uniform(0, 360)
    thickness = random.randint(1, 3)
    resized = cv2.resize(template, None, fx=scale, fy=scale, interpolation=cv2.INTER_LINEAR)
    h, w = resized.shape[:2]
    center = (w / 2, h / 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    cos = abs(matrix[0, 0])
    sin = abs(matrix[0, 1])
    new_w = int((h * sin) + (w * cos))
    new_h = int((h * cos) + (w * sin))
    matrix[0, 2] += new_w / 2 - center[0]
    matrix[1, 2] += new_h / 2 - center[1]
    rotated = cv2.warpAffine(resized, matrix, (new_w, new_h), flags=cv2.INTER_LINEAR, borderValue=0)
    _, rotated = cv2.threshold(rotated, 16, 255, cv2.THRESH_BINARY)
    if thickness > 1:
      kernel = np.ones((thickness, thickness), dtype=np.uint8)
      rotated = cv2.dilate(rotated, kernel, iterations=1)
    return rotated


def place_symbol(background: np.ndarray, symbol: np.ndarray, occupied: list[tuple[int, int, int, int]]) -> tuple[tuple[int, int, int, int], bool]:
    h, w = symbol.shape[:2]
    size = background.shape[0]
    if w >= size or h >= size:
        return (0, 0, 0, 0), False
    for _ in range(30):
        x = random.randint(4, size - w - 4)
        y = random.randint(4, size - h - 4)
        box = (x, y, x + w, y + h)
        if any(not (box[2] < other[0] or other[2] < box[0] or box[3] < other[1] or other[3] < box[1]) for other in occupied):
            continue
        color = random.randint(220, 255)
        mask = symbol > 0
        background[y:y + h, x:x + w][mask] = (color, color, color)
        stem_len = random.randint(12, 60)
        stem_dir = random.choice([(0, 1), (0, -1), (1, 0), (-1, 0)])
        cx = x + w // 2
        cy = y + h // 2
        sx = max(0, min(size - 1, cx + stem_dir[0] * stem_len))
        sy = max(0, min(size - 1, cy + stem_dir[1] * stem_len))
        cv2.line(background, (cx, cy), (sx, sy), (255, 255, 255), random.randint(1, 2), cv2.LINE_AA)
        occupied.append(box)
        return box, True
    return (0, 0, 0, 0), False


def save_sample(image_path: Path, label_path: Path, templates: list[np.ndarray], image_size: int) -> None:
    background = random_background(image_size)
    occupied: list[tuple[int, int, int, int]] = []
    labels: list[str] = []
    for _ in range(random.randint(1, 5)):
        template = random.choice(templates)
        symbol = augment_template(template)
        box, ok = place_symbol(background, symbol, occupied)
        if not ok:
            continue
        x1, y1, x2, y2 = box
        cx = ((x1 + x2) / 2) / image_size
        cy = ((y1 + y2) / 2) / image_size
        bw = (x2 - x1) / image_size
        bh = (y2 - y1) / image_size
        labels.append(f"0 {cx:.6f} {cy:.6f} {bw:.6f} {bh:.6f}")
    cv2.imwrite(str(image_path), background)
    label_path.write_text("\n".join(labels), encoding="utf-8")


def build_dataset(template_dir: Path, output_dir: Path, train_count: int, val_count: int, image_size: int) -> None:
    templates = load_templates(template_dir)
    if output_dir.exists():
        shutil.rmtree(output_dir)
    for split in ("train", "val"):
        (output_dir / "images" / split).mkdir(parents=True, exist_ok=True)
        (output_dir / "labels" / split).mkdir(parents=True, exist_ok=True)

    for idx in range(train_count):
        save_sample(
            output_dir / "images" / "train" / f"train_{idx:04d}.png",
            output_dir / "labels" / "train" / f"train_{idx:04d}.txt",
            templates,
            image_size,
        )
    for idx in range(val_count):
        save_sample(
            output_dir / "images" / "val" / f"val_{idx:04d}.png",
            output_dir / "labels" / "val" / f"val_{idx:04d}.txt",
            templates,
            image_size,
        )

    data_yaml = {
        "path": str(output_dir.resolve()),
        "train": "images/train",
        "val": "images/val",
        "names": {0: "triangle_head"},
    }
    with (output_dir / "dataset.yaml").open("w", encoding="utf-8") as handle:
        yaml.safe_dump(data_yaml, handle, sort_keys=False, allow_unicode=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template-dir", type=Path, default=Path("data/head_templates"))
    parser.add_argument("--output-dir", type=Path, default=Path("data/triangle_head_dataset"))
    parser.add_argument("--train-count", type=int, default=700)
    parser.add_argument("--val-count", type=int, default=140)
    parser.add_argument("--image-size", type=int, default=640)
    args = parser.parse_args()
    random.seed(42)
    np.random.seed(42)
    build_dataset(args.template_dir, args.output_dir, args.train_count, args.val_count, args.image_size)
    print(f"Dataset created at {args.output_dir}")


if __name__ == "__main__":
    main()
