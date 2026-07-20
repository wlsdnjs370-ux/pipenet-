"""Sprinkler 헤드/알람밸브 YOLO 학습용 합성 데이터셋 생성.

목적: 5개 심볼 (헤드 4종 + 알람밸브) 을 다양한 회전/스케일/위치로 합성
→ layer 분류와 human 검토가 놓친 객체도 잡을 수 있는 일반화 모델 학습

출력 (YOLO 표준 구조):
    data/sprinkler_yolo_dataset/
        images/train/img_0001.jpg
        images/val/img_0001.jpg
        labels/train/img_0001.txt   (cx, cy, w, h - normalized)
        labels/val/img_0001.txt
        dataset.yaml

클래스 (다중 클래스):
    0 - head_yellow_circle  (헤드1/4: 노란 원 + 가로선, 측면 헤드)
    1 - head_red_triangle   (헤드2: 빨간 ▽, 하향식 헤드)
    2 - head_red_dot        (헤드3: 빨간 작은 원, 일반 헤드)
    3 - alarm_valve         (알람밸브: ▶ + 원 + 가로선)
"""

from __future__ import annotations

import argparse
import random
from pathlib import Path

import cv2
import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parent.parent
HEAD_TEMPLATES = PROJECT_ROOT / "data" / "head_templates"
ALARM_TEMPLATES = PROJECT_ROOT / "data" / "alarm_templates"
DATASET_DIR = PROJECT_ROOT / "data" / "sprinkler_yolo_dataset"

# 심볼 → 클래스 매핑
SYMBOLS = [
    {"file": HEAD_TEMPLATES / "head_yellow_circle_side.png",  "cls": 0, "size_range": (20, 60)},
    {"file": HEAD_TEMPLATES / "head_yellow_circle_side2.png", "cls": 0, "size_range": (20, 60)},
    {"file": HEAD_TEMPLATES / "head_red_triangle_down.png",   "cls": 1, "size_range": (20, 70)},
    {"file": HEAD_TEMPLATES / "head_red_dot.png",             "cls": 2, "size_range": (15, 50)},
    {"file": ALARM_TEMPLATES / "alarm_valve_triangle.png",    "cls": 3, "size_range": (80, 200)},
]
CLASS_NAMES = ["head_yellow_circle", "head_red_triangle", "head_red_dot", "alarm_valve"]


def load_template(path: Path) -> np.ndarray:
    """PNG → BGRA. 배경(어두운 회색)을 알파로 처리."""
    img = cv2.imread(str(path), cv2.IMREAD_UNCHANGED)
    if img is None:
        raise FileNotFoundError(path)
    if img.shape[2] == 3:
        # BGR → BGRA, 배경 회색을 투명으로
        bgr = img
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        # 배경(어두운 회색 ~20-40) 을 투명, 나머지를 불투명
        alpha = np.where(gray > 50, 255, 0).astype(np.uint8)
        # 모폴로지로 알파 정리
        kernel = np.ones((2, 2), np.uint8)
        alpha = cv2.morphologyEx(alpha, cv2.MORPH_CLOSE, kernel)
        img = cv2.merge([bgr[:, :, 0], bgr[:, :, 1], bgr[:, :, 2], alpha])
    return img


def random_rotate_scale(img: np.ndarray, target_size: int, rotation: float) -> np.ndarray:
    """심볼을 target_size 픽셀에 맞춰 리사이즈 + 임의 회전."""
    h, w = img.shape[:2]
    scale = target_size / max(h, w)
    new_w, new_h = int(w * scale), int(h * scale)
    if new_w < 4 or new_h < 4:
        return img
    img = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_AREA)
    # 회전
    center = (new_w // 2, new_h // 2)
    M = cv2.getRotationMatrix2D(center, rotation, 1.0)
    # 회전 후 bbox 크기 계산
    cos_v = abs(M[0, 0]); sin_v = abs(M[0, 1])
    bound_w = int(new_h * sin_v + new_w * cos_v)
    bound_h = int(new_h * cos_v + new_w * sin_v)
    M[0, 2] += (bound_w - new_w) / 2
    M[1, 2] += (bound_h - new_h) / 2
    rotated = cv2.warpAffine(img, M, (bound_w, bound_h),
                              flags=cv2.INTER_LINEAR,
                              borderMode=cv2.BORDER_CONSTANT,
                              borderValue=(0, 0, 0, 0))
    return rotated


def paste_with_alpha(background: np.ndarray, foreground: np.ndarray, x: int, y: int) -> None:
    """foreground(BGRA) 를 background(BGR) 의 (x, y) 위치에 알파 합성. in-place."""
    fh, fw = foreground.shape[:2]
    bh, bw = background.shape[:2]
    x0 = max(0, x); y0 = max(0, y)
    x1 = min(bw, x + fw); y1 = min(bh, y + fh)
    if x0 >= x1 or y0 >= y1:
        return
    fg_x0 = x0 - x; fg_y0 = y0 - y
    fg_x1 = fg_x0 + (x1 - x0); fg_y1 = fg_y0 + (y1 - y0)
    fg = foreground[fg_y0:fg_y1, fg_x0:fg_x1]
    bg_roi = background[y0:y1, x0:x1]
    if fg.shape[2] == 4:
        alpha = fg[:, :, 3:4].astype(np.float32) / 255.0
        bg_roi[:] = (fg[:, :, :3].astype(np.float32) * alpha + bg_roi.astype(np.float32) * (1 - alpha)).astype(np.uint8)
    else:
        bg_roi[:] = fg


def random_background(width: int, height: int, rng: random.Random) -> np.ndarray:
    """sprinkler 도면 분위기의 합성 배경 생성.
    어두운 회색 + 무작위 흰/회색 선들 (벽/배관 흉내).
    """
    bg = np.full((height, width, 3), rng.randint(20, 40), dtype=np.uint8)
    # 회색 선 (벽)
    n_lines = rng.randint(8, 25)
    for _ in range(n_lines):
        x1, y1 = rng.randint(0, width), rng.randint(0, height)
        x2, y2 = rng.randint(0, width), rng.randint(0, height)
        color = rng.randint(100, 200)
        cv2.line(bg, (x1, y1), (x2, y2), (color, color, color), rng.randint(1, 2))
    # 작은 잡음 사각 (텍스트 흉내)
    for _ in range(rng.randint(3, 8)):
        x, y = rng.randint(20, width-40), rng.randint(20, height-20)
        w = rng.randint(10, 30); h = rng.randint(6, 12)
        cv2.rectangle(bg, (x, y), (x + w, y + h), (180, 180, 180), 1)
    return bg


def generate_one(rng: random.Random, templates: list[dict], img_size: int = 640) -> tuple[np.ndarray, list[tuple]]:
    """이미지 1장 + YOLO 라벨 리스트 [(cls, cx, cy, w, h), ...] (normalized).
    """
    bg = random_background(img_size, img_size, rng)
    labels = []
    # 알람밸브 0-1개, 헤드 5-15개
    n_alarm = rng.choices([0, 1], weights=[3, 1])[0]
    n_heads = rng.randint(5, 18)

    placed_boxes = []

    def try_place(sym):
        for _ in range(10):  # max 10 try
            tgt = rng.randint(*sym["size_range"])
            rot = rng.uniform(-30, 30) if sym["cls"] != 1 else rng.uniform(-180, 180)
            tpl = load_template(sym["file"])
            rotated = random_rotate_scale(tpl, tgt, rot)
            rh, rw = rotated.shape[:2]
            if rw >= img_size - 20 or rh >= img_size - 20:
                continue
            x = rng.randint(10, img_size - rw - 10)
            y = rng.randint(10, img_size - rh - 10)
            box = (x, y, x + rw, y + rh)
            # overlap check (간단)
            ok = True
            for pb in placed_boxes:
                if not (box[2] < pb[0] or pb[2] < box[0] or box[3] < pb[1] or pb[3] < box[1]):
                    ok = False; break
            if not ok:
                continue
            paste_with_alpha(bg, rotated, x, y)
            placed_boxes.append(box)
            cx = (box[0] + box[2]) / 2 / img_size
            cy = (box[1] + box[3]) / 2 / img_size
            w = (box[2] - box[0]) / img_size
            h = (box[3] - box[1]) / img_size
            labels.append((sym["cls"], cx, cy, w, h))
            return True
        return False

    head_syms = [s for s in templates if s["cls"] != 3]
    alarm_syms = [s for s in templates if s["cls"] == 3]
    for _ in range(n_alarm):
        try_place(rng.choice(alarm_syms))
    for _ in range(n_heads):
        try_place(rng.choice(head_syms))

    return bg, labels


def write_yolo_label(path: Path, labels: list[tuple]) -> None:
    lines = []
    for cls, cx, cy, w, h in labels:
        lines.append(f"{cls} {cx:.6f} {cy:.6f} {w:.6f} {h:.6f}")
    path.write_text("\n".join(lines), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--n_train", type=int, default=400)
    parser.add_argument("--n_val", type=int, default=80)
    parser.add_argument("--img_size", type=int, default=640)
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    rng = random.Random(args.seed)
    # 출력 폴더 정리
    for sub in ["images/train", "images/val", "labels/train", "labels/val"]:
        (DATASET_DIR / sub).mkdir(parents=True, exist_ok=True)

    # 템플릿 로드 (사전 검증)
    for s in SYMBOLS:
        if not s["file"].exists():
            raise FileNotFoundError(f"심볼 누락: {s['file']}")

    print(f"Generating {args.n_train} train + {args.n_val} val images...")
    for split, n in [("train", args.n_train), ("val", args.n_val)]:
        for i in range(n):
            img, labels = generate_one(rng, SYMBOLS, args.img_size)
            img_path = DATASET_DIR / f"images/{split}/img_{i:04d}.jpg"
            lbl_path = DATASET_DIR / f"labels/{split}/img_{i:04d}.txt"
            cv2.imwrite(str(img_path), img, [cv2.IMWRITE_JPEG_QUALITY, 90])
            write_yolo_label(lbl_path, labels)
        print(f"  {split}: {n} images written")

    # dataset.yaml
    yaml_path = DATASET_DIR / "dataset.yaml"
    yaml_path.write_text(
        f"path: {DATASET_DIR.as_posix()}\n"
        f"train: images/train\n"
        f"val: images/val\n"
        f"nc: {len(CLASS_NAMES)}\n"
        f"names: {CLASS_NAMES}\n",
        encoding="utf-8",
    )
    print(f"  dataset.yaml -> {yaml_path}")
    print("Done.")


if __name__ == "__main__":
    main()
