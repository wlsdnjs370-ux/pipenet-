"""Sprinkler 심볼 다중 클래스 YOLO 학습 (헤드 3종 + 알람밸브).

데이터셋: data/sprinkler_yolo_dataset/dataset.yaml
출력: models/sprinkler_yolo/weights/best.pt
"""

from __future__ import annotations

import argparse
from pathlib import Path

from ultralytics import YOLO

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATASET_YAML = PROJECT_ROOT / "data" / "sprinkler_yolo_dataset" / "dataset.yaml"
OUTPUT_DIR = PROJECT_ROOT / "models" / "sprinkler_yolo"


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--base", default="yolo11n.pt", help="베이스 모델 가중치 (yolo11n.pt 등)")
    p.add_argument("--epochs", type=int, default=30)
    p.add_argument("--imgsz", type=int, default=640)
    p.add_argument("--batch", type=int, default=8)
    p.add_argument("--device", default="cpu")
    p.add_argument("--patience", type=int, default=10)
    args = p.parse_args()

    if not DATASET_YAML.exists():
        raise SystemExit(f"dataset.yaml 없음 — 먼저 generate_sprinkler_yolo_dataset.py 실행: {DATASET_YAML}")

    print(f"[train] dataset: {DATASET_YAML}")
    print(f"[train] base: {args.base}")
    print(f"[train] epochs={args.epochs}  imgsz={args.imgsz}  batch={args.batch}  device={args.device}")

    model = YOLO(args.base)
    results = model.train(
        data=str(DATASET_YAML),
        epochs=args.epochs,
        imgsz=args.imgsz,
        batch=args.batch,
        device=args.device,
        patience=args.patience,
        project=str(OUTPUT_DIR.parent),
        name=OUTPUT_DIR.name,
        exist_ok=True,
        verbose=True,
        plots=False,
        save=True,
        amp=False,
    )
    # 학습 끝나면 best.pt 위치 출력
    best_path = OUTPUT_DIR / "weights" / "best.pt"
    if best_path.exists():
        print(f"[train] DONE. best weights: {best_path}")
    else:
        print(f"[train] DONE but best.pt not found at {best_path}")


if __name__ == "__main__":
    main()
