from __future__ import annotations

import argparse
from pathlib import Path

import torch
from ultralytics import YOLO


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", type=Path, default=Path("data/triangle_head_dataset/dataset.yaml"))
    parser.add_argument("--project", type=Path, default=Path("models"))
    parser.add_argument("--name", type=str, default="triangle_head_yolo")
    parser.add_argument("--model", type=str, default="yolo11n.pt")
    parser.add_argument("--epochs", type=int, default=80)
    parser.add_argument("--imgsz", type=int, default=640)
    parser.add_argument("--batch", type=int, default=16)
    args = parser.parse_args()

    if not torch.cuda.is_available():
        raise RuntimeError("CUDA GPU is required for this training script.")

    model = YOLO(args.model)
    model.train(
        data=str(args.data),
        project=str(args.project),
        name=args.name,
        device=0,
        epochs=args.epochs,
        imgsz=args.imgsz,
        batch=args.batch,
        workers=4,
        pretrained=True,
        close_mosaic=10,
        degrees=15.0,
        scale=0.3,
        translate=0.08,
        fliplr=0.5,
        mosaic=0.4,
        mixup=0.0,
        hsv_h=0.0,
        hsv_s=0.0,
        hsv_v=0.0,
        patience=20,
        single_cls=True,
    )
    print(f"Best model: {args.project / args.name / 'weights' / 'best.pt'}")


if __name__ == "__main__":
    main()
