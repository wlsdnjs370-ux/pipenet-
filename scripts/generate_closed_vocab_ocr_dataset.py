from __future__ import annotations

import argparse
import json
import random
import shutil
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont


IMAGE_SIZE = (160, 56)


def build_vocabulary() -> list[str]:
    diameters = ["25A", "32A", "40A", "50A", "65A", "80A", "100A", "125A", "150A", "200A"]
    floors = [f"B{i}" for i in range(1, 7)] + [f"{i}F" for i in range(1, 41)]
    tags = ["HSP", "MSP", "LSP", "LLSP", "AV", "PV", "PRV", "FX", "ESFR", "RTI", "QR"]
    return diameters + floors + tags


def load_fonts() -> list[Path]:
    candidates = [
        Path(r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\arialbd.ttf"),
        Path(r"C:\Windows\Fonts\bahnschrift.ttf"),
        Path(r"C:\Windows\Fonts\calibri.ttf"),
        Path(r"C:\Windows\Fonts\calibrib.ttf"),
        Path(r"C:\Windows\Fonts\malgun.ttf"),
        Path(r"C:\Windows\Fonts\gulim.ttc"),
        Path(r"C:\Windows\Fonts\HANDotum.ttf"),
    ]
    return [p for p in candidates if p.exists()]


def render_token(token: str, font_paths: list[Path], rng: random.Random) -> Image.Image:
    bg = 255 if rng.random() < 0.75 else rng.randint(215, 245)
    fg = rng.randint(0, 35)
    image = Image.new("L", IMAGE_SIZE, color=bg)
    draw = ImageDraw.Draw(image)

    font_path = rng.choice(font_paths)
    font_size = rng.randint(24, 38)
    font = ImageFont.truetype(str(font_path), font_size)
    bbox = draw.textbbox((0, 0), token, font=font, stroke_width=rng.randint(0, 1))
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = rng.randint(6, max(6, IMAGE_SIZE[0] - text_w - 6))
    y = rng.randint(4, max(4, IMAGE_SIZE[1] - text_h - 4))

    draw.text(
        (x, y),
        token,
        font=font,
        fill=fg,
        stroke_width=rng.randint(0, 1),
        stroke_fill=max(0, fg - 10),
    )

    if rng.random() < 0.35:
        for _ in range(rng.randint(1, 3)):
            x1 = rng.randint(0, IMAGE_SIZE[0] - 1)
            y1 = rng.randint(0, IMAGE_SIZE[1] - 1)
            x2 = rng.randint(0, IMAGE_SIZE[0] - 1)
            y2 = rng.randint(0, IMAGE_SIZE[1] - 1)
            draw.line((x1, y1, x2, y2), fill=rng.randint(160, 220), width=1)

    if rng.random() < 0.25:
        image = image.filter(ImageFilter.GaussianBlur(radius=rng.uniform(0.2, 0.8)))

    arr = np.array(image, dtype=np.uint8)
    if rng.random() < 0.6:
        noise = rng.normalvariate(0, 6)
        arr = np.clip(arr.astype(np.float32) + noise, 0, 255).astype(np.uint8)

    speckle = rng.randint(10, 45)
    ys = np.random.randint(0, IMAGE_SIZE[1], size=speckle)
    xs = np.random.randint(0, IMAGE_SIZE[0], size=speckle)
    arr[ys, xs] = np.random.randint(0, 255, size=speckle)

    image = Image.fromarray(arr, mode="L")
    angle = rng.uniform(-4.0, 4.0)
    image = image.rotate(angle, resample=Image.Resampling.BILINEAR, expand=False, fillcolor=bg)

    if rng.random() < 0.2:
        image = Image.fromarray(255 - np.array(image), mode="L")
    return image


def generate_dataset(output_dir: Path, train_per_class: int, val_per_class: int, seed: int) -> dict:
    rng = random.Random(seed)
    np.random.seed(seed)
    fonts = load_fonts()
    if not fonts:
        raise RuntimeError("No fonts found for synthetic OCR dataset generation.")

    vocab = build_vocabulary()
    if output_dir.exists():
        shutil.rmtree(output_dir)
    for split in ("train", "val"):
        for token in vocab:
            (output_dir / split / token).mkdir(parents=True, exist_ok=True)

    for token in vocab:
        for idx in range(train_per_class):
            image = render_token(token, fonts, rng)
            image.save(output_dir / "train" / token / f"{idx:04d}.png")
        for idx in range(val_per_class):
            image = render_token(token, fonts, rng)
            image.save(output_dir / "val" / token / f"{idx:04d}.png")

    meta = {
        "vocabulary": vocab,
        "train_per_class": train_per_class,
        "val_per_class": val_per_class,
        "image_size": list(IMAGE_SIZE),
        "font_count": len(fonts),
        "fonts": [str(p) for p in fonts],
    }
    (output_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return meta


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=Path("data/closed_vocab_ocr_dataset"))
    parser.add_argument("--train-per-class", type=int, default=180)
    parser.add_argument("--val-per-class", type=int, default=40)
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    meta = generate_dataset(args.output, args.train_per_class, args.val_per_class, args.seed)
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
