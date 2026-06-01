from __future__ import annotations

import json
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

import torch
from PIL import Image
from torch import nn
from torchvision import transforms
from torchvision.models import resnet18


class ClosedVocabOCRNet(nn.Module):
    def __init__(self, num_classes: int) -> None:
        super().__init__()
        self.backbone = resnet18(weights=None)
        self.backbone.conv1 = nn.Conv2d(1, 64, kernel_size=7, stride=2, padding=3, bias=False)
        self.backbone.fc = nn.Linear(self.backbone.fc.in_features, num_classes)

    def forward(self, x: torch.Tensor) -> torch.Tensor:
        return self.backbone(x)


@dataclass(slots=True)
class OCRPrediction:
    token: str
    confidence: float
    scores: dict[str, float]


class ClosedVocabOCR:
    def __init__(self, *, base_dir: Path) -> None:
        self.base_dir = base_dir
        self.model_dir = base_dir / "models" / "closed_vocab_ocr"
        self.vocab_path = self.model_dir / "vocab.json"
        self.weights_path = self.model_dir / "model.pt"
        self.metrics_path = self.model_dir / "metrics.json"
        self.device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        self.transform = transforms.Compose(
            [
                transforms.Grayscale(num_output_channels=1),
                transforms.Resize((56, 160)),
                transforms.ToTensor(),
                transforms.Normalize((0.5,), (0.5,)),
            ]
        )
        self.vocab = json.loads(self.vocab_path.read_text(encoding="utf-8"))
        self.index_to_token = {int(k): v for k, v in self.vocab["index_to_token"].items()}
        self.token_to_index = {k: int(v) for k, v in self.vocab["token_to_index"].items()}
        self.model = ClosedVocabOCRNet(len(self.token_to_index))
        state = torch.load(self.weights_path, map_location=self.device)
        self.model.load_state_dict(state)
        self.model.to(self.device)
        self.model.eval()

    def predict(self, image_path: Path, candidates: list[str] | None = None) -> OCRPrediction:
        image = Image.open(image_path).convert("L")
        x = self.transform(image).unsqueeze(0).to(self.device)
        with torch.no_grad():
            logits = self.model(x)[0]
            probs = torch.softmax(logits, dim=0).detach().cpu()

        if candidates:
            valid = [c for c in candidates if c in self.token_to_index]
            if valid:
                candidate_scores = {c: float(probs[self.token_to_index[c]]) for c in valid}
                token = max(candidate_scores, key=candidate_scores.get)
                confidence = candidate_scores[token]
                return OCRPrediction(token=token, confidence=confidence, scores=candidate_scores)

        full_scores = {self.index_to_token[i]: float(probs[i]) for i in range(len(probs))}
        best_idx = int(torch.argmax(probs).item())
        return OCRPrediction(
            token=self.index_to_token[best_idx],
            confidence=float(probs[best_idx]),
            scores=full_scores,
        )


@lru_cache(maxsize=1)
def get_cached_closed_vocab_ocr(base_dir: str) -> ClosedVocabOCR:
    return ClosedVocabOCR(base_dir=Path(base_dir))
