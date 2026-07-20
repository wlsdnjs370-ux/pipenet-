"""헤드 노드 이진 분류 — MLP / GraphSAGE / GAT 비교 학습.

문제 정의
==========
입력: DXF → 배관망 그래프 G = (V, E)
  V = snap-grouped 파이프 끝점/교차점/헤드 위치
  E = real LINE edges + 컴포넌트 가상 다리 + 헤드 drop line

학습 목표
----------
각 노드 v ∈ V 가 "헤드 위치" 인지 이진 분류.
  y = 1 if v ∈ detect_heads() 출력의 클러스터 중심,  else 0

노드 feature x_v ∈ ℝ^F (F=13)
---------------------------------
0  degree (정규화: tanh(deg/5))
1  is_alarm_layer    (인근 50mm 내 '-소화(배관-SP 2차)' LINE endpoint)
2  is_secondary_pipe (인근 200mm 내 '-소화(SP가지관)' LINE endpoint)
3  is_flexible       (인근 200mm 내 'SP 후렉시블' LWPOLYLINE endpoint)
4  is_head_layer     (인근 100mm 내 HEAD 카테고리 entity)
5  has_text_nearby   (인근 300mm 내 TEX 레이어 TEXT)
6  has_circle_nearby (인근 200mm 내 작은 CIRCLE r<200)
7  has_hatch_nearby  (인근 200mm 내 HATCH)
8  has_insert_nearby (인근 100mm 내 INSERT)
9  x_norm (전체 bbox 기준 0~1 정규화)
10 y_norm (전체 bbox 기준 0~1 정규화)
11 dist_to_centroid (모든 노드 중심으로부터 정규화 거리)
12 dist_to_alarm    (자동 식별된 알람밸브로부터 정규화 거리)

모델 비교
----------
1. MLP        — feature 만 사용, 그래프 구조 무시. baseline.
2. GraphSAGE  — 메시지 패싱 (이웃 평균 + 자기 변환). 2 layer.
3. GAT        — 어텐션 기반 (다중 head). 2 layer.

평가
-----
5-fold stratified CV. 매 fold:
  학습 80%, 검증 20% (positive 라벨 비율 균등)
  손실: BCEWithLogitsLoss(class weight = neg/pos)
  최적화: AdamW lr=1e-3, weight_decay=5e-4
  epoch=200, early stopping=30 patience on val F1

지표
-----
- Accuracy / Precision / Recall / F1 (positive class)
- AUROC, AUPRC
- Confusion matrix
fold 별 + 전체 평균 ± std 출력. JSON + Markdown 결과 저장.
"""

from __future__ import annotations

import io
import json
import math
import random
import sys
import time
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import torch
import torch.nn as nn
import torch.nn.functional as F
from sklearn.metrics import (
    accuracy_score,
    average_precision_score,
    confusion_matrix,
    f1_score,
    precision_score,
    recall_score,
    roc_auc_score,
)
from sklearn.model_selection import StratifiedKFold
from torch_geometric.data import Data
from torch_geometric.nn import GATConv, SAGEConv

# UTF-8 콘솔
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

# 프로젝트 root + remote30_prototype 사용
PROJECT = Path(r"C:\Users\admin\PycharmProjects\JupyterProject")
sys.path.insert(0, str(PROJECT))
import remote30_prototype as rp  # noqa: E402

SEED = 42
random.seed(SEED); np.random.seed(SEED); torch.manual_seed(SEED)
DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")
print(f"[env] device={DEVICE}  torch={torch.__version__}")


# ────────────────────────────────────────────────────────────────────────────
# 1) 데이터셋 빌드 — DXF → 그래프 + node feature + label
# ────────────────────────────────────────────────────────────────────────────


@dataclass
class GraphDataset:
    node_pos: list[tuple[float, float]]
    node_idx: dict[tuple[float, float], int]
    edges: list[tuple[int, int]]
    node_features: np.ndarray  # (N, F)
    labels: np.ndarray         # (N,) binary 0/1
    head_positions: set[tuple[float, float]]
    src_pos: tuple[float, float]
    bbox: tuple[float, float, float, float]


def _nearby_entities(pipe_ents: list[dict], cx: float, cy: float, radius: float,
                     layer_filter: str | None = None, type_filter: str | None = None,
                     layer_keyword: str | None = None,
                     extra_check=None) -> int:
    """주어진 위치 (cx, cy) 의 radius 내 entity 수. 필터: layer 정확 일치 또는 layer keyword."""
    cnt = 0
    for en in pipe_ents:
        if type_filter and en.get("t") != type_filter:
            continue
        if layer_filter and en.get("l") != layer_filter:
            continue
        if layer_keyword and layer_keyword not in en.get("l", ""):
            continue
        # 위치 추정 — t 별로 다름
        t = en.get("t")
        if t == "L":
            p = en["p"]
            xs = [p[0], p[2]]; ys = [p[1], p[3]]
            for ex, ey in zip(xs, ys):
                if math.hypot(ex - cx, ey - cy) <= radius:
                    cnt += 1
                    break
        elif t == "PL":
            for px, py in en.get("p", []):
                if math.hypot(px - cx, py - cy) <= radius:
                    cnt += 1
                    break
        elif t == "C":
            c = en["c"]
            r = en.get("r", 0)
            if extra_check and not extra_check(r):
                continue
            if math.hypot(c[0] - cx, c[1] - cy) <= radius:
                cnt += 1
        elif t == "I" or t == "T":
            p = en["p"]
            if math.hypot(p[0] - cx, p[1] - cy) <= radius:
                cnt += 1
        elif t == "H":
            for px, py in en.get("p", []):
                if math.hypot(px - cx, py - cy) <= radius:
                    cnt += 1
                    break
    return cnt


def build_dataset(dxf_path: Path) -> GraphDataset:
    print(f"[data] parsing {dxf_path.name} ...")
    bundle = rp.parse_dxf_bundle(dxf_path)
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = rp.filter_pipenet_only(bundle)

    # 그래프 빌드 (snap+bridge+head_drop — Stage 3 와 동일)
    graph, edge_len = rp._build_graph(pipe_ents)
    for tol in (200.0, 500.0, 1000.0, 2000.0):
        rp._bridge_components(graph, edge_len, max_bridge_mm=tol)

    # 헤드 위치
    head_detections = rp.detect_heads(pipe_ents, layer_cat)
    head_pos_snapped = {rp._round_pt(h.pos[0], h.pos[1]) for h in head_detections}
    for hp in head_pos_snapped:
        nearest = rp._nearest_graph_node(graph, hp)
        if nearest is None or hp == nearest:
            continue
        d = math.hypot(hp[0] - nearest[0], hp[1] - nearest[1])
        if d > 1e-3 and d <= rp.HEAD_BRIDGE_MAX_MM:
            graph.setdefault(hp, set()).add(nearest)
            graph[nearest].add(hp)
            edge_len[(min(hp, nearest), max(hp, nearest))] = d

    # 알람밸브 — feature 거리 계산용
    src_raw, _ = rp._find_source(pipe_ents, layer_cat)
    src_pos = rp._nearest_graph_node(graph, src_raw) if src_raw else None
    if src_pos is None and graph:
        src_pos = max(graph, key=lambda n: len(graph[n]))

    # 노드 인덱싱
    node_pos = sorted(graph.keys())
    node_idx = {n: i for i, n in enumerate(node_pos)}
    edges: list[tuple[int, int]] = []
    seen_edges = set()
    for u, nbrs in graph.items():
        for v in nbrs:
            i, j = node_idx[u], node_idx[v]
            key = (min(i, j), max(i, j))
            if key in seen_edges:
                continue
            seen_edges.add(key)
            edges.append((i, j))

    # bbox
    xs = [p[0] for p in node_pos]; ys = [p[1] for p in node_pos]
    bbox = (min(xs), min(ys), max(xs), max(ys))
    bw = max(bbox[2] - bbox[0], 1.0); bh = max(bbox[3] - bbox[1], 1.0)
    cx_g = (bbox[0] + bbox[2]) / 2; cy_g = (bbox[1] + bbox[3]) / 2
    max_diag = math.hypot(bw, bh)

    # feature matrix (N, F=13)
    F_dim = 13
    feat = np.zeros((len(node_pos), F_dim), dtype=np.float32)
    print(f"[data] extracting features for {len(node_pos)} nodes ...")
    for i, (nx, ny) in enumerate(node_pos):
        deg = len(graph[(nx, ny)])
        feat[i, 0] = math.tanh(deg / 5.0)
        feat[i, 1] = min(1.0, _nearby_entities(pipe_ents, nx, ny, 50, layer_keyword="배관-SP 2차") / 2.0)
        feat[i, 2] = min(1.0, _nearby_entities(pipe_ents, nx, ny, 200, layer_keyword="-소화(SP가지관)") / 5.0)
        feat[i, 3] = min(1.0, _nearby_entities(pipe_ents, nx, ny, 200, layer_filter="SP 후렉시블") / 3.0)
        # is_head_layer — HEAD 카테고리 layer 의 entity 가 100mm 내 있나
        cnt_head = 0
        for en in pipe_ents:
            cat = layer_cat.get(en.get("l", ""), "OTHER")
            if cat != "HEAD":
                continue
            t = en.get("t")
            if t == "I":
                p = en["p"]
                if math.hypot(p[0] - nx, p[1] - ny) <= 100:
                    cnt_head += 1
            elif t == "C":
                c = en["c"]
                if math.hypot(c[0] - nx, c[1] - ny) <= 100:
                    cnt_head += 1
        feat[i, 4] = min(1.0, cnt_head / 3.0)
        feat[i, 5] = min(1.0, _nearby_entities(pipe_ents, nx, ny, 300, layer_filter="TEX", type_filter="T") / 3.0)
        feat[i, 6] = min(1.0, _nearby_entities(pipe_ents, nx, ny, 200, type_filter="C",
                                                extra_check=lambda r: r < 200) / 3.0)
        feat[i, 7] = min(1.0, _nearby_entities(pipe_ents, nx, ny, 200, type_filter="H") / 2.0)
        feat[i, 8] = min(1.0, _nearby_entities(pipe_ents, nx, ny, 100, type_filter="I") / 2.0)
        feat[i, 9] = (nx - bbox[0]) / bw
        feat[i, 10] = (ny - bbox[1]) / bh
        feat[i, 11] = math.hypot(nx - cx_g, ny - cy_g) / max_diag
        if src_pos:
            feat[i, 12] = math.hypot(nx - src_pos[0], ny - src_pos[1]) / max_diag

    # 라벨 — 헤드 위치와 일치하는 노드 = 1
    labels = np.zeros(len(node_pos), dtype=np.int64)
    for i, pos in enumerate(node_pos):
        if pos in head_pos_snapped:
            labels[i] = 1

    pos_count = int(labels.sum())
    print(f"[data] nodes={len(node_pos)}  edges={len(edges)}  positive(head)={pos_count}  "
          f"negative={len(node_pos) - pos_count}  positive_ratio={pos_count/len(node_pos)*100:.1f}%")
    return GraphDataset(
        node_pos=node_pos, node_idx=node_idx, edges=edges,
        node_features=feat, labels=labels,
        head_positions=head_pos_snapped, src_pos=src_pos or (0, 0), bbox=bbox,
    )


# ────────────────────────────────────────────────────────────────────────────
# 2) 모델 정의
# ────────────────────────────────────────────────────────────────────────────


class MLPClassifier(nn.Module):
    """그래프 무시, 노드 feature 만 사용 — baseline."""

    def __init__(self, in_dim: int, hidden: int = 64):
        super().__init__()
        self.net = nn.Sequential(
            nn.Linear(in_dim, hidden), nn.ReLU(), nn.Dropout(0.3),
            nn.Linear(hidden, hidden), nn.ReLU(), nn.Dropout(0.3),
            nn.Linear(hidden, 1),
        )

    def forward(self, x, edge_index=None):
        return self.net(x).squeeze(-1)


class GraphSAGENet(nn.Module):
    """SAGEConv 2-layer."""

    def __init__(self, in_dim: int, hidden: int = 64):
        super().__init__()
        self.conv1 = SAGEConv(in_dim, hidden)
        self.conv2 = SAGEConv(hidden, hidden)
        self.head = nn.Sequential(nn.Linear(hidden, hidden // 2), nn.ReLU(), nn.Dropout(0.3),
                                  nn.Linear(hidden // 2, 1))

    def forward(self, x, edge_index):
        h = F.relu(self.conv1(x, edge_index)); h = F.dropout(h, p=0.3, training=self.training)
        h = F.relu(self.conv2(h, edge_index))
        return self.head(h).squeeze(-1)


class GATNet(nn.Module):
    """GATConv 2-layer 4 heads."""

    def __init__(self, in_dim: int, hidden: int = 32, heads: int = 4):
        super().__init__()
        self.conv1 = GATConv(in_dim, hidden, heads=heads, dropout=0.2)
        self.conv2 = GATConv(hidden * heads, hidden, heads=1, concat=False, dropout=0.2)
        self.head = nn.Linear(hidden, 1)

    def forward(self, x, edge_index):
        h = F.elu(self.conv1(x, edge_index)); h = F.dropout(h, p=0.3, training=self.training)
        h = F.elu(self.conv2(h, edge_index))
        return self.head(h).squeeze(-1)


# ────────────────────────────────────────────────────────────────────────────
# 3) 학습 + 평가
# ────────────────────────────────────────────────────────────────────────────


def train_one_fold(model: nn.Module, data: Data, train_mask: torch.Tensor, val_mask: torch.Tensor,
                   class_weight: float, epochs: int = 200, patience: int = 30) -> dict:
    model = model.to(DEVICE)
    data = data.to(DEVICE)
    train_mask = train_mask.to(DEVICE); val_mask = val_mask.to(DEVICE)

    optimizer = torch.optim.AdamW(model.parameters(), lr=1e-3, weight_decay=5e-4)
    # class weight in BCE — positive weight = neg/pos
    pos_weight = torch.tensor([class_weight], dtype=torch.float32, device=DEVICE)
    criterion = nn.BCEWithLogitsLoss(pos_weight=pos_weight)

    best_val_f1 = -1.0; best_metrics = None; bad_epochs = 0
    for epoch in range(1, epochs + 1):
        model.train()
        optimizer.zero_grad()
        logits = model(data.x, data.edge_index)
        loss = criterion(logits[train_mask], data.y[train_mask].float())
        loss.backward(); optimizer.step()

        model.eval()
        with torch.no_grad():
            logits = model(data.x, data.edge_index)
            probs = torch.sigmoid(logits)
        # val metrics
        y_true = data.y[val_mask].cpu().numpy()
        y_prob = probs[val_mask].cpu().numpy()
        y_pred = (y_prob >= 0.5).astype(np.int32)
        val_f1 = f1_score(y_true, y_pred, zero_division=0)
        if val_f1 > best_val_f1:
            best_val_f1 = val_f1
            best_metrics = compute_metrics(y_true, y_pred, y_prob)
            best_metrics["epoch"] = epoch
            bad_epochs = 0
        else:
            bad_epochs += 1
            if bad_epochs >= patience:
                break

    return best_metrics or compute_metrics(
        data.y[val_mask].cpu().numpy(),
        np.zeros(int(val_mask.sum().item()), dtype=np.int32),
        np.zeros(int(val_mask.sum().item()), dtype=np.float32),
    )


def compute_metrics(y_true: np.ndarray, y_pred: np.ndarray, y_prob: np.ndarray) -> dict:
    pos = int(y_true.sum())
    acc = accuracy_score(y_true, y_pred)
    p = precision_score(y_true, y_pred, zero_division=0)
    r = recall_score(y_true, y_pred, zero_division=0)
    f1 = f1_score(y_true, y_pred, zero_division=0)
    try:
        if len(set(y_true)) > 1:
            auroc = roc_auc_score(y_true, y_prob)
            auprc = average_precision_score(y_true, y_prob)
        else:
            auroc = float("nan"); auprc = float("nan")
    except Exception:
        auroc = float("nan"); auprc = float("nan")
    cm = confusion_matrix(y_true, y_pred, labels=[0, 1]).tolist()
    return {
        "n_pos_true": pos, "n_pos_pred": int(y_pred.sum()),
        "accuracy": float(acc), "precision": float(p), "recall": float(r), "f1": float(f1),
        "auroc": float(auroc) if not math.isnan(auroc) else None,
        "auprc": float(auprc) if not math.isnan(auprc) else None,
        "confusion_matrix": cm,
    }


def cross_validate(dataset: GraphDataset, model_name: str, k: int = 5, epochs: int = 200) -> dict:
    print(f"\n[cv] {model_name}  k={k}  epochs={epochs}")
    X = dataset.node_features
    y = dataset.labels
    n = len(y); pos = int(y.sum()); neg = n - pos
    class_weight = max(1.0, neg / max(1, pos))

    edge_index = torch.tensor(dataset.edges, dtype=torch.long).t().contiguous() if dataset.edges else \
        torch.empty((2, 0), dtype=torch.long)
    # 무방향 → 양방향
    if edge_index.numel():
        edge_index = torch.cat([edge_index, edge_index.flip(0)], dim=1)
    data = Data(x=torch.tensor(X, dtype=torch.float32),
                edge_index=edge_index,
                y=torch.tensor(y, dtype=torch.long))

    skf = StratifiedKFold(n_splits=k, shuffle=True, random_state=SEED)
    fold_metrics = []
    t0 = time.time()
    for fold_i, (tr_idx, val_idx) in enumerate(skf.split(np.zeros(n), y), start=1):
        train_mask = torch.zeros(n, dtype=torch.bool); train_mask[tr_idx] = True
        val_mask = torch.zeros(n, dtype=torch.bool); val_mask[val_idx] = True

        if model_name == "MLP":
            model = MLPClassifier(X.shape[1])
        elif model_name == "GraphSAGE":
            model = GraphSAGENet(X.shape[1])
        elif model_name == "GAT":
            model = GATNet(X.shape[1])
        else:
            raise ValueError(model_name)

        m = train_one_fold(model, data, train_mask, val_mask, class_weight, epochs=epochs)
        fold_metrics.append(m)
        print(f"  fold {fold_i}/{k}  ep={m['epoch']:3d}  acc={m['accuracy']:.3f}  "
              f"f1={m['f1']:.3f}  prec={m['precision']:.3f}  rec={m['recall']:.3f}  "
              f"auroc={m['auroc']}  auprc={m['auprc']}")

    elapsed = time.time() - t0

    def agg(key):
        vals = [m[key] for m in fold_metrics if m[key] is not None]
        if not vals: return {"mean": None, "std": None}
        return {"mean": float(np.mean(vals)), "std": float(np.std(vals))}

    summary = {
        "model": model_name, "k_folds": k, "elapsed_s": elapsed,
        "class_weight": class_weight,
        "metrics_mean_std": {
            "accuracy": agg("accuracy"),
            "precision": agg("precision"),
            "recall": agg("recall"),
            "f1": agg("f1"),
            "auroc": agg("auroc"),
            "auprc": agg("auprc"),
        },
        "fold_metrics": fold_metrics,
    }
    print(f"  → {model_name} 평균 F1={summary['metrics_mean_std']['f1']['mean']:.3f}±"
          f"{summary['metrics_mean_std']['f1']['std']:.3f}  "
          f"AUROC={summary['metrics_mean_std']['auroc']['mean']}  ({elapsed:.1f}s)")
    return summary


# ────────────────────────────────────────────────────────────────────────────
# 메인
# ────────────────────────────────────────────────────────────────────────────


def main():
    DXF = PROJECT / "대명동201동 단위세대_layer정리.dxf"
    if not DXF.exists():
        raise FileNotFoundError(DXF)

    OUT_DIR = PROJECT / "data" / "gnn_training"
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # 데이터셋 준비 (1번만)
    ds = build_dataset(DXF)

    # 모델별 학습
    summaries = {}
    for model_name in ["MLP", "GraphSAGE", "GAT"]:
        summaries[model_name] = cross_validate(ds, model_name, k=5, epochs=200)

    # 결과 비교 표
    print(f"\n{'='*70}\n[final] 5-fold CV 결과 비교 (mean ± std)\n{'='*70}")
    cols = ["model", "f1", "precision", "recall", "auroc", "auprc", "accuracy"]
    header = f"{cols[0]:<12s} | " + " | ".join(f"{c:>14s}" for c in cols[1:])
    print(header); print("-" * len(header))
    rows = []
    for m, s in summaries.items():
        ms = s["metrics_mean_std"]
        def fmt(k):
            v = ms[k]
            if v["mean"] is None: return "  n/a"
            return f"{v['mean']:.3f}±{v['std']:.3f}"
        row_vals = [m] + [fmt(c) for c in cols[1:]]
        print(f"{row_vals[0]:<12s} | " + " | ".join(f"{v:>14s}" for v in row_vals[1:]))
        rows.append(row_vals)

    # JSON 저장
    result_json = OUT_DIR / "metrics.json"
    result_json.write_text(json.dumps({
        "dataset": {
            "dxf": DXF.name,
            "nodes": int(len(ds.labels)),
            "edges": int(len(ds.edges)),
            "positive_count": int(ds.labels.sum()),
            "positive_ratio": float(ds.labels.mean()),
            "feature_dim": int(ds.node_features.shape[1]),
        },
        "models": summaries,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n[save] {result_json}")

    # Markdown 보고서
    md = OUT_DIR / "report.md"
    lines = [
        "# Remote 30 헤드 노드 분류 — GNN 학습 보고서",
        "",
        f"- 데이터: `{DXF.name}`",
        f"- 노드 수: **{len(ds.labels)}**, edge 수: **{len(ds.edges)}**",
        f"- 양성 (HEAD) 노드: **{int(ds.labels.sum())}** ({ds.labels.mean()*100:.1f}%)",
        f"- 노드 feature 차원: **{ds.node_features.shape[1]}**",
        f"- 디바이스: `{DEVICE}`",
        "",
        "## 5-fold CV 결과 (mean ± std)",
        "",
        "| Model | F1 | Precision | Recall | AUROC | AUPRC | Accuracy |",
        "|---|---|---|---|---|---|---|",
    ]
    for m, s in summaries.items():
        ms = s["metrics_mean_std"]
        def fmt(k):
            v = ms[k]
            if v["mean"] is None: return "n/a"
            return f"{v['mean']:.3f}±{v['std']:.3f}"
        lines.append(f"| **{m}** | {fmt('f1')} | {fmt('precision')} | {fmt('recall')} | "
                     f"{fmt('auroc')} | {fmt('auprc')} | {fmt('accuracy')} |")
    # 결론 자동 생성
    best_f1 = max(summaries.values(), key=lambda s: s["metrics_mean_std"]["f1"]["mean"] or 0)
    lines += [
        "",
        f"### Best F1: **{best_f1['model']}** (mean={best_f1['metrics_mean_std']['f1']['mean']:.3f})",
        "",
        "## 시사점",
        "- MLP vs GraphSAGE 차이: GraphSAGE 가 의미있게 우세하면 그래프 구조가 헤드 분류에 기여.",
        "- GAT vs GraphSAGE: 어텐션 효과 — 일부 이웃에 가중치 부여가 도움이 되는지.",
        "- 양성 비율이 극히 낮으면 (~10%) precision/recall trade-off 관찰 필수.",
        "",
        f"## Fold 별 상세 — JSON: `{result_json.name}`",
    ]
    md.write_text("\n".join(lines), encoding="utf-8")
    print(f"[save] {md}")
    print("\n[done]")


if __name__ == "__main__":
    main()
