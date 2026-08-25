# -*- coding: utf-8 -*-
"""[G6] SDF 방출 — 템플릿 SDF + 표준 SLF 를 한 쌍으로.

템플릿을 써야 Graphics 블록(아이소매트릭 표시 메타·schemes·Display-options)이
보존되고, 표준 SLF 가 옆에 있어야 PIPENET 이 호칭경↔내경을 lookup 한다.
**둘 중 하나라도 없으면 파일을 만들지 않고 실패한다**(§G6 · §T5).

모듈 A 는 자산이 없어도 경고만 내고 진행하지만, G 는 사람이 그 자리에서 창을
보고 있다 — 조용히 이상한 파일이 나가면 안 된다. 관경이 "Unset" 으로 뜨는
SDF 를 «만들어졌다» 고 돌려주는 것이 가장 나쁜 결과다.

자산 경로 해석은 모듈 A 와 같은 순서다:
    ① 환경변수 REMOTE30_TEMPLATE_SDF / REMOTE30_STANDARD_SLF
    ② G 트리의 `design/assets/`
    ③ 저장소의 `assets/`  (모듈 A 와 같은 파일을 함께 쓴다)
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

_G_ROOT = Path(__file__).resolve().parents[3]
_REPO = Path(__file__).resolve().parents[4]
_ASSETS = Path(__file__).resolve().parent / "assets"

TEMPLATE_SDF_FILENAME = ("3-1형_자연낙차_LSP_4F_OA_지하층포함_120m~200m미만_"
                         "6.6K로 감압_알람밸브.sdf")
STANDARD_SLF_FILENAME = "2. Pipenet_hand_FX28.slf"


class AssetMissing(RuntimeError):
    """템플릿·표준 라이브러리가 없다. 경고가 아니라 실패다(§T5)."""


def _resolve(env_key: str, filename: str, role: str) -> Path:
    override = os.environ.get(env_key)
    if override:
        p = Path(override).expanduser()
        if not p.is_file():
            raise AssetMissing(
                f"{role}: 환경변수 {env_key} 가 가리키는 파일이 없습니다 — {p}")
        return p.resolve()
    for base in (_ASSETS, _REPO / "assets"):
        p = base / filename
        if p.is_file():
            return p.resolve()
    raise AssetMissing(
        f"{role} 을 찾지 못했습니다: {filename}\n"
        f"  찾은 곳: {_ASSETS} · {_REPO / 'assets'}\n"
        f"  → 환경변수 {env_key} 로 절대 경로를 지정하거나 위 폴더에 두세요.\n"
        f"  (이 자산이 없으면 관경이 'Unset' 으로 뜨거나 표시 메타가 빠집니다)")


def resolve_template_sdf() -> Path:
    return _resolve("REMOTE30_TEMPLATE_SDF", TEMPLATE_SDF_FILENAME,
                    "템플릿 SDF")


def resolve_standard_slf() -> Path:
    return _resolve("REMOTE30_STANDARD_SLF", STANDARD_SLF_FILENAME,
                    "표준 SLF")


def _pc_models():
    """pipenet_converter 는 src layout 이라 경로를 붙여 준다."""
    src = _REPO / "pipenet_converter" / "src"
    if src.is_dir() and str(src) not in sys.path:
        sys.path.insert(0, str(src))
    from pipenet_converter import models as m
    from pipenet_converter.sdf_writer import write_sdf
    return m, write_sdf


def tables_to_network(tables, *, project_title: str):
    """PipeTablesG → pipenet_converter 의 PipeNetwork.

    단위 되돌리기(§T3): 표의 노드 좌표는 mm, 모델은 m 를 쓴다. 호칭경 mm 는
    지름 m 로 바꾼다 — 여기서 자리수를 틀리면 SDF 가 통째로 어긋난다.
    """
    m, _ = _pc_models()
    net = m.PipeNetwork(title=project_title)

    for row in tables.nodes:
        lab = str(row["label"])
        net.nodes[lab] = m.Node(
            node_id=lab,
            x=float(row.get("x", 0)) / 1000.0,      # mm → m
            y=float(row.get("y", 0)) / 1000.0,
            z=float(row.get("elevation", 0.0)),     # 이미 m
            node_type=("input" if row.get("io_node") == "Input" else "base"),
        )

    fit_by_pipe: dict = {}
    for f in tables.fittings:
        fit_by_pipe.setdefault(str(f["pipe"]), []).append(
            m.Fitting(fitting_type=str(f["type"]), count=int(f.get("count", 1))))

    eq_by_pipe: dict = {}
    for e in tables.equipment:
        eq_by_pipe.setdefault(str(e["pipe"]), []).append(
            m.Equipment(equipment_id=str(e.get("label", "")),
                        description=str(e.get("desc", "")),
                        equivalent_length_m=float(e.get("eq_len", 0.0)),
                        rel_position=float(e.get("rel_pos", 0.5))))

    for row in tables.pipes:
        pid = str(row["label"])
        pipe = m.Pipe(
            pipe_id=pid,
            from_node=str(row["in"]), to_node=str(row["out"]),
            diameter_m=float(row.get("dia", 0)) / 1000.0,   # 호칭경 mm → m
            length_m=float(row.get("length", 0.0)),
            rise_m=float(row.get("elev", 0.0)),
            c_factor=float(row.get("c", 120)),
            material=str(row.get("type", "")) or None,
            status=str(row.get("status", "Normal")).lower(),
            fittings=fit_by_pipe.get(pid, []),
        )
        if hasattr(pipe, "equipment") and eq_by_pipe.get(pid):
            pipe.equipment = eq_by_pipe[pid]
        net.pipes[pid] = pipe

    for row in tables.nozzles:
        lab = str(row["label"])
        # 유량은 표의 L/min 에서 유도한다 — 손으로 자른 상수를 쓰면 되돌릴 때
        # 값이 어긋난다(모듈 A 도 같은 이유로 L/min 을 원본으로 둔다).
        flow_m3s = row.get("flow_m3s")
        if flow_m3s is None:
            flow_m3s = float(row.get("flow_lmin", 0.0)) / 60000.0
        net.nozzles[lab] = m.Nozzle(
            nozzle_id=lab,
            input_node=str(row["in"]), output_node=str(row["out"]),
            flow_m3s=float(flow_m3s),
            status=int(row.get("status", 1)),
            library_item=str(row.get("lib", "SP-HEAD")),
        )
    return net


def emit_design_sdf(tables, out_path, *,
                    project_title: str = "Module G 수리계산 입력") -> Path:
    """SDF + SLF 를 **한 쌍으로** 저장한다. 자산이 없으면 아무것도 안 만든다.

    반환: 쓴 .sdf 경로. .slf 는 같은 stem 으로 옆에 놓인다.
    """
    # ★자산 확인이 먼저다. 파일을 반쯤 써 놓고 실패하면 안 된다(§T5).
    template = resolve_template_sdf()
    slf_src = resolve_standard_slf()

    _, write_sdf = _pc_models()
    net = tables_to_network(tables, project_title=project_title)

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    write_sdf(net, out, template_path=template)

    slf_dst = out.with_suffix(".slf")
    shutil.copyfile(slf_src, slf_dst)

    print(f"[G6] SDF {out.name} · {out.stat().st_size:,} bytes "
          f"(노드 {len(net.nodes)} · 배관 {len(net.pipes)} · "
          f"노즐 {len(net.nozzles)})")
    print(f"[G6] SLF {slf_dst.name} · {slf_dst.stat().st_size:,} bytes "
          f"— SDF 는 이 라이브러리 없이 열면 관경이 'Unset' 이 된다")
    return out
