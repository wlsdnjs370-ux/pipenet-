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

import copy as _copy
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
        io = "Input" if row.get("io_node") == "Input" else "No"
        node = m.Node(
            node_id=lab,
            # ★/1000 을 하지 않는다. 좌표는 `normalize_node_coords` 가 이미
            #   캔버스 단위로 맞춰 놓았다 — 여기서 또 나누면 다시 한 점에 뭉친다.
            x=float(row.get("x", 0)),
            y=float(row.get("y", 0)),
            z=float(row.get("elevation", 0.0)),     # 수리 표고(m) — 손대지 않는다
            node_type=("input" if io == "Input" else "base"),
        )
        # PIPENET 규약은 "No"/"Input" 이다. writer 는 metadata 가 없으면
        # node_type("base")를 그대로 흘려보내 규약 밖 값이 나간다(§증상 3).
        try:
            node.metadata["io_node"] = io
        except Exception:      # noqa: BLE001 — 메타를 못 실어도 좌표는 살린다
            pass
        net.nodes[lab] = node

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
                    project_title: str = "Module G 수리계산 입력",
                    iso: bool = False, iso_z_scale: float = 1.0,
                    canvas_units: float = 3000.0,
                    iso_ref_label=None, iso_no_lift_labels=None) -> Path:
    """SDF + SLF 를 **한 쌍으로** 저장한다. 자산이 없으면 아무것도 안 만든다.

    `iso` / `iso_z_scale` / `canvas_units` 는 **표시 전용**이다(§G12). 수리계산은
    length·elevation·rise 로 하므로 이 값을 어떻게 두어도 계산 결과는 같다.
    기본값으로 부르면 종전 호출부가 그대로 동작한다(§3).

    반환: 쓴 .sdf 경로. .slf 는 같은 stem 으로 옆에 놓인다.
    """
    # ★자산 확인이 먼저다. 파일을 반쯤 써 놓고 실패하면 안 된다(§T5).
    template = resolve_template_sdf()
    slf_src = resolve_standard_slf()

    from services.cad_import.design.sdf_post import (
        bake_isometric, check_schedule, inject_pipe_types,
        normalize_node_coords, sanitize_template)

    # 관종 이름을 먼저 검사한다 — 오타면 파일을 만들지 않는다(§G10).
    sched_by_pipe = {}
    for row in getattr(tables, "pipes", None) or ():
        sched_by_pipe[str(row.get("label"))] = check_schedule(
            row.get("type") or "")

    # ★부른 쪽의 표를 건드리지 않는다. 정규화·베이크는 노드 좌표를 in-place 로
    #   바꾸는데, 창은 `self._tables` 를 들고 있다가 **다시 저장**할 수 있다.
    #   그대로 두면 두 번째 저장에서 이미 굽힌 좌표를 또 굽어 망이 어긋나고,
    #   아이소를 껐다 켠 저장은 조용히 등각 그림이 된다. 노드만 복제하면 된다 —
    #   배관·노즐·부속·장비 행은 여기서 바뀌지 않는다.
    tables = _copy.copy(tables)
    tables.nodes = [dict(n) for n in (getattr(tables, "nodes", None) or ())]

    # ★순서가 중요하다: 정규화 → 베이크. 바꾸면 lift 배율이 어긋난다(§G12).
    #   좌표를 바꾸므로 이 복제본은 이 시점부터 «표시용» 이다 — 길이·표고는 안 건드린다.
    normalize_node_coords(tables, canvas_units=canvas_units)
    if iso:
        bake_isometric(tables, iso_z_scale=iso_z_scale,
                       ref_label=iso_ref_label,
                       no_lift_labels=iso_no_lift_labels)

    _, write_sdf = _pc_models()
    net = tables_to_network(tables, project_title=project_title)

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    write_sdf(net, out, template_path=template)
    # writer 는 <Pipe-type> 을 안 쓴다 — 규격 바인딩은 여기서 얹는다(§G9).
    inject_pipe_types(out, sched_by_pipe)

    # ★SLF 를 «먼저» 옆에 놓는다. 정리가 그 파일명을 가리키게 하는데, 아직 없는
    #   파일을 가리키면 정리한 의미가 없다.
    slf_dst = out.with_suffix(".slf")
    shutil.copyfile(slf_src, slf_dst)

    # 템플릿에서 묻어온 남의 경로·제목을 지우고 라이브러리를 옆의 SLF 로 돌린다(§G14).
    cleaned = sanitize_template(out, slf_dst.name)

    print(f"[G6] SDF {out.name} · {out.stat().st_size:,} bytes "
          f"(노드 {len(net.nodes)} · 배관 {len(net.pipes)} · "
          f"노즐 {len(net.nozzles)})")
    print(f"[G6] SLF {slf_dst.name} · {slf_dst.stat().st_size:,} bytes "
          f"— SDF 는 이 라이브러리 없이 열면 관경이 'Unset' 이 된다")
    print(f"[G14] 템플릿 잔재 정리 · 남의 라이브러리 경로 {cleaned['user_lib']}건 → "
          f"'{slf_dst.name}' · 주기 {cleaned['text_element']}건 · "
          f"제목 {cleaned['title']}건 · 설명 {cleaned['net_desc']}건 지움")
    return out
