# -*- coding: utf-8 -*-
"""[오너 2026-09-22 · 그림 42~44] 계통도 여러 장을 «한 줄» 로 잇는다.

오너 지시: «계통도의 역할은 평면도의 배관망과 기계실의 배관망을 이어주는 것 …
하나의 계통도만으로 제대로 된 경로를 이끌지 못할 수도 있어 … 연결되는 공통 노드의
연결축은 평면도-계통도1-계통도2-계통도n-기계실».

각 계통도 칸은 지금처럼 두 점을 찍어 경로를 뽑는다.
  ① 기계실 쪽 끝 = `input_node_label` (한 장일 때의 «펌프»)
  ② 평면도 쪽 끝 = `av_node_label`    (한 장일 때의 «알람밸브»)
계통도 k 의 ① 점과 계통도 k+1 의 ② 점은 **같은 한 노드**(공통 노드)다.

여기서는 결합(`merge_network`) **앞** 에서 계통도들을 이어 계통도 하나(dict)를
만든다. 그 뒤는 지금 길을 그대로 탄다 — 평면도(알람밸브 10) · 긴 계통도 ·
기계실(연결점)의 결합, 통합 화면, .sdf · .kfp · .has 산출 모두.
계통도가 한 장이면 받은 것을 **그대로**(같은 객체) 돌려준다.

잇는 규칙(오너 승인 2026-09-22):
  · 계통도 k+1 을 통째로 평행이동해 ② 점을 앞 계통도의 ① 점에 겹친다.
    돌리거나 늘리지 않는다 — 각 도면의 모양·배관 길이는 그대로다.
  · 표고는 앞(평면도 쪽) 계통도 기준 — 계통도 k+1 전체를 올리거나 내려
    공통 노드 표고를 맞추고, 옮긴 양을 단계 기록에 남긴다.
  · 계통도 2 부터는 노드·배관 이름 앞에 `s<k>` 를 붙인다(예: 계통도 2 의
    `r3` → `s2r3`). PIPENET 라벨은 영문자로 시작하고 영문자·숫자만 쓸 수
    있어 밑줄을 쓰지 않는다. 계통도 1 은 지금 이름 그대로다.
  · 비었거나 끝점이 없는 계통도는 잇지 않고 멈춘다(S340 — 임의로 메우지 않는다).
"""
from __future__ import annotations

from routes.module_f.merge import MergeError

# 도면에서 뽑은 경로만 잇는다 — 배치(`system_layout`)가 같은 규칙으로 서는 것.
_EXTRACTED = frozenset({"dxf", "dxf_clean_network"})
_ROWS = ("fittings", "valves", "pumps")


def chain_label(k: int, label) -> str:
    """계통도 k(≥2)의 이름 — 앞에 `s<k>` 를 붙인다."""
    return f"s{int(k)}{label}"


def _xy(riser: dict, node: dict, flat_section: bool) -> tuple[float, float]:
    """잇기에 쓸 도면 자리. 단면(x·표고) 계통도는 y 를 0 으로 본다
    (`system_layout` 이 그 모드에서 y 를 0 으로 읽는 것과 같다)."""
    y = (0.0 if flat_section and riser.get("system_coordinate_mode") == "section_xz"
         else float(node.get("y", 0) or 0))
    return float(node.get("x", 0) or 0), y


def chain_risers(risers, names=None) -> dict:
    """계통도 1 … n (평면도 쪽부터) → 이어 붙인 계통도 하나.

    반환 dict 는 계통도 1 의 설정을 이어받고, `input_node_label` 은 계통도 n 의
    ① 점(새 이름)이다. `chain` 에 이음 기록(장 수 · 공통 노드 · 표고 이동)을 둔다.
    """
    risers = list(risers or ())
    names = list(names or [f"계통도 {i + 1}" for i in range(len(risers))])
    if not risers:
        raise MergeError("이을 계통도가 없습니다.")
    for name, r in zip(names, risers):
        if not r or not r.get("nodes") or not r.get("pipes"):
            raise MergeError(f"{name} 의 경로가 아직 없습니다 — 그 칸에서 두 점을 찍어 "
                             "«경로 추출» 을 하거나, 쓰지 않을 칸이면 × 로 지우세요.")
    if len(risers) == 1:
        return risers[0]
    for name, r in zip(names, risers):
        labels = {str(n.get("label")) for n in r["nodes"]}
        for key, what in (("av_node_label", "② 평면도 쪽 끝"),
                          ("input_node_label", "① 기계실 쪽 끝")):
            if str(r.get(key) or "") not in labels:
                raise MergeError(f"{name}: {what} 노드를 찾지 못했습니다 — 경로를 다시 뽑으세요.")
        if r.get("extracted_from") not in _EXTRACTED:
            raise MergeError(f"{name}: 도면에서 두 점으로 뽑은 경로만 이을 수 있습니다.")

    # 단면(x·표고) 계통도끼리만이면 종전 모드를 그대로 둔다. 섞이면 단면 쪽 y 를
    # 0 으로 펴서 잇는다 — 배치가 그 모드에서 하는 일과 같다.
    sections = [r.get("system_coordinate_mode") == "section_xz" for r in risers]
    flat = not all(sections)

    base = risers[0]
    nodes, at = [], {}
    for n in base["nodes"]:
        m = dict(n)
        m["x"], m["y"] = _xy(base, n, flat)
        nodes.append(m)
        at[str(m["label"])] = m
    pipes = [dict(p) for p in base["pipes"]]
    rows = {key: [dict(x) for x in (base.get(key) or ())] for key in _ROWS}
    pipe_names = {str(p.get("label")) for p in pipes}
    joint = str(base["input_node_label"])
    joints, shifts = [], []
    notices = ([f"{names[0]} {base['elevation_notice']}"]
               if base.get("elevation_notice") else [])

    for k, (name, r) in enumerate(zip(names[1:], risers[1:]), start=2):
        av, inp = str(r["av_node_label"]), str(r["input_node_label"])
        raw = {str(n.get("label")): n for n in r["nodes"]}
        here = at[joint]
        jx, jy = float(here.get("x", 0) or 0), float(here.get("y", 0) or 0)
        ax, ay = _xy(r, raw[av], flat)
        dx, dy = jx - ax, jy - ay
        dz = float(here.get("elevation", 0) or 0) - float(raw[av].get("elevation", 0) or 0)
        new = {lab: (joint if lab == av else chain_label(k, lab)) for lab in raw}
        clash = {v for lab, v in new.items() if lab != av and v in at}
        if clash:
            raise MergeError(f"{name}: 노드 이름이 앞 계통도와 겹칩니다 — {sorted(clash)[:3]}")
        # 공통 노드는 이제 두 계통도 사이의 노드다 — 급수원 표시 · 경계 압력을 뗀다.
        here["io_node"] = "No"
        here.pop("pressure_pa", None)
        for n in r["nodes"]:
            lab = str(n.get("label"))
            if lab == av:
                continue
            m = dict(n)
            x, y = _xy(r, n, flat)
            m.update(label=new[lab], x=x + dx, y=y + dy,
                     elevation=round(float(n.get("elevation", 0) or 0) + dz, 6))
            if lab != inp:
                m["io_node"] = "No"
                m.pop("pressure_pa", None)
            nodes.append(m)
            at[m["label"]] = m
        for p in r["pipes"]:
            q = dict(p)
            q["label"] = chain_label(k, p.get("label"))
            if q["label"] in pipe_names:
                raise MergeError(f"{name}: 배관 이름이 앞 계통도와 겹칩니다 — {q['label']}")
            pipe_names.add(q["label"])
            q["in"], q["out"] = new[str(p["in"])], new[str(p["out"])]
            pipes.append(q)
        for key in _ROWS:
            for row in r.get(key) or ():
                q = dict(row)
                if "pipe" in q:
                    q["pipe"] = chain_label(k, q["pipe"])
                for end in ("in", "out"):
                    if end in q and str(q[end]) in new:
                        q[end] = new[str(q[end])]
                if key != "fittings" and "label" in q:
                    q["label"] = chain_label(k, q["label"])
                rows[key].append(q)
        joints.append(joint)
        if abs(dz) > 1e-6:
            shifts.append({"name": name, "dz_m": round(dz, 3)})
        if r.get("elevation_notice"):
            notices.append(f"{name} {r['elevation_notice']}")
        joint = new[inp]

    out = dict(base)
    out.update(nodes=nodes, pipes=pipes, input_node_label=joint, **rows)
    if flat:
        out.pop("system_coordinate_mode", None)
    if notices:
        out["elevation_notice"] = " · ".join(notices)
    step = (f"S720 계통도 {len(risers)}장 잇기 · 공통 노드 "
            + " · ".join(joints)
            + "".join(f" · {s['name']} 표고 {s['dz_m']:+.3f} m 맞춤" for s in shifts))
    out["chain"] = {"count": len(risers), "names": names, "joints": joints,
                    "shifts": shifts, "step": step}
    return out


def system_names(got: dict) -> list:
    """결합망이 쓴 계통도 이름들 — 평면도 쪽부터. 한 장이면 칸 이름 그대로 «계통도».

    이름은 칸 이름에서 띄어쓰기만 뺀 것이다(오너 표기 «공통노드 : 평면도-계통도1»).
    장 수는 이은 자리(`chain_joints`) 수 + 1 이라 결합할 때의 장 수가 그대로 나온다 —
    결합 뒤에 칸을 더해도 «지금 화면에 있는 그 결합망» 의 이름이 어긋나지 않는다.
    """
    n = len(got.get("chain_joints") or ()) + 1
    return ["계통도"] if n == 1 else [f"계통도{i}" for i in range(1, n + 1)]


def joint_names(got: dict) -> dict:
    """[오너 2026-09-22] 공통 노드마다 «어느 두 도면이 만나는 자리인가».

    평면도 쪽부터 차례로 넣는다(넣은 차례가 곧 연결 축이다):
        평면도 ∩ 계통도 1            → 평면도-계통도1
        계통도 k 의 ① = 계통도 k+1 의 ②  → 계통도k-계통도k+1
        기계실이 붙은 자리             → 계통도n-기계실
    번호(10 · 1 · s21)는 파일 안 라벨이라 여기서 쓰지 않는다 — 화면 글자만 이 이름을 쓴다.
    """
    parts = got.get("parts") or {}
    plan = {str(x) for x in (parts.get("plan") or ())}
    system = {str(x) for x in (parts.get("system") or ())}
    room = {str(x) for x in (parts.get("machineroom") or ())}
    names = system_names(got)
    out: dict = {}

    def put(label, name):
        label = str(label)
        if not label:
            return
        if label in out:
            if name not in out[label].split(" · "):
                out[label] += f" · {name}"
        else:
            out[label] = name

    for lab in sorted(plan & system):
        put(lab, f"평면도-{names[0]}")
    for lab in sorted(plan & room):
        put(lab, "평면도-기계실")
    for i, lab in enumerate(got.get("chain_joints") or ()):
        if i + 1 < len(names):
            put(lab, f"{names[i]}-{names[i + 1]}")
    if got.get("attached") and got.get("pump_junction"):
        put(got["pump_junction"], f"{names[-1]}-기계실")
    return out
