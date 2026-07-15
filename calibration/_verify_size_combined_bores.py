"""역할별 내경 배정(size_combined_bores) 불변식 검증.

토폴로지(수원→헤드), io_node Input = m1:
  기계실  m1 - m2 - "1"(펌프 junction)
  입상관  1 - 2 - 3 - 4 - 5 - "10"(AV)
  평면도  10 - h1(트렁크, 전량 30헤드) ─┬─ h2  (헤드 15)
                                        └─ h3  (헤드 15)

검증 불변식:
  1) 평면도 배관은 규약 입력값 이상(never-shrink, 유속 승급 허용)
  2) 입상관 전 구간 단일 균일경
  3) mr_bore >= riser_bore
  4) min(riser) >= max(plane)
  5) 멱등: 두 번 돌려도 changed==0
  6) 유속 초과 0 (≤50A 6 · ≥65A 10 m/s) — annotate_pipe_velocity 로 확인
  7) 기계실 없는 케이스(자연낙차)에서도 6) 성립
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from remote30_full_network import (  # noqa: E402
    size_combined_bores, annotate_pipe_velocity)


def _mk():
    labels = ["m1", "m2", "1", "2", "3", "4", "5", "10", "h1", "h2", "h3"]
    nodes = [{"label": l} for l in labels]
    nodes[0]["io_node"] = "Input"  # m1 = 수원
    plane_pipes = [
        {"in": "10", "out": "h1", "dia": 65},   # 트렁크: 30헤드 전량 → 유속 승급 대상
        {"in": "h1", "out": "h2", "dia": 50},   # 15헤드
        {"in": "h1", "out": "h3", "dia": 50},   # 15헤드
    ]
    riser_pipes = [
        {"in": "1", "out": "2", "dia": 125},
        {"in": "2", "out": "3", "dia": 150},
        {"in": "3", "out": "4", "dia": 100},
        {"in": "4", "out": "5", "dia": 125},
        {"in": "5", "out": "10", "dia": 125},
    ]
    mr_pipes = [
        {"in": "m1", "out": "m2", "dia": 80},
        {"in": "m2", "out": "1", "dia": 80},
    ]
    pipes = plane_pipes + riser_pipes + mr_pipes
    # 헤드 30개: h2 에 15, h3 에 15 (각 80 L/min → 트렁크 2400 L/min)
    nozzles = ([{"in": "h2", "flow_lmin": 80.0} for _ in range(15)]
               + [{"in": "h3", "flow_lmin": 80.0} for _ in range(15)])
    return nodes, pipes, nozzles


def _dias(pipes, keys):
    return [int(p["dia"]) for p in pipes if (p["in"], p["out"]) in keys]


def main():
    plane_keys = {("10", "h1"), ("h1", "h2"), ("h1", "h3")}
    riser_keys = {("1", "2"), ("2", "3"), ("3", "4"), ("4", "5"), ("5", "10")}
    mr_keys = {("m1", "m2"), ("m2", "1")}
    riser_labels = ["1", "2", "3", "4", "5", "10"]
    mr_labels = ["m1", "m2"]

    nodes, pipes, nozzles = _mk()
    plane_in = {k: d for k, d in zip(plane_keys, _dias(pipes, plane_keys))}  # noqa: F841
    plane_before = {(p["in"], p["out"]): int(p["dia"])
                    for p in pipes if (p["in"], p["out"]) in plane_keys}

    r = size_combined_bores(nodes, pipes, nozzles,
                            riser_labels=riser_labels,
                            machine_room_labels=mr_labels, safety=1.2)
    print("RESULT:", r)

    plane_after = {(p["in"], p["out"]): int(p["dia"])
                   for p in pipes if (p["in"], p["out"]) in plane_keys}
    riser_after = _dias(pipes, riser_keys)
    mr_after = _dias(pipes, mr_keys)
    print("plane :", plane_after)
    print("riser :", riser_after)
    print("mr    :", mr_after)

    ok = True

    def check(name, cond):
        nonlocal ok
        ok = ok and cond
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")

    print("\n불변식:")
    check("1) 평면도 never-shrink (>=규약입력)",
          all(plane_after[k] >= plane_before[k] for k in plane_before))
    check("2) 입상관 단일 균일경", len(set(riser_after)) == 1)
    check("3) mr_bore >= riser_bore", min(mr_after) >= max(riser_after))
    check("4) min(riser) >= max(plane)", min(riser_after) >= max(plane_after.values()))

    r2 = size_combined_bores(nodes, pipes, nozzles,
                             riser_labels=riser_labels,
                             machine_room_labels=mr_labels, safety=1.2)
    check("5) 멱등(2회차 changed==0)", r2["changed"] == 0)

    va = annotate_pipe_velocity(nodes, pipes, nozzles)
    print("  velocity annotate:", va)
    check("6) 유속 초과 0", va["violations"] == 0)
    check("6b) size 리턴 violations_after==0", r["violations_after"] == 0)

    # 기계실 없는 케이스
    nodes2, pipes2, nozzles2 = _mk()
    pipes2 = [p for p in pipes2 if (p["in"], p["out"]) not in mr_keys]
    nodes2 = [n for n in nodes2 if n["label"] not in ("m1", "m2")]
    nodes2[0] = dict(nodes2[0]); nodes2[0]["io_node"] = "Input"  # "1" = 수원
    r3 = size_combined_bores(nodes2, pipes2, nozzles2,
                             riser_labels=riser_labels,
                             machine_room_labels=[], safety=1.2)
    va3 = annotate_pipe_velocity(nodes2, pipes2, nozzles2)
    print("\nMR 없음:", r3, "| velocity:", va3)
    check("7) MR 없으면 mr_bore==0", r3["mr_bore"] == 0)
    check("7b) MR 없어도 유속 초과 0", va3["violations"] == 0)

    print("\n", "ALL PASS" if ok else "SOME FAILED")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
