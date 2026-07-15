"""역할별 내경 배정(size_combined_bores) 불변식 검증.

토폴로지(수원→헤드):
  기계실  m1 - m2 - (pump junction "1")
  입상관  1 - 2 - 3 - 4 - 5 - AV("10")
  평면도  10 - h1 - h2 (규약배관 별표1 값, 여기선 트렁크 65A, 가지 50A/40A)

검증 불변식:
  1) 평면도 배관 dia 불변(규약 유지)
  2) 입상관 전 구간 단일 균일경
  3) mr_bore >= riser_bore (기계실이 라이저보다 굵거나 같음)
  4) min(riser) >= max(plane)  ("가장 얇은 입상관 ≥ 평면도 최대경")
  5) 멱등: 두 번 돌려도 동일
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from remote30_full_network import size_combined_bores  # noqa: E402


def _mk():
    nodes = [{"label": l} for l in
             ["m1", "m2", "1", "2", "3", "4", "5", "10", "h1", "h2"]]
    # 평면도 규약값: 트렁크 65A, 가지 50A/40A
    plane_pipes = [
        {"in": "10", "out": "h1", "dia": 65},
        {"in": "h1", "out": "h2", "dia": 50},
        {"in": "h2", "out": "h2b", "dia": 40},
    ]
    # 입상관: 편집으로 뒤죽박죽(125/150/100) — 균일화 대상
    riser_pipes = [
        {"in": "1", "out": "2", "dia": 125},
        {"in": "2", "out": "3", "dia": 150},
        {"in": "3", "out": "4", "dia": 100},
        {"in": "4", "out": "5", "dia": 125},
        {"in": "5", "out": "10", "dia": 125},
    ]
    # 기계실: 얇게(80) — 라이저보다 굵어져야 함
    mr_pipes = [
        {"in": "m1", "out": "m2", "dia": 80},
        {"in": "m2", "out": "1", "dia": 80},
    ]
    pipes = plane_pipes + riser_pipes + mr_pipes
    nozzles = [{"flow_lmin": 80.0} for _ in range(30)]  # 30 헤드 × 80 = 2400 L/min
    return nodes, pipes, nozzles


def _dias(pipes, keys):
    return [int(p["dia"]) for p in pipes if (p["in"], p["out"]) in keys]


def main():
    nodes, pipes, nozzles = _mk()
    plane_keys = {("10", "h1"), ("h1", "h2"), ("h2", "h2b")}
    riser_keys = {("1", "2"), ("2", "3"), ("3", "4"), ("4", "5"), ("5", "10")}
    mr_keys = {("m1", "m2"), ("m2", "1")}
    riser_labels = ["1", "2", "3", "4", "5", "10"]
    mr_labels = ["m1", "m2"]

    plane_before = _dias(pipes, plane_keys)
    r = size_combined_bores(nodes, pipes, nozzles,
                            riser_labels=riser_labels,
                            machine_room_labels=mr_labels, safety=1.2)
    print("RESULT:", r)

    plane_after = _dias(pipes, plane_keys)
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
    check("1) 평면도 규약값 불변", plane_before == plane_after)
    check("2) 입상관 단일 균일경", len(set(riser_after)) == 1)
    check("3) mr_bore >= riser_bore", min(mr_after) >= max(riser_after))
    check("4) min(riser) >= max(plane)", min(riser_after) >= max(plane_after))

    # 5) 멱등
    r2 = size_combined_bores(nodes, pipes, nozzles,
                             riser_labels=riser_labels,
                             machine_room_labels=mr_labels, safety=1.2)
    check("5) 멱등(2회차 changed==0)", r2["changed"] == 0)

    # 기계실 없는 케이스(자연낙차)
    nodes2, pipes2, nozzles2 = _mk()
    pipes2 = [p for p in pipes2 if (p["in"], p["out"]) not in mr_keys]
    r3 = size_combined_bores(nodes2, pipes2, nozzles2,
                             riser_labels=riser_labels,
                             machine_room_labels=[], safety=1.2)
    print("\nMR 없음:", r3)
    check("6) MR 없으면 mr_bore==0", r3["mr_bore"] == 0)
    check("7) MR 없어도 riser>=plane", r3["riser_bore"] >= r3["plane_max"])

    print("\n", "ALL PASS" if ok else "SOME FAILED")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
