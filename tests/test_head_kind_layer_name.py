# -*- coding: utf-8 -*-
"""[헤드 상하향] 도면 «이름» 이 확정적으로 말하면 그것이 권위다.

■ 무엇이 문제였나 (대명동 단위세대 평면도 · 사용자 지목으로 드러남)

  사용자가 수리계산 표에서 네 헤드(118·119·59·50)를 지목했다. 넷 다 표고
  **+0.3 m**(가지관 «위»)였다. 그런데 그 헤드들이 사는 레이어 이름은

      -소화(SP헤드하향)

  이다. 도면 작성자는 «하향» 이라고 적어 두었는데 전개는 상향식으로 읽었다.
  실측: **111개 전부** 그랬다.

■ 원인 사슬

      classify_head_kind()   반호가 없으면 → "상향식"     (근거 없는 기본값)
                             `return "하향식" if owned else "상향식"`
      kind_of_head()         판정이 «미지정» 일 때만 이름 폴백
                             → 폴백이 **영영 안 불린다**
      HEAD_DZ_M["상향식"]   = +0.3   (하향식이면 −0.3)
      bake_isometric         de ≥ 0 → 스텁이 **위로**

  ★그림 문제가 아니다. 상하향은 헤드가 가지관 «위» 냐 «아래» 냐라 표고의
  **부호**를 가른다. 평면에서는 헤드와 부모가 같은 점이라 안 보이고, 등각을
  켜야 스텁 방향으로 드러난다 — 사용자가 「아이소를 누를 때」라고 한 이유다.

■ 결정 (2026-09-09 오너)

  「도면 이름이 맞다 (하향식)」. 이름이 상향/하향을 **명시하면** 그것이 권위이고,
  기하 판정은 근거 없는 기본값으로 떨어질 수 있으므로 진다. 이름이 아무 말도
  안 하는 레이어는 종전 그대로 기하 판정을 쓴다 — 그런 도면은 안 바뀐다.

  조치 뒤 실측(대명동 K=30): 헤드 표고 {0.3: 22} → **{0.0: 22}**,
  등각 스텁 «위로 22» → **«아래로 22»**, 표고차↔접속관 길이 불일치 0 유지.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _fw():
    from services.cad_import.pipeline import flow
    return flow


# ─────────────────────────────── 이름이 말하면 이름이 권위
def test_이름이_하향이면_기하_상향식을_덮는다():
    """★대명동 실측 111/111 — 이 한 줄이 표고 부호를 뒤집는다."""
    fw = _fw()
    got, flipped = fw.kind_with_layer_name("상향식", "-소화(SP헤드하향)")
    assert got == "하향식" and flipped is True


def test_이름이_상향이면_기하_하향식을_덮는다():
    fw = _fw()
    got, flipped = fw.kind_with_layer_name("하향식", "-소화(SP헤드상향)")
    assert got == "상향식" and flipped is True


def test_이름이_상하향이면_그것이다():
    fw = _fw()
    got, flipped = fw.kind_with_layer_name("상향식", "SP헤드상하향")
    assert got == "상하향식" and flipped is True


def test_이름이_아무_말도_안_하면_기하_그대로다():
    """★그런 도면의 산출은 한 바이트도 안 바뀐다 — 이 시험이 그것을 지킨다."""
    fw = _fw()
    for layer in ("SPRINKLER", "FIRE", "", None, "-소화(SP가지관)"):
        got, flipped = fw.kind_with_layer_name("상향식", layer)
        assert got == "상향식" and flipped is False, layer
        got, flipped = fw.kind_with_layer_name("하향식", layer)
        assert got == "하향식" and flipped is False, layer


def test_같으면_뒤집었다고_말하지_않는다():
    fw = _fw()
    assert fw.kind_with_layer_name("하향식", "SP헤드하향") == ("하향식", False)


def test_기하가_미지정이면_이름을_쓴다():
    """종전 폴백 — 이름이 있으면 그것, 없으면 미지정 그대로."""
    fw = _fw()
    assert fw.kind_with_layer_name("미지정", "SP헤드하향")[0] == "하향식"
    assert fw.kind_with_layer_name("미지정", "SPRINKLER")[0] == "미지정"


# ─────────────────────────────── 권위 자리가 하나여야 한다
def test_1_1_분류가_그_규칙을_쓴다():
    """★표·산출·화면이 한 값을 보려면 «1-1 분류» 에서 뒤집어야 한다.

    이 저장소의 권위는 「1-1 분류 → 편집 kind_overrides」다(kinds.py 머리말).
    `kind_of_head` 만 고치면 물길은 하향으로 보는데 표는 상향으로 남는다.
    """
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "pipeline" / "flow.py").read_text(encoding="utf-8")
    i = src.index("def stage11_classify_heads(")
    seg = src[i:i + 3000]
    assert "kind_with_layer_name(" in seg, "1-1 분류가 이름을 안 본다"
    assert "kind_by_geometry" in seg, "무엇을 덮었는지 안 남긴다"
    # `kind_of_head` 도 같은 함수를 써야 두 곳이 안 갈린다.
    j = src.index("def kind_of_head(")
    assert "kind_with_layer_name(" in src[j:j + 400]


def test_뒤집은_것을_세어서_말한다():
    """조용히 뒤집으면 돌던 산출이 통째로 달라진 것을 아무도 모른다(S340)."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "pipeline" / "flow.py").read_text(encoding="utf-8")
    i = src.index('flipped = [r for r in out if r.get("kind_by_geometry")]')
    seg = src[i:i + 1400]
    assert "도면 이름을 따라 종류를 바꾼 헤드" in seg
    assert "표고의 **부호**" in seg, "왜 중요한지를 안 말한다"
    assert '_head_kind_name_flipped' in seg, "센 것을 안 실어 보낸다"


# ─────────────────────────────── ★이름 판정이 «접속» 까지 막으면 안 된다
def test_5단계_접속은_기하_판정을_쓴다():
    """★상하향은 두 가지를 가르는데, 이름이 권위인 것은 **하나뿐** 이다.

        ⑴ 헤드가 가지관 «위» 냐 «아래» 냐  → 표고의 부호 (이름이 권위)
        ⑵ 원 밑 통과관에 **붙일 수 있나**   → 5단계 접속 (기하 문제)

    `upright_disks` 는 「상향식 INCLUDE · 하향식 EXCLUDE」다. 이름 판정을 여기
    까지 끌고 오면 `-소화(SP헤드하향)` 같은 도면에서 5단계 접속이 **통째로**
    사라져 헤드가 배관에서 떨어진다 — 사용자가 「찍기에서 헤드 범주가 좁아진
    것 같다」고 한 그 자리다.

    (대명동은 중심접속 111/111 이라 안 드러났다. 원 밑 통과에 기대는 도면에서
     비로소 손실이 된다 — 안 드러난다고 없는 것이 아니다.)
    """
    from services.cad_import.pipeline.flow import upright_disks
    st = {"w": None, "spec": {"material_picks": []}}
    hcov = [(100.0, 100.0, 40.0)]
    #  이름은 하향식인데 기하는 상향식 — 5단계는 **기하** 를 봐야 한다.
    kinds = [{"c": (100.0, 100.0), "head_r": 40.0,
              "kind": "하향식", "kind_by_geometry": "상향식"}]
    assert upright_disks(st, hcov, kinds, arm_index={}) == hcov


def test_기하도_하향식이면_5단계에서_뺀다():
    """종전 규칙은 그대로다 — 기하가 하향이라 판정한 것은 계속 뺀다."""
    from services.cad_import.pipeline.flow import upright_disks
    st = {"w": None, "spec": {"material_picks": []}}
    hcov = [(100.0, 100.0, 40.0)]
    kinds = [{"c": (100.0, 100.0), "head_r": 40.0,
              "kind": "하향식", "kind_by_geometry": "하향식"}]
    assert upright_disks(st, hcov, kinds, arm_index={}) == []


def test_기하값이_없으면_kind_를_쓴다():
    """이름이 안 덮은 헤드는 `kind_by_geometry` 가 없다 — 그때는 종전 그대로."""
    from services.cad_import.pipeline.flow import upright_disks
    st = {"w": None, "spec": {"material_picks": []}}
    hcov = [(100.0, 100.0, 40.0)]
    assert upright_disks(st, hcov,
                         [{"c": (100.0, 100.0), "head_r": 40.0,
                           "kind": "상향식"}], arm_index={}) == hcov
    assert upright_disks(st, hcov,
                         [{"c": (100.0, 100.0), "head_r": 40.0,
                           "kind": "하향식"}], arm_index={}) == []


def test_판정을_끌_수_있다():
    """어느 조치가 무엇을 바꿨는지 가르려면 꺼 볼 수 있어야 한다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "pipeline" / "flow.py").read_text(encoding="utf-8")
    assert "MF_NO_LAYER_KIND" in src


# ─────────────────────────────── 상하향이 표고 부호를 가른다
def test_상하향이_표고_부호다():
    """★이 표가 «그림» 이 아니라 «수리계산» 을 가르는 근거다."""
    from services.cad_import.convert.engine import HEAD_DZ_M
    assert HEAD_DZ_M["상향식"] > 0
    assert HEAD_DZ_M["하향식"] < 0
    assert abs(HEAD_DZ_M["상향식"]) == abs(HEAD_DZ_M["하향식"]) == 0.3


def test_등각_스텁은_표고_부호를_따른다():
    """부호가 뒤집히면 스텁이 서는 방향도 뒤집힌다 — 등각에서만 보인다."""
    from services.cad_import.design.sdf_post import bake_isometric

    class _T:
        nodes = [{"label": "1", "x": 0.0, "y": 0.0, "elevation": 0.0},
                 {"label": "2", "x": 100.0, "y": 0.0, "elevation": 0.3},
                 {"label": "H", "x": 100.0, "y": 0.0, "elevation": 0.0}]
    got = bake_isometric(_T(), units_per_m=100.0,
                         head_nodes=["H"], head_parent={"H": "2"})
    at = {n["label"]: (n["x"], n["y"]) for n in _T.nodes}
    _ = got
    # 하향식 — 헤드(0.0)가 부모(0.3)보다 낮으니 화면에서 **아래** 다.
    assert at["H"][1] < at["2"][1], at


# ─────────────────────────────── ★사람이 찍은 것은 이름이 못 덮는다
def test_사람이_찍은_상하향식은_이름이_못_덮는다():
    """★권위의 차례 — 사람의 픽 > 도면 이름 > 기하 기본값.

    `classify_head_kind` 가 «상하향식» 을 돌려주는 갈래는 1)2)3) 뿐이고
    **셋 다 근거가 스펙의 상하향 x칸**이다. 기하가 추측한 값이 아니라
    사람이 02.찍기에서 그렇게 정한 값이다.

    오너가 「도면 이름이 맞다」고 한 것은 근거 없이 상향식으로 떨어지는
    «기하 기본값» 을 두고 한 말이다. 여기를 빼먹으면 `-소화(SP헤드하향)`
    도면에서 x칸으로 찍은 상하향식 헤드가 **전부 하향식**이 된다 — 사람이
    찍기에서 정한 것이 수리계산에 반영되지 않는 그 자리다.
    """
    fw = _fw()
    for layer in ("-소화(SP헤드하향)", "-소화(SP헤드상향)", "SP헤드하향"):
        got, flipped = fw.kind_with_layer_name("상하향식", layer)
        assert (got, flipped) == ("상하향식", False), layer


def test_상하향식은_오직_사람의_픽에서만_나온다():
    """★위 규칙의 전제 — 기하가 «상하향식» 을 지어내면 규칙이 무너진다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "pipeline" / "flow.py").read_text(encoding="utf-8")
    i = src.index("def classify_head_kind(")
    j = src.index("def kind_with_layer_name(")
    body = src[i:j]
    # «상하향식» 은 스펙(사람의 픽) 갈래에서만 나온다 — 기하 꼬리(5~6)에는
    #   단 한 번도 없어야 한다. 있으면 기하가 상하향식을 «지어내는» 것이라
    #   위 규칙(이름이 못 덮는다)이 기하 기본값까지 보호하게 된다.
    head, _, tail = body.partition("    # 5~6)")
    assert tail, "갈래 표시가 사라졌다 — 구조가 바뀌었으니 규칙을 다시 봐야 한다"
    assert 'return "상하향식"' not in tail, tail[:400]
    assert head.count('return "상하향식"') >= 3
    assert 'spec.get("heads")' in head and "dual_marks_of(spec)" in head
    assert 'normalize_head_slot(slot_src) == "상하향"' in head


def test_기하_기본값은_여전히_이름에_진다():
    """조치가 종전 결정을 되돌리지 않았는지 — 대명동 111/111 이 그대로다."""
    fw = _fw()
    assert fw.kind_with_layer_name("상향식", "-소화(SP헤드하향)") \
        == ("하향식", True)
