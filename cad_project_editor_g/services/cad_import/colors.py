"""CAD 색표 SSOT. 화면·matplotlib 없음.

대화형 다크 캔버스와 확인 그림(라이트)이 같은 표를 쓴다.
"""
import zlib

# ACI 이름 — 로그·스펙 표시. 모르는 번호는 색N.
ACI = {1: "빨강", 2: "노랑", 3: "초록", 4: "하늘", 5: "파랑", 6: "분홍",
       7: "흰검", 8: "회색", 9: "연회색", 80: "진초록80", 10: "빨강10",
       30: "주황30"}

# ACI 표준색 — 오너 CAD 화면과 같은 색으로 그려야 대조가 된다(흰검7=검정).
ACI_RGB = {1: "#ff0000", 2: "#cccc00", 3: "#00a000", 4: "#00b0b0",
           5: "#0000ff", 6: "#ff00ff", 7: "#000000", 8: "#808080",
           9: "#bbbbbb", 10: "#ff0000", 16: "#a0522d", 30: "#ff7f00",
           80: "#008040", 253: "#606060"}

C = {
    "재료":       "#1155cc",
    "의심조각":    "#ff9900",
    "합친토막":    "#c000c0",
    "도장안빠짐":   "#e02020",
    "헤드후보_켬":  "#ff00ff",
    "헤드후보_끔":  "#888888",
    "작은원":      "#00a651",
    "건축배경":    "#dddddd",
    "안붙은끝점":   "#ff0000",
    "배관기호":    "#ff9900",
    "헤드원":      "#00aa00",
    "헤드중심":    "#006600",
}

BG_DARK = "#000000"
ACI_RGB_DARK = dict(ACI_RGB)
ACI_RGB_DARK.update({
    2: "#ffff00",
    3: "#00e000",
    4: "#00e0e0",
    5: "#5070ff",
    7: "#ffffff",
    8: "#a0a0a0", 9: "#c8c8c8", 253: "#9a9a9a",
})
C_DARK = {"재료": "#b366ff", "헤드원": "#b366ff", "배관기호": "#ffaa00",
          "안붙은끝점": "#ff5050", "글자": "#dddddd"}

PALETTE = ["#3cb44b", "#911eb4", "#008080", "#9a6324", "#800000",
           "#808000", "#000075", "#e75480", "#2f4f4f"]
# 손질 물덩이 색 — 크기 순위. wrap 하면 큰 덩이 둘이 같은 색이 된다.
BODY_COLORS = ("#3ba7ff", "#ff5a3c", "#4ddb63", "#ffd21e", "#c86bff", "#ff8ad0",
               "#2ee6d0", "#ffa040", "#8fd400", "#7f9bff", "#ff4f8b", "#00d0a0")
# 헤드 종류 — 물흐름 전·젖은 = 밝은색 / 마른 = 같은 색상 어두운 톤
KIND_COLORS = {
    "상향식": "#ffa040",
    "하향식": "#3ba7ff",
    "상하향식": "#c86bff",
    "미지정": "#9a9a9a",
}
KIND_COLORS_DRY = {
    "상향식": "#6a4018",
    "하향식": "#1a4a78",
    "상하향식": "#4a2868",
    "미지정": "#3a3a3a",
}
KIND_COLOR_NAME = {
    "상향식": "주황", "하향식": "파랑", "상하향식": "보라", "미지정": "회색",
}
EDIT_SOURCE = "#ffffff"
EDIT_VALVE = "#9b59b6"
EDIT_PENDING = "#ffffff"
EDIT_WET_PIPE = "#22b573"
JOIN_COLOR = {2: "#0080ff", 3: "#ff8000", 4: "#ff0000"}
JOIN_NAME = {2: "파랑", 3: "주황", 4: "빨강"}
MMPP = {"훑기": 20.0, "확인": 15.0, "부검": 3.0}
MMPP_DEFAULT = MMPP["확인"]


def cname(c):
    return ACI.get(c, f"색{c}")


def rgb(c):
    """색: ACI 번호(int) 또는 '#rrggbb'(str)."""
    if isinstance(c, str):
        return c
    return ACI_RGB.get(c, "#996633")


def _aci_full(c):
    """표에 없는 ACI 번호 → **AutoCAD 표준색**. 지어내지 않는다.

    ★손으로 적은 표(`ACI_RGB`)는 흔한 번호 열댓 개뿐이라, 실도면의 51·105·
      255 같은 번호가 전부 같은 대체색으로 떨어졌다. 레이어가 달라도 화면에서
      같은 색이 되니 «색으로 구분» 이 성립하지 않는다(실측: 계통도 11색 중
      3개, 기계실 13색 중 6개가 한 색으로 뭉쳤다).

      ezdxf 가 256색 표준표를 이미 갖고 있다(`DXF_DEFAULT_COLORS`) — 이 저장소가
      DXF 를 읽는 데 쓰는 바로 그 라이브러리다. 근사식을 새로 쓰지 않고 그것을
      부른다. 없으면(라이브러리 부재) 종전 대체색으로 조용히 돌아간다.
    """
    try:
        from ezdxf.colors import aci2rgb
        r, g, b = aci2rgb(int(c))
    except Exception:  # noqa: BLE001 — 색 하나 때문에 화면이 죽으면 안 된다
        return None
    return f"#{r:02x}{g:02x}{b:02x}"


def rgb_dark(c):
    """다크 캔버스용 색 — ACI 번호(int) 또는 '#rrggbb'(str)."""
    if isinstance(c, str):
        return c
    # 손으로 고른 다크용 색이 먼저다 — 검정(7)을 흰색으로 바꾸는 등, 어두운
    # 캔버스에서 읽히게 손본 값이라 표준색보다 이쪽이 맞다.
    if c in ACI_RGB_DARK:
        return ACI_RGB_DARK[c]
    return _aci_full(c) or "#c8a064"


def blob_color(pts):
    """덩어리 색 — 순번이 아니라 자리. hash() 금지(실행마다 달라짐)."""
    ax, ay = min(pts)
    return PALETTE[zlib.crc32(f"{ax:.0f},{ay:.0f}".encode()) % len(PALETTE)]
