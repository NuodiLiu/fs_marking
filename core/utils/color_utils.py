# GPT has deprecaded info, Font contains ObjectThemeColor, RGB, TintAndShade

import colorsys
import win32com.client as win32
from win32com.client import constants
import math
try:
    from colormath.color_objects import sRGBColor, LabColor, LCHabColor
    from colormath.color_conversions import convert_color
    _HAS_COLORMATH = True
except ImportError:
    _HAS_COLORMATH = False

try:
    from colour import XYZ_to_OKLab, sRGB_to_XYZ
    _HAS_OKLAB = True
except ImportError:
    _HAS_OKLAB = False


def apply_tint(base_bgr: int, tint: float) -> int:
    """
    把主题色的 BGR 基色 和 tint/shade 应用到一起，
    返回最终的 BGR（0xBBGGRR）。
    """
    b = (base_bgr >> 16) & 0xFF
    g = (base_bgr >>  8) & 0xFF
    r =  base_bgr        & 0xFF
    if tint >= 0:
        r = int(r + (255 - r) * tint)
        g = int(g + (255 - g) * tint)
        b = int(b + (255 - b) * tint)
    else:
        r = int(r * (1 + tint))
        g = int(g * (1 + tint))
        b = int(b * (1 + tint))
    return (b << 16) | (g << 8) | r

def apply_tint_rgb(r: int, g: int, b: int, tint: float) -> tuple[int, int, int]:
    """
    接受 RGB 分量，返回应用 tint 后的新 RGB。
    """
    if tint >= 0:
        r = int(r + (255 - r) * tint)
        g = int(g + (255 - g) * tint)
        b = int(b + (255 - b) * tint)
    else:
        r = int(r * (1 + tint))
        g = int(g * (1 + tint))
        b = int(b * (1 + tint))
    return r, g, b

def bgr_to_rgb_tuple(bgr24: int) -> tuple[int,int,int]:
    b = (bgr24 >> 16) & 0xFF
    g = (bgr24 >>  8) & 0xFF
    r =  bgr24        & 0xFF
    return r, g, b

def int_to_rgb(val: int) -> tuple[int,int,int]:
    return ((val >> 16) & 0xFF,
            (val >>  8) & 0xFF,
            (val      ) & 0xFF)

def get_ui_color_rgb(font) -> tuple[int,int,int]:
    if font.TextColor.ObjectThemeColor != -1:
        r, g, b = int_to_rgb(font.Color)
        return apply_tint_rgb(r, g, b, font.TextColor.TintAndShade)

    # —— 2. 回退到标准/自定义色
    raw = None
    if hasattr(font, "TextColor"):
        try:
            raw = font.TextColor.RGB
        except Exception:
            raw = None
    if raw is None:
        raw = font.Color
    bgr24 = raw & 0x00FFFFFF
    return bgr_to_rgb_tuple(bgr24)

def is_blue_hsl(r: int, g: int, b: int) -> bool:
    """
    用 HSL Hue(135-255) + Saturation>=0.15 判断蓝色范围。
    """
    rn, gn, bn = r/255.0, g/255.0, b/255.0
    h, l, s = colorsys.rgb_to_hls(rn, gn, bn)
    deg = h * 360
    return 135 <= deg <= 255 and s >= 0.15

def is_blue_hsl(rn, gn, bn) -> bool:
    h, l, s = colorsys.rgb_to_hls(rn, gn, bn)
    deg = h * 360
    return 135 <= deg <= 255 and s >= 0.15

def is_blue_lab(r, g, b) -> bool:
    if not _HAS_COLORMATH:
        return False
    rgb = sRGBColor(r, g, b, is_upscaled=True)
    lab: LabColor = convert_color(rgb, LabColor)
    # b* 轴负代表偏蓝
    return lab.lab_b < -5

def is_blue_ycbcr(rn, gn, bn) -> bool:
    # BT.601
    y  = 0.299*rn + 0.587*gn + 0.114*bn
    cb = (bn - y)*0.564 + 0.5
    cr = (rn - y)*0.713 + 0.5
    return cb > cr and cb >= 0.4

# for para range and heading style font only, wordart and textframe should use fillcolor
def is_font_color_blueish(font, para_format=None) -> bool:
    """
    判断给定 Font 的 UI 颜色是否“蓝色”。
    参数:
      font         -- win32com.client Dispatch 的 Font 对象
      para_format  -- win32com.client Dispatch 的 ParagraphFormat（可忽略）
    返回:
      True if blue-ish, else False
    """
    # —— 获取 UI 上真正的 (r, g, b)
    r, g, b = get_ui_color_rgb(font)  # 复用之前写好的函数

    # —— 打印调试
    block = f"\x1b[48;2;{r};{g};{b}m  \x1b[0m"
    print(f"{block} UI RGB=({r},{g},{b}) hex=#{r:02X}{g:02X}{b:02X}")

    # —— 三空间判断
    rn, gn, bn = r/255.0, g/255.0, b/255.0
    vote_hsl   = is_blue_hsl(rn, gn, bn)
    vote_lab   = is_blue_lab(r, g, b)
    vote_ycbcr = is_blue_ycbcr(rn, gn, bn)

    votes = {"HSL": vote_hsl, "Lab": vote_lab, "YCbCr": vote_ycbcr}
    count = sum(votes.values())
    total = len(votes)
    print(votes, "→", count, "/", total)

    return count >= math.ceil(total / 2)