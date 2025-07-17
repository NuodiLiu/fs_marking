import colorsys
import math

# 可选依赖
try:
    from colormath.color_objects import sRGBColor, LabColor, LCHabColor
    from colormath.color_conversions import convert_color
    _HAS_COLORMATH = True
except ImportError:
    _HAS_COLORMATH = False

# 如果要用 Oklab，需要安装 colour-science 库
try:
    from colour import XYZ_to_OKLab, sRGB_to_XYZ
    _HAS_OKLAB = True
except ImportError:
    _HAS_OKLAB = False


def int_to_rgb(val: int) -> tuple[int,int,int]:
    return ((val >> 16) & 0xFF,
            (val >>  8) & 0xFF,
            (val      ) & 0xFF)

def is_blue_hsv(rn, gn, bn) -> bool:
    h, s, v = colorsys.rgb_to_hsv(rn, gn, bn)
    deg = h * 360
    return 180 <= deg <= 240 and s >= 0.25 and v >= 0.2

def is_blue_hsl(rn, gn, bn) -> bool:
    h, l, s = colorsys.rgb_to_hls(rn, gn, bn)
    deg = h * 360
    return 200 <= deg <= 240 and s >= 0.15

def is_blue_lch(r, g, b) -> bool:
    if not _HAS_COLORMATH:
        return False
    rgb = sRGBColor(r, g, b, is_upscaled=True)
    lch: LCHabColor = convert_color(rgb, LCHabColor)
    return 200 <= lch.lch_h <= 260 and lch.lch_c >= 20

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

def is_blue_oklab(rn, gn, bn) -> bool:
    if not _HAS_OKLAB:
        return False
    # colourscience: sRGB->XYZ->Oklab
    xyz = sRGB_to_XYZ([rn, gn, bn])
    oklab = XYZ_to_OKLab(xyz)
    # oklab = [L, a, b], b 负代表偏蓝
    return oklab[2] < -0.02


def is_blueish(val: int) -> bool:
    r, g, b = int_to_rgb(val)
    rn, gn, bn = r/255., g/255., b/255.

    # 各空间判断
    checks = {
        "HSV":   is_blue_hsv(rn, gn, bn),
        "HSL":   is_blue_hsl(rn, gn, bn),
        "LCh":   is_blue_lch(r, g, b),
        "Lab":   is_blue_lab(r, g, b),
        "YCbCr": is_blue_ycbcr(rn, gn, bn),
        "Oklab": is_blue_oklab(rn, gn, bn),
    }

    # 少于两项可用时，自动去掉对应空间
    available = [v for v in checks.values() if isinstance(v, bool)]
    vote_count = sum(available)
    threshold = math.ceil(len(available) / 2)

    # 调试时可以打印每个空间的投票结果：
    # print({k:v for k,v in checks.items() if isinstance(v,bool)}, "→ votes:", vote_count, "/", len(available))

    return vote_count >= threshold


def print_color_blocks(color_values: list[int]) -> None:
    for val in color_values:
        r, g, b = int_to_rgb(val)
        block = f"\x1b[48;2;{r};{g};{b}m  \x1b[0m"
        result = get_basic_color_category(val)
        print(f"{block}  0x{val:06X}  →  is_blueish: {result}")


if __name__ == "__main__":
    SAMPLE_COLORS = [
        0x000080,  # Navy
        0x00008B,  # DarkBlue
        0x0000CD,  # MediumBlue
        0x0000FF,  # Blue
        0x00BFFF,  # DeepSkyBlue
        0x1E90FF,  # DodgerBlue
        0x00CED1,  # DarkTurquoise
        0x00FFFF,  # Aqua / Cyan
        0x40E0D0,  # Turquoise
        0x48D1CC,  # MediumTurquoise
        0x20B2AA,  # LightSeaGreen
        0x7FFFD4,  # Aquamarine
        0x5F9EA0,  # CadetBlue
        0x4682B4,  # SteelBlue
        0x6495ED,  # CornflowerBlue
        0x87CEFA,  # LightSkyBlue
        0x87CEEB,  # SkyBlue
        0xB0E0E6,  # PowderBlue
        0xADD8E6,  # LightBlue
        0xAFEEEE,  # PaleTurquoise
        0x6A5ACD,  # SlateBlue
        0x7B68EE,  # MediumSlateBlue
        0x483D8B,  # DarkSlateBlue
        0x191970,  # MidnightBlue
    ]
    print_color_blocks(SAMPLE_COLORS)
