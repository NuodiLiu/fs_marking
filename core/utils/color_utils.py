import colorsys

def bgr_to_rgb(color_value: int) -> tuple[int, int, int]:
    r = (color_value >> 16) & 0xFF
    g = (color_value >> 8) & 0xFF
    b = color_value & 0xFF
    return r, g, b

def get_basic_color_category(color_value: int) -> str:
    """返回 'pure black' / 'pure white' / 'red-ish' / 'green-ish' / 'blue-ish' / 'unknown'"""
    r, g, b = bgr_to_rgb(color_value)

    # 1) 直接判断黑白
    if max(r, g, b) < 30:                     # 非常暗 -> 黑
        return "pure black"
    if min(r, g, b) > 225:                   # 非常亮且 R≈G≈B -> 白
        return "pure white"

    # 2) 转 HSV 判断主色调
    h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
    h_deg = h * 360

    if h_deg >= 330 or h_deg < 30:
        return "red-ish"
    elif 90 <= h_deg < 150:
        return "green-ish"
    elif 180 <= h_deg < 270:
        return "blue-ish"
    else:
        return "unknown"