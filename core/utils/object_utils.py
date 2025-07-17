from win32com.client import constants # type: ignore

def font_has_effects(font):
    try:
        return (
            font.Bold or
            font.Italic or
            font.Shadow or
            font.Outline or
            font.Emboss or
            (hasattr(font, "Glow") and font.Glow.Radius > 0) or
            (hasattr(font, "ThreeD") and font.ThreeD.Visible)
        )
    except Exception as e:
        print(f"⚠️ Font effect check failed: {e}")
        return False
