from core.rules.style_setting_rules.style_checkpoints import StyleCheckPoint
from core.utils.color_utils import is_blueish
from win32com.client import constants

def get_all_style_checkpoints():
    return [
        # --- Normal ---
        StyleCheckPoint("Normal", "Font should be Arial", lambda f, p: f.Name == "Arial"),
        StyleCheckPoint("Normal", "Font size should be 11pt", lambda f, p: f.Size == 11),
        StyleCheckPoint("Normal", "Line spacing should be single", lambda f, p: p.LineSpacingRule == constants.wdLineSpaceSingle),
        StyleCheckPoint("Normal", "Should be left aligned", lambda f, p: p.Alignment == constants.wdAlignParagraphLeft),

        # --- Heading 1 ---
        StyleCheckPoint("Heading 1", "Font size should be 20pt", lambda f, p: f.Size == 20),
        StyleCheckPoint("Heading 1", "Should be left aligned", lambda f, p: p.Alignment == constants.wdAlignParagraphLeft),
        StyleCheckPoint("Heading 1", "Font color should be blue-ish", lambda f, p: is_blueish(f.Color) == "blue-ish"),
        StyleCheckPoint("Heading 1", "Spacing should be 12pt before and 6pt after ", lambda f, p: p.SpaceBefore == 12 and p.SpaceAfter == 6),

        # --- Heading 2 ---
        StyleCheckPoint("Heading 2", "Should be left aligned", lambda f, p: p.Alignment == constants.wdAlignParagraphLeft),
        StyleCheckPoint("Heading 2", "Should be bold", lambda f, p: f.Bold),
        StyleCheckPoint("Heading 2", "Should be italic", lambda f, p: f.Italic),
    ]
