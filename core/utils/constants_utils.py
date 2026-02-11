# core/utils/constants_utils.py

"""
Utilities for converting human-readable strings to Word COM constants.
"""

from win32com.client import constants


# 行距规则映射（字符串到常量）
LINE_SPACING_RULES = {
    "single": constants.wdLineSpaceSingle,
    "1.5": constants.wdLineSpace1pt5,
    "1pt5": constants.wdLineSpace1pt5,
    "double": constants.wdLineSpaceDouble,
    "multiple": constants.wdLineSpaceMultiple,
    "at_least": constants.wdLineSpaceAtLeast,
    "exactly": constants.wdLineSpaceExactly
}

# 行距规则常量到名称（用于错误消息）
LINE_SPACING_RULE_NAMES = {
    constants.wdLineSpaceSingle: "Single",
    constants.wdLineSpace1pt5: "1.5 lines",
    constants.wdLineSpaceDouble: "Double",
    constants.wdLineSpaceMultiple: "Multiple",
    constants.wdLineSpaceAtLeast: "At least",
    constants.wdLineSpaceExactly: "Exactly"
}

# 对齐方式映射（字符串到常量）
ALIGNMENT_MAP = {
    "left": constants.wdAlignParagraphLeft,
    "center": constants.wdAlignParagraphCenter,
    "right": constants.wdAlignParagraphRight,
    "justify": constants.wdAlignParagraphJustify
}

# 对齐方式常量到名称（用于错误消息）
ALIGNMENT_NAMES = {
    constants.wdAlignParagraphLeft: "Left",
    constants.wdAlignParagraphCenter: "Center",
    constants.wdAlignParagraphRight: "Right",
    constants.wdAlignParagraphJustify: "Justify"
}


def parse_alignment(value):
    """
    转换alignment参数为常量
    
    Args:
        value: "left", "center", "right", "justify" 或常量值，None表示不检查
    
    Returns:
        对应的常量值或None
    
    Raises:
        ValueError: 如果是无效的字符串
    """
    if value is None:
        return None
    
    if isinstance(value, str):
        result = ALIGNMENT_MAP.get(value.lower())
        if result is None:
            valid_options = ", ".join(f"'{k}'" for k in ALIGNMENT_MAP.keys())
            raise ValueError(f"Invalid alignment: '{value}'. Valid options: {valid_options}")
        return result
    
    return value


def parse_line_spacing_rule(value):
    """
    转换line_spacing_rule参数为常量
    
    Args:
        value: "single", "1.5", "double", "multiple", "at_least", "exactly" 或常量值，None表示不检查
    
    Returns:
        对应的常量值或None
    
    Raises:
        ValueError: 如果是无效的字符串
    """
    if value is None:
        return None
    
    if isinstance(value, str):
        result = LINE_SPACING_RULES.get(value.lower())
        if result is None:
            valid_options = ", ".join(f"'{k}'" for k in LINE_SPACING_RULES.keys())
            raise ValueError(f"Invalid line_spacing_rule: '{value}'. Valid options: {valid_options}")
        return result
    
    return value


def get_alignment_name(constant_value):
    """
    获取对齐方式常量的人类可读名称
    
    Args:
        constant_value: 对齐方式常量
    
    Returns:
        人类可读的名称，如"Left", "Center"等
    """
    return ALIGNMENT_NAMES.get(constant_value, f"Unknown({constant_value})")


def get_line_spacing_rule_name(constant_value):
    """
    获取行距规则常量的人类可读名称
    
    Args:
        constant_value: 行距规则常量
    
    Returns:
        人类可读的名称，如"Single", "1.5 lines"等
    """
    return LINE_SPACING_RULE_NAMES.get(constant_value, f"Unknown({constant_value})")
