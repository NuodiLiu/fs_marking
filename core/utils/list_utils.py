from win32com.client import constants
from typing import List, Any
from core.utils.utils import cm_to_points
import re
from math import isclose

def is_multilevel(paragraphs) -> bool:
    """判断段落中是否有 Word 原生多级列表"""
    for para in paragraphs:
        lf = para.Range.ListFormat
        if lf.ListType == constants.wdListOutlineNumbering:  # multilevel list
            return True
    return False

def is_list_format_matching(paragraphs: List[Any], regex: str, level: int = 1) -> bool:
    """
    判断指定级别的列表编号是否匹配某个正则模式。
    
    :param paragraphs: Word 段落列表（win32com）
    :param regex: 要匹配的编号正则，比如 r"^[A-Z]\.?$" 匹配 A. / B.
    :param level: 要检查的列表层级（默认为一级）
    :return: 如果有任意段落匹配则返回 True
    """
    pattern = re.compile(regex)
    for para in paragraphs:
        lf = para.Range.ListFormat
        if lf.ListLevelNumber == level and lf.ListType != 0:
            list_string = lf.ListString.strip()
            if pattern.fullmatch(list_string):
                return True
    return False

def check_level_indent_exact(
    paragraphs: List[Any],
    level: int,
    expected_align: float,
    expected_indent: float,
    use_cm: bool = True
) -> bool:
    """
    精准检查所有某个 level 的段落是否严格匹配指定缩进设置。

    :param paragraphs: Word 段落列表（win32com）
    :param level: 要检查的 ListLevelNumber（1-based）
    :param expected_align_cm: 段落左对齐位置（对应 LeftIndent）
    :param expected_indent_cm: 段落首行缩进位置（对应 FirstLineIndent）
    :param use_cm: 是否以 cm 为单位，默认 True（否则视为 pt）
    :return: True 表示该级别的段落缩进全部符合要求
    """
    align = cm_to_points(expected_align) if use_cm else expected_align
    indent = cm_to_points(expected_indent) if use_cm else expected_indent

    for para in paragraphs:
        lf = para.Range.ListFormat
        if lf.ListLevelNumber == level:
            left_indent = para.LeftIndent
            first_line_indent = para.FirstLineIndent

            if not (isclose(left_indent, align, abs_tol=0.1) and
                    isclose(first_line_indent, indent, abs_tol=0.1)):
                return False
    return True