from math import isclose
from core.utils.utils import cm_to_points

def check_paragraph_indent(paragraph, indent_type: str, expected_indent: float,  use_cm: bool = True) -> bool:
    """
    检查段落缩进类型是否与预期一致。

    :param paragraph: Word 段落对象
    :param expected_indent_cm: 缩进值（单位 cm）
    :param indent_type: 缩进类型，可选值：hanging / normal / none / firstline
    :param use_cm: 是否使用 cm 作为单位
    :return: 是否匹配指定缩进
    """
    indent = cm_to_points(expected_indent) if use_cm else expected_indent
    left = paragraph.LeftIndent
    first = paragraph.FirstLineIndent

    if indent_type == "hanging":
        return isclose(left, indent, abs_tol=0.1) and isclose(first, -indent, abs_tol=0.1)
    elif indent_type == "normal":
        return isclose(left, indent, abs_tol=0.1) and isclose(first, indent, abs_tol=0.1)
    elif indent_type == "none":
        return isclose(left, 0, abs_tol=0.1) and isclose(first, 0, abs_tol=0.1)
    elif indent_type == "firstline":
        return isclose(left, 0, abs_tol=0.1) and isclose(first, indent, abs_tol=0.1)
    else:
        raise ValueError(f"Unsupported indent_type: {indent_type}")
