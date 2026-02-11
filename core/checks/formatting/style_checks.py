# core/checks/formatting/style_checks.py

from core.checks.base_check import BaseCheck
from core.utils.constants_utils import (
    parse_alignment,
    parse_line_spacing_rule,
    get_alignment_name,
    get_line_spacing_rule_name
)
from win32com.client import constants


class StyleFontCheck(BaseCheck):
    """检查样式的字体属性"""
    
    def __init__(self, style_name="Normal", font_name=None, font_size=None, bold=None, italic=None, color=None):
        """
        Args:
            style_name: 样式名称（如"Normal", "Heading 1"）
            font_name: 字体名称（如"Arial"），None表示不检查
            font_size: 字号（点数），None表示不检查
            bold: 是否加粗（True/False），None表示不检查
            italic: 是否斜体（True/False），None表示不检查
            color: 字体颜色检查函数名（如"is_font_color_blueish"），None表示不检查
        """
        super().__init__(style_name=style_name, font_name=font_name, font_size=font_size,
                         bold=bold, italic=italic, color=color)
        self.style_name = style_name
        self.font_name = font_name
        self.font_size = font_size
        self.bold = bold
        self.italic = italic
        self.color = color
    
    def check(self, doc):
        try:
            # 获取样式对象
            try:
                style_obj = doc.Styles(self.style_name)
            except:
                return (False, f"❌ Style '{self.style_name}' not found in document")
            
            font = style_obj.Font
            errors = []
            
            # 检查字体名称
            if self.font_name is not None:
                try:
                    actual_font = font.Name
                    if actual_font != self.font_name:
                        errors.append(f"font name is '{actual_font}', expected '{self.font_name}'")
                except Exception as e:
                    errors.append(f"cannot read font name: {e}")
            
            # 检查字号
            if self.font_size is not None:
                try:
                    actual_size = font.Size
                    if abs(actual_size - self.font_size) > 0.5:  # 允许0.5的误差
                        errors.append(f"font size is {actual_size}pt, expected {self.font_size}pt")
                except Exception as e:
                    errors.append(f"cannot read font size: {e}")
            
            # 检查加粗
            if self.bold is not None:
                try:
                    actual_bold = bool(font.Bold)
                    if actual_bold != self.bold:
                        errors.append(f"bold is {actual_bold}, expected {self.bold}")
                except Exception as e:
                    errors.append(f"cannot read bold: {e}")
            
            # 检查斜体
            if self.italic is not None:
                try:
                    actual_italic = bool(font.Italic)
                    if actual_italic != self.italic:
                        errors.append(f"italic is {actual_italic}, expected {self.italic}")
                except Exception as e:
                    errors.append(f"cannot read italic: {e}")
            
            # 检查颜色（如果提供了颜色检查函数）
            if self.color is not None:
                try:
                    from core.utils.color_utils import is_font_color_blueish
                    if self.color == "is_font_color_blueish":
                        if not is_font_color_blueish(font):
                            errors.append("font color is not blue-ish")
                except Exception as e:
                    errors.append(f"cannot check font color: {e}")
            
            if errors:
                return (False, f"❌ Style '{self.style_name}': {', '.join(errors)}")
            
            return (True, f"✅ Style '{self.style_name}' font properties correct")
        
        except Exception as e:
            return (False, f"❌ Error checking style '{self.style_name}' font: {str(e)}")


class StyleParagraphFormatCheck(BaseCheck):
    """检查样式的段落格式属性"""
    
    def __init__(self, style_name="Normal", alignment=None, line_spacing=None, line_spacing_rule=None,
                 space_before=None, space_after=None, first_line_indent=None, left_indent=None, right_indent=None):
        """
        Args:
            style_name: 样式名称（如"Normal", "Heading 1"）
            alignment: 对齐方式（"left", "center", "right", "justify"或常量值），None表示不检查
            line_spacing: 行距值（点数或倍数，取决于line_spacing_rule），None表示不检查
            line_spacing_rule: 行距规则（"single", "1.5", "double", "multiple", "at_least", "exactly"或常量值），None表示不检查
            space_before: 段前间距（点数），None表示不检查
            space_after: 段后间距（点数），None表示不检查
            first_line_indent: 首行缩进（点数），None表示不检查
            left_indent: 左缩进（点数），None表示不检查
            right_indent: 右缩进（点数），None表示不检查
        """
        super().__init__(style_name=style_name, alignment=alignment, line_spacing=line_spacing,
                         line_spacing_rule=line_spacing_rule, space_before=space_before, space_after=space_after,
                         first_line_indent=first_line_indent, left_indent=left_indent, right_indent=right_indent)
        self.style_name = style_name
        
        # 使用utils函数转换字符串到常量
        self.alignment = parse_alignment(alignment)
        self.line_spacing = line_spacing
        self.line_spacing_rule = parse_line_spacing_rule(line_spacing_rule)
        
        self.space_before = space_before
        self.space_after = space_after
        self.first_line_indent = first_line_indent
        self.left_indent = left_indent
        self.right_indent = right_indent
    
    def check(self, doc):
        try:
            # 获取样式对象
            try:
                style_obj = doc.Styles(self.style_name)
            except:
                return (False, f"❌ Style '{self.style_name}' not found in document")
            
            para_format = style_obj.ParagraphFormat
            errors = []
            
            # 检查对齐方式
            if self.alignment is not None:
                try:
                    actual_alignment = para_format.Alignment
                    if actual_alignment != self.alignment:
                        actual_name = get_alignment_name(actual_alignment)
                        expected_name = get_alignment_name(self.alignment)
                        errors.append(f"alignment is {actual_name}, expected {expected_name}")
                except Exception as e:
                    errors.append(f"cannot read alignment: {e}")
            
            # 检查行距规则
            if self.line_spacing_rule is not None:
                try:
                    actual_rule = para_format.LineSpacingRule
                    if actual_rule != self.line_spacing_rule:
                        actual_name = get_line_spacing_rule_name(actual_rule)
                        expected_name = get_line_spacing_rule_name(self.line_spacing_rule)
                        errors.append(f"line spacing rule is {actual_name}, expected {expected_name}")
                except Exception as e:
                    errors.append(f"cannot read line spacing rule: {e}")
                            constants.wdLineSpaceExactly: "Exactly"
                        }
                        actual_name = rule_names.get(actual_rule, f"Unknown({actual_rule})")
                        expected_name = rule_names.get(self.line_spacing_rule, f"Unknown({self.line_spacing_rule})")
                        errors.append(f"line spacing rule is {actual_name}, expected {expected_name}")
                except Exception as e:
                    errors.append(f"cannot read line spacing rule: {e}")
            
            # 检查行距值
            if self.line_spacing is not None:
                try:
                    actual_spacing = para_format.LineSpacing
                    if abs(actual_spacing - self.line_spacing) > 0.5:
                        errors.append(f"line spacing is {actual_spacing}, expected {self.line_spacing}")
                except Exception as e:
                    errors.append(f"cannot read line spacing: {e}")
            
            # 检查段前间距
            if self.space_before is not None:
                try:
                    actual_space = para_format.SpaceBefore
                    if abs(actual_space - self.space_before) > 0.5:
                        errors.append(f"space before is {actual_space}pt, expected {self.space_before}pt")
                except Exception as e:
                    errors.append(f"cannot read space before: {e}")
            
            # 检查段后间距
            if self.space_after is not None:
                try:
                    actual_space = para_format.SpaceAfter
                    if abs(actual_space - self.space_after) > 0.5:
                        errors.append(f"space after is {actual_space}pt, expected {self.space_after}pt")
                except Exception as e:
                    errors.append(f"cannot read space after: {e}")
            
            # 检查首行缩进
            if self.first_line_indent is not None:
                try:
                    actual_indent = para_format.FirstLineIndent
                    if abs(actual_indent - self.first_line_indent) > 0.5:
                        errors.append(f"first line indent is {actual_indent}pt, expected {self.first_line_indent}pt")
                except Exception as e:
                    errors.append(f"cannot read first line indent: {e}")
            
            # 检查左缩进
            if self.left_indent is not None:
                try:
                    actual_indent = para_format.LeftIndent
                    if abs(actual_indent - self.left_indent) > 0.5:
                        errors.append(f"left indent is {actual_indent}pt, expected {self.left_indent}pt")
                except Exception as e:
                    errors.append(f"cannot read left indent: {e}")
            
            # 检查右缩进
            if self.right_indent is not None:
                try:
                    actual_indent = para_format.RightIndent
                    if abs(actual_indent - self.right_indent) > 0.5:
                        errors.append(f"right indent is {actual_indent}pt, expected {self.right_indent}pt")
                except Exception as e:
                    errors.append(f"cannot read right indent: {e}")
            
            if errors:
                return (False, f"❌ Style '{self.style_name}': {', '.join(errors)}")
            
            return (True, f"✅ Style '{self.style_name}' paragraph format correct")
        
        except Exception as e:
            return (False, f"❌ Error checking style '{self.style_name}' paragraph format: {str(e)}")
