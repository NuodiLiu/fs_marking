# core/checks/formatting/paragraph_checks.py

from core.checks.base_check import BaseCheck
from typing import Tuple


class ParagraphFormatCheck(BaseCheck):
    """通过文本内容查找段落并检查其格式"""
    
    def __init__(self, text_start, left_indent=None, right_indent=None, 
                 first_line_indent=None, fuzzy=True, threshold=0.8):
        """
        Args:
            text_start: 段落开头的文字（用于定位段落）
            left_indent: 左缩进（cm），None表示不检查
            right_indent: 右缩进（cm），None表示不检查
            first_line_indent: 首行缩进（cm），None表示不检查
            fuzzy: 是否使用模糊匹配
            threshold: 模糊匹配阈值（0-1）
        """
        self.text_start = text_start
        self.left_indent = left_indent
        self.right_indent = right_indent
        self.first_line_indent = first_line_indent
        self.fuzzy = fuzzy
        self.threshold = threshold
    
    def check(self, doc) -> Tuple[bool, str]:
        """查找段落并检查格式"""
        from core.utils.utils import fuzzy_match, cm_to_points
        
        # 1. 查找段落
        target_para = None
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if not text:
                continue
            
            # 取段落开头部分进行匹配（更多字符以提高匹配准确性）
            match_length = min(len(text), len(self.text_start) + 20)
            text_head = text[:match_length]
            
            if self.fuzzy:
                # 使用模糊匹配，比较text_head和text_start
                if fuzzy_match(self.text_start, text_head, self.threshold):
                    target_para = para
                    break
            else:
                if text.startswith(self.text_start):
                    target_para = para
                    break
        
        if target_para is None:
            # 调试信息：显示前3个非空段落的开头
            debug_paras = []
            count = 0
            for para in doc.Paragraphs:
                text = para.Range.Text.strip()
                if text and count < 3:
                    debug_paras.append(text[:50])
                    count += 1
            
            debug_info = "; ".join(debug_paras) if debug_paras else "No paragraphs found"
            return False, f"❌ Cannot find paragraph starting with '{self.text_start[:30]}...'. First 3 paragraphs: {debug_info}"
        
        # 2. 检查格式
        para_format = target_para.Range.ParagraphFormat
        errors = []
        
        if self.left_indent is not None:
            expected = cm_to_points(self.left_indent)
            actual = para_format.LeftIndent
            if abs(actual - expected) > 1.0:  # 容差1点
                errors.append(f"left indent is {actual/28.35:.2f}cm, expected {self.left_indent}cm")
        
        if self.right_indent is not None:
            expected = cm_to_points(self.right_indent)
            actual = para_format.RightIndent
            if abs(actual - expected) > 1.0:
                errors.append(f"right indent is {actual/28.35:.2f}cm, expected {self.right_indent}cm")
        
        if self.first_line_indent is not None:
            expected = cm_to_points(self.first_line_indent)
            actual = para_format.FirstLineIndent
            if abs(actual - expected) > 1.0:
                errors.append(f"first line indent is {actual/28.35:.2f}cm, expected {self.first_line_indent}cm")
        
        if errors:
            return False, f"❌ Paragraph indent: {', '.join(errors)}"
        
        return True, f"✅ Paragraph indent correct"
