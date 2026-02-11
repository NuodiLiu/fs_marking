# core/checks/formatting/list_checks.py

from core.checks.base_check import BaseCheck
from typing import Tuple
from win32com.client import constants


class SimpleBulletListCheck(BaseCheck):
    """检查段落是否在simple bullet list中（非multilevel）"""
    
    def __init__(self, text_start, align_at_left=None, exclude_bullets=None, 
                 fuzzy=True, threshold=0.8):
        """
        Args:
            text_start: 段落开头的文字（用于定位段落）
            align_at_left: bullet符号对齐位置（cm），None表示不检查
            exclude_bullets: 禁止使用的bullet符号列表（如["\u2022", "\u25E6"]）
            fuzzy: 是否使用模糊匹配
            threshold: 模糊匹配阈值
        """
        self.text_start = text_start
        self.align_at_left = align_at_left
        self.exclude_bullets = exclude_bullets or []
        self.fuzzy = fuzzy
        self.threshold = threshold
    
    def check(self, doc) -> Tuple[bool, str]:
        """查找段落并检查是否在simple bullet list中"""
        from core.utils.utils import fuzzy_match, cm_to_points
        
        # 1. 查找段落
        target_para = None
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if not text:
                continue
            
            # 取段落开头部分，长度略大于text_start
            match_length = int(len(self.text_start) * 1.5)
            text_head = text[:match_length]
            
            if self.fuzzy:
                if fuzzy_match(self.text_start, text_head, self.threshold):
                    target_para = para
                    break
            else:
                if text.startswith(self.text_start):
                    target_para = para
                    break
        
        if target_para is None:
            return False, f"❌ Cannot find paragraph starting with '{self.text_start[:30]}...'"
        
        # 2. 检查列表格式
        lf = target_para.Range.ListFormat
        errors = []
        
        # 必须在列表中
        if lf.ListType == 0:
            errors.append("not in a list")
            return False, f"❌ '{self.text_start[:20]}...': {', '.join(errors)}"
        
        # 必须是simple list，不能是multilevel
        if lf.ListType == constants.wdListOutlineNumbering:
            errors.append("is multilevel list, expected simple bullet list")
        
        # 检查对齐
        if self.align_at_left is not None:
            expected_align = cm_to_points(self.align_at_left)
            bullet_position = target_para.LeftIndent + target_para.Range.ParagraphFormat.FirstLineIndent
            
            if abs(bullet_position - expected_align) > 2.0:  # 容差2点
                errors.append(f"bullet aligned at {bullet_position/28.35:.2f}cm, expected {self.align_at_left}cm")
        
        # 检查bullet符号
        if self.exclude_bullets:
            bullet_char = lf.ListString.strip()
            if bullet_char in self.exclude_bullets:
                bullet_names = {"\u2022": "dot(•)", "\u25E6": "circle(○)", "\u25CB": "circle(○)"}
                bullet_name = bullet_names.get(bullet_char, repr(bullet_char))
                errors.append(f"uses {bullet_name}, not allowed")
        
        if errors:
            return False, f"❌ '{self.text_start[:20]}...': {', '.join(errors)}"
        
        return True, f"✅ '{self.text_start[:20]}...' is in simple bullet list"


class MultilevelListItemCheck(BaseCheck):
    """检查段落是否在multilevel list的指定层级"""
    
    def __init__(self, text_start, level=1, align_at_left=None, 
                 exclude_bullets=None, fuzzy=True, threshold=0.8):
        """
        Args:
            text_start: 段落开头的文字（用于定位段落）
            level: 要求的列表层级（1, 2, 3...）
            align_at_left: bullet/编号对齐位置（cm），None表示不检查
            exclude_bullets: 禁止使用的符号列表
            fuzzy: 是否使用模糊匹配
            threshold: 模糊匹配阈值
        """
        self.text_start = text_start
        self.level = level
        self.align_at_left = align_at_left
        self.exclude_bullets = exclude_bullets or []
        self.fuzzy = fuzzy
        self.threshold = threshold
    
    def check(self, doc) -> Tuple[bool, str]:
        """查找段落并检查是否在multilevel list指定层级"""
        from core.utils.utils import fuzzy_match, cm_to_points
        
        # 1. 查找段落
        target_para = None
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if not text:
                continue
            
            # 取段落开头部分，长度略大于text_start
            match_length = int(len(self.text_start) * 1.5)
            text_head = text[:match_length]
            
            if self.fuzzy:
                if fuzzy_match(self.text_start, text_head, self.threshold):
                    target_para = para
                    break
            else:
                if text.startswith(self.text_start):
                    target_para = para
                    break
        
        if target_para is None:
            return False, f"❌ Cannot find paragraph starting with '{self.text_start[:30]}...'"
        
        # 2. 检查列表格式
        lf = target_para.Range.ListFormat
        errors = []
        
        # 必须在列表中
        if lf.ListType == 0:
            errors.append("not in a list")
            return False, f"❌ '{self.text_start[:20]}...': {', '.join(errors)}"
        
        # 必须是multilevel list
        if lf.ListType != constants.wdListOutlineNumbering:
            errors.append("not a multilevel list")
        else:
            # 检查层级
            actual_level = lf.ListLevelNumber
            if actual_level != self.level:
                errors.append(f"level {actual_level}, expected level {self.level}")
        
        # 检查对齐
        if self.align_at_left is not None:
            expected_align = cm_to_points(self.align_at_left)
            bullet_position = target_para.LeftIndent + target_para.Range.ParagraphFormat.FirstLineIndent
            
            if abs(bullet_position - expected_align) > 2.0:
                errors.append(f"aligned at {bullet_position/28.35:.2f}cm, expected {self.align_at_left}cm")
        
        # 检查符号
        if self.exclude_bullets:
            bullet_char = lf.ListString.strip()
            if bullet_char in self.exclude_bullets:
                bullet_names = {"\u2022": "dot(•)", "\u25E6": "circle(○)", "\u25CB": "circle(○)"}
                bullet_name = bullet_names.get(bullet_char, repr(bullet_char))
                errors.append(f"uses {bullet_name}, not allowed")
        
        if errors:
            return False, f"❌ '{self.text_start[:20]}...': {', '.join(errors)}"
        
        return True, f"✅ '{self.text_start[:20]}...' is in multilevel list level {self.level}"
