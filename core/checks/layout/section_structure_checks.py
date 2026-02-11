# core/checks/layout/section_structure_checks.py

from core.checks.base_check import BaseCheck
from typing import Tuple


class SectionCountCheck(BaseCheck):
    """检查文档section数量（通用检查）"""
    
    def __init__(self, count, mode="min"):
        """
        Args:
            count: 目标数量
            mode: 比较模式
                - "min": 至少有count个section（count不足返回False）
                - "max": 最多有count个section（超过返回False）
                - "exact": 恰好有count个section（不等返回False）
                - "less_than": 少于count个section（少于返回True，用于扣分）
        """
        self.count = count
        self.mode = mode
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        检查section数量是否满足条件
        
        Returns:
            (True/False, message) 根据mode返回不同的判断结果
        """
        try:
            actual_count = doc.Sections.Count
            
            if self.mode == "min":
                if actual_count < self.count:
                    return False, f"❌ Document has {actual_count} section(s), expected at least {self.count}"
                return True, f"✅ Document has {actual_count} section(s)"
                
            elif self.mode == "max":
                if actual_count > self.count:
                    return False, f"❌ Document has {actual_count} section(s), expected at most {self.count}"
                return True, f"✅ Document has {actual_count} section(s)"
                
            elif self.mode == "exact":
                if actual_count != self.count:
                    return False, f"❌ Document has {actual_count} section(s), expected exactly {self.count}"
                return True, f"✅ Document has exactly {actual_count} section(s)"
                
            elif self.mode == "less_than":
                if actual_count < self.count:
                    return True, f"❌ Document has only {actual_count} section(s), expected at least {self.count}"
                return False, f"✅ Document has {actual_count} section(s)"
                
            else:
                return False, f"❌ Invalid mode '{self.mode}'"
            
        except Exception as e:
            return False, f"❌ Error checking section count: {str(e)}"

