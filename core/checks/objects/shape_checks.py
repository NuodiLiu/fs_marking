# core/checks/objects/shape_checks.py

from core.checks.base_check import BaseCheck
from core.utils.utils import fuzzy_match
from win32com.client import constants


class PageHasShapeCheck(BaseCheck):
    """检查指定页面是否有shape"""
    
    def __init__(self, page_num=1):
        """
        Args:
            page_num: 页码（1-based）
        """
        self.page_num = page_num
    
    def check(self, doc):
        try:
            for shape in doc.Shapes:
                shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                if shape_page == self.page_num:
                    return (True, f"✅ Page {self.page_num} has shape")
            
            return (False, f"❌ Page {self.page_num} has no shape")
        
        except Exception as e:
            return (False, f"❌ Error checking shapes: {str(e)}")


class ShapeContainsTextCheck(BaseCheck):
    """检查指定页面的shape是否包含指定文字"""
    
    def __init__(self, page_num=1, expected_text="", fuzzy=True, threshold=0.8):
        """
        Args:
            page_num: 页码（1-based）
            expected_text: 期望的文字内容
            fuzzy: 是否使用模糊匹配
            threshold: 模糊匹配阈值（0-1）
        """
        self.page_num = page_num
        self.expected_text = expected_text
        self.fuzzy = fuzzy
        self.threshold = threshold
    
    def check(self, doc):
        try:
            for shape in doc.Shapes:
                shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                if shape_page != self.page_num:
                    continue
                
                # 尝试获取shape中的文字
                try:
                    if shape.TextFrame.HasText:
                        actual_text = shape.TextFrame.TextRange.Text.strip()
                        
                        if self.fuzzy:
                            if fuzzy_match(self.expected_text, actual_text, self.threshold):
                                return (True, f"✅ Shape contains text matching '{self.expected_text}'")
                        else:
                            if self.expected_text in actual_text:
                                return (True, f"✅ Shape contains text '{self.expected_text}'")
                except:
                    # 某些shape类型可能没有TextFrame
                    continue
            
            return (False, f"❌ No shape on page {self.page_num} contains text '{self.expected_text}'")
        
        except Exception as e:
            return (False, f"❌ Error checking shape text: {str(e)}")


class ShapeCountCheck(BaseCheck):
    """检查指定页面的shape数量"""
    
    def __init__(self, page_num=1, expected_count=1, mode="exact"):
        """
        Args:
            page_num: 页码（1-based）
            expected_count: 期望的shape数量
            mode: "exact"（精确）, "min"（至少）, "max"（最多）
        """
        self.page_num = page_num
        self.expected_count = expected_count
        self.mode = mode
    
    def check(self, doc):
        try:
            count = 0
            for shape in doc.Shapes:
                shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                if shape_page == self.page_num:
                    count += 1
            
            if self.mode == "exact":
                if count == self.expected_count:
                    return (True, f"✅ Page {self.page_num} has exactly {count} shape(s)")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} shape(s), expected {self.expected_count}")
            
            elif self.mode == "min":
                if count >= self.expected_count:
                    return (True, f"✅ Page {self.page_num} has {count} shape(s) (>= {self.expected_count})")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} shape(s), expected at least {self.expected_count}")
            
            elif self.mode == "max":
                if count <= self.expected_count:
                    return (True, f"✅ Page {self.page_num} has {count} shape(s) (<= {self.expected_count})")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} shape(s), expected at most {self.expected_count}")
            
            return (False, f"❌ Invalid mode: {self.mode}")
        
        except Exception as e:
            return (False, f"❌ Error checking shape count: {str(e)}")
