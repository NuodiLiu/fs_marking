# core/checks/content/toc_checks.py

from core.checks.base_check import BaseCheck
from win32com.client import constants
from typing import Tuple


class HasTOCCheck(BaseCheck):
    """检查文档是否包含 TOC（目录）"""
    
    def __init__(self, page_num=None):
        """
        Args:
            page_num: 可选，指定TOC应该在哪一页（1-based）
                     None表示不限制页码
        """
        self.page_num = page_num
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        检查文档是否有TOC
        
        Returns:
            (True, "") if TOC exists (and on correct page if specified)
            (False, error_message) otherwise
        """
        try:
            toc_found = False
            toc_page = None
            
            for field in doc.Fields:
                try:
                    if field.Type == constants.wdFieldTOC:
                        toc_found = True
                        # 获取 TOC 所在页码
                        toc_page = field.Result.Information(constants.wdActiveEndPageNumber)
                        
                        # 如果指定了页码，检查是否匹配
                        if self.page_num is not None:
                            if toc_page == self.page_num:
                                return True, f"✅ TOC found on page {self.page_num}"
                        else:
                            return True, f"✅ TOC found on page {toc_page}"
                except:
                    continue
            
            if not toc_found:
                return False, "❌ No TOC found in document"
            
            # TOC存在但不在指定页面
            if self.page_num is not None:
                return False, f"❌ TOC found on page {toc_page}, expected page {self.page_num}"
            
            return False, "❌ No valid TOC found"
            
        except Exception as e:
            return False, f"❌ Error checking TOC: {str(e)}"


class TOCOnSeparatePageCheck(BaseCheck):
    """检查 TOC 是否独占一页（该页只有TOC内容）"""
    
    def __init__(self, page_num=2):
        """
        Args:
            page_num: TOC所在的页码（1-based），默认第2页
        """
        self.page_num = page_num
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        检查TOC是否独占一页
        
        Returns:
            (True, "") if TOC is on a separate page
            (False, error_message) otherwise
        """
        try:
            # 1. 找到指定页面的 TOC field
            toc_field = None
            for field in doc.Fields:
                try:
                    if (field.Type == constants.wdFieldTOC and
                        field.Result.Information(constants.wdActiveEndPageNumber) == self.page_num):
                        toc_field = field
                        break
                except:
                    continue
            
            if toc_field is None:
                return False, f"❌ No TOC found on page {self.page_num}"
            
            toc_range = toc_field.Result
            toc_text = toc_range.Text
            
            # 2. 获取该页面所有内容
            start = end = None
            for para in doc.Paragraphs:
                try:
                    if para.Range.Information(constants.wdActiveEndPageNumber) == self.page_num:
                        if start is None:
                            start = para.Range.Start
                        end = para.Range.End
                except:
                    continue
            
            if start is None:
                return False, f"❌ Cannot locate content on page {self.page_num}"
            
            page_range = doc.Range(start, end)
            page_text = page_range.Text
            
            # 3. 找出 TOC 之前的标题（如 "Contents", "Table of Contents"）
            header_text = ""
            for para in doc.Paragraphs:
                try:
                    if (para.Range.Information(constants.wdActiveEndPageNumber) == self.page_num and
                        para.Range.End <= toc_range.Start):
                        header_text = para.Range.Text
                except:
                    continue
            
            # 4. 从页面内容中移除标题和TOC文本
            cleaned = page_text
            for fragment in (header_text, toc_text):
                if fragment:
                    cleaned = cleaned.replace(fragment, "", 1)
            
            # 5. 检查剩余内容是否为空（允许空白字符）
            if cleaned.strip():
                return False, f"❌ Page {self.page_num} contains other content besides TOC"
            
            return True, f"✅ TOC is on a separate page (page {self.page_num})"
            
        except Exception as e:
            return False, f"❌ Error checking TOC page: {str(e)}"
