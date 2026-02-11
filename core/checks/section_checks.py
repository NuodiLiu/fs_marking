"""
Atomic check units for section-level header and footer validation.

These are granular checks for specific section properties.
"""

from core.checks.base_check import BaseCheck
from core.utils.section_utils import get_section_for_page, is_border_enabled_for_page
from typing import Tuple
from win32com.client import constants


class SectionHasHeaderCheck(BaseCheck):
    """
    Check if a specific section has header content.
    
    MOST ATOMIC: Checks only ONE section's header existence.
    """
    
    def __init__(self, section_index: int, should_exist: bool = True):
        """
        Args:
            section_index: Which section to check (1-based)
            should_exist: True if header should exist, False if it should not
        """
        super().__init__(section_index=section_index, should_exist=should_exist)
        self.section_index = section_index
        self.should_exist = should_exist
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        Check if the section's header meets expectations.
        
        Returns:
            (True, "") if header status matches expectation
            (False, error_message) otherwise
        """
        try:
            # Check if section exists
            if doc.Sections.Count < self.section_index:
                return False, f"Document has only {doc.Sections.Count} section(s), cannot check section {self.section_index}"
            
            section = doc.Sections(self.section_index)
            header = section.Headers(constants.wdHeaderFooterPrimary)
            
            # Check if header has content
            has_content = bool(header.Range.Text.strip())
            
            if self.should_exist and not has_content:
                return False, f"Section {self.section_index} should have header but doesn't"
            elif not self.should_exist and has_content:
                return False, f"Section {self.section_index} should not have header but has: '{header.Range.Text.strip()[:50]}...'"
            
            return True, ""
            
        except Exception as e:
            return False, f"Error checking header in section {self.section_index}: {str(e)}"


class SectionHasFooterCheck(BaseCheck):
    """
    Check if a specific section has footer content.
    
    MOST ATOMIC: Checks only ONE section's footer existence.
    """
    
    def __init__(self, section_index: int, should_exist: bool = True):
        """
        Args:
            section_index: Which section to check (1-based)
            should_exist: True if footer should exist, False if it should not
        """
        super().__init__(section_index=section_index, should_exist=should_exist)
        self.section_index = section_index
        self.should_exist = should_exist
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        Check if the section's footer meets expectations.
        
        Returns:
            (True, "") if footer status matches expectation
            (False, error_message) otherwise
        """
        try:
            # Check if section exists
            if doc.Sections.Count < self.section_index:
                return False, f"Document has only {doc.Sections.Count} section(s), cannot check section {self.section_index}"
            
            section = doc.Sections(self.section_index)
            footer = section.Footers(constants.wdHeaderFooterPrimary)
            
            # Check if footer has content
            has_content = bool(footer.Range.Text.strip())
            
            if self.should_exist and not has_content:
                return False, f"Section {self.section_index} should have footer but doesn't"
            elif not self.should_exist and has_content:
                return False, f"Section {self.section_index} should not have footer but has: '{footer.Range.Text.strip()[:50]}...'"
            
            return True, ""
            
        except Exception as e:
            return False, f"Error checking footer in section {self.section_index}: {str(e)}"


class HeaderContentCheck(BaseCheck):
    """
    Check if header contains specific content or pattern.
    
    ATOMIC: Checks header content in ONE section.
    """
    
    def __init__(self, section_index: int, min_length: int = 1, pattern: str = None):
        """
        Args:
            section_index: Which section to check (1-based)
            min_length: Minimum length of header content (default 1)
            pattern: Optional regex pattern to match (if None, just checks length)
        """
        super().__init__(section_index=section_index, min_length=min_length, pattern=pattern)
        self.section_index = section_index
        self.min_length = min_length
        self.pattern = pattern
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        Check if header content meets requirements.
        
        Returns:
            (True, "") if header content is valid
            (False, error_message) otherwise
        """
        try:
            if doc.Sections.Count < self.section_index:
                return False, f"Document has only {doc.Sections.Count} section(s)"
            
            section = doc.Sections(self.section_index)
            header = section.Headers(constants.wdHeaderFooterPrimary)
            content = header.Range.Text.strip()
            
            # Check minimum length
            if len(content) < self.min_length:
                return False, f"Section {self.section_index} header is too short (length {len(content)}, minimum {self.min_length})"
            
            # Check pattern if provided
            if self.pattern:
                import re
                if not re.search(self.pattern, content):
                    return False, f"Section {self.section_index} header does not match pattern '{self.pattern}'"
            
            return True, ""
            
        except Exception as e:
            return False, f"Error checking header content in section {self.section_index}: {str(e)}"


class FooterContentCheck(BaseCheck):
    """
    Check if footer contains specific content or pattern.
    
    ATOMIC: Checks footer content in ONE section.
    """
    
    def __init__(self, section_index: int, min_length: int = 1, pattern: str = None):
        """
        Args:
            section_index: Which section to check (1-based)
            min_length: Minimum length of footer content (default 1)
            pattern: Optional regex pattern to match
        """
        super().__init__(section_index=section_index, min_length=min_length, pattern=pattern)
        self.section_index = section_index
        self.min_length = min_length
        self.pattern = pattern
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        Check if footer content meets requirements.
        
        Returns:
            (True, "") if footer content is valid
            (False, error_message) otherwise
        """
        try:
            if doc.Sections.Count < self.section_index:
                return False, f"Document has only {doc.Sections.Count} section(s)"
            
            section = doc.Sections(self.section_index)
            footer = section.Footers(constants.wdHeaderFooterPrimary)
            content = footer.Range.Text.strip()
            
            # Check minimum length
            if len(content) < self.min_length:
                return False, f"Section {self.section_index} footer is too short (length {len(content)}, minimum {self.min_length})"
            
            # Check pattern if provided
            if self.pattern:
                import re
                if not re.search(self.pattern, content):
                    return False, f"Section {self.section_index} footer does not match pattern '{self.pattern}'"
            
            return True, ""
            
        except Exception as e:
            return False, f"Error checking footer content in section {self.section_index}: {str(e)}"


class PageBorderShadowCheck(BaseCheck):
    """
    检查指定页面是否有阴影边框（自动处理section映射）
    
    会自动：
    1. 找到页面所属的section
    2. 判断是否为section第一页
    3. 根据EnableFirstPageInSection/EnableOtherPagesInSection判断该页是否启用边框
    4. 检查Borders.Shadow
    """
    
    def __init__(self, page_num=1, should_have_shadow=True):
        """
        Args:
            page_num: 页码（1-based）
            should_have_shadow: True检查有阴影，False检查无阴影
        """
        super().__init__(page_num=page_num, should_have_shadow=should_have_shadow)
        self.page_num = page_num
        self.should_have_shadow = should_have_shadow
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        检查页面是否有阴影边框
        
        Returns:
            (True, "") if border shadow status matches expectation
            (False, error_message) otherwise
        """
        try:
            # 1. 找到页面所属的section
            section, is_first_page = get_section_for_page(doc, self.page_num)
            
            if section is None:
                return False, f"Cannot find section for page {self.page_num}"
            
            # 2. 判断该页是否启用了边框
            border_enabled = is_border_enabled_for_page(section, is_first_page)
            
            if not border_enabled:
                # 该页未启用边框
                if self.should_have_shadow:
                    return False, f"❌ Page {self.page_num} has no border enabled (cannot have shadow)"
                else:
                    return True, f"✅ Page {self.page_num} has no border (no shadow as expected)"
            
            # 3. 检查是否有阴影
            try:
                has_shadow = section.Borders.Shadow
            except:
                has_shadow = False
            
            # 4. 比较期望值
            if self.should_have_shadow:
                if has_shadow:
                    return True, f"✅ Page {self.page_num} has shadow border"
                else:
                    return False, f"❌ Page {self.page_num} has no shadow border"
            else:
                if not has_shadow:
                    return True, f"✅ Page {self.page_num} has no shadow border"
                else:
                    return False, f"❌ Page {self.page_num} has shadow border (should not have)"
            
        except Exception as e:
            return False, f"❌ Error checking page border: {str(e)}"


class SectionBorderShadowCheck(BaseCheck):
    """
    检查指定section是否有阴影边框（低层检查，直接操作section）
    """
    
    def __init__(self, section_index=1, should_have_shadow=True):
        """
        Args:
            section_index: Section索引（1-based）
            should_have_shadow: True检查有阴影，False检查无阴影
        """
        super().__init__(section_index=section_index, should_have_shadow=should_have_shadow)
        self.section_index = section_index
        self.should_have_shadow = should_have_shadow
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        检查section是否有阴影边框
        
        Returns:
            (True, "") if border shadow status matches expectation
            (False, error_message) otherwise
        """
        try:
            if doc.Sections.Count < self.section_index:
                return False, f"Document has only {doc.Sections.Count} section(s)"
            
            section = doc.Sections(self.section_index)
            
            try:
                has_shadow = section.Borders.Shadow
            except:
                has_shadow = False
            
            if self.should_have_shadow:
                if has_shadow:
                    return True, f"✅ Section {self.section_index} has shadow border"
                else:
                    return False, f"❌ Section {self.section_index} has no shadow border"
            else:
                if not has_shadow:
                    return True, f"✅ Section {self.section_index} has no shadow border"
                else:
                    return False, f"❌ Section {self.section_index} has shadow border (should not have)"
            
        except Exception as e:
            return False, f"❌ Error checking section border: {str(e)}"
