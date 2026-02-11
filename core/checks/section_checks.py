"""
Atomic check units for section-level header and footer validation.

These are granular checks for specific section properties.
"""

from core.checks.base_check import BaseCheck
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
