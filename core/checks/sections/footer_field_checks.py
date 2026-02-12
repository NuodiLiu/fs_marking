"""
Footer field validation checks.

Checks for specific field types (e.g., AutoFileName, PageNumber) in footer.
"""

from core.checks.base_check import BaseCheck
from typing import Tuple
from win32com.client import constants


# Common Word field type constants for reference:
# wdFieldFileName = 29  (AutoFileName / FILENAME)
# wdFieldPage     = 33  (Page number / PAGE)

class FooterHasFieldCheck(BaseCheck):
    """
    Check if a section's footer contains a specific field type.
    
    Validates that the footer has a field of the given type (e.g., FileName or Page).
    Works on both Windows and Mac-created documents since we check via COM field types.
    """
    
    def __init__(self, section_index: int = 1, field_type: int = 29, field_name: str = ""):
        """
        Args:
            section_index: Which section to check (1-based)
            field_type: Word field type constant (29=FileName, 33=Page, etc.)
            field_name: Human-readable name for error messages (auto-detected if empty)
        """
        self.section_index = section_index
        self.field_type = field_type
        self.field_name = field_name or self._get_field_name(field_type)
    
    @staticmethod
    def _get_field_name(field_type: int) -> str:
        """Map common field types to readable names."""
        names = {
            29: "AutoFileName",
            33: "PageNumber",
        }
        return names.get(field_type, f"FieldType({field_type})")
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        Check if the footer contains the expected field type.
        
        Returns:
            (True, message) if field found
            (False, error_message) if not found
        """
        try:
            if doc.Sections.Count < self.section_index:
                return False, f"❌ Document has only {doc.Sections.Count} section(s), cannot check section {self.section_index}"
            
            section = doc.Sections(self.section_index)
            footer = section.Footers(constants.wdHeaderFooterPrimary)
            
            # Check if footer has any content at all
            footer_text = footer.Range.Text.strip()
            if not footer_text:
                return False, f"❌ Section {self.section_index} footer is empty, no {self.field_name} field"
            
            # Search for the field type in footer fields
            for field in footer.Range.Fields:
                if field.Type == self.field_type:
                    return True, f"✅ Footer has {self.field_name} field"
            
            # Field not found in Fields collection
            return False, f"❌ Footer does not contain {self.field_name} field (found text: '{footer_text[:50]}')"
        
        except Exception as e:
            return False, f"❌ Error checking footer field: {str(e)}"
