"""
Registry for all available check units.

This module maintains a central registry of all check types that can be
used in configuration files.
"""

from core.checks.margin_checks import MarginCheck, AllMarginsCheck
from core.checks.section_checks import (
    SectionHasHeaderCheck,
    SectionHasFooterCheck,
    HeaderContentCheck,
    FooterContentCheck,
    PageBorderShadowCheck,
    SectionBorderShadowCheck
)
from core.checks.objects.shape_checks import (
    PageHasShapeCheck,
    ShapeContainsTextCheck,
    ShapeCountCheck
)
from core.checks.objects.table_checks import (
    PageHasTableCheck,
    TableDimensionCheck,
    TableCellNotEmptyCheck,
    TableCellContainsTextCheck,
    TableHasStyleCheck,
    TableCountCheck
)
from core.checks.formatting.style_checks import (
    StyleFontCheck,
    StyleParagraphFormatCheck
)


# Central registry of all available check types
CHECK_REGISTRY = {
    # Margin checks
    "margin": MarginCheck,
    "all_margins": AllMarginsCheck,
    
    # Section header/footer existence checks
    "section_has_header": SectionHasHeaderCheck,
    "section_has_footer": SectionHasFooterCheck,
    "section_no_header": lambda **kwargs: SectionHasHeaderCheck(should_exist=False, **kwargs),
    "section_no_footer": lambda **kwargs: SectionHasFooterCheck(should_exist=False, **kwargs),
    
    # Header/footer content checks
    "header_content": HeaderContentCheck,
    "footer_content": FooterContentCheck,
    
    # Border checks
    "page_border_shadow": PageBorderShadowCheck,
    "section_border_shadow": SectionBorderShadowCheck,
    
    # Shape checks
    "page_has_shape": PageHasShapeCheck,
    "shape_contains_text": ShapeContainsTextCheck,
    "shape_count": ShapeCountCheck,
    
    # Table checks
    "page_has_table": PageHasTableCheck,
    "table_dimension": TableDimensionCheck,
    "table_cell_not_empty": TableCellNotEmptyCheck,
    "table_cell_contains_text": TableCellContainsTextCheck,
    "table_has_style": TableHasStyleCheck,
    "table_count": TableCountCheck,
    
    # Style checks
    "style_font": StyleFontCheck,
    "style_paragraph_format": StyleParagraphFormatCheck,
}


def register_check(name: str, check_class):
    """
    Register a new check type to the registry.
    
    This allows for dynamic extension of available check types.
    
    Args:
        name: The identifier to use in configuration files
        check_class: The check class (must inherit from BaseCheck)
    
    Example:
        >>> register_check("custom_margin", CustomMarginCheck)
        >>> # Now can use in config: ("custom_margin", {"param": value})
    """
    if name in CHECK_REGISTRY:
        raise ValueError(f"Check type '{name}' is already registered")
    
    CHECK_REGISTRY[name] = check_class
    

def get_check(name: str):
    """
    Get a check class from the registry.
    
    Args:
        name: The identifier of the check type
    
    Returns:
        The check class
    
    Raises:
        KeyError: If check type is not found
    """
    if name not in CHECK_REGISTRY:
        available = ", ".join(sorted(CHECK_REGISTRY.keys()))
        raise KeyError(f"Check type '{name}' not found. Available types: {available}")
    
    return CHECK_REGISTRY[name]


def list_available_checks():
    """
    List all available check types.
    
    Returns:
        List of check type names
    """
    return sorted(CHECK_REGISTRY.keys())
