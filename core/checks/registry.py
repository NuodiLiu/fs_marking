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
    FooterContentCheck
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
