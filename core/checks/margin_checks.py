"""
Atomic check units for margin validation.

These are the most granular checks for page margins.
"""

from core.checks.base_check import BaseCheck
from core.utils.utils import points_to_cm
from typing import Tuple


class MarginCheck(BaseCheck):
    """
    Check if a single margin (Top/Bottom/Left/Right) matches expected value.
    
    This is the MOST ATOMIC unit for margin checking.
    """
    
    def __init__(self, side: str, expected_cm: float, tolerance: float = 0.01, section_index: int = 1):
        """
        Args:
            side: "Top", "Bottom", "Left", or "Right"
            expected_cm: Expected margin in centimeters
            tolerance: Allowed deviation in cm (default 0.01cm to handle rounding)
            section_index: Which section to check (default 1)
        """
        super().__init__(
            side=side,
            expected_cm=expected_cm,
            tolerance=tolerance,
            section_index=section_index
        )
        self.side = side
        self.expected_cm = expected_cm
        self.tolerance = tolerance
        self.section_index = section_index
        
        # Validate side parameter
        valid_sides = ["Top", "Bottom", "Left", "Right"]
        if self.side not in valid_sides:
            raise ValueError(f"side must be one of {valid_sides}, got '{self.side}'")
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        Check if the specified margin matches the expected value.
        
        Returns:
            (True, "") if margin is correct
            (False, error_message) if margin is incorrect or error occurred
        """
        try:
            # Get the section
            if doc.Sections.Count < self.section_index:
                return False, f"Document has only {doc.Sections.Count} section(s), cannot check section {self.section_index}"
            
            section = doc.Sections(self.section_index)
            page_setup = section.PageSetup
            
            # Get the margin value in points
            # Word uses points (pt), 1 cm ≈ 28.35 pt
            margin_attr = f"{self.side}Margin"
            if not hasattr(page_setup, margin_attr):
                return False, f"Invalid margin side: {self.side}"
            
            actual_pt = getattr(page_setup, margin_attr)
            actual_cm = points_to_cm(actual_pt)
            
            # Check if within tolerance
            difference = abs(actual_cm - self.expected_cm)
            if difference > self.tolerance:
                return False, f"{self.side} margin is {actual_cm:.2f}cm, expected {self.expected_cm}cm"
            
            return True, ""
            
        except Exception as e:
            return False, f"Error checking {self.side} margin: {str(e)}"


class AllMarginsCheck(BaseCheck):
    """
    Check if all four margins match expected value.
    
    This is a convenience check that combines 4 atomic MarginCheck units.
    """
    
    def __init__(self, expected_cm: float, tolerance: float = 0.01, section_index: int = 1):
        """
        Args:
            expected_cm: Expected margin for all sides in centimeters
            tolerance: Allowed deviation in cm
            section_index: Which section to check (default 1)
        """
        super().__init__(
            expected_cm=expected_cm,
            tolerance=tolerance,
            section_index=section_index
        )
        self.expected_cm = expected_cm
        self.tolerance = tolerance
        self.section_index = section_index
        
        # Create 4 atomic checks
        self.checks = [
            MarginCheck("Top", expected_cm, tolerance, section_index),
            MarginCheck("Bottom", expected_cm, tolerance, section_index),
            MarginCheck("Left", expected_cm, tolerance, section_index),
            MarginCheck("Right", expected_cm, tolerance, section_index),
        ]
    
    def check(self, doc) -> Tuple[bool, str]:
        """
        Check if all four margins match the expected value.
        
        Returns:
            (True, "") if all margins are correct
            (False, error_message) with combined errors if any margin is incorrect
        """
        errors = []
        
        for margin_check in self.checks:
            success, error_msg = margin_check.check(doc)
            if not success:
                errors.append(error_msg)
        
        if errors:
            return False, "; ".join(errors)
        
        return True, ""
