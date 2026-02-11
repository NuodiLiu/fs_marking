"""
Base class for all atomic check units.
"""

from abc import ABC, abstractmethod
from typing import Tuple, Any


class BaseCheck(ABC):
    """
    Base class for atomic check units.
    
    Each check unit should:
    1. Be as small and focused as possible
    2. Check ONE specific thing
    3. Return clear success/failure status with error messages
    4. Be independently testable
    
    Note: Subclasses should define their own __init__ with explicit parameters
    for better type checking and IDE support.
    """
    
    @abstractmethod
    def check(self, doc) -> Tuple[bool, str]:
        """
        Execute the check on a document.
        
        Args:
            doc: Word document object (win32com)
        
        Returns:
            Tuple[bool, str]: (success, error_message)
                - success: True if check passes, False otherwise
                - error_message: Empty string if success, descriptive error if failed
        """
        pass
    
    def __repr__(self):
        # Get all instance attributes for representation
        attrs = {k: v for k, v in self.__dict__.items() if not k.startswith('_')}
        return f"{self.__class__.__name__}({attrs})"
