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
    """
    
    def __init__(self, **params):
        """
        Initialize check with parameters.
        
        Args:
            **params: Check-specific parameters
        """
        self.params = params
    
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
        return f"{self.__class__.__name__}({self.params})"
