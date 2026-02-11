"""
Atomic check units for document validation.

This module provides the smallest, reusable check units that can be composed
into complex rules through configuration.
"""

from core.checks.base_check import BaseCheck
from core.checks.registry import CHECK_REGISTRY, register_check

__all__ = ['BaseCheck', 'CHECK_REGISTRY', 'register_check']
