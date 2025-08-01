
# core/__init__.py
"""
Core business logic modules for Revenue Watchdog
"""

from .data_parser import DataParser
from .llm_interface import LLMInterface

__all__ = ['DataParser', 'LLMInterface']