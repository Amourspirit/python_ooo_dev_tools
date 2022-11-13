from dataclasses import dataclass


@dataclass(frozen=True)
class WindowTitle:
    """Window Title Info"""

    title: str
    """Window Title"""
    is_regex: bool = False
    """Determines if title is treated as regular expression."""
    class_name: str = "SALFRAME"
