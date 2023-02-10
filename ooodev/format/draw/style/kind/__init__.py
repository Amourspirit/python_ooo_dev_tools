from __future__ import annotations

from enum import Enum


class DrawStyleFamilyKind(Enum):
    """Default Draw Style Families"""

    DEFAULT = "default"
    CELL = "cell"
    GRAPHICS = "graphics"
    TABLE = "table"

    def __str__(self) -> str:
        return self.value
