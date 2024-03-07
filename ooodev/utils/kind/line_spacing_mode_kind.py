from __future__ import annotations
from typing import TYPE_CHECKING
from enum import Enum

if TYPE_CHECKING:
    from ooo.dyn.style.line_spacing import LineSpacing


class ModeKind(Enum):
    """Mode Kind for line spacing"""

    # Enum value, mode, default
    SINGLE = (0, 0, 100)  # zero value,
    """Single Line Spacing ``1mm``"""
    LINE_1_15 = (1, 0, 115)  # zero value, value
    """Line Spacing ``1.15mm``"""
    LINE_1_5 = (2, 0, 150)
    """Line Spacing ``1.5mm``"""
    DOUBLE = (3, 0, 200)
    """Double line spacing ``2mm``"""
    PROPORTIONAL = (4, 0, 0)  # PERCENTAGE, No conversion onf height value 98 % = 98 MM100
    """Proportional line spacing"""
    AT_LEAST = (5, 1, 0)  # IN 1/100 MM
    """At least line spacing"""
    LEADING = (6, 2, 0)  # in 1/100 MM
    """Leading Line Spacing"""
    FIXED = (7, 3, 0)  # in 1/100 MM
    """Fixed Line Spacing"""

    def __int__(self) -> int:
        return self.value[2]

    def get_mode(self) -> int:
        return self.value[1]

    def get_enum_val(self) -> int:
        return self.value[0]

    @staticmethod
    def from_uno(ls: LineSpacing) -> ModeKind:
        """Converts UNO ``LineSpacing`` struct to ``ModeKind`` enum."""
        mode = ls.Mode
        val = ls.Height
        if mode == 0:
            if val == 100:
                return ModeKind.SINGLE
            if val == 115:
                return ModeKind.LINE_1_15
            if val == 150:
                return ModeKind.LINE_1_5
            return ModeKind.DOUBLE if val == 200 else ModeKind.PROPORTIONAL
        if mode == 1:
            return ModeKind.AT_LEAST
        if mode == 2:
            return ModeKind.LEADING
        if mode == 3:
            return ModeKind.FIXED
        raise ValueError("Unable to convert uno LineSpacing object to ModeKind Enum")
