from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.calc.calc_cell import CalcCell
else:
    CalcSheet = Any


class CalcCellPropPartial:
    """A partial class for Calc Cell."""

    def __init__(self, obj: CalcCell) -> None:
        self.__calc_cell = obj

    @property
    def calc_cell(self) -> CalcCell:
        """Chart Document."""
        return self.__calc_cell
