from __future__ import annotations

from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ..calc_sheet import CalcSheet
else:
    CalcSheet = Any


class CalcSheetPropPartial:
    """A partial class for Calc Document."""

    def __init__(self, obj: CalcSheet) -> None:
        self.__calc_sheet = obj

    @property
    def calc_sheet(self) -> CalcSheet:
        """Chart Document."""
        return self.__calc_sheet
