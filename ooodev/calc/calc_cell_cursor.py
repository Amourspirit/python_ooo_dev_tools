from __future__ import annotations
from typing import TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from com.sun.star.sheet import XSheetCellCursor
    from .calc_sheet import CalcSheet
else:
    XSheetCellCursor = object

from ooodev.office import calc as mCalc
from ooodev.adapter.sheet.sheet_cell_cursor_comp import SheetCellCursorComp
from . import calc_cell_range as mCalcCellRange


class CalcCellCursor(SheetCellCursorComp):
    def __init__(self, owner: CalcSheet, cursor: XSheetCellCursor) -> None:
        self.__owner = owner
        super().__init__(cursor)  # type: ignore

    def find_used_cursor(self) -> mCalcCellRange.CalcCellRange:
        """
        Find used cursor

        Raises:
            MissingInterfaceError: if unable to find interface

        Returns:
            CalcCellRange: Cell range
        """
        found = mCalc.Calc.find_used_cursor(self.component)
        return mCalcCellRange.CalcCellRange(self.calc_sheet, found)

    # region Properties
    @property
    def calc_sheet(self) -> CalcSheet:
        """Sheet that owns this cell."""
        return self.__owner

    # endregion Properties
