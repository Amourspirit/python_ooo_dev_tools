from __future__ import annotations
from typing import TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetView
    from .calc_doc import CalcDoc

from ooodev.adapter.sheet.spreadsheet_view_comp import SpreadsheetViewComp


class CalcSheetView(SpreadsheetViewComp):
    def __init__(self, owner: CalcDoc, view: XSpreadsheetView) -> None:
        super().__init__(view)  # type: ignore
        self.__owner = owner

    # region Properties
    @property
    def calc_doc(self) -> CalcDoc:
        """
        Returns:
            CalcDoc: Calc doc
        """
        return self.__owner

    # endregion Properties
