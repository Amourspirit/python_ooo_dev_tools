# coding: utf-8
from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ..event_args import EventArgs

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheet

class CellArgs(EventArgs):
    def __init__(self, source: Any) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
        """
        super().__init__(source)

    @property
    def sheet(self) -> XSpreadsheet | None:
        """
        Gets/Sets spreadsheet of the event
        """
        try:
            return self._sheet
        except AttributeError:
            return None

    @sheet.setter
    def sheet(self, value: XSpreadsheet):
        self._sheet = value
    
    @property
    def cells(self) -> Any:
        """
        Gets/Sets the cells for the event.
        
        Depending on the event can be any cell value such as a cell name, range, XCell, XCellRange etc.
        """
        return self._cells

    @cells.setter
    def cells(self, value: Any):
        self._cells = value