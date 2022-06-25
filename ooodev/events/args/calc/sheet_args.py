# coding: utf-8
from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ..event_args import EventArgs

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetDocument
    from com.sun.star.sheet import XSpreadsheet

class SheetArgs(EventArgs):
    def __init__(self, source: Any) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
        """
        super().__init__(source)

    @property
    def index(self) -> int | None:
        """
        Gets/Sets the index of the event
        """
        try:
            return self._index
        except AttributeError:
            return None

    @index.setter
    def index(self, value: int):
        self._index = value
    
    @property
    def name(self) -> str | None:
        """
        Gets/Sets name of the event
        """
        try:
            return self._name
        except AttributeError:
            return None

    @name.setter
    def name(self, value: str):
        self._name = value
    
    
    @property
    def doc(self) -> XSpreadsheetDocument | None:
        """
        Gets/Sets document of the event
        """
        try:
            return self._doc
        except AttributeError:
            return None

    @doc.setter
    def doc(self, value: XSpreadsheetDocument):
        self._doc = value
    
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