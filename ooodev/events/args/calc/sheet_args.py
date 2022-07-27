# coding: utf-8
from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ..event_args import AbstractEvent

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetDocument
    from com.sun.star.sheet import XSpreadsheet


class AbstractSheetArgs(AbstractEvent):
    __slots__ = ()

    def __init__(self, source: Any) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
        """
        super().__init__(source)
        self.index = None
        self.name = None
        self.doc = None
        self.sheet = None

    index: int | None
    """Gets/Sets the index of the event"""
    name: str | None
    """Gets/Sets name of the event"""
    doc: XSpreadsheetDocument | None
    """Gets/Sets document of the event"""
    sheet: XSpreadsheet | None
    """Gets/Sets spreadsheet of the event"""


class SheetArgs(AbstractSheetArgs):
    __slots__ = ("source", "_event_name", "event_data", "name", "index", "doc", "sheet", "_event_source")

    @staticmethod
    def from_args(args: SheetArgs) -> SheetArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (SheetArgs): Existing Instance

        Returns:
            SheetArgs: args
        """
        eargs = SheetArgs(source=args.source)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.doc = args.doc
        eargs.event_data = args.event_data
        eargs.index = args.index
        eargs.name = args.name
        eargs.sheet = args.sheet
        return eargs
