from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.events.args.event_args import AbstractEvent


if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheet


class AbstractCellArgs(AbstractEvent):
    __slots__ = ()

    def __init__(self, source: Any) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
        """
        super().__init__(source)
        self.sheet = None
        self.cells = None

    sheet: XSpreadsheet | None
    """Gets/Sets spreadsheet of the event"""
    cells: Any
    """
    Gets/Sets the cells for the event.

    Depending on the event can be any cell value such as a cell name, range, XCell, XCellRange etc.
    """


class CellArgs(AbstractCellArgs):
    __slots__ = ("source", "_event_name", "event_data", "sheet", "cells", "_event_source", "_kv_data")

    @staticmethod
    def from_args(args: AbstractCellArgs) -> CellArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (AbstractCellArgs): Existing Instance

        Returns:
            CellArgs: args
        """
        eargs = CellArgs(source=args.source)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.sheet = args.sheet
        eargs.cells = args.cells
        eargs.event_data = args.event_data
        return eargs
