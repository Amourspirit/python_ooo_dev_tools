# coding: utf-8
from __future__ import annotations
from typing import Any
from .cell_args import AbstractCellArgs
from ..cancel_event_args import AbstractCancelEventArgs


class AbstractCellCancelArgs(AbstractCancelEventArgs, AbstractCellArgs):
    __slots__ = ()

    def __init__(self, source: Any, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source=source, cancel=cancel)


class CellCancelArgs(AbstractCellCancelArgs):
    """
    Sheet Cancel Event Args
    """

    __slots__ = ("source", "_event_name", "event_data", "sheet", "cells", "cancel", "_event_source")

    @staticmethod
    def from_args(args: CellCancelArgs) -> CellCancelArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (CellArgs): Existing Instance

        Returns:
            CellArgs: args
        """
        eargs = CellCancelArgs(source=args.source)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.sheet = args.sheet
        eargs.cells = args.cells
        eargs.event_data = args.event_data
        eargs.cancel = args.cancel
        return eargs
