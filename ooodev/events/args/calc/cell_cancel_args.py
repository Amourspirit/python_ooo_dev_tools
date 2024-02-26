from __future__ import annotations
from typing import Any
from ooodev.events.args.calc.cell_args import AbstractCellArgs
from ooodev.events.args.cancel_event_args import AbstractCancelEventArgs


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
        self.handled = False


class CellCancelArgs(AbstractCellCancelArgs):
    """
    Sheet Cancel Event Args
    """

    __slots__ = (
        "source",
        "_event_name",
        "event_data",
        "sheet",
        "cells",
        "cancel",
        "handled",
        "_event_source",
        "_kv_data",
    )

    @staticmethod
    def from_args(args: AbstractCellCancelArgs) -> CellCancelArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (AbstractCellCancelArgs): Existing Instance

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
        eargs.handled = args.handled
        return eargs
