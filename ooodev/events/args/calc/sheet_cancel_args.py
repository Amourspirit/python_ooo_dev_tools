# coding: utf-8
from __future__ import annotations
from typing import Any
from .sheet_args import AbstractSheetArgs
from ..cancel_event_args import AbstractCancelEventArgs


class AbstractSheetCancelArgs(AbstractCancelEventArgs, AbstractSheetArgs):
    __slots__ = ()

    def __init__(self, source: Any, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source=source, cancel=cancel)


class SheetCancelArgs(AbstractSheetCancelArgs):
    """
    Sheet Cancel Event Args
    """

    __slots__ = ("source", "_event_name", "event_data", "name", "index", "doc", "sheet", "cancel", "_event_source")

    @staticmethod
    def from_args(args: SheetCancelArgs) -> SheetCancelArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (SheetCancelArgs): Existing Instance

        Returns:
            SheetCancelArgs: args
        """
        eargs = SheetCancelArgs(source=args.source)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.doc = args.doc
        eargs.event_data = args.event_data
        eargs.index = args.index
        eargs.name = args.name
        eargs.sheet = args.sheet
        eargs.cancel = args.cancel
        return eargs
