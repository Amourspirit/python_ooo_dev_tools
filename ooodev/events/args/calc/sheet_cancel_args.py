# coding: utf-8
from __future__ import annotations
from typing import Any
from .sheet_args import SheetArgs
from ..cancel_event_args import CancelEventArgs

class SheetCancelArgs(SheetArgs, CancelEventArgs):
    """
    Sheet Cancel Event Args
    """
    def __init__(self, source: Any, sheet_arg: str | int, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            sheet_arg (str): Sheet index or name
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source=source, sheet_arg=sheet_arg)
        self.cancel = cancel
        # self.cancel = cancel
        
