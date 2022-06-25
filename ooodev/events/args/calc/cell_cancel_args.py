# coding: utf-8
from __future__ import annotations
from typing import Any
from .cell_args import CellArgs
from ..cancel_event_args import CancelEventArgs

class CellCancelArgs(CancelEventArgs, CellArgs):
    """
    Sheet Cancel Event Args
    """
    def __init__(self, source: Any, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source=source, cancel=cancel)        
