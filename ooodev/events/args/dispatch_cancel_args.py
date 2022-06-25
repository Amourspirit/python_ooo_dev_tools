# coding: utf-8
from typing import Any
from .dispatch_args import DispatchArgs
from .cancel_event_args import CancelEventArgs

class DispatchCancelArgs(DispatchArgs, CancelEventArgs):
    """
    Dispatch Cancel Args
    """
    def __init__(self, source: Any, cmd: str, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cmd (str): Event Dispatch Command
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source=source, cmd=cmd)
        self.cancel = cancel
