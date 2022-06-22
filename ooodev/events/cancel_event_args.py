# coding: utf-8
from typing import Any
from .event_args import EventArgs

class CancelEventArgs(EventArgs):
    """Cancel Event Arguments"""
    def __init__(self, source: Any, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source)
        self._cancel = cancel

    @property
    def cancel(self) -> bool:
        """
        Gets/Sets cancel value
        """
        return self._cancel

    @cancel.setter
    def cancel(self, value: bool):
        self._cancel = value