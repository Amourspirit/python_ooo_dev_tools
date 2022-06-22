# coding: utf-8
from typing import Any
from .dispatch_event import DispatchEvent

class DispatchCancelEvent(DispatchEvent):
    """
    Dispatch Cancel Event
    """
    def __init__(self, source: Any, cmd: str, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cmd (str): Event Dispatch Command
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source, cmd)
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