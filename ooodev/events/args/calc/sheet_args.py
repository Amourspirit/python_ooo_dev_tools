# coding: utf-8
from typing import Any
from ..event_args import EventArgs

class SheetArgs(EventArgs):
    def __init__(self, source: Any, sheet_arg: str | int) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            sheet_arg (str): Sheet index or name
        """
        super().__init__(source)
        self._sheet_arg = sheet_arg

    @property
    def sheet_arg(self) -> str:
        """
        Gets/Sets the dispatch cmd of the event
        """
        return self._sheet_arg

    @sheet_arg.setter
    def sheet_arg(self, value: str):
        self._sheet_arg = value