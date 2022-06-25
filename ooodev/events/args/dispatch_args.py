# coding: utf-8
from __future__ import annotations
from typing import Any
from .event_args import EventArgs

class DispatchArgs(EventArgs):
    def __init__(self, source: Any, cmd: str) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cmd (str): Event Dispatch Command
        """
        super().__init__(source)
        self._cmd = cmd

    @property
    def cmd(self) -> str:
        """
        Gets/Sets the dispatch cmd of the event
        """
        return self._cmd

    @cmd.setter
    def cmd(self, value: str):
        self._cmd = value