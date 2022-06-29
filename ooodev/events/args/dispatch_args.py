# coding: utf-8
from __future__ import annotations
from typing import Any
from .event_args import AbstractEvent


class AbstractDispacthArgs(AbstractEvent):
    # https://stackoverflow.com/questions/472000/usage-of-slots
    __slots__ = ()

    def __init__(self, source: Any, cmd: str) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cmd (str): Event Dispatch Command
        """
        super().__init__(source)
        self.cmd = cmd

    cmd: str
    """Gets/Sets the dispatch cmd of the event"""


class DispatchArgs(AbstractDispacthArgs):
    __slots__ = ("source", "_event_name", "event_data", "cmd")

    @staticmethod
    def from_args(args: DispatchArgs) -> DispatchArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (DispatchArgs): Existing Instance

        Returns:
            DispatchArgs: args
        """
        eargs = DispatchArgs(source=args.source, cmd=args.cmd)
        eargs.event_data = args.event_data
        return args
