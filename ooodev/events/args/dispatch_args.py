from __future__ import annotations
from typing import Any
from ooodev.events.args.event_args import AbstractEvent

# pylint: disable=protected-access
# pylint: disable=assigning-non-slot


class AbstractDispatchArgs(AbstractEvent):
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


class DispatchArgs(AbstractDispatchArgs):
    __slots__ = ("source", "_event_name", "event_data", "cmd", "_event_source", "_kv_data")

    @staticmethod
    def from_args(args: AbstractDispatchArgs) -> DispatchArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (AbstractDispatchArgs): Existing Instance

        Returns:
            DispatchArgs: args
        """
        eargs = DispatchArgs(source=args.source, cmd=args.cmd)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        return eargs
