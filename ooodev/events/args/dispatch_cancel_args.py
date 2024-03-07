from __future__ import annotations
from typing import Any
from ooodev.events.args.dispatch_args import AbstractDispatchArgs
from ooodev.events.args.cancel_event_args import AbstractCancelEventArgs

# pylint: disable=protected-access
# pylint: disable=assigning-non-slot


class AbstractDispatchCancelArgs(AbstractDispatchArgs, AbstractCancelEventArgs):
    __slots__ = ()

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
        self.handled = False


class DispatchCancelArgs(AbstractDispatchCancelArgs):
    """
    Dispatch Cancel Args
    """

    __slots__ = ("cmd", "source", "_event_name", "event_data", "cancel", "handled", "_event_source", "_kv_data")

    @staticmethod
    def from_args(args: AbstractDispatchCancelArgs) -> DispatchCancelArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (AbstractDispatchCancelArgs): Existing Instance

        Returns:
            DispatchCancelArgs: args
        """
        eargs = DispatchCancelArgs(source=args.source, cmd=args.cmd)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        eargs.cancel = args.cancel
        eargs.handled = args.handled
        return eargs
