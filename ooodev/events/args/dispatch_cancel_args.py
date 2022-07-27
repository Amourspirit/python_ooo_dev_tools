# coding: utf-8
from __future__ import annotations
from typing import Any
from .dispatch_args import AbstractDispacthArgs
from .cancel_event_args import AbstractCancelEventArgs


class AbstractDispatchCancelArgs(AbstractDispacthArgs, AbstractCancelEventArgs):
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


class DispatchCancelArgs(AbstractDispatchCancelArgs):
    """
    Dispatch Cancel Args
    """

    __slots__ = ("cmd", "source", "_event_name", "event_data", "cancel", "_event_source")

    @staticmethod
    def from_args(args: DispatchCancelArgs) -> DispatchCancelArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (DispatchCancelArgs): Existing Instance

        Returns:
            DispatchCancelArgs: args
        """
        eargs = DispatchCancelArgs(source=args.source, cmd=args.cmd)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        eargs.cancel = args.cancel
        return eargs
