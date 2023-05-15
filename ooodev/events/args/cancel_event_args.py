from __future__ import annotations
from typing import Any
from .event_args import AbstractEvent


class AbstractCancelEventArgs(AbstractEvent):
    # https://stackoverflow.com/questions/472000/usage-of-slots
    __slots__ = ()

    def __init__(self, source: Any, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source)
        self.cancel = cancel
        self.handled = False

    cancel: bool
    """Gets/Sets cancel value"""
    handled: bool
    """Get/Set Handled value. Typically if set to ``True`` then ``cancel`` is ignored."""


class CancelEventArgs(AbstractCancelEventArgs):
    """Cancel Event Arguments"""

    __slots__ = ("source", "_event_name", "event_data", "cancel", "handled", "_event_source", "_kv_data")

    @staticmethod
    def from_args(args: AbstractCancelEventArgs) -> CancelEventArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (AbstractCancelEventArgs): Existing Instance

        Returns:
            CancelEventArgs: args
        """
        eargs = CancelEventArgs(source=args.source)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        eargs.cancel = args.cancel
        eargs.handled = args.handled
        return eargs
