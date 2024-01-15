from __future__ import annotations
from typing import Any, Generic, TypeVar
from .event_args_generic import EventArgsGeneric

_T = TypeVar("_T")


class CancelEventArgsGeneric(EventArgsGeneric[_T], Generic[_T]):
    """Cancel Event Arguments"""

    __slots__ = ("cancel", "handled")

    def __init__(
        self,
        source: Any,
        event_data: _T,
        cancel=False,
    ) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source, event_data)
        self.cancel = cancel
        self.handled = False

    cancel: bool
    """Gets/Sets cancel value"""
    handled: bool
    """Get/Set Handled value. Typically if set to ``True`` then ``cancel`` is ignored."""

    @staticmethod
    def from_args(args: CancelEventArgsGeneric) -> CancelEventArgsGeneric:
        """
        Gets a new instance from existing instance

        Args:
            args (AbstractCancelEventArgs): Existing Instance

        Returns:
            CancelEventArgs: args
        """
        eargs = CancelEventArgsGeneric(source=args.source, event_data=args.event_data)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        eargs.cancel = args.cancel
        eargs.handled = args.handled
        return eargs
