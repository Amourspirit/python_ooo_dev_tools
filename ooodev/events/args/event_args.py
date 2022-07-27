# coding: utf-8
from __future__ import annotations
from typing import Any
from abc import ABC


class AbstractEvent(ABC):
    # https://stackoverflow.com/questions/472000/usage-of-slots
    __slots__ = ()

    def __init__(self, source: Any) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
        """
        self.source = source
        self._event_name = ""
        self._event_source = None
        self.event_data = None

    source: Any
    """Gets/Sets Event source"""
    event_data: Any
    """Gets/Sets any extra data associated with the event"""

    @property
    def event_name(self) -> str:
        """
        Gets the event name for these args
        """
        return self._event_name

    @property
    def event_source(self) -> Any | None:
        """
        Gets the event source for these args
        """
        return self._event_source

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}: {self.event_name}>"


class EventArgs(AbstractEvent):
    """
    Event Arguments Class
    """

    __slots__ = ("source", "_event_name", "event_data", "_event_source")

    @staticmethod
    def from_args(args: EventArgs) -> EventArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (EventArgs): Existing Instance

        Returns:
            EventArgs: args
        """
        eargs = EventArgs(source=args.source)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        return eargs


e = EventArgs(None)
