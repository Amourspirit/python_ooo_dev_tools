# coding: utf-8
from __future__ import annotations
from typing import Any


class EventArgs:
    """Event Arguments Class"""

    def __init__(self, source: Any) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
        """
        self._source = source
        self._event_name = ""

    @property
    def source(self) -> Any:
        """
        Gets/Sets Event source
        """
        return self._source

    @source.setter
    def source(self, value: Any):
        self._source = value

    @property
    def event_name(self) -> str:
        """
        Gets the event name for these args
        """
        return self._event_name

    @property
    def event_data(self) -> Any:
        """
        Gets/Sets any extra data associated with the event
        """
        try:
            return self._event_data
        except AttributeError:
            return None

    @event_data.setter
    def event_data(self, value: Any):
        self._event_data = value

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}: {self.event_name}>"

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
        eargs.event_data = args.event_data
        return args
