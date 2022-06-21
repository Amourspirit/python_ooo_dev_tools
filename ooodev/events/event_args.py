# coding: utf-8
from typing  import Any

class EventArgs:
    """Event Arguments Class"""
    def __init__(self, source: Any) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
        """
        self._source = source

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
        Gets/Sets Event that raised these args
        """
        return self._event_name

    @event_name.setter
    def event_name(self, value: str):
        self._event_name = value
    
    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}: {self.event_name}>"