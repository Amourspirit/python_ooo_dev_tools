from __future__ import annotations
from typing import Any, Protocol
from ooodev.utils.gen_util import NULL_OBJ


class EventArgsT(Protocol):
    # def __init__(self, source: Any, *args, **kwargs) -> None:
    #     ...
    source: Any
    """Gets/Sets Event source"""
    ...
    event_data: Any
    """Gets/Sets any extra data associated with the event"""
    ...

    _event_name: str
    ...
    _event_source: Any

    @property
    def event_name(self) -> str:
        """
        Gets the event name for these args
        """
        ...

    @property
    def event_source(self) -> Any:
        """
        Gets the event source for these args
        """
        ...

    def get(self, key: str, default: Any = NULL_OBJ) -> Any:
        """
        Gets user data from event.

        Args:
            key (str): Key used to store data
            default (Any, optional): Default value to return if ``key`` is not found.

        Returns:
            Any: Data for ``key`` if found; Otherwise, if ``default`` is set then ``default``.

        .. versionadded:: 0.9.0
        """
        ...

    def set(self, key: str, value: Any, allow_overwrite: bool = True) -> bool:
        """
        Sets a key value pair for event instance.

        Args:
            key (str): Key
            value (Any): Value
            allow_overwrite (bool, optional): If ``True`` and a ``key`` already exist then its ``value`` will be over written; Otherwise ``value`` will not be over written. Defaults to ``True``.

        Returns:
            bool: ``True`` if values is written; Otherwise, ``False``
        """
        ...

    def has(self, key: str) -> bool:
        """
        Gets if a key exist in the instance

        Args:
            key (str): key

        Returns:
            bool: ``True`` if key exist; Otherwise ``False``
        """
        ...

    def remove(self, key: str) -> bool:
        """
        Removes key value pair from instance

        Args:
            key (str): key

        Returns:
            bool: ``True`` if key was found and removed; Otherwise, ``False``
        """
        ...
