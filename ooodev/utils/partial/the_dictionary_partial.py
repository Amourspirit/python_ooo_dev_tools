from __future__ import annotations
from typing import Any, Dict
from ooodev.utils.gen_util import NULL_OBJ


class TheDict:
    """A partial class for implementing dictionary values"""

    def __init__(self, values: Dict[str, Any] | None = None) -> None:
        self._kv_data = values if values is not None else {}

    def __getitem__(self, key: str) -> Any:
        return self._kv_data.get(key, NULL_OBJ)

    def __setitem__(self, key: str, value: Any) -> None:
        self._kv_data[key] = value

    def __delitem__(self, key: str) -> None:
        if key in self._kv_data:
            del self._kv_data[key]

    def get(self, key: str, default: Any = NULL_OBJ) -> Any:
        """
        Gets user data from event.

        Args:
            key (str): Key used to store data
            default (Any, optional): Default value to return if ``key`` is not found.

        Returns:
            Any: Data for ``key`` if found; Otherwise, if ``default`` is set then ``default``.
        """
        if self._kv_data is None:
            if default is NULL_OBJ:
                raise KeyError(f'"{key}" not found. Maybe you want to include a default value.')
            return default
        if default is NULL_OBJ:
            return self._kv_data[key]
        return self._kv_data.get(key, default)

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
        if self._kv_data is None:
            self._kv_data = {}
        if allow_overwrite:
            self._kv_data[key] = value
            return True
        if key not in self._kv_data:
            self._kv_data[key] = value
            return True
        return False

    def has(self, key: str) -> bool:
        """
        Gets if a key exist in the instance

        Args:
            key (str): key

        Returns:
            bool: ``True`` if key exist; Otherwise ``False``
        """
        return False if self._kv_data is None else key in self._kv_data

    def remove(self, key: str) -> bool:
        """
        Removes key value pair from instance

        Args:
            key (str): key

        Returns:
            bool: ``True`` if key was found and removed; Otherwise, ``False``
        """
        if self._kv_data is None:
            return False
        if key in self._kv_data:
            del self._kv_data[key]
            return True
        return False


class TheDictionaryPartial:
    """A partial class for implementing dictionary values"""

    def __init__(self, values: Dict[str, Any] | None = None) -> None:
        """
        Constructor

        Args:
            values (Dict[str, Any] | None, optional): Dictionary Values. Defaults to None.
        """
        self.__the_dict = None if values is None else TheDict(values)

    @property
    def extra_data(self) -> TheDict:
        """Extra Data Key Value Pair Dictionary"""
        if self.__the_dict is None:
            self.__the_dict = TheDict()
        return self.__the_dict
