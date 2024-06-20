from __future__ import annotations
from typing import Any, Dict
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.utils.helper.dot_dict import DotDict


class TheDict:
    """A partial class for implementing dictionary values"""

    def __init__(self, values: Dict[str, Any] | None = None) -> None:
        if values:
            self.__dict__.update(values)

    def __getitem__(self, key: str) -> Any:
        return self.__dict__[key]

    def __setitem__(self, key: str, value: Any) -> None:
        self.__dict__[key] = value

    def __delitem__(self, key: str) -> None:
        if key in self.__dict__:
            del self.__dict__[key]

    def __getattr__(self, key: str):
        try:
            return self.__dict__[key]
        except KeyError:
            raise AttributeError(f"'{self.__class__.__name__}' object has no attribute '{key}'")

    def __setattr__(self, key: str, value: Any):
        self.__dict__[key] = value

    def __delattr__(self, key: str):
        del self.__dict__[key]

    def __contains__(self, key: str):
        return key in self.__dict__

    def __len__(self):
        return len(self.__dict__)

    def __copy__(self):
        return self.copy()

    def get(self, key: str, default: Any = NULL_OBJ) -> Any:
        """
        Gets user data from event.

        Args:
            key (str): Key used to store data
            default (Any, optional): Default value to return if ``key`` is not found.

        Returns:
            Any: Data for ``key`` if found; Otherwise, if ``default`` is set then ``default``.

        Raises:
            KeyError: If ``key`` is not found and ``default`` is not set.
        """

        if default is NULL_OBJ:
            return self.__dict__[key]
        return self.__dict__.get(key, default)

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
        if allow_overwrite:
            self.__dict__[key] = value
            return True
        if key not in self.__dict__:
            self.__dict__[key] = value
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
        return key in self

    def remove(self, key: str) -> bool:
        """
        Removes key value pair from instance

        Args:
            key (str): key

        Returns:
            bool: ``True`` if key was found and removed; Otherwise, ``False``
        """
        if key in self.__dict__:
            del self.__dict__[key]
            return True
        return False

    def items(self):
        """Returns all items in the dictionary in a set like object."""
        return self.__dict__.items()

    def keys(self):
        """Returns all keys in the dictionary in a set like object."""
        return self.__dict__.keys()

    def values(self):
        """Returns an object providing a view on the dictionary's values."""
        return self.__dict__.values()

    def copy(self):
        """Returns a shallow copy of the dictionary."""
        return self.__dict__.copy()

    def copy_dict(self) -> dict:
        """Returns a shallow copy of the dictionary."""
        return self.__dict__.copy()

    def update(self, other: dict | TheDict | DotDict):
        """
        Update dictionary with another dictionary.

        Args:
            other (dict, TheDict, DotDict): Dictionary to update with.
        """
        if isinstance(other, (TheDict, DotDict)):
            self.__dict__.update(other.__dict__)
        else:
            self.__dict__.update(other)

    def clear(self):
        """Clears the dictionary"""
        self.__dict__.clear()


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
        """
        Extra Data Key Value Pair Dictionary.

        Properties can be assigned properties and access like a dictionary and with dot notation.

        Note:
            This is a dictionary object that can be used to store key value pairs.
            Generally speaking this data is not part of the object's main data structure and is not saved with the object (document).

            This property is used to store data that is not part of the object's main data structure and can be used however the developer sees fit.
        """
        if self.__the_dict is None:
            self.__the_dict = TheDict()
        return self.__the_dict
