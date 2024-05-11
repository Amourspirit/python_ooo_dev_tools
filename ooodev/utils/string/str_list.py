from __future__ import annotations
from typing import Any, Iterable, overload


class StrList:
    """
    String List Class.

    .. versionadded:: 0.41.0
    """

    def __init__(self, strings: Iterable[str] | None = None, sep: str = ";"):
        """
        Constructor

        Args:
            strings (Iterable[str], optional): Iterable String such as a list or a tuple. If a str in passed int the each char become an item. Defaults to None.
            sep (str, optional): _description_. Defaults to ";".
        """
        if strings is None:
            strings = []
        else:
            strings = list(strings)
        self._strings = strings
        self._sep = sep
        self._iter_index = 0

    # region Methods
    def append(self, value: str) -> StrList:
        """
        Add a string to the list

        Args:
            value (str): String to add

        Returns:
            StrList: Self.
        """
        self._strings.append(value)
        return self

    def remove(self, value: str) -> StrList:
        """
        Remove a string from the list

        Args:
            value (str): String to remove.

        Returns:
            StrList: Self.
        """
        self._strings.remove(value)
        return self

    def to_string(self) -> str:
        """Convert list to string."""
        return str(self)

    def clear(self) -> StrList:
        """Clear the list."""
        self._strings.clear()
        return self

    def copy(self) -> StrList:
        """Copy the list."""
        return StrList(strings=self._strings.copy(), sep=self._sep)

    def extend(self, strings: Iterable[str]) -> StrList:
        """Extend the list."""
        self._strings.extend(strings)
        return self

    def insert(self, index: int, value: str) -> StrList:
        """Insert a value into the list."""
        self._strings.insert(index, value)
        return self

    def pop(self, index: int) -> str:
        """Pop a value from the list."""
        return self._strings.pop(index)

    def index(self, value: str) -> int:
        """Get the index of a value."""
        return self._strings.index(value)

    def count(self, value: str) -> int:
        """Get the count of a value."""
        return self._strings.count(value)

    def reverse(self) -> StrList:
        """Reverse the list."""
        self._strings.reverse()
        return self

    def sort(self, key=None, reverse=False) -> StrList:
        """Sort the list."""
        self._strings.sort(key=key, reverse=reverse)
        return self

    def remove_duplicates(self) -> StrList:
        """Remove duplicates from the list."""
        self._strings = list(dict.fromkeys(self._strings))
        return self

    # endregion Methods

    # region Dunder Methods

    def __add__(self, other: StrList) -> StrList:
        """Add two lists."""
        return StrList(strings=self._strings + other._strings, sep=self._sep)

    def __iadd__(self, other: StrList) -> StrList:
        """Add two lists."""
        self._strings += other._strings
        return self

    def __eq__(self, other: StrList) -> bool:
        """Check if two lists are equal."""
        return self._strings == other._strings

    def __contains__(self, value: str):
        """Get if the value is in the list."""
        return value in self._strings

    def __iter__(self):
        """Iterator for the list."""
        self._iter_index = 0
        length = len(self)
        while self._iter_index < length:
            yield self._strings[self._iter_index]
            self._iter_index += 1

    def __reversed__(self):
        """Reverse iterator for the list."""
        self._iter_index = len(self) - 1
        while self._iter_index >= 0:
            yield self._strings[self._iter_index]
            self._iter_index -= 1

    def __str__(self):
        """Convert list to string."""
        if not self._strings:
            return ""
        return self._sep.join(self._strings)

    def __len__(self):
        """Get the length of the list."""
        return len(self._strings)

    def __delitem__(self, index):
        """Delete an item from the list."""
        del self._strings[index]

    # region __getitem__
    @overload
    def __getitem__(self, index: int) -> str: ...

    @overload
    def __getitem__(self, index: slice) -> StrList: ...

    def __getitem__(self, index: int | slice) -> Any:
        """
        Get an item from the list.

        Supports slicing. When sliced a new StrList is returned.
        """

        if isinstance(index, slice):
            return type(self)(self._strings[index], self._sep)
        else:
            return self._strings[index]

    # endregion __getitem__

    def __repr__(self) -> str:
        """Get the string representation of the object."""
        return f"<{self.__class__.__name__}(count={len(self._strings)}, sep={self._sep!r})?"

    # endregion Dunder Methods

    # region Static Methods
    @staticmethod
    def from_str(value: str, sep: str = ";") -> StrList:
        """Create a StrList from a string."""
        return StrList(strings=value.split(sep), sep=sep)

    # endregion Static Methods

    # region Properties
    @property
    def separator(self) -> str:
        """Gets/Sets the separator."""
        return self._sep

    @separator.setter
    def separator(self, value: str):
        """Set the separator."""
        self._sep = value
