from __future__ import annotations
from typing import Any, TypeVar, Generic, Dict, cast
from ooodev.utils.gen_util import NULL_OBJ

T = TypeVar("T")


class DotDict(Generic[T]):
    """
    Generic class for accessing dictionary keys as attributes or keys as attributes.

    Type Parameters:
        T: Value type

    Args:
        missing_attr_val (Any, optional): Value to return if attribute is not found.
            If omitted then AttributeError is raised if attribute is not found.
        kwargs (T): Keyword arguments.

    Example:

        .. code-block:: python

            # String values
            d1 = DotDict[str](a="hello", b="world")

            # Integer values
            d2 = DotDict[int](a=1, b=2)

            # Mixed values with Union
            d3 = DotDict[Union[str, int]](a="hello", b=2)

            # Mixed values with object
            d4 = DotDict[object](a="hello", b=2)

            # Mixed values with no generic type
            d5 = DotDict(a="hello", b=2)

            # Mixed values missing attribute value
            d6 = DotDict[object](None, a="hello", b=2)
            assert d6.missing is None
    """

    def __init__(self, missing_attr_val: Any = NULL_OBJ, **kwargs: T) -> None:
        """
        Constructor

        Args:
            missing_attr_val (Any, optional): Value to return if attribute is not found.
                If omitted then AttributeError is raised.
            kwargs (T): Keyword arguments.
        """
        self._missing_attrib_value = missing_attr_val  # kwargs.pop("_missing_attrib_value", NULL_OBJ)
        self._dict: Dict[str, T] = {}
        self._dict.update(cast(Dict[str, T], kwargs))

    def __getitem__(self, key: str) -> T:
        return self._dict[key]

    def __setitem__(self, key: str, value: T) -> None:
        self._dict[key] = value

    def __delitem__(self, key: str) -> None:
        del self._dict[key]

    def __getattr__(self, key: str) -> T:
        try:
            return self._dict[key]  # type: ignore
        except KeyError:
            if self._missing_attrib_value is not NULL_OBJ:
                return self._missing_attrib_value  # type: ignore
            raise AttributeError(f"'{self.__class__.__name__}' object has no attribute '{key}'")

    def __setattr__(self, key: str, value: Any) -> None:
        if key.startswith("_"):
            super().__setattr__(key, value)
        else:
            self._dict[key] = value  # type: ignore

    def __delattr__(self, key: str) -> None:
        if key.startswith("_"):
            super().__delattr__(key)
        else:
            del self._dict[key]  # type: ignore

    def __contains__(self, key: str) -> bool:
        return key in self._dict

    def __len__(self) -> int:
        return len(self._dict)

    def __copy__(self) -> DotDict[T]:
        return self.copy()

    def get(self, key: str, default: T | None = None) -> T | None:
        """
        Get value from dictionary.

        Args:
            key (KT): Key to get value.
            default (T | None, optional): Default value if key not found. Defaults to None.

        Returns:
            T | None: Value of key or default value.
        """
        return self._dict.get(key, default)

    def items(self) -> Any:
        """Returns all items in the dictionary in a set like object."""
        return self._dict.items()

    def keys(self) -> Any:
        """Returns all keys in the dictionary in a set like object."""
        return self._dict.keys()

    def values(self) -> Any:
        """Returns an object providing a view on the dictionary's values."""
        return self._dict.values()

    def update(self, other: Dict[str, T] | DotDict[T]) -> None:
        """
        Update dictionary with another dictionary.

        Args:
            other (Dict[KT, T] | DotDict[KT, T]): Dictionary to update with.
        """
        if isinstance(other, DotDict):
            self._dict.update(other._dict)
        else:
            self._dict.update(other)

    def copy(self) -> DotDict[T]:
        """Returns a shallow copy of the dictionary."""
        return DotDict(**self._dict)

    def copy_dict(self) -> Dict[str, T]:
        """Returns a shallow copy of the dictionary."""
        return self._dict.copy()

    def clear(self) -> None:
        """Clears the dictionary"""
        self._dict.clear()
