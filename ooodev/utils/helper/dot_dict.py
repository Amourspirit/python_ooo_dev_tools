from __future__ import annotations
from typing import Any, TypeVar, Generic, Dict, cast, Generator
from collections import OrderedDict
from ooodev.utils.gen_util import NULL_OBJ

T = TypeVar("T")

# Protected attributes that should not be included in dictionary operations
_PROTECTED_ATTRIBS = ("_missing_attrib_value", "_internal_keys", "_is_protocol")


class DotDict(Generic[T]):
    """
    Generic class for accessing dictionary keys as attributes or keys as attributes.

    Type Parameters:
        T: Value type

    Args:
        missing_attr_val (Any, optional): Value to return if attribute is not found.
            If omitted then AttributeError is raised if attribute is not found.
        kwargs (T): Keyword arguments.

    Note:
        It is possible to override class attributes such as keys, copy, and items attributes.
        This is not recommended.

        .. code-block:: python

            d = DotDict[str](a="hello", keys="world")
            assert d.keys == "world"

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
        self._missing_attrib_value = missing_attr_val
        self._internal_keys: OrderedDict[str, None] = OrderedDict()
        self.__dict__.update(cast(Dict[str, T], kwargs))
        for key in kwargs:
            self._internal_keys[key] = None

    def __bool__(self) -> bool:
        """Returns True if the dictionary is not empty."""
        return len(self._internal_keys) > 0

    def __getitem__(self, key: str) -> T:
        """Gets item by key using dictionary syntax."""
        return self.__dict__[key]

    def __setitem__(self, key: str, value: T) -> None:
        """Sets item by key using dictionary syntax."""
        self.__dict__[key] = value
        self._internal_keys[key] = None

    def __delitem__(self, key: str) -> None:
        """Deletes item by key using dictionary syntax."""
        del self.__dict__[key]
        if key in self._internal_keys:
            del self._internal_keys[key]

    def __getattr__(self, key: str) -> T:
        """Gets item by key using attribute syntax."""
        try:
            return self.__dict__[key]  # type: ignore
        except KeyError:
            if self._missing_attrib_value is not NULL_OBJ:
                return self._missing_attrib_value  # type: ignore
            raise AttributeError(f"'{self.__class__.__name__}' object has no attribute '{key}'")

    def __setattr__(self, key: str, value: Any) -> None:
        """Sets item by key using attribute syntax."""
        if key.startswith("__"):
            super().__setattr__(key, value)
        else:
            if key not in _PROTECTED_ATTRIBS:
                self._internal_keys[key] = None
            self.__dict__[key] = value  # type: ignore

    def __delattr__(self, key: str) -> None:
        """Deletes item by key using attribute syntax."""
        if key in self._internal_keys:
            del self._internal_keys[key]
        if key.startswith("__"):
            super().__delattr__(key)
        else:
            del self.__dict__[key]  # type: ignore

    def __contains__(self, key: str) -> bool:
        """Returns True if key exists in dictionary."""
        return key in self.__dict__

    def __len__(self) -> int:
        """Returns the number of items in the dictionary."""
        return len(self._internal_keys)

    def __copy__(self) -> DotDict[T]:
        """Returns a shallow copy of the dictionary."""
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
        return self.__dict__.get(key, default)

    def items(self) -> Generator[tuple[str, T], None, None]:
        """Returns all items in the dictionary in a set like object."""
        for key in self._internal_keys.keys():
            if key not in self.__dict__:
                continue
            yield key, self.__dict__[key]

    def keys(self) -> Generator[str, None, None]:
        """Returns all keys in the dictionary in a set like object."""
        # filter out _PROTECTED_ATTRIBS and return a generator expression
        for key in self._internal_keys.keys():
            yield key

    def values(self) -> Generator[T, None, None]:
        """Returns an object providing a view on the dictionary's values."""
        # filter out _PROTECTED_ATTRIBS and return a generator expression
        for key in self._internal_keys.keys():
            if key not in self.__dict__:
                continue
            yield self.__dict__[key]

    def update(self, other: Dict[str, T] | DotDict[T]) -> None:
        """
        Update dictionary with another dictionary.

        Args:
            other (Dict[KT, T] | DotDict[KT, T]): Dictionary to update with.

        Raises:
            TypeError: If other is not a dict or DotDict
        """
        if isinstance(other, DotDict):
            self.__dict__.update(other.__dict__)
            self._internal_keys.update(other._internal_keys)
        elif isinstance(other, dict):
            self.__dict__.update(other)
            for key in other.keys():
                self._internal_keys[key] = None
        else:
            raise TypeError(f"Expected dict or DotDict, got {type(other)}")

    def copy(self) -> DotDict[T]:
        """Returns a shallow copy of the dictionary."""
        copy_dict = {}
        for key in self._internal_keys.keys():
            if key not in self.__dict__:
                continue
            copy_dict[key] = self.__dict__[key]
        copy_dict["missing_attr_val"] = self._missing_attrib_value
        inst = DotDict[T](**copy_dict)
        return inst

    def copy_dict(self) -> Dict[str, T]:
        """Returns a shallow copy as a standard dictionary."""
        copy_dict = {}
        for key in self._internal_keys.keys():
            if key not in self.__dict__:
                continue
            copy_dict[key] = self.__dict__[key]
        return copy_dict

    def clear(self) -> None:
        """Clears all items from the dictionary while preserving protected attributes."""
        protected = {}
        for attr in _PROTECTED_ATTRIBS:
            if attr in self.__dict__:
                protected[attr] = self.__dict__[attr]
        self._internal_keys.clear()
        self.__dict__.clear()
        self.__dict__.update(protected)
