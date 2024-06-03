from __future__ import annotations
from typing import Any


class DotDict:
    """
    Class for accessing dictionary keys as attributes or keys as attributes.

    Example:

            .. code-block:: python

                d = DotDict(a=1, b=2)
                print(d.a)  # Outputs: 1
                print(d['b'])  # Outputs: 2
                d['c'] = 3
                print(d.c)  # Outputs: 3
                print ('a' in d)  # Outputs: True
                del d['a']
                print ('a' in d)  # Outputs: False
                print(d.a)  # Raises AttributeError
                d.a = 1
                print(d.a)  # Outputs: 1
    """

    def __init__(self, **kwargs: Any):
        self.__dict__.update(kwargs)

    def __getitem__(self, key: str):
        return self.__dict__[key]

    def __setitem__(self, key: str, value: Any):
        self.__dict__[key] = value

    def __delitem__(self, key: str):
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

    def get(self, key: str, default: Any = None) -> Any:
        """
        Get value from dictionary.

        Args:
            key (str): Key to get value.
            default (Any, optional): Default value if key not found. Defaults to None.

        Returns:
            Any: Value of key or default value.
        """
        return self.__dict__.get(key, default)

    def items(self):
        return self.__dict__.items()

    def keys(self):
        return self.__dict__.keys()

    def values(self):
        return self.__dict__.values()
