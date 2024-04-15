from __future__ import annotations
from collections import OrderedDict
from typing import Any


class LRUCache:
    """
    Least Recently Used (LRU) Cache
    """

    def __init__(self, capacity: int):
        """
        Least Recently Used (LRU) Cache

        Args:
            capacity (int): Cache capacity.
        """
        self._cache = OrderedDict()
        self._capacity = capacity

    def clear(self) -> None:
        """
        Clear cache.
        """
        self._cache.clear()

    def get(self, key: Any) -> Any:
        """
        Get value by key.

        Args:
            key (Any): Any Hashable object.

        Returns:
            Any: Value or ``None`` if not found.

        Note:
            The ``get`` method is an alias for the ``__getitem__`` method.
            So you can use ``cache_inst.get(key)`` or ``cache_inst[key]`` interchangeably.
        """
        return self.__getitem__(key)

    def put(self, key: Any, value: Any) -> None:
        """
        Put value by key.

        Args:
            key (Any): Any Hashable object.
            value (Any): Any object.

        Note:
            The ``put`` method is an alias for the ``__setitem__`` method.
            So you can use ``cache_inst.put(key, value)`` or ``cache_inst[key] = value`` interchangeably.
        """
        self.__setitem__(key, value)

    def remove(self, key: Any) -> None:
        """
        Remove key.

        Args:
            key (Any): Any Hashable object.

        Note:
            The ``remove`` method is an alias for the ``__delitem__`` method.
            So you can use ``cache_inst.remove(key)`` or ``del cache_inst[key]`` interchangeably.
        """
        self.__delitem__(key)

    def __getitem__(self, key: Any) -> Any:
        if self._capacity <= 0:
            return None
        if key not in self._cache:
            return None
        else:
            self._cache.move_to_end(key)
            return self._cache[key]

    def __setitem__(self, key: Any, value: Any) -> None:
        if self._capacity <= 0:
            return
        self._cache[key] = value
        if len(self._cache) > self._capacity:
            self._cache.popitem(last=False)

    def __contains__(self, key: Any) -> bool:
        return key in self._cache

    def __delitem__(self, key: Any) -> None:
        if key in self._cache:
            del self._cache[key]

    def __repr__(self) -> str:
        return f"LRUCache({self._capacity})"

    def __str__(self) -> str:
        return f"LRUCache({self._capacity})"
