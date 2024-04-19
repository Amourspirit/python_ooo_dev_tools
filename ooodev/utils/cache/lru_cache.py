from __future__ import annotations
from collections import OrderedDict
from typing import Any


class LRUCache:
    """
    Least Recently Used (LRU) Cache
    """

    # region Initialization
    def __init__(self, capacity: int):
        """
        Least Recently Used (LRU) Cache

        Args:
            capacity (int): Cache capacity.
        """
        self._cache = OrderedDict()
        self._capacity = max(capacity, 0)

    # endregion Initialization

    # region Dictionary Methods

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
        return self[key]

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
        self[key] = value

    def remove(self, key: Any) -> None:
        """
        Remove key.

        Args:
            key (Any): Any Hashable object.

        Note:
            The ``remove`` method is an alias for the ``__delitem__`` method.
            So you can use ``cache_inst.remove(key)`` or ``del cache_inst[key]`` interchangeably.
        """
        del self[key]

    # endregion Dictionary Methods

    # region Dunder Methods

    def __getitem__(self, key: Any) -> Any:
        if key is None:
            raise TypeError("Key must not be None.")
        if self._capacity <= 0:
            return None
        if key not in self._cache:
            return None
        self._cache.move_to_end(key)
        return self._cache[key]

    def __setitem__(self, key: Any, value: Any) -> None:
        if key is None or value is None:
            raise TypeError("Key and value must not be None.")
        if self._capacity <= 0:
            return
        self._cache[key] = value
        if len(self._cache) > self._capacity:
            self._cache.popitem(last=False)

    def __contains__(self, key: Any) -> bool:
        return False if key is None else key in self._cache

    def __delitem__(self, key: Any) -> None:
        if key is None:
            raise TypeError("Key must not be None.")
        if key in self._cache:
            del self._cache[key]

    def __repr__(self) -> str:
        return f"LRUCache({self._capacity})"

    def __str__(self) -> str:
        return f"LRUCache({self._capacity})"

    def __len__(self) -> int:
        return len(self._cache)

    # endregion Dunder Methods

    # region Properties
    @property
    def capacity(self) -> int:
        """
        Gets/Sets Cache capacity.

        Setting the capacity to 0 or less will clear the cache and effectively turn caching off.
        Setting the capacity to a lower value will remove the least recently used items.

        Returns:
            int: Cache capacity.
        """
        return self._capacity

    @capacity.setter
    def capacity(self, value: int) -> None:
        """
        Cache capacity.

        Args:
            value (int): Cache capacity.
        """
        self._capacity = value
        self._capacity = max(self._capacity, 0)
        if self._capacity == 0:
            self.clear()
            return
        while len(self._cache) > self._capacity:
            self._cache.popitem(last=False)

    # endregion Properties
