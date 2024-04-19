from __future__ import annotations
from typing import Any

from ooodev.utils.cache.lru_cache import LRUCache
from ooodev.utils.cache.time_cache import TimeCache


class TLRUCache:
    """Time and Least Recently Used (LRU) Cache."""

    def __init__(self, capacity: int, seconds: float):
        """
        Time and Least Recently Used (LRU) Cache.

        Args:
            capacity (int): Cache capacity.
            seconds (float): Time in seconds before item expires.
        """
        self._cache = LRUCache(capacity)
        self._ttl_cache = TimeCache(seconds)
        self._dummy = "x"

    def clear(self):
        """
        Clear cache.
        """
        self._cache.clear()
        self._ttl_cache.clear()

    def get(self, key: Any):
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

    # region Dunder Methods
    def __getitem__(self, key: Any) -> Any:
        if key is None:
            raise TypeError("Key cannot be None.")
        if key not in self._cache:
            if key in self._ttl_cache:
                # Key must be valid in both caches
                del self._ttl_cache[key]
            return None
        if key not in self._ttl_cache:
            del self._cache[key]
            return None
        value = self._ttl_cache[key]
        # update LRU to keep it fresh
        self._cache[key] = self._dummy
        return value

    def __setitem__(self, key: Any, value: Any) -> None:
        if key is None or value is None:
            raise TypeError("Key and value cannot be None.")
        self._cache[key] = self._dummy
        self._ttl_cache[key] = value

    def __contains__(self, key: Any) -> bool:
        return False if key is None else self[key] is not None

    def __delitem__(self, key: Any) -> None:
        if key is None:
            raise TypeError("Key must not be None.")
        if key in self._ttl_cache:
            del self._ttl_cache[key]
        if key in self._cache:
            del self._cache[key]

    def __repr__(self) -> str:
        return f"TLRUCache({self._cache.capacity}, {self._ttl_cache.seconds})"

    def __str__(self) -> str:
        return f"TLRUCache({self._cache.capacity}, {self._ttl_cache.seconds})"

    def __len__(self) -> int:
        return len(self._cache)

    # endregion Dunder Methods
