from __future__ import annotations
from typing import Any
import threading
from ooodev.utils.cache.lru_cache import LRUCache
from ooodev.utils.cache.time_cache import TimeCache
from ooodev.events.args.event_args import EventArgs


class TLRUCache:
    """
    Time and Least Recently Used (LRU) Cache.

    When time expires, the item is removed from the cache automatically.
    """

    def __init__(self, capacity: int, seconds: float):
        """
        Time and Least Recently Used (LRU) Cache.

        Args:
            capacity (int): Cache capacity.
            seconds (float): Time in seconds before item expires.
        """
        self._lock = threading.Lock()
        self._fn_on_ttl_expired = self._on_ttl_expired
        self._lru_cache = LRUCache(capacity)
        self._seconds = seconds

        self._ttl_cache = TimeCache(seconds, self._get_ttl_seconds())
        self._dummy = object()
        self._ttl_cache.subscribe_event("cache_items_expired", self._fn_on_ttl_expired)

    def _get_ttl_seconds(self) -> float:
        # set seconds to a minimum of 10 seconds
        # also set as a separate method so it can be overridden for testing.
        return min(self._seconds, 10.0)

    def _on_ttl_expired(self, source: Any, event: EventArgs) -> None:
        with self._lock:
            keys: list = event.event_data.keys
            for key in keys:
                if key in self._lru_cache:
                    del self._lru_cache[key]

    def clear(self):
        """
        Clear cache.
        """
        self._lru_cache.clear()
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

    def clear_expired(self) -> None:
        """
        Clear expired items.
        """
        # this will trigger event to clear from LRU cache if needed.
        self._ttl_cache.clear_expired()

    # region Dunder Methods
    def __getitem__(self, key: Any) -> Any:
        if key is None:
            raise TypeError("Key cannot be None.")
        if key not in self._lru_cache:
            if key in self._ttl_cache:
                # Key must be valid in both caches
                del self._ttl_cache[key]
            return None
        if key not in self._ttl_cache:
            del self._lru_cache[key]
            return None
        value = self._ttl_cache[key]
        # update LRU to keep it fresh
        self._lru_cache[key] = self._dummy
        return value

    def __setitem__(self, key: Any, value: Any) -> None:
        if key is None or value is None:
            raise TypeError("Key and value cannot be None.")
        self._lru_cache[key] = self._dummy
        self._ttl_cache[key] = value

    def __contains__(self, key: Any) -> bool:
        return False if key is None else self[key] is not None

    def __delitem__(self, key: Any) -> None:
        if key is None:
            raise TypeError("Key must not be None.")
        if key in self._ttl_cache:
            del self._ttl_cache[key]
        if key in self._lru_cache:
            del self._lru_cache[key]

    def __repr__(self) -> str:
        return f"TLRUCache({self._lru_cache.capacity}, {self._ttl_cache.seconds})"

    def __str__(self) -> str:
        return f"TLRUCache({self._lru_cache.capacity}, {self._ttl_cache.seconds})"

    def __len__(self) -> int:
        return len(self._lru_cache)

    # endregion Dunder Methods
