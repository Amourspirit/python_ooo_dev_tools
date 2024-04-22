from __future__ import annotations
from typing import Any
from datetime import datetime, timedelta, timezone


class TimeCache:
    """
    Time based Cache.

    Cached items expire after a specified time.
    """

    def __init__(self, seconds: float):
        """
        Time based Cache.

        Args:
            seconds (float): Cache expiration time in seconds.
        """
        self._delta = timedelta(seconds=max(seconds, 0))
        self._expiration_time = datetime.now(timezone.utc) + self._delta
        self._cache = {}

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

    def clear_expired(self) -> None:
        """
        Clear expired items from the cache.
        """
        seconds = self.seconds
        del_keys = [
            key
            for key, (_, timestamp) in self._cache.items()
            if (datetime.now(timezone.utc) - timestamp).total_seconds() >= seconds
        ]
        for key in del_keys:
            del self._cache[key]

    # region Dunder Methods
    def __getitem__(self, key: Any) -> Any:
        if key is None:
            raise TypeError("Key must not be None.")
        if key in self._cache:
            value, timestamp = self._cache[key]
            if (datetime.now(timezone.utc) - timestamp).total_seconds() < self.seconds:
                return value
            else:
                del self._cache[key]  # remove expired item
        return None

    def __setitem__(self, key: Any, value: Any) -> None:
        if key is None or value is None:
            raise TypeError("Key and value must not be None.")
        current_time = datetime.now(timezone.utc)
        self._cache[key] = (value, current_time)

    def __contains__(self, key: Any) -> bool:
        return False if key is None else self[key] is not None

    def __delitem__(self, key: Any) -> None:
        if key is None:
            raise TypeError("Key must not be None.")
        if key in self:
            del self._cache[key]

    def __repr__(self) -> str:
        return f"TimeBasedCache({self._delta.seconds})"

    def __str__(self) -> str:
        return f"TimeBasedCache({self._delta.seconds})"

    def __len__(self) -> int:
        return len(self._cache)

    # endregion Dunder Methods

    # region Properties
    @property
    def seconds(self) -> float:
        """
        Gets/Sets Cache expiration time in seconds.
        """
        return self._delta.total_seconds()

    @seconds.setter
    def seconds(self, value: float) -> None:
        self._delta = timedelta(seconds=value)
        self._expiration_time = datetime.now(timezone.utc) + self._delta
        self.clear_expired()

    # endregion Properties
