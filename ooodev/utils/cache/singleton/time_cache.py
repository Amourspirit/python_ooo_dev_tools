from __future__ import annotations
from typing import Any

from ooodev.utils.cache.time_cache import TimeCache as InstTimeCache
from ooodev.meta.constructor_singleton import ConstructorSingleton


class TimeCache(InstTimeCache, metaclass=ConstructorSingleton):
    def __init__(self, *, seconds: float = 300.0, cleanup_interval: float = 60.0, **kwargs: Any) -> None:
        """
        Time based Cache.

        Args:
            seconds (float, optional): Cache expiration time in seconds. Defaults to ``300.0``.
            cleanup_interval (float, optional): Cache cleanup interval in seconds.
                If set to ``0`` then the cleanup is disabled. Defaults to ``60.0``.
            kwargs (Any): Additional keyword arguments. The arguments are used to create a unique instance of the singleton class.
        """
        super().__init__(seconds, cleanup_interval)
