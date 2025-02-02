from __future__ import annotations
from typing import Any

from ooodev.utils.cache.time_cache import TimeCache as InstTimeCache
from ooodev.meta.constructor_singleton import ConstructorSingleton


class TimeCache(InstTimeCache, metaclass=ConstructorSingleton):
    """
    Time based singleton Cache.

    Cached items expire after a specified time. If ``cleanup_interval`` is set, then the cache is cleaned up at regular intervals;
    Otherwise, the cache is only cleaned up when an item is accessed.

    Each time an element is accessed, the timestamp is updated. If the element has expired, it is removed from the cache.

    When an item expires, the event ``cache_items_expired`` is triggered.
    This event is called on a separate thread. for this reason it is important to make sure that the event handler is thread safe.\
    
    See Also :ref:`ooodev.utils.cache.singleton.time_cache`
    """

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
