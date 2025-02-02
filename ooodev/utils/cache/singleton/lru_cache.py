from __future__ import annotations
from typing import Any

from ooodev.utils.cache.lru_cache import LRUCache as InstLRUCache
from ooodev.meta.constructor_singleton import ConstructorSingleton


class LRUCache(InstLRUCache, metaclass=ConstructorSingleton):
    """
    Least Recently Used (LRU) singleton Cache

    See Also :ref:`ooodev.utils.cache.singleton.lru_cache`

    .. versionadded:: 0.52.0
    """

    def __init__(self, *, capacity: int = 128, **kwargs: Any) -> None:
        """
        Least Recently Used (LRU) Cache

        Args:
            capacity (int): Cache capacity. Defaults to ``128``.
            kwargs (Any): Additional keyword arguments. The arguments are used to create a unique instance of the singleton class.
        """
        super().__init__(capacity)
