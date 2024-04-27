from __future__ import annotations
from ooodev.utils.cache.time_cache import TimeCache


class GTC(TimeCache):
    """
    Global Time based Cache.

    This cache is a singleton and is shared across all instances of the application.

    The cache as a lifetime of ``5`` minutes and expired items are cleaned up every ``60`` seconds.
    """

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not isinstance(cls._instance, cls):
            cls._instance = super(GTC, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self) -> None:
        """
        Time based Cache.

        Args:
            seconds (float): Cache expiration time in seconds.
            cleanup_interval (float, optional): Cache cleanup interval in seconds.
                If set to ``0`` then the cleanup is disabled. Defaults to ``60.0``.
        """
        if self._initialized:
            return
        super().__init__(300, cleanup_interval=60)
        self._initialized = True

    # region Override properties
    @property
    def cleanup_interval(self):
        """Gets the cache cleanup interval in seconds."""
        return super().cleanup_interval

    @property
    def seconds(self):
        """Gets the cache expiration time in seconds."""
        return super().seconds

    # endregion Override properties
