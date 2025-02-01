from __future__ import annotations
import warnings
from ooodev.utils.cache.file_cache.file_cache import FileCache


# Alias for backward compatibility
class PickleCache(FileCache):
    """
    Please use ``ooodev.utils.cache.file_cache.file_cache.FileCache`` instead.

    .. versionchanged:: 0.52.0
        Deprecated
    """

    def __init__(self, *args, **kwargs):
        warnings.warn(
            "PickleCache is deprecated and will be removed in a future release. "
            "Please use ooodev.utils.cache.file_cache.file_cache.FileCache instead.",
            DeprecationWarning,
            stacklevel=2,
        )
        super().__init__(*args, **kwargs)
