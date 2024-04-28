# see : ooodev.utils.cache.file_cache.cache_base.CacheBase for an example
class ConstructorSingleton(type):
    """
    Singleton class that uses constructor arguments to determine if an instance should be created.

    Only keyword arguments are supported.
    Keyword arguments must be hashable.

    .. versionadded:: 0.41.0
    """

    _instances = {}

    def __call__(cls, *args, **kwargs):
        # convert kwargs into a tuple of items
        if not kwargs:
            key = "default"
        else:
            t_kwargs = tuple(kwargs.items())
            key = hash((t_kwargs))
        if key not in cls._instances:
            if args:
                raise ValueError("ConstructorSingleton does not support positional arguments.")
            cls._instances[key] = super().__call__(**kwargs)
        return cls._instances[key]
