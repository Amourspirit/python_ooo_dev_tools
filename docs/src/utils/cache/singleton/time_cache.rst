.. _ooodev.utils.cache.singleton.time_cache:

Singleton TimeCache
===================

Introduction
------------

This class is a singleton class that stores data to memory.
The data is stored with a time to live (TTL) value.

This class is more of a dynamic singleton class.
When the same parameter is passed to the constructor, the same instance is returned.
If the parameter is different, a new instance is created.
Custom key value pairs can be passed to the constructor to create a new instance.
Custom key value pairs must be hashable.

This class is functionally the same as the instance :ref:`ooodev.utils.cache.time_cache` with the exception that it is a singleton class.


.. versionadded:: 0.52.0

Examples
--------

This example creates an instance of the TimeCache class with a TTL of ``2`` seconds.

.. code-block:: python

    import time
    from ooodev.utils.cache.singleton import TimeCache

    cache = TimeCache(seconds=2, cleanup_interval=1)  # 60 seconds
    cache["key"] = "value"
    assert "key" in cache # True
    assert cache["key"] == "value"

    time.sleep(1)
    assert "key" in cache # True

    time.sleep(3)
    assert "key" not in cache # True



This example demonstrates the use of custom key value pairs to create a new singleton instance of the TimeCache class.

.. code-block:: python

    from ooodev.utils.cache.singleton import TimeCache

    cache = TimeCache(seconds=300, cleanup_interval=60, custom1="custom1", custom2="custom2")

    cache["key1"] = "value1"
    cache["key2"] = "value2"
    cache["key3"] = "value3"

    print(cache["key1"]) # prints "value1"
    print(cache["key2"]) # prints "value2"
    print(cache["key3"]) # prints "value3"

    cache2 = TimeCache()
    assert cache not is cache2 # True

    cache3 = TimeCache(seconds=300, cleanup_interval=60, custom1="custom1", custom2="custom2")
    assert cache is cache3 # True
    print(cache3["key1"]) # prints "value1"

Class
-----

.. autoclass:: ooodev.utils.cache.singleton.TimeCache
    :members:
    :undoc-members:
    :special-members: __setitem__, __getitem__, __delitem__, __contains__, __len__, 