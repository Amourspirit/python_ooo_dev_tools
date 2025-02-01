.. _ooodev.utils.cache.singleton.lru_cache:

LRUCache
========

Introduction
------------

This class is a singleton class that stores data to memory.

This class is more of a dynamic singleton class.
When the same parameter is passed to the constructor, the same instance is returned.
If the parameter is different, a new instance is created.
Custom key value pairs can be passed to the constructor to create a new instance.
Custom key value pairs must be hashable.

This class is functionally the same as the :ref:`ooodev.utils.cache.lru_cache` with the exception that it is a singleton class.


Examples
--------

This example creates an instance of the LRUCache class with a maximum size of ``10``.

Each time an item is accessed, it is moved to the front of the cache.
If the cache is full, the least recently used item is removed and the new item is added to the front of the cache.

.. code-block:: python

    from ooodev.utils.cache.singleton import LRUCache

    cache = LRUCache(max_size=10)

    cache["key1"] = "value1"
    cache["key2"] = "value2"
    cache["key3"] = "value3"

    print(cache["key1"]) # prints "value1"
    print(cache["key2"]) # prints "value2"
    print(cache["key3"]) # prints "value3"

    del cache["key1"]
    assert "key1" not in cache

    cache2 = LRUCache(max_size=10)
    assert cache is cache2 # True


This example demonstrates the use of custom key value pairs to create a new singleton instance of the LRUCache class.

.. code-block:: python

    from ooodev.utils.cache.singleton import LRUCache

    cache = LRUCache(custom1="custom1", custom2="custom2")

    cache["key1"] = "value1"
    cache["key2"] = "value2"
    cache["key3"] = "value3"

    print(cache["key1"]) # prints "value1"
    print(cache["key2"]) # prints "value2"
    print(cache["key3"]) # prints "value3"

    del cache["key1"]
    assert "key1" not in cache

    cache2 = LRUCache()
    assert cache Not is cache2 # True

    cache3 = LRUCache(custom1="custom1", custom2="custom2")
    assert cache is cache3 # True
    print(cache3["key1"]) # prints "value1"

Class
-----

.. autoclass:: ooodev.utils.cache.singleton.LRUCache
    :members:
    :undoc-members:
    :special-members: __setitem__, __getitem__, __delitem__, __contains__, __len__, 