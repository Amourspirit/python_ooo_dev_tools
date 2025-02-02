.. _ooodev.utils.cache.lru_cache:

Class LRUCache
==============

Introduction
------------

This class is a class that stores data to memory.
Unlike :ref:`ooodev.utils.cache.singleton.lru_cache`, this class is not a singleton class.
This means you must be sure to keep the instance of the class alive as long as you need the cache.

Examples
--------

This example creates an instance of the LRUCache class with a maximum size of ``10``.

Each time an item is accessed, it is moved to the front of the cache.
If the cache is full, the least recently used item is removed and the new item is added to the front of the cache.

.. code-block:: python

    from ooodev.utils.cache import LRUCache

    cache = LRUCache(max_size=10)

    cache["key1"] = "value1"
    cache["key2"] = "value2"
    cache["key3"] = "value3"

    print(cache["key1"]) # prints "value1"
    print(cache["key2"]) # prints "value2"
    print(cache["key3"]) # prints "value3"

    del cache["key1"]
    assert "key1" not in cache

Class
-----

.. autoclass:: ooodev.utils.cache.LRUCache
    :members:
    :undoc-members:
    :special-members: __setitem__, __getitem__, __delitem__, __contains__, __len__, 