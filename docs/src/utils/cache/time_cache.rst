.. _ooodev.utils.cache.time_cache:

TimeCache
=========

Introduction
------------

This class is a class that stores data to memory.
The data is stored with a time to live (TTL) value.
When the TTL expires, the data is removed from the cache.

Unlike :ref:`ooodev.utils.cache.singleton.time_cache`, this class is not a singleton class.


Examples
--------

This example creates an instance of the TimeCache class with a TTL of ``2`` seconds.

.. code-block:: python

    import time
    from ooodev.utils.cache import TimeCache

    cache = TimeCache(seconds=2)  # 60 seconds
    cache["key"] = "value"
    assert "key" in cache # True
    assert cache["key"] == "value"

    time.sleep(1)
    assert "key" in cache # True

    time.sleep(3)
    assert "key" not in cache # True

This example demonstrates the use of the cache_items_expired event.
A callback function is called when the TTL expires.

.. code-block:: python

    import threading
    from ooodev.utils.cache import TimeCache

    LOCK = threading.Lock()

    def on_items_expired(source, event):
        with LOCK:
            # event.event_data is a DotDict with an attribute keys
            # that contains a list of keys that have expired.
            keys = event.event_data.keys
            for key in keys:
                print(f"Expired: {key}")

    cache = TimeCache(60.0)  # 60 seconds
    cache.subscribe_event("cache_items_expired", on_items_expired)
    cache["key"] = "value"
    value = cache["key"]

Class
-----

.. autoclass:: ooodev.utils.cache.TimeCache
    :members:
    :undoc-members:
    :special-members: __setitem__, __getitem__, __delitem__, __contains__, __len__, 