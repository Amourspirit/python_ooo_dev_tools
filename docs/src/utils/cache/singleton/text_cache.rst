.. _ooodev.utils.cache.singleton.text_cache:

Singleton TextCache
===================

Introduction
------------

This class is a singleton class that stores text data to file.

It is used to store data that is not sensitive and can be shared between different instances of the same class.

This class is more of a dynamic singleton class. When the same parameter is passed to the constructor, the same instance is returned.
If the parameter is different, a new instance is created. Custom key value pairs can be passed to the constructor to create a new instance.
Custom key value pairs must be hashable.

.. versionadded:: 0.52.0

Examples
--------

This example store text data in a the LibreOffice tmp directory.
Something like ``/tmp/ooo_uno_tmpl/my_tmp/tmp_data.txt``.

The cache is set to expire after ``300`` seconds.

.. code-block:: python

    from ooodev.utils.cache.singleton import TextCache

    cache = TextCache(tmp_dir="my_tmp", lifetime=300)

    file_name = "temp_data.txt"
    cache[file_name] = "Hello World!"
    print(cache[file_name]) # prints "Hello World!"
    if file_name in cache:
        del cache[file_name]

This example store text data in a the LibreOffice tmp directory.
Something like ``/tmp/ooo_uno_tmpl/txt_only/text_data.txt``.

The cache is set to expire after ``300`` seconds.

.. code-block:: python

    from ooodev.utils.cache.singleton import TextCache

    cache = TextCache(tmp_dir="txt_only", lifetime=300, custom1="custom1", custom2="custom2")

    file_name = "text_data.txt"
    cache[file_name] = "Hello World!"
    print(cache[file_name]) # prints "Hello World!"
    if file_name in cache:
        del cache[file_name]

Class
-----

.. autoclass:: ooodev.utils.cache.singleton.TextCache
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members:
    :special-members: __setitem__, __getitem__, __delitem__, __contains__, __len__, 