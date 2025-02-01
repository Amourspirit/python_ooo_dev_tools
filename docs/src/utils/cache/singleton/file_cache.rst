.. _ooodev.utils.cache.singleton.file_cache:

Class FileCache
===============

Introduction
------------

This class is a singleton class that stores data to file. The data can be anything that can be pickled.

It is used to store data that is not sensitive and can be shared between different instances of the same class.

This class is more of a dynamic singleton class. When the same parameter is passed to the constructor, the same instance is returned.
If the parameter is different, a new instance is created.
Custom key value pairs can be passed to the constructor to create a new instance.
Custom key value pairs must be hashable.

Examples
--------

This example store data in a the LibreOffice tmp directory.
Something like ``/tmp/ooo_uno_tmpl/my_tmp/tmp_data.pkl``.

The cache is set to expire after ``300`` seconds.

.. code-block:: python

    from ooodev.utils.cache.singleton import FileCache

    cache = FileCache(tmp_dir="my_tmp", lifetime=300)

    data = get_json_data() # Get JSON data from somewhere as a dictionary
    file_name = "tmp_data.pkl"
    cache[file_name] = data
    assert cache[file_name] == data
    if file_name in cache:
        del cache[file_name]

This example store data in a the LibreOffice tmp directory.
Something like ``/tmp/ooo_uno_tmpl/json/data.pkl``.

The cache is set to expire after ``300`` seconds.

.. code-block:: python

    import json
    from ooodev.utils.cache.singleton import FileCache

    cache = FileCache(tmp_dir="json", lifetime=300, custom1="custom1", custom2="custom2")
    data = get_json_data() # Get JSON data from somewhere as a list of dictionary

    file_name = "data.pkl"
    cache[file_name] = data
    assert cache[file_name] == data
    if file_name in cache:
        del cache[file_name]

Class
-----

.. autoclass:: ooodev.utils.cache.singleton.FileCache
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members:
    :special-members: __setitem__, __getitem__, __delitem__, __contains__, __len__, 