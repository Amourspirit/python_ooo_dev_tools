.. _ooodev.utils.cache.singleton.file_change_aware_cache:

FileChangeAwareCache
====================

Introduction
------------

This class is a singleton class that stores data to file.
Also the cache is invalidated when the file is changed.
The data can be anything that can be pickled.

It is used to store data that is not sensitive and can be shared between different instances of the same class.

This class is more of a dynamic singleton class.
When the same parameter is passed to the constructor, the same instance is returned.
If the parameter is different, a new instance is created.
Custom key value pairs can be passed to the constructor to create a new instance.
Custom key value pairs must be hashable.

Examples
--------

This example store data in a the LibreOffice tmp directory.
Something like ``/tmp/ooo_uno_tmpl/my_tmp/x75c002da1ba8f255013111f3084243be.pkl``.

The cache is set to expire when file changes.

.. code-block:: python

    from __future__ import annotations
    from typing import Any
    import json
    from pathlib import Path
    from ooodev.utils.cache.singleton import FileChangeAwareCache

    def get_json_data(json_file: str | Path) -> Any:
        with open(json_file, 'r', encoding="utf-8") as json_file:
            data = json.load(json_file)
        return data
    
    def save_json_data(json_file: str | Path, data: Any) -> None:
        with open(json_file, 'w', encoding="utf-8") as json_file:
            json.dump(data, json_file)

    def get_cache_data(cache: FileChangeAwareCache, json_file: str | Path) -> Any:
        data = cache[json_file]
        if data is None:
            data = get_json_data(json_file)
            cache[json_file] = data
        return data

    cache = FileChangeAwareCache(tmp_dir="my_tmp")
    json_file = Path("/user/me/Documents/User.json")

    data = get_cache_data(json_file)

    data["new_key"] = "new_value"
    # once the file is changed, the cache is invalidated
    save_json_data(json_file, data)
    
    data = get_cache_data(json_file)
    assert data["new_key"] == "new_value"



Class
-----

.. autoclass:: ooodev.utils.cache.singleton.FileChangeAwareCache
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members:
    :special-members: __setitem__, __getitem__, __delitem__, __contains__, __len__, 