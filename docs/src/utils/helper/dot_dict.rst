
DotDict
=======

The ``DotDict`` class is a generic dictionary implementation that allows accessing dictionary keys using both dictionary syntax (``d["key"]``) and attribute syntax (``d.key``).

Features
--------

- Generic type support for values
- Optional default value for missing attributes
- Dictionary-like operations (get, items, keys, values, update, etc.)
- Attribute-style access to dictionary items
- Protected attributes preservation
- Shallow copy support

Type Parameters
---------------

- ``T``: The type of values stored in the dictionary

Constructor
----------

.. code-block:: python

    DotDict[T](missing_attr_val: Any = NULL_OBJ, **kwargs: T)

Parameters:
    - ``missing_attr_val``: Value to return when accessing non-existent attributes. If not provided, raises AttributeError
    - ``**kwargs``: Initial key-value pairs for the dictionary

Usage Examples
--------------

Basic Usage:

.. code-block:: python

    # String values
    d1 = DotDict[str](a="hello", b="world")
    print(d1.a)  # "hello"
    print(d1["b"])  # "world"

    # Integer values
    d2 = DotDict[int](a=1, b=2)
    print(d2.a)  # 1

    # Mixed values with Union
    from typing import Union
    d3 = DotDict[Union[str, int]](a="hello", b=2)

    # Mixed values with object
    d4 = DotDict[object](a="hello", b=2)

Default Values:

.. code-block:: python

    # With default value for missing attributes
    d = DotDict[str]("default", a="hello")
    print(d.missing)  # "default"

    # With None as default
    d = DotDict[str](None, a="hello")
    print(d.missing)  # None

Dictionary Operations:

.. code-block:: python

    d = DotDict[str](a="hello", b="world")
    
    # Get value
    value = d.get("a")  # "hello"
    default = d.get("missing", "default")  # "default"
    
    # Update
    d.update({"c": "!"})
    
    # Items, keys, values
    items = dict(d.items())
    keys = list(d.keys())
    values = list(d.values())

    # Copy
    d_copy = d.copy()
    d_dict = d.copy_dict()  # returns standard dict

Notes
-----

1. Protected Attributes:
   The following attributes are protected and not included in dictionary operations:
   - ``_missing_attrib_value``
   - ``_internal_keys``
   - ``_is_protocol``

2. Overriding Built-ins:
   It's possible but not recommended to override built-in attributes like ``keys``, ``copy``, and ``items``:

   .. code-block:: python

       d = DotDict[str](a="hello", keys="world")
       print(d.keys)  # "world" (not the keys() method)

Version History
---------------

- Version 0.52.6: Added ``__bool__`` method
- Version 0.52.2: Added generic type support and missing attribute value feature


Class DotDict
-------------

.. autoclass:: ooodev.utils.helper.dot_dict.DotDict
    :members:
    :undoc-members: