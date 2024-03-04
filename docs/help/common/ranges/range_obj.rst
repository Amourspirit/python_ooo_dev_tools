
.. _help_ooodev.utils.data_type.range_obj.RangeObj:

Range Objects
=============

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Working with the :py:class:`ooodev.utils.data_type.range_obj.RangeObj` class.

Rules
-----

- Can add and subtract ``int``
- Can add and subtract ``str`` where ``str`` is treated as column name
- Can add and subtract ``RowObj``
- Can add and subtract ``ColObj``
- Can add and subtract ``CellObj``
- Adding or subtracting results in a new ``RangeObj``
- Adding/subtracting a number to a ``RangeObj`` is treated differently then adding/subtracting a ``RangeObj`` to a number.
- Adding/subtracting a string (column) to a ``RangeObj`` is treated differently then adding/subtracting a ``RangeObj`` to a string.
- Adding/subtracting ``RowObj``, ``ColObj``, or ``CellObj`` to ``RangeObj`` is treated differently then adding/subtracting ``RangeObj`` to ``RowObj``, ``ColObj``, or ``CellObj``.
- ``RangeObj`` can be combined using the ``/`` operator.


Addition
--------

Example
^^^^^^^

Adding Rows using integers
""""""""""""""""""""""""""

Adding positive number to ``RangeObj`` results in rows being append to end or range.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="A", col_end="C", row_start=1, row_end=3, sheet_idx=0)
    >>> str(rng1)
    'A1:C3'
    >>> rng1.row_count 
    3
    >>> rng2 = rng1 + 5
    >>> str(rng2)
    'A1:C8'
    >>> repr(rng2)
    "RangeObj(col_start='A', col_end='C', row_start=1, row_end=8, sheet_idx=0)"
    >>> rng2.row_count  
    8

Start Range ``A1:C3``

.. image:: https://user-images.githubusercontent.com/4193389/213174796-1fc4e447-116b-458f-95c6-94971ea331b8.png
    :alt: Start Range A1:C3

Result `A1:C8`

.. image:: https://user-images.githubusercontent.com/4193389/213175212-9065fd3c-3c84-46c3-aa7f-edb00c7007df.png
    :alt: Result A1:C8

Adding ``RangeObj`` to a positive number result is new rows being added at the start of the range.

Starting with ``A10:C15`` we end up with ``A5:C15``.

Note that in this example the number comes before ``rng1``.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="A", col_end="C", row_start=10, row_end=15, sheet_idx=0)
    >>> str(rng1)
    'A10:C15'
    >>> rng1.row_count
    6
    >>> rng2 = 5 + rng1
    >>> str(rng2)
    'A5:C15'
    >>> repr(rng2)
    "RangeObj(col_start='A', col_end='C', row_start=5, row_end=15, sheet_idx=0)"
    >>> rng2.row_count 
    11

Start Range ``A10:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213176936-a584dda5-48be-428c-9d5a-83677cb85584.png
    :alt: Start Range A10:C15

Result ``A5:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213177289-f249bab6-9f52-4a12-8dfa-27adbdff08f6.png
    :alt: Result A5:C15

**Adding a negative number.**

Adding a negative number to  ``RangeObj`` results in rows being removed from the end of the range.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="A", col_end="C", row_start=10, row_end=15, sheet_idx=0)
    >>> str(rng1)
    'A10:C15'
    >>> rng1.row_count 
    6
    >>> rng2 = rng1 + -5
    >>> str(rng2)
    'A10:C10'
    >>> repr(rng2)
    "RangeObj(col_start='A', col_end='C', row_start=10, row_end=10, sheet_idx=0)"
    >>> rng2.row_count
    1

Start Range ``A10:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213176936-a584dda5-48be-428c-9d5a-83677cb85584.png
    :alt: Start Range A10:C15

Result ``A10:C10``

.. image:: https://user-images.githubusercontent.com/4193389/213177732-96aef023-fa29-4e5b-a5e9-7676ee938a39.png
    :alt: Result A10:C10

Adding a ``RangeObj`` to a negative number results in rows being added to the end of the range.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="A", col_end="C", row_start=10, row_end=15, sheet_idx=0)
    >>> str(rng1)
    'A10:C15'
    >>> rng1.row_count 
    6
    >>> rng2 = -5 - rng1
    >>> str(rng2)
    'A5:C15'
    >>> repr(rng2)
    "RangeObj(col_start='A', col_end='C', row_start=5, row_end=15, sheet_idx=0)"
    >>> rng2.row_count
    11

Start Range ``A10:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213176936-a584dda5-48be-428c-9d5a-83677cb85584.png
    :alt: Start Range A10:C15

Result ``A5:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213177289-f249bab6-9f52-4a12-8dfa-27adbdff08f6.png
    :alt: Result A5:C15

Adding Rows using RowObj
""""""""""""""""""""""""

``RowObj`` instances can also be used to add rows to a ``RangeObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.row_obj import RowObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.row_count 
    6
    >>> rng2 = rng1 + RowObj.from_int(2)
    >>> str(rng2)
    'F10:H17'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='H', row_start=10, row_end=17, sheet_idx=0)"
    >>> rng2.row_count
    8

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15


Result ``F10:H17``

.. image:: https://user-images.githubusercontent.com/4193389/213185159-80af94ed-da18-41dc-b097-efca3d4bf0b2.png
    :alt: Result F10:H17


Adding RangeObj instance to RowObj instance
"""""""""""""""""""""""""""""""""""""""""""

``RangeObj`` instances can also be used to add rows to a ``RowObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.row_obj import RowObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.row_count 
    6
    >>> rng2 = RowObj.from_int(2) + rng1
    >>> str(rng2)
    'F8:H15'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='H', row_start=8, row_end=15, sheet_idx=0)"
    >>> rng2.row_count
    8

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``F8:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213184844-081d21d7-24d6-406c-bfdd-05fe380c8090.png
    :alt: Result F8:H15

Adding Columns using col string
"""""""""""""""""""""""""""""""

Adding columns is accomplished by adding a column letter such as ``C`` to add three columns.

Adding column to ``RangeObj`` results in columns being added to the right of the range.

Add 3 columns to the right of ``RangeObj``.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.col_count
    3
    >>> rng2 = rng1 + "C"
    >>> str(rng2)
    'F10:K15'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='K', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.col_count
    6

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``F10:K15``

.. image:: https://user-images.githubusercontent.com/4193389/213186367-211f861b-31b1-45be-8e86-5e973ebbd91c.png
    :alt: Result F10:K15

Adding column to ``RangeObj`` results in columns being added to the right of the range.

Add 3 columns to the left of ``RangeObj``.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.col_count
    3
    >>> rng2 = "C" + rng1
    >>> str(rng2)
    'C10:H15'
    >>> repr(rng2)
    "RangeObj(col_start='C', col_end='H', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.col_count
    6

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``C10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182305-fcc0cbd6-8c3b-42e0-adf5-52f85e112dfa.png)
    :alt: Result C10:H15

Adding Columns using ColObj
"""""""""""""""""""""""""""

``ColObj`` instances can also be used to add rows to a ``RangeObj`` instance.

Adding ``ColObj`` instance to ``RangeObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.col_obj import ColObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.col_count
    3
    >>> rng2 = rng1 + ColObj.from_int(2)
    >>> str(rng2)
    'F10:J15'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='J', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.col_count 
    5

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``F10:J15``

.. image:: https://user-images.githubusercontent.com/4193389/213183668-e007fcb0-5f86-4be2-82e3-376ab7c79097.png
    :alt: Result F10:J15

Adding RangeObj instance to ColObj instance
"""""""""""""""""""""""""""""""""""""""""""

``RangeObj`` instances can also be used to add rows to a ``ColObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.col_obj import ColObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.col_count
    3
    >>> rng2 = rng1 + ColObj.from_int(2)
    >>> str(rng2)
    'F10:J15'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='J', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.col_count 
    5

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``F10:J15``

.. image:: https://user-images.githubusercontent.com/4193389/213183668-e007fcb0-5f86-4be2-82e3-376ab7c79097.png
    :alt: Result F10:J15

Adding Columns and rows using CellObj
"""""""""""""""""""""""""""""""""""""

Add CellObj instance to RowObj instance
"""""""""""""""""""""""""""""""""""""""

``CellObj`` instances can also be used to add rows to a ``RowObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.cell_obj import CellObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.row_count
    6
    >>> rng1.col_count
    3
    >>> rng2 = rng1 + CellObj.from_idx(1, 1)
    >>> str(rng2)
    'F10:J17'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='J', row_start=10, row_end=17, sheet_idx=0)"
    >>> rng2.row_count 
    8
    >>> rng2.col_count 
    5

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``F10:J17``

.. image:: https://user-images.githubusercontent.com/4193389/213184226-7beb6e1a-fb7c-4317-9991-06660b01782d.png
    :alt: Result F10:J17


Add RowObj instance to CellObj instance
"""""""""""""""""""""""""""""""""""""""

``RowObj`` instances can also be used to add rows to a ``CellObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.cell_obj import CellObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.row_count
    6
    >>> rng1.col_count
    3
    >>> rng2 = CellObj.from_idx(1, 1) + rng1
    >>> str(rng2)
    'D8:H15'
    >>> repr(rng2)
    "RangeObj(col_start='D', col_end='H', row_start=8, row_end=15, sheet_idx=0)"
    >>> rng2.row_count 
    8
    >>> rng2.col_count 
    5

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``D8:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213184613-1bc906f9-e82c-44ba-86f0-b3e635762fb6.png
    :alt: Result D8:H15

Subtraction
-----------

Example
^^^^^^^

Subtracting Rows using integer
""""""""""""""""""""""""""""""

Subtracting positive number  from ``RangeObj`` results in rows being removed from end of range.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="A", col_end="C", row_start=10, row_end=20, sheet_idx=0)
    >>> str(rng1)
    'A10:C20'
    >>> rng1.row_count
    11
    >>> rng2 = rng1 - 5
    >>> str(rng2)
    'A10:C15'
    >>> repr(rng2)
    "RangeObj(col_start='A', col_end='C', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.row_count
    6

Start Range ``A10:C20``

.. image:: https://user-images.githubusercontent.com/4193389/213186678-78627305-cfe3-4925-b05c-7d258c81447c.png
    :alt: Start Range A10:C20

Result ``A10:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213186921-05cafa40-0f7a-49cf-91e8-e960ba81f122.png
    :alt: Result A10:C15

Subtracting ``RangeObj`` from a positive number results in rows being remove from start of range.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="A", col_end="C", row_start=10, row_end=20, sheet_idx=0)
    >>> str(rng1)
    'A10:C20'
    >>> rng1.row_count
    11
    >>> rng2 = 5 - rng1
    >>> str(rng2)
    'A15:C20'
    >>> repr(rng2)
    "RangeObj(col_start='A', col_end='C', row_start=15, row_end=20, sheet_idx=0)"
    >>> rng2.row_count
    6

Start Range ``A10:C20``

.. image:: https://user-images.githubusercontent.com/4193389/213186678-78627305-cfe3-4925-b05c-7d258c81447c.png
    :alt: Start Range A10:C20

Result ``A15:C20``

.. image:: https://user-images.githubusercontent.com/4193389/213187281-3e4ece14-3bc8-491d-a1ae-2b676a080219.png
    :alt: Result A15:C20

Subtracting negative number  from ``RangeObj`` results in rows being added to end of range.
The same as adding a positive number.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="A", col_end="C", row_start=10, row_end=15, sheet_idx=0)
    >>> str(rng1)
    'A10:C15'
    >>> rng1.row_count
    6
    >>> rng2 = rng1 - -5
    >>> str(rng2)
    'A10:C20'
    >>> repr(rng2)
    "RangeObj(col_start='A', col_end='C', row_start=10, row_end=20, sheet_idx=0)"
    >>> rng2.row_count
    11

Start Range ``A10:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213186921-05cafa40-0f7a-49cf-91e8-e960ba81f122.png
    :alt: Start Range A10:C15

Result ``A10:C20``

.. image:: https://user-images.githubusercontent.com/4193389/213186678-78627305-cfe3-4925-b05c-7d258c81447c.png
    :alt: Result A10:C20

Subtracting ``RangeObj``  from negative number results in rows being subtracted from start of range.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> 
    >>> rng1 = RangeObj(col_start="A", col_end="C", row_start=10, row_end=15, sheet_idx=0)
    >>> str(rng1)
    'A10:C15'
    >>> rng1.row_count
    6
    >>> rng2 = -3 - rng1
    >>> str(rng2)
    'A7:C15'
    >>> repr(rng2)
    "RangeObj(col_start='A', col_end='C', row_start=7, row_end=15, sheet_idx=0)"
    >>> rng2.row_count
    9

Start Range ``A10:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213186921-05cafa40-0f7a-49cf-91e8-e960ba81f122.png
    :alt: Start Range A10:C15

Result ``A7:C15``

.. image:: https://user-images.githubusercontent.com/4193389/213187948-2873c43f-432c-43b7-8418-0e1f718f9356.png
    :alt: Result A7:C15

Subtracting Rows using RowObj
"""""""""""""""""""""""""""""

``RowObj`` instances can also be used to subtract rows from a ``RangeObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.row_obj import RowObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.row_count 
    6
    >>> rng2 = rng1 - RowObj.from_int(2)
    >>> str(rng2)
    'F10:H13'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='H', row_start=10, row_end=13, sheet_idx=0)"
    >>> rng2.row_count
    4

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``F10:H13``

.. image:: https://user-images.githubusercontent.com/4193389/213188390-19250f8b-237e-452d-9f60-1289531cf805.png
    :alt: Result F10:H13

Subtracting ``RangeObj`` instance from ``RowObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.row_obj import RowObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.row_count 
    6
    >>> rng2 = RowObj.from_int(2) - rng1
    >>> str(rng2)
    'F12:H15'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='H', row_start=12, row_end=15, sheet_idx=0)"
    >>> rng2.row_count
    4

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``F12:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213188838-394915fd-6e55-4c22-a76b-303b9a666d4d.png
    :alt: Result F12:H15

Subtracting Columns using column string
"""""""""""""""""""""""""""""""""""""""

Subtract column string from ``RangeObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.col_count 
    3
    >>> rng2 = rng1 - "B"
    >>> str(rng2)
    'F10:F15'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='F', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.col_count 
    1

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``F10:F15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Result F10:F15

Subtract ``RangeObj`` instance from column string.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.col_count 
    3
    >>> rng2 = "B" - rng1
    >>> str(rng2)
    'H10:H15'
    >>> repr(rng2)
    "RangeObj(col_start='H', col_end='H', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.col_count
    1

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``H10:H15``

.. image:: https://user-images.githubuserc.com/4193389/213189756-f97a388a-dac6-46b1-8021-f75c161f0523.png
    :alt: Result H10:H15

Subtracting Columns using ColObj
""""""""""""""""""""""""""""""""

Subtract ``ColObj`` instance from ``RangeObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.col_obj import ColObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.col_count 
    3
    >>> rng2 = rng1 - ColObj.from_int(2)
    >>> str(rng2)
    'F10:F15'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='F', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.col_count
    1

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15


Result ``F10:F15``

.. image:: https://user-images.githubusercontent.com/4193389/213190272-88ac2548-8454-4cf4-b964-4e14083d2012.png
    :alt: Result F10:F15

Subtracting ``RangeObj`` instance from ``ColObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.col_obj import ColObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.col_count 
    3
    >>> rng2 = ColObj.from_int(2) - rng1
    >>> str(rng2)
    'H10:H15'
    >>> repr(rng2)
    "RangeObj(col_start='H', col_end='H', row_start=10, row_end=15, sheet_idx=0)"
    >>> rng2.col_count
    1

Start Range ``F10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213182893-42576a34-1cab-468f-b258-575a41c7974e.png
    :alt: Start Range F10:H15

Result ``H10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213189756-f97a388a-dac6-46b1-8021-f75c161f0523.png
    :alt: Result H10:H15

Subtracting Columns using CellObj
"""""""""""""""""""""""""""""""""

Subtracting ``CellObj`` instance from ``RangeObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.cell_obj import CellObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.row_count
    6
    >>> rng1.col_count 
    3
    >>> rng2 = rng1 - CellObj.from_idx(1, 1)
    >>> str(rng2)
    'F10:F13'
    >>> repr(rng2)
    "RangeObj(col_start='F', col_end='F', row_start=10, row_end=13, sheet_idx=0)"
    >>> rng2.row_count
    4
    >>> rng2.col_count
    1

Start Range ``H10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213189756-f97a388a-dac6-46b1-8021-f75c161f0523.png
    :alt: Start Range H10:H15

Result ``F10:F13``

.. image:: https://user-images.githubusercontent.com/4193389/213190980-c307a840-4cde-435a-8b8b-f66ab9efd418.png
    :alt: Result F10:F13

Subtracting ``RangeObj`` instance from ``CellObj`` instance.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>> from ooodev.utils.data_type.cell_obj import CellObj
    >>>
    >>> rng1 = RangeObj(col_start="F", col_end="H", row_start=10, row_end=15, sheet_idx=0) 
    >>> str(rng1)
    'F10:H15'
    >>> rng1.row_count
    6
    >>> rng1.col_count 
    3
    >>> rng2 = CellObj.from_idx(1, 1) - rng1
    >>> str(rng2)
    'H12:H15'
    >>> repr(rng2)
    "RangeObj(col_start='H', col_end='H', row_start=12, row_end=15, sheet_idx=0)"
    >>> rng2.row_count
    4
    >>> rng2.col_count
    1

Start Range ``H10:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213189756-f97a388a-dac6-46b1-8021-f75c161f0523.png
    :alt: Start Range H10:H15

Result ``H12:H15``

.. image:: https://user-images.githubusercontent.com/4193389/213191322-c8ff9854-fd7f-4450-96a4-df0205376799.png
    :alt: Result H12:H15


Combining
---------

Combine (merging) of ``RangeObj`` is done using the ``/`` operator (similar to ``Path``).

Example
^^^^^^^

Combing two RowObj's
""""""""""""""""""""

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>>
    >>> rng1 = RangeObj(col_start="C", col_end="F", row_start=3, row_end=6, sheet_idx=0)
    >>> str(rng1)
    'C3:F6'
    >>> rng2 = RangeObj(col_start="C", col_end="F", row_start=1, row_end=2, sheet_idx=0)
    >>> str(rng2)
    'C1:F2'
    >>> rng3 = rng1 / rng2
    >>> str(rng3)
    'C1:F6'

First Range ``C3:F6``

.. image:: https://user-images.githubusercontent.com/4193389/213155290-72f9f679-b1f6-4295-b66a-bf37f8e7009f.png
    :alt: First Range C3:F6

Second Range ``C1:F2``

.. image:: https://user-images.githubusercontent.com/4193389/213154603-28a7e490-a455-4c09-bffa-d8ed373073f8.png
    :alt: Second Range C1:F2

Combined: ``C1:F6``

.. image:: https://user-images.githubusercontent.com/4193389/213154884-dbb653f7-e7c4-4a85-a6e3-f25970ea8407.png
    :alt: Combined C1:F6

Combining Separate Ranges
"""""""""""""""""""""""""

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>>
    >>> rng1 = RangeObj(col_start="A", col_end="B", row_start=2, row_end=4, sheet_idx=0)
    >>> str(rng1)
    'A2:B4'
    >>> rng2 = RangeObj(col_start="C", col_end="F", row_start=6, row_end=8, sheet_idx=0)
    >>> str(rng2)
    'C6:F8'
    >>> rng3 = rng1 / rng2
    >>> str(rng3)
    'A2:F8'

First Range ``A2:B4``

.. image:: https://user-images.githubusercontent.com/4193389/213160729-d1739d5d-549a-4c43-a620-bd363377a103.png
    :alt: First Range A2:B4

Second Range ``C6:F8``

.. image:: https://user-images.githubusercontent.com/4193389/213161053-f1217c41-6ee7-47d2-b825-36f8d1dd151d.png
    :alt: Second Range C6:F8

Combined: ``A2:F8``

.. image:: https://user-images.githubusercontent.com/4193389/213161509-741834b9-8094-452e-b57f-2bf8cc68e5a7.png
    :alt: Combined A2:F8

Combining Ranges With String Range
""""""""""""""""""""""""""""""""""

``RangeObj`` can be combined with String ranges and vice versa.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>>
    >>> rng1 = RangeObj(col_start="A", col_end="B", row_start=2, row_end=4, sheet_idx=0)
    >>> str(rng1)
    'A2:B4'
    >>> rng2 = rng1 / 'C6:F8'
    >>> str(rng2)
    'A2:F8'

``rng2 = 'C6:F8' / rng1`` is also valid.

First Range ``A2:B4``

.. image:: https://user-images.githubusercontent.com/4193389/213160729-d1739d5d-549a-4c43-a620-bd363377a103.png
    :alt: First Range A2:B4

Second Range ``C6:F8``

.. image:: https://user-images.githubusercontent.com/4193389/213161053-f1217c41-6ee7-47d2-b825-36f8d1dd151d.png
    :alt: Second Range C6:F8

Combined: ``A2:F8``

.. image:: https://user-images.githubusercontent.com/4193389/213161509-741834b9-8094-452e-b57f-2bf8cc68e5a7.png
    :alt: Combined A2:F8

Combining multiple ranges
"""""""""""""""""""""""""

Multiple ranges can be combined.

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>>
    >>> rng1 = RangeObj(col_start="A", col_end="B", row_start=2, row_end=4, sheet_idx=0)
    >>> rng2 = RangeObj(col_start="C", col_end="F", row_start=6, row_end=8, sheet_idx=0)
    >>> rng3 = RangeObj(col_start="J", col_end="L", row_start=7, row_end=14, sheet_idx=0)
    >>> rng4 = rng1 / rng2 / rng3 / "K12:O22"
    >>> str(rng4)
    'A2:O22'

.. _help_ooodev.utils.data_type.range_obj.RangeObj.contains:

Checking if value in Range
--------------------------

To check if a cell is in a range, use the ``in`` operator.

The ``in`` operator checks if a cell is in a range and can check the same values as the
:py:meth:`~ooodev.utils.data_type.range_obj.contains` method.

Acceptable values are:

- ``CellObj``
- ``CellAddress``
- ``CellValues``
- ``str`` in the format ``"A1"``

Example:

.. code-block:: python

    >>> from ooodev.utils.data_type.range_obj import RangeObj
    >>>
    >>> rng = RangeObj.from_range("AA2:AB7")
    >>> print("AA3" in rng)
    True

.. _help_ooodev.utils.data_type.range_obj.RangeObj.__iter__:

Iterating over Range
---------------------

To iterate over a range, use the ``for`` loop.

The iteration is done in a column-major order, meaning that the cells are iterated over by column, then by row.

.. code-block:: python

    # each cell is an instance of CellObj
    >>> rng = RangeObj.from_range("A1:C4")
    >>> for cell in rng:
    >>>     print(cell)
    A1
    B1
    C1
    A2
    B2
    C2
    A3
    B3
    C3
    A4
    B4
    C4
