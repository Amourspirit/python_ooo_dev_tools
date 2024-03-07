.. _help_ooodev.utils.data_type.cell_obj.CellObj:

Cell Objects
============

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Working with the :py:class:`ooodev.utils.data_type.cell_obj.CellObj` class.

Comparison
----------

Rules
^^^^^

- only ``==`` and ``!=`` comparisons are supported
- ``CellObj`` can be compared to ``CellObj``
- ``CellObj`` can be compared to ``str``

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.cell_obj import CellObj
    >>> 
    >>> b2 = CellObj.from_cell("B2")
    >>> b2
    CellObj(col='B', row=2, sheet_idx=-1)

    >>> b4 = CellObj("b", 4)
    >>> b4 == b2
    False

    >>> b4 != b2 
    True

    >>> b4 == "b4"
    True

    >>> b4 == "B4" 
    True

Right, Left, Up, Down
---------------------

Rules
^^^^^

- Going out of range results in ``IndexError``

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.cell_obj import CellObj 
    >>>
    >>> b4 = CellObj("B", 4, 0)
    >>> b4.right # gets cell to the right
    CellObj(col='C', row=4, sheet_idx=0)

    >>> b4.left # get cell to the left
    CellObj(col='A', row=4, sheet_idx=0)

    >>> b4.up # get cell above
    CellObj(col='B', row=3, sheet_idx=0)

    >>> b4.down # get cell below
    CellObj(col='B', row=5, sheet_idx=0)

    >>> b4.down.right
    CellObj(col='C', row=5, sheet_idx=0)

    >>> b4.down.right.right
    CellObj(col='D', row=5, sheet_idx=0)

    >>> b4.left.left # out of range
    IndexError

Addition and Subtraction
------------------------

Rules
^^^^^

- Adding an ``int`` to ``CellObj`` gets the cell to the down by the ``int`` amount
- Subtracting an ``int`` from ``CellObj`` gets the cell to the up by the ``int`` amount
- Adding a ``str`` (column) to ``CellObj`` gets the cell to the right by the ``str`` amount
- Subtracting a ``str`` (column) from ``CellObj`` gets the cell to the left by the ``str`` amount
- ``RowObj`` and ``ColObj`` can be added and subtracted
- ``CellObj`` can be added and subtracted
- Going out of range results in ``IndexError``

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.cell_obj import CellObj 
    >>> 
    >>> b4 = CellObj("B", 4, 0)
    >>> c6 = CellObj("C", 6, 0) 
    >>> b4 + 3
    CellObj(col='B', row=7, sheet_idx=0)

    >>> b4 - 3
    CellObj(col='B', row=1, sheet_idx=0)

    >>> b4 + "C" # get E4 by adding 3 col
    CellObj(col='E', row=4, sheet_idx=0)

    >>> b4 - "A" # get A4 by subtracting 1 col
    CellObj(col='A', row=4, sheet_idx=0)

    >>> b4 + (b4.row_obj + 5) # get B9
    CellObj(col='B', row=9, sheet_idx=0)

    >>> b4 + b4.row_obj.next # same as b4.down
    CellObj(col='B', row=5, sheet_idx=0)

    >>> b4 + (b4.col_obj + 5) # get G4
    CellObj(col='G', row=4, sheet_idx=0)

    >>> b4 + b4.col_obj.next # same as b4.right
    CellObj(col='C', row=4, sheet_idx=0)

    >>> b4.right            
    CellObj(col='C', row=4, sheet_idx=0)

    >>> b4 - b4.col_obj.prev # same as b4.left
    CellObj(col='A', row=4, sheet_idx=0)

    >>> b4.left
    CellObj(col='A', row=4, sheet_idx=0)

    >>> b4 + c6 # get E10 add 3 col and 6 row to b4
    CellObj(col='E', row=10, sheet_idx=0)

    >>> c6 - b4 # get A2 subtract 2 col and 4 row from c6
    CellObj(col='A', row=2, sheet_idx=0)

    >>> b4 - (b4.col_obj - 2)
    IndexError

Multiply and Divide
-------------------

Rules
^^^^^

- Multiplying an ``int`` to ``CellObj`` gets the cell to the down
- Dividing an ``int`` from ``CellObj`` gets the cell to the up
- Multiplying a ``str`` (column) to ``CellObj`` gets the cell to the right
- Dividing a ``str`` (column) from ``CellObj`` gets the cell to the left
- ``RowObj`` and ``ColObj`` can be multiplied and divided
- ``CellObj`` can be multiplied and divided
- If ``CellObj`` division results in a fraction then rounding is used (9 / 4 = 2)
- Going out of range results in ``IndexError``

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.cell_obj import CellObj 
    >>> 
    >>> f10 = CellObj("F", 10, 0)
    >>> b4 = CellObj("B", 4, 0)
    >>> b2 = CellObj("B", 2, 0)
    >>> f10 * 3 # multiply row by 3
    CellObj(col='F', row=30, sheet_idx=0)

    >>> f10  / 2 # divide row by 2
    CellObj(col='F', row=5, sheet_idx=0)

    >>> f10  * "C" # multiply col by 3
    CellObj(col='R', row=10, sheet_idx=0)

    >>> f10  / "B" # divided col by 2
    CellObj(col='C', row=10, sheet_idx=0)

    >>> f10  * (f10.row_obj * 10) # times 10 rows
    CellObj(col='F', row=100, sheet_idx=0)

    >>> f10  * (f10.col_obj * 10) # times 10 cols
    CellObj(col='BH', row=10, sheet_idx=0)

    >>> f10  / (f10.row_obj / 2) # get F5
    CellObj(col='F', row=5, sheet_idx=0)

    >>> f10  / (f10.col_obj / 2) # get C10
    CellObj(col='C', row=10, sheet_idx=0)

    >>> b4 * f10 # b(2) X f(6), 4 X 10
    CellObj(col='L', row=40, sheet_idx=0)

    >>> f10 / b2 # f(6) / b(2), 10 / 2
    CellObj(col='C', row=5, sheet_idx=0)

    >>> f10 / b4 # f(6) / b(4), 10 / 4, Rounding is used
    CellObj(col='C', row=2, sheet_idx=0)

    >>> b2 / f10 
    IndexError