.. _class_calc_calc_sheet:

Class CalcSheet
===============

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The CalcSheet class represents a Calc spreadsheet document.

This class has index access to cells using the [] method. The index can be:

- a string in the form "A1" or "B2" (column letter followed by row number)
- a tuple (column, row) where column is an integer and row is an integer
- a :py:class:`~ooodev.utils.data_type.cell_obj.CellObj` object
- a UNO ``CellAddress`` object
- a UNO ``XCell`` object

Getting access via cell name
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. code-block:: python

    >>> doc.sheets[0]["A2"].get_val()
    1.0

Getting access via column and row
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Column and row are 0-based indexes.

.. code-block:: python

    >>> doc.sheets[0][(0, 1)].get_val()
    1.0

Getting access via CellObj
^^^^^^^^^^^^^^^^^^^^^^^^^^

``CellObj`` is 1-based for row.

.. code-block:: python

    >>> doc.sheets[0][CellObj("A", 2)].get_val()
    1.0

Class Declaration
-----------------

.. autoclass:: ooodev.calc.CalcSheet
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: