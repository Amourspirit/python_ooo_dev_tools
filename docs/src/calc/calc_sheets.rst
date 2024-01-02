.. _class_calc_calc_sheets:

Class CalcSheets
================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The CalcSheets class represents the collection of sheets in a Calc document.

This class contains several python magic methods to make it behave like a collection.

Getting Number of Pages
^^^^^^^^^^^^^^^^^^^^^^^

To get the number of pages in a draw document, use the built in ``len()`` method:

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> len(doc.sheets)
    1

Getting a sheet
^^^^^^^^^^^^^^^

There are several ways to get a page from a Calc document.
The simplest is to use the ``[]`` method:

Get Sheet by Index.

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> doc.sheets[0]
    <ooodev.calc.CalcSheet object at 0x7f8b1c0b4a90>

Get Sheet by Name.

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> doc.sheets['Sheet1']
    <ooodev.calc.CalcSheet object at 0x7f8b1c0b4a90>

To get the last sheet in a document, use the ``-1`` index:

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> doc.insert_sheet(name="Sheet2", idx=1)
    >>> doc.sheets[-1]
    <ooodev.calc.CalcSheet object at 0x7f8b1c0b4a90>

Deleting a sheet
^^^^^^^^^^^^^^^^

To delete a sheet, use the ``del`` keyword:

Delete by sheet index.

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> del doc.sheets[0]

Delete by sheet name.

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> del doc.sheets['Sheet1']

Iterating over sheets
^^^^^^^^^^^^^^^^^^^^^

To iterate over the sheets in a document, use the ``for`` keyword:

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> doc.insert_sheet(name="Sheet2")
    >>> doc.insert_sheet(name="Sheet3")
    >>> for sheet in doc.sheets:
    ...     print(sheet.name)
    Sheet1
    Sheet2
    Sheet3

Class Declaration
-----------------

.. autoclass:: ooodev.calc.CalcSheets
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: