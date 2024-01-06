.. _class_calc_spreadsheet_draw_page:

Class SpreadsheetDrawPage
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The ``SpreadsheetDrawPage`` class is for accessing the draw page of a spreadsheet document.

This class has a ``forms`` property that gives access to :ref:`class_calc_calc_forms`
which in turn gives access to :ref:`class_calc_calc_form` that represents a form in the spreadsheet document.

Adding a shape to a sheet
^^^^^^^^^^^^^^^^^^^^^^^^^

The ``SpreadsheetDrawPage`` class has many methods for adding shapes to a sheet.

Here is an example of adding a rectangle to a sheet.

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> sheet = doc.sheets[0]
    >>> rect = sheet.draw_page.draw_rectangle(x=100, y=10, width=50, height=20)
    >>> rect
    <ooodev.draw.shapes.rectangle_shape.RectangleShape object at 0x000001B636EA3490>


Class Declaration
-----------------

.. autoclass:: ooodev.calc.SpreadsheetDrawPage
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: