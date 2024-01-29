.. _ooodev.write.WriteDrawPage:

Class WriteDrawPage
===================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The ``WriteDrawPage`` class is for accessing the draw page of a Writer document.

This class has a ``forms`` property that gives access to :ref:`ooodev.write.WriteForms`
which in turn gives access to :ref:`ooodev.write.WriteForm` that represents a form in the document.

Adding a shape to a sheet
^^^^^^^^^^^^^^^^^^^^^^^^^

The ``WriteDrawPage`` class has many methods for adding shapes to a sheet.

Here is an example of adding a rectangle to a sheet.

.. code-block:: python

    >>> doc = doc = WriteDoc(Write.create_doc())
    >>> rect = doc.draw_page.draw_rectangle(x=100, y=10, width=50, height=20)
    >>> rect
    <ooodev.draw.shapes.rectangle_shape.RectangleShape object at 0x000001B636EA3490>


Class Declaration
-----------------

.. autoclass:: ooodev.write.WriteDrawPage
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: