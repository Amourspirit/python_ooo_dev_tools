.. _ooodev.draw.DrawForms:

Class DrawForms
================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The DrawForms class represents the collection of forms in a Draw document.

This class contains several python magic methods to make it behave like a collection.

Getting Number of Forms
^^^^^^^^^^^^^^^^^^^^^^^

To get the number of pages in a draw document, use the built in ``len()`` method:

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc())
    >>> draw_page = doc.slides[0]
    >>> len(draw_page.forms)
    0
    >>>draw_page.forms.add_form()
    >>> len(draw_page.forms)
    1

Getting a form
^^^^^^^^^^^^^^

There are several ways to get a form from a Calc sheet.
The simplest is to use the ``[]`` method:

Get Form by Index.

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc())
    >>> draw_page = doc.slides[0]
    >>> if len(draw_page.forms) == 0:
    ...   draw_page.forms.add_form()
    >>> form = draw_page.forms[0]
    <ooodev.draw.DrawForm object at 0x7f8b1c0b4a90>

Get Form by Name.

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc())
    >>> draw_page = doc.slides[0]
    >>> if len(draw_page.forms) == 0:
    ...   draw_page.forms.add_form("MyForm")
    >>> form = draw_page.forms["MyForm"]
    <ooodev.draw.DrawForm object at 0x7f8b1c0b4a90>

To get the last form in a sheet, use the ``-1`` index:

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc())
    >>> draw_page = doc.slides[0]
    >>> if len(draw_page.forms) == 0:
    ...   draw_page.forms.add_form()
    >>> form = draw_page.forms[-1]
    <ooodev.draw.DrawForm object at 0x7f8b1c0b4a90>

Deleting a form
^^^^^^^^^^^^^^^

To delete a form, use the ``del`` keyword:

Delete by form index.

.. code-block:: python

    >>> del doc.slides[0].draw_page.forms[1]

Delete by form name.

.. code-block:: python

    >>> del doc.slides[0].draw_page.forms["MyForm"]

Iterating over forms
^^^^^^^^^^^^^^^^^^^^

To iterate over the forms in a sheet, use the ``for`` keyword:

.. code-block:: python

    >>> for form in doc.slides[0].draw_page.forms:
    ...     print(form.name)
    MyForm

Class Declaration
-----------------

.. autoclass:: ooodev.draw.DrawForms
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: