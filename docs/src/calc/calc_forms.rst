.. _ooodev.calc.CalcForms:

Class CalcForms
===============

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The CalcForms class represents the collection of forms in a Calc document.

This class contains several python magic methods to make it behave like a collection.

Getting Number of Forms
^^^^^^^^^^^^^^^^^^^^^^^

To get the number of pages in a draw document, use the built in ``len()`` method:

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> sheet = doc.sheets[0]
    >>> draw_page = sheet.draw_page
    >>> len(draw_page.forms)
    0
    >>> draw_page.forms.add_form()
    >>> len(draw_page.forms)
    1

Getting a form
^^^^^^^^^^^^^^

There are several ways to get a form from a Calc sheet.
The simplest is to use the ``[]`` method:

Get Form by Index.

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> sheet = doc.sheets[0]
    >>> if len(sheet.draw_page.forms) == 0:
    ...    sheet.draw_page.forms.add_form()
    >>> form = sheet.draw_page.forms[0]
    <ooodev.calc.CalcForm object at 0x7f8b1c0b4a90>

Get Form by Name.

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> sheet = doc.sheets[0]
    >>> if len(sheet.draw_page.forms) == 0:
    ...    sheet.draw_page.forms.add_form("MyForm")
    >>> form = sheet.draw_page.forms["MyForm"]
    <ooodev.calc.CalcForm object at 0x7f8b1c0b4a90>

To get the last form in a sheet, use the ``-1`` index:

.. code-block:: python

    >>> doc = CalcDoc(Calc.create_doc(loader))
    >>> sheet = doc.sheets[0]
    >>> if len(sheet.draw_page.forms) == 0:
    ...    sheet.draw_page.forms.add_form()
    >>> form = sheet.draw_page.forms[-1]
    <ooodev.calc.CalcForm object at 0x7f8b1c0b4a90>

Deleting a form
^^^^^^^^^^^^^^^

To delete a form, use the ``del`` keyword:

Delete by form index.

.. code-block:: python

    >>> sheet = doc.sheets[0]
    >>> del sheet.draw_page.forms[1]

Delete by form name.

.. code-block:: python

    >>> sheet = doc.sheets[0]
    >>> del sheet.draw_page.forms["MyForm"]

Iterating over forms
^^^^^^^^^^^^^^^^^^^^

To iterate over the forms in a sheet, use the ``for`` keyword:

.. code-block:: python

    >>> sheet = doc.sheets[0]
    >>> for form in sheet.draw_page.forms:
    ...     print(form.name)
    MyForm

Class Declaration
-----------------

.. autoclass:: ooodev.calc.CalcForms
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: