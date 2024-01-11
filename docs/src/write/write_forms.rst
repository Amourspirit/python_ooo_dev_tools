.. _class_write_write_forms:

Class WriteForms
================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The WriteForms class represents the collection of forms in a Writer document.

This class contains several python magic methods to make it behave like a collection.

Getting Number of Forms
^^^^^^^^^^^^^^^^^^^^^^^

To get the number of pages in a draw document, use the built in ``len()`` method:

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> len(doc.draw_page.forms)
    0
    >>>doc.draw_page.forms.add_form()
    >>> len(doc.draw_page.forms)
    1

Getting a form
^^^^^^^^^^^^^^

There are several ways to get a form from a Calc sheet.
The simplest is to use the ``[]`` method:

Get Form by Index.

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> if len(doc.draw_page.forms) == 0:
    ...   doc.draw_page.forms.add_form()
    >>> form = doc.draw_page.forms[0]
    <ooodev.write.WriteForm object at 0x7f8b1c0b4a90>

Get Form by Name.

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> if len(doc.draw_page.forms) == 0:
    ...    doc.draw_page.forms.add_form("MyForm")
    >>> form = doc.draw_page.forms["MyForm"]
    <ooodev.write.WriteForm object at 0x7f8b1c0b4a90>

To get the last form in a sheet, use the ``-1`` index:

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> if len(doc.draw_page.forms) == 0:
    ...    doc.draw_page.forms.add_form()
    >>> form = doc.draw_page.forms[-1]
    <ooodev.write.WriteForm object at 0x7f8b1c0b4a90>

Deleting a form
^^^^^^^^^^^^^^^

To delete a form, use the ``del`` keyword:

Delete by form index.

.. code-block:: python

    >>> del doc.draw_page.forms[1]

Delete by form name.

.. code-block:: python

    >>> del doc.draw_page.forms["MyForm"]

Iterating over forms
^^^^^^^^^^^^^^^^^^^^

To iterate over the forms in a sheet, use the ``for`` keyword:

.. code-block:: python

    >>> for form in doc.draw_page.forms:
    ...     print(form.name)
    MyForm

Class Declaration
-----------------

.. autoclass:: ooodev.write.WriteForms
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: