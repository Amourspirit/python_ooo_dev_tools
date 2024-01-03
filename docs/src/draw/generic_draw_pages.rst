.. _class_draw_generic_draw_pages:

Class GenericDrawPages
======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The GenericDrawPages class represents the collection of pages in a document.

This class contains several python magic methods to make it behave like a collection.

Getting Number of Pages
^^^^^^^^^^^^^^^^^^^^^^^

To get the number of pages in a draw document, use the built in ``len()`` method:

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc(loader))
    >>> len(doc.draw_pages)
    1

Getting a page
^^^^^^^^^^^^^^

There are several ways to get a page from a draw document.
The simplest is to use the ``[]`` method:

Get by page index.

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc(loader))
    >>> doc.draw_pages[0]
    <ooodev.draw.GenericDrawPage object at 0x7f7f0c0b2b90>

To get the last page in a draw document, use the built in ``[-1]`` method:

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc(loader))
    >>> doc.draw_pages[-1]
    <ooodev.draw.GenericDrawPage object at 0x7f7f0c0b2b90>

Deleting a page
^^^^^^^^^^^^^^^

There are several ways to delete a page from a draw document.
The simplest is to use the ``del`` method:

Delete by page index.

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc(loader))
    >>> del doc.draw_pages[0]

Or by ``DrawPage`` object.

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc(loader))
    >>> pg = doc.draw_pages[0]
    >>> del doc.draw_pages[pg]

or by ``XDrawPage`` object.

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc(loader))
    >>> pg = doc.draw_pages[0]
    >>> del doc.draw_pages[pg.component]

Iterating over pages
^^^^^^^^^^^^^^^^^^^^

To iterate over the pages in a draw document, use the built in ``for`` loop:

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc(loader))
    >>> for pg in doc.slides:
    ...     print(pg)
    <ooodev.draw.GenericDrawPage object at 0x7f7f0c0b2b90>

See Also
--------

- :ref:`class_draw_draw_pages`
- :ref:`class_write_write_doc`

Class Declaration
-----------------

.. autoclass:: ooodev.draw.GenericDrawPages
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: