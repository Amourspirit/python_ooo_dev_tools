.. _class_draw_impress_pages:

Class ImpressPages
==================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The ImpressPages class represents the collection of pages in a document.

This class contains several python magic methods to make it behave like a collection.

Getting Number of Pages
^^^^^^^^^^^^^^^^^^^^^^^

To get the number of pages in a draw document, use the built in ``len()`` method:

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> len(doc.slides)
    1

Getting a page
^^^^^^^^^^^^^^

There are several ways to get a page from a draw document.
The simplest is to use the ``[]`` method:

Get by page index.

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> doc.slides[0]
    <ooodev.draw.ImpressPage object at 0x7f7f0c0b2b90>

Or by page name.

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> doc.slides["page1"]
    <ooodev.draw.ImpressPage object at 0x7f7f0c0b2b90>

To get the last page in a draw document, use the built in ``[-1]`` method:

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> doc.slides[-1]
    <ooodev.draw.ImpressPage object at 0x7f7f0c0b2b90>

Deleting a page
^^^^^^^^^^^^^^^

There are several ways to delete a page from an Impress document.
The simplest is to use the ``del`` method:

Delete by page index.

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> del doc.slides[0]

Or by page name.

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> del doc.slides["page1"]

Or by ``ImpressPage`` object.

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> slide = doc.slides[0]
    >>> del doc.slides[slide]

or by ``XDrawPage`` object.

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> slide = doc.slides[0]
    >>> del doc.slides[slide.component]

Iterating over pages
^^^^^^^^^^^^^^^^^^^^

To iterate over the pages in a draw document, use the built in ``for`` loop:

.. code-block:: python

    >>> doc = ImpressDoc.open_doc("test.odp")
    >>> for slide in doc.slides:
    ...     print(slide)
    <ooodev.draw.ImpressPage object at 0x7f7f0c0b2b90>

Class Declaration
-----------------

.. autoclass:: ooodev.draw.ImpressPages
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: