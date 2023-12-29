.. _class_draw_draw_pages:

Class DrawPages
===============

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The DrawPages class represents the collection of pages in a document.

This class contains several python magic methods to make it behave like a collection.

Getting Number of Pages
^^^^^^^^^^^^^^^^^^^^^^^

To get the number of pages in a draw document, use the built in ``len()`` method:

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> len(doc.slides)
    1

Getting a page
^^^^^^^^^^^^^^

There are several ways to get a page from a draw document.
The simplest is to use the ``[]`` method:

Get by page index.

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> doc.slides[0]
    <ooodev.draw.DrawPage object at 0x7f7f0c0b2b90>

Or by page name.

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> doc.slides["page1"]
    <ooodev.draw.DrawPage object at 0x7f7f0c0b2b90>

To get the last page in a draw document, use the built in ``[-1]`` method:

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> doc.slides[-1]
    <ooodev.draw.DrawPage object at 0x7f7f0c0b2b90>

Deleting a page
^^^^^^^^^^^^^^^

There are several ways to delete a page from a draw document.
The simplest is to use the ``del`` method:

Delete by page index.

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> del doc.slides[0]

Or by page name.

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> del doc.slides["page1"]

Or by ``DrawPage`` object.

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> slide = doc.slides[0]
    >>> del doc.slides[slide]

or by ``XDrawPage`` object.

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> slide = doc.slides[0]
    >>> del doc.slides[slide.component]

Iterating over pages
^^^^^^^^^^^^^^^^^^^^

To iterate over the pages in a draw document, use the built in ``for`` loop:

.. code-block:: python

    >>> doc = DrawDoc(Draw.create_draw_doc(loader))
    >>> for slide in doc.slides:
    ...     print(slide)
    <ooodev.draw.DrawPage object at 0x7f7f0c0b2b90>

Class Hierarchy
---------------

.. autoclass:: ooodev.draw.DrawPages
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: