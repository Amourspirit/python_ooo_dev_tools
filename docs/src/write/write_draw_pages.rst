.. _class_write_draw_pages:

Class WriteDrawPages
====================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The ``WriteDrawPages`` class represents the collection of draw pages in a Writer document.

This class contains several python magic methods to make it behave like a collection.

Getting Number of Pages
^^^^^^^^^^^^^^^^^^^^^^^

To get the number of pages in a draw document, use the built in ``len()`` method:

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> len(doc.draw_pages)
    1

Getting a Draw Page
^^^^^^^^^^^^^^^^^^^

Get Draw Page by Index.

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> doc.draw_pages[0]
    <ooodev.write.WriteDrawPage object at 0x7f8b1c0b4a90>



To get the last sheet in a document, use the ``-1`` index:

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> doc.draw_pages[-1]
    <ooodev.write.WriteDrawPage object at 0x7f8b1c0b4a90>

Alternatively, you can get a draw page direct from :ref:`class_write_write_doc` using the ``draw_page`` property.

Deleting a Draw Page
^^^^^^^^^^^^^^^^^^^^

To delete a sheet, use the ``del`` keyword:

Delete by draw page index.

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> # other code
    >>> del doc.draw_pages[0]


Iterating over Draw Pages
^^^^^^^^^^^^^^^^^^^^^^^^^

To iterate over the draw pages in a document, use the ``for`` keyword:

.. code-block:: python

    >>> doc = WriteDoc(Write.create_doc())
    >>> for draw_page in doc.draw_pages:
    ...     print(draw_page.name)

Class Declaration
-----------------

.. autoclass:: ooodev.write.WriteDrawPages
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: