.. _help_ooodev.utils.data_type.row_obj.RowObj:

Row Objects
===========

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Working with the :py:class:`ooodev.utils.data_type.row_obj.RowObj` class.

Comparison
----------

Rules
^^^^^

- can be compared to other ``RowObj``
- can be compare to ``int``


Example
^^^^^^^

.. code-block:: python

        >>> from ooodev.utils.data_type.row_obj import RowObj
        >>>
        >>> r1 = RowObj.from_int(1)
        >>> r2 = RowObj(2)
        >>> r2 > r1
        True

        >>> r1 < r2
        True

        >>> r2 <= r2
        True

        >>> r2 >= r2
        True

        >>> r2 >= 2
        True

        >>> 2 <= r2
        True

        >>> r2 == 2
        True

        >>> 2 == r2
        True

        >>> r1 > r2
        False

Previous and Next
-----------------


.. code-block:: python

    >>> from ooodev.utils.data_type.row_obj import RowObj
    >>>
    >>> r6 = RowObj(6)
    >>> r6
    RowObj(value=6)

    >>> r5 = r6.prev
    >>> r5
    RowObj(value=5)

    >>> r5.prev
    RowObj(value=4)

    >>> r2 = r5.prev.prev.prev
    >>> r2
    RowObj(value=2)

    >>> r8 = r6.next.next
    >>> r8
    RowObj(value=8)


Addition and Subtraction
------------------------

Rules
^^^^^

- can add and subtract to other ``RowObj``
- can add and subtract ``int``
- Attempt to make negative ``RowObj`` result in ``IndexError``

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.row_obj import RowObj
    >>>
    >>> r2 = RowObj(2)
    >>> r5 = r2 + 3
    >>> r5
    RowObj(value=5)

    >>> r3 = r5 - 2
    >>> r3
    RowObj(value=3)

    >>> r8 = r5 + r3
    >>> r8
    RowObj(value=8)

    >>> r6 = sum([r2, r2, r2])
    >>> r6
    RowObj(value=6)

    >>> r2 - r5
    IndexError

Multiply and Divide
-------------------

Rules
^^^^^

- can multiply and divide to other RowObj
- can multiply and divide int
- Attempt to make negative RowObj result in IndexError

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.row_obj import RowObj
    >>>
    >>> r2 = RowObj(2)
    >>> r20 = r2 * 10
    >>> r20
    RowObj(value=20)

    >>> r40 = r2 * r20
    >>> r40
    RowObj(value=40)

    >>> r20 = r40 / 2
    >>> r20
    RowObj(value=20)

    >>> r10 = r20 / r2
    >>> r10
    RowObj(value=10)

    >>> r2 / r10
    IndexError
