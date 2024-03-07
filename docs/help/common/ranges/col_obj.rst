.. _help_ooodev.utils.data_type.col_obj.ColObj:

Column Objects
==============

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Working with the :py:class:`ooodev.utils.data_type.col_obj.ColObj` class.

Comparison
----------

Rules
^^^^^

- Can be compared to other ``ColObj`` instances
- Can be compared ``int`` where ``int`` is treated as ``one-based`` column number
- Can be compared to ``str`` where ``str`` is treated as a column name

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.col_obj import ColObj
    >>> 
    >>> a_col = ColObj('A')
    >>> c_col  = ColObj("C") 
    >>> c_col > a_col
    >>> 
    True

    >>> c_col >= a_col 
    True

    >>> c_col >= c_col 
    True

    >>> c_col < a_col  
    False

    >>> a_col < c_col
    True

    >>> a_col > c_col 
    False

    >>> c_col > "B"   
    True

    >>> c_col > "C" 
    False

    >>> c_col >= "C"
    True

    >>> "d" > c_col 
    True

    >>> c_col == 3 
    True

    >>> 4 < c_col  
    False


Previous and Next
-----------------

Rules
^^^^^

- Going out of range results in ``IndexError``

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.col_obj import ColObj
    >>> 
    >>> c_col  = ColObj("C")
    >>> c_col.next
    ColObj(value='D')

    >>> c_col.next.next
    ColObj(value='E')

    >>> c_col.prev     
    ColObj(value='B')

    >>> a_col = c_col.prev.prev
    >>> a_col.prev
    IndexError

Addition and Subtraction
------------------------

Rules
^^^^^

- Can add and subtract to other ``ColObj`` instances
- Can add and subtract ``int``
- Can add and subtract ``str`` where ``str`` is treated as column name
- Attempt to make negative ``ColObj`` result in ``IndexError``

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.col_obj import ColObj
    >>> a_col = ColObj("A")
    >>> a_col + 2
    ColObj(value='C')

    >>> e_col = a_col + 4
    >>> a_col + e_col
    ColObj(value='F')

    >>> e_col - a_col 
    ColObj(value='D')

    >>> e_col - 2
    ColObj(value='C')

    >>> e_col - "B" # minus 2 col
    ColObj(value='C')

    >>> e_col + 6  
    ColObj(value='K')

    >>> e_col + "F" # add 6 col
    ColObj(value='K')

    >>> "F" + e_col
    ColObj(value='K')

    >>> 12 - e_col
    ColObj(value='G')

    >>> "C" + e_col # add 3 col
    ColObj(value='H')

    >>> a_col - e_col 
    IndexError

Multiply and Divide
-------------------

Rules
^^^^^

- Can multiply and divide to other ``ColObj`` instances
- Can multiply and divide ``int``
- Can multiply and divide ``str`` where ``str`` is treated as column name
- Attempt to make negative ``ColObj`` result in ``IndexError``

Example
^^^^^^^

.. code-block:: python

    >>> from ooodev.utils.data_type.col_obj import ColObj
    >>>
    >>> b_col = ColObj("B")
    >>> f_col = b_col * 3 # 2 X 3
    >>> f_col
    ColObj(value='F') # col 6

    >>> f_col / 3 # 6 / 2
    ColObj(value='B') # col 2

    >>> f_col / b_col # 6 / 2
    ColObj(value='C') # col 3

    >>> f_col * b_col # 6 X 2
    ColObj(value='L') # col 12

    >>> f_col * "C" # 6 X 3
    ColObj(value='R') # col 18

    >>> f_col / "C"  # 6 / 3
    ColObj(value='B') # col 2

    >>> f_col / 7 # 6 / 7
    IndexError
