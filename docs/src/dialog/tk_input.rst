.. _dialog_tk_input:

Class Window
============

Input Box
---------

For inputting text and passwords in a Pythonic way.


Example Usage:

.. code-block:: python

    from ooodev.dialog.tk_input import Window

    pass_inst = Window(title="Password", input_msg="Input Password:", is_password=True)
    pwd = pass_inst.get_input()
    print(pwd)

.. note::

    LibreOffice does not ship with tkinter_ and therefore this option will not work in all setups.
    Generally speaking this can be used in Linux.

.. seealso::

    :py:meth:`.GUI.get_password`

.. autoclass:: ooodev.dialog.tk_input.Window
    :members:
    :undoc-members:

.. _tkinter: https://docs.python.org/3/library/tkinter.html