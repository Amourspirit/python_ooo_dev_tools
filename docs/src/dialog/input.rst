Class Input
===========

Input Box
---------

For inputting text and passwords.

.. code-block:: python

    from ooodev.dialog.input import Input

    pwd = Input.get_input(title="Password", msg="Input Password:", is_password=True)
    print(pwd)

.. autoclass:: ooodev.dialog.input.Input
    :members:
    :undoc-members: