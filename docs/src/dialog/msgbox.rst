Class MsgBox
============

Straight forward message box for displaying messages on screen.

.. code-block:: python

    from ooodev.dialog.msgbox import MsgBox

    result = MsgBox.msgbox("Are you sure?", boxtype=MsgBox.Type.QUERYBOX, buttons=MsgBox.Buttons.BUTTONS_YES_NO)
    if result == MsgBox.Results.YES:
        print("All is ok")
    elif result == MsgBox.Results.NO:
        print("Cancel is the choice!")

.. autoclass:: ooodev.dialog.msgbox.MsgBox
    :members:
    :undoc-members: