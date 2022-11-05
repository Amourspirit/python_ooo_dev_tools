.. _class_msg_box:

Class MsgBox
============

Straight forward message box for displaying messages on screen.

.. code-block:: python

    from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum

    result = MsgBox.msgbox(
        "Are you sure?",
        boxtype=MessageBoxType.QUERYBOX,
        buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO
        )

    if result == MessageBoxResultsEnum.Results.YES:
        print("All is OK")
    elif result == MessageBoxResultsEnum.Results.NO:
        print("No is the choice!")

.. autoclass:: ooodev.dialog.msgbox.MsgBox
    :members:
    :undoc-members: