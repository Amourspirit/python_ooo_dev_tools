from __future__ import annotations
from typing import overload, TYPE_CHECKING


if TYPE_CHECKING:
    from ooodev.dialog import Dialog
    from ooo.dyn.awt.message_box_type import MessageBoxType
    from ooo.dyn.awt.message_box_buttons import MessageBoxButtonsEnum
    from ooo.dyn.awt.message_box_results import MessageBoxResultsEnum
    from typing_extensions import Protocol
else:
    Protocol = object


class CreateDialogPartialT(Protocol):
    """Type for CreateDialogPartial"""

    def create_dialog(self, x: int, y: int, width: int, height: int, title: str) -> Dialog:
        """
        Creates a dialog.

        Args:
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int): Height. If ``-1``, the dialog Size is not set.
            title (str): Dialog title.

        Returns:
            Dialog: An empty dialog. The dialog contains methods for adding controls.
        """
        ...

    # region msgbox
    @overload
    def msgbox(self, msg: str) -> MessageBoxResultsEnum:
        """
        Simple message box. With a title of "Message" and buttons of ``Buttons.BUTTONS_OK``.

        Args:
            msg (str): the message for display.

        Returns:
            Results: MessageBoxResultsEnum.
        """
        ...

    @overload
    def msgbox(self, msg: str, title: str) -> MessageBoxResultsEnum:
        """
        Simple message box. With buttons of ``Buttons.BUTTONS_OK``.

        Args:
            msg (str): the message for display.
            title (str): the title of the message box.

        Returns:
            Results: MessageBoxResultsEnum.
        """
        ...

    @overload
    def msgbox(self, msg: str, title: str, *, buttons: MessageBoxButtonsEnum | int) -> MessageBoxResultsEnum:
        """
        Simple message box.

        Args:
            msg (str): the message for display.
            title (str, optional): the title of the message box. Defaults to "Message".
            buttons (MessageBoxButtonsEnum, int, optional): determines what buttons to display.

        Returns:
            Results: MessageBoxResultsEnum.
        """
        ...

    @overload
    def msgbox(self, msg: str, title: str, boxtype: MessageBoxType | int) -> MessageBoxResultsEnum:
        """
        Simple message box.

        Args:
            msg (str): the message for display.
            title (str): the title of the message box.
            boxtype (MessageBoxType): determines the type of message box to display.

        Returns:
            Results: MessageBoxResultsEnum.

        Note:
            If BoxType is an integer, the following values are valid:

            - 0: ``MESSAGEBOX``
            - 1: ``INFOBOX``
            - 2: ``WARNINGBOX``
            - 3: ``ERRORBOX``
            - 4: ``QUERYBOX``
        """
        ...

    @overload
    def msgbox(
        self, msg: str, title: str, boxtype: MessageBoxType | int, buttons: MessageBoxButtonsEnum | int
    ) -> MessageBoxResultsEnum:
        """
        Simple message box.

        Args:
            msg (str): the message for display.
            title (str): the title of the message box.
            boxtype (MessageBoxType): determines the type of message box to display.
            buttons (MessageBoxButtonsEnum, int): determines what buttons to display.

        Returns:
            Results: MessageBoxResultsEnum.

        Note:
            If BoxType is an integer, the following values are valid:

            - 0: ``MESSAGEBOX``
            - 1: ``INFOBOX``
            - 2: ``WARNINGBOX``
            - 3: ``ERRORBOX``
            - 4: ``QUERYBOX``
        """
        ...

    # endregion msgbox

    # region input
    def input_box(
        self,
        title: str,
        msg: str,
        input_value: str = "",
        ok_lbl: str = "OK",
        cancel_lbl: str = "Cancel",
        is_password: bool = False,
    ) -> str:
        """
        Displays an input box and returns the results.

        Args:
            title (str): Title for the dialog
            msg (str): Message to display such as "Input your Name"
            input_value (str, optional): Value of input box when first displayed.
            ok_lbl (str, optional): OK button Label. Defaults to "OK".
            cancel_lbl (str, optional): Cancel Button Label. Defaults to "Cancel".
            is_password (bool, optional): Determines if the input box is masked for password input. Defaults to ``False``.

        Returns:
            str: The value of input or empty string.
        """
        ...

    # endregion input
