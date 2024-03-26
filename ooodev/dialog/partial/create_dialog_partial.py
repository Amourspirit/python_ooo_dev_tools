from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno
from ooo.dyn.awt.message_box_type import MessageBoxType
from ooo.dyn.awt.message_box_buttons import MessageBoxButtonsEnum
from ooo.dyn.awt.message_box_results import MessageBoxResultsEnum
from ooodev.mock import mock_g
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import event_ctx
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.dialog.msgbox import MsgBox

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.dialog.dialog import Dialog
    from ooodev.events.args.cancel_event_args import CancelEventArgs


class CreateDialogPartial:
    """Partial for working with dialogs."""

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        """
        CreateDialogPartial Constructor.

        Args:
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst

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
        # pylint: disable=import-outside-toplevel
        from ooodev.dialog.dialog import Dialog

        dlg = Dialog(x=x, y=y, width=width, height=height, title=title, lo_inst=self.__lo_inst)
        return dlg

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

    def msgbox(
        self,
        msg: str,
        title: str = "Message",
        boxtype: MessageBoxType | int = MessageBoxType.MESSAGEBOX,
        buttons: MessageBoxButtonsEnum | int = MessageBoxButtonsEnum.BUTTONS_OK,
    ) -> MessageBoxResultsEnum:
        """
        Simple message box.

        Args:
            msg (str): the message for display.
            title (str, optional): the title of the message box. Defaults to "Message".
            boxtype (MessageBoxType, optional): determines the type of message box to display. Defaults to ``Type.MESSAGEBOX``.
            buttons (MessageBoxButtonsEnum, int, optional): determines what buttons to display. Defaults to ``Buttons.BUTTONS_OK``.

        Returns:
            Results: MessageBoxResultsEnum

        Note:
            If BoxType is an integer, the following values are valid:

            - 0: ``MESSAGEBOX``
            - 1: ``INFOBOX``
            - 2: ``WARNINGBOX``
            - 3: ``ERRORBOX``
            - 4: ``QUERYBOX``

        Note:

            - Button press ``Abort`` return ``MessageBoxResultsEnum.CANCEL``
            - Button press ``Cancel`` return ``MessageBoxResultsEnum.CANCEL``
            - Button press ``Ignore`` returns ``MessageBoxResultsEnum.IGNORE``
            - Button press ``No`` returns ``MessageBoxResultsEnum.NO``
            - Button press ``OK`` returns ``MessageBoxResultsEnum.OK``
            - Button press ``Retry`` returns ``MessageBoxResultsEnum.RETRY``
            - Button press ``Yes`` returns ``MessageBoxResultsEnum.YES``

        Note:
            Raises a global event ``GblNamedEvent.MSG_BOX_CREATING`` before creating the dialog.
            The event args are of type ``CancelEventArgs``.
            The ``event_data`` is a dictionary that contains the following key:

            - ``msg``: The message to display.
            - ``title``: The title of the dialog.
            - ``boxtype``: The type of message box to display.
            - ``buttons``: The buttons to display.

            If the event is cancelled, the ``result`` value of ``event_data` if set will be returned.
            Otherwise if the event is not handled, a ``CancelEventError`` is raised.
        """

        def on_dialog_creating(source: Any, event_args: CancelEventArgs, *args, **kwargs) -> None:
            if isinstance(self, EventsPartial):
                self.trigger_event(GblNamedEvent.MSG_BOX_CREATING, event_args)

        with LoContext(inst=self.__lo_inst):
            with event_ctx((GblNamedEvent.MSG_BOX_CREATING, on_dialog_creating), source=self, lo_observe=True):
                result = MsgBox.msgbox(msg=msg, title=title, boxtype=boxtype, buttons=buttons)
        return result

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

        Note:
            Raises a global event ``GblNamedEvent.INPUT_BOX_CREATING`` before creating the dialog.
            The event args are of type ``CancelEventArgs``.
            The ``event_data`` is a dictionary that contains the following key:

            - ``msg``: The message to display.
            - ``title``: The title of the dialog.
            - ``input_value``: The value of the input box when first displayed.
            - ``ok_lbl``: The label for the OK button.
            - ``cancel_lbl``: The label for the Cancel button.
            - ``is_password``: Determines if the input box is masked for password input.
            - ``frame``: The frame of the dialog. If not set, the frame of the current document is used.

            The default ``frame`` is ``None``. If set value must be a ``XFrame`` object.

            If the event is cancelled, the ``result`` value of ``event_data` if set will be returned.
            Otherwise if the event is not handled, a ``CancelEventError`` is raised.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.dialog.input import Input

        def on_dialog_creating(source: Any, event_args: CancelEventArgs, *args, **kwargs) -> None:
            if isinstance(self, GuiPartial):
                event_args.event_data["frame"] = self.get_frame()
            if isinstance(self, EventsPartial):
                self.trigger_event(GblNamedEvent.INPUT_BOX_CREATING, event_args)

        with LoContext(inst=self.__lo_inst):
            with event_ctx((GblNamedEvent.INPUT_BOX_CREATING, on_dialog_creating), source=self, lo_observe=True):
                result = Input.get_input(
                    title=title,
                    msg=msg,
                    input_value=input_value,
                    ok_lbl=ok_lbl,
                    cancel_lbl=cancel_lbl,
                    is_password=is_password,
                )
        return result

    # endregion input


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dialog import Dialog
    from ooodev.dialog.input import Input
