# coding: utf-8
from __future__ import annotations
from typing import cast
import uno  # pylint: disable=unused-import

from com.sun.star.awt import XToolkit2

# pylint: disable=useless-import-alias
from ooo.dyn.awt.message_box_results import MessageBoxResultsEnum as MessageBoxResultsEnum
from ooo.dyn.awt.message_box_buttons import MessageBoxButtonsEnum as MessageBoxButtonsEnum
from ooo.dyn.awt.message_box_type import MessageBoxType as MessageBoxType
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.event_singleton import _Events
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


class MsgBox:
    @staticmethod
    def msgbox(
        msg: str,
        title: str = "Message",
        boxtype: MessageBoxType = MessageBoxType.MESSAGEBOX,
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
            Results: MessageBoxResultsEnum.

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
        cargs = CancelEventArgs(MsgBox.msgbox.__qualname__)
        cargs.event_data = {"msg": msg, "title": title, "boxtype": boxtype, "buttons": buttons}

        _Events().trigger(GblNamedEvent.MSG_BOX_CREATING, cargs)
        if cargs.cancel is True:
            if "result" in cargs.event_data:
                return cast(MessageBoxResultsEnum, cargs.event_data["result"])
            if cargs.handled is False:
                raise mEx.CancelEventError(cargs, "Dialog creation was cancelled.")
        msg = cast(str, cargs.event_data["msg"])
        title = cast(str, cargs.event_data["title"])
        boxtype = cast(MessageBoxType, cargs.event_data["boxtype"])
        buttons = int(cargs.event_data["buttons"])  # type: ignore

        # sourcery skip: remove-unnecessary-cast
        if boxtype == MessageBoxType.INFOBOX:
            # this is the default behavior anyways. So assigning ok to make it official here
            _buttons = MessageBoxButtonsEnum.BUTTONS_OK.value
        else:
            _buttons = buttons

        tk = mLo.Lo.create_instance_mcf(XToolkit2, "com.sun.star.awt.Toolkit", raise_err=True)
        parent = tk.getDesktopWindow()
        box = tk.createMessageBox(parent, boxtype, _buttons, str(title), str(msg))  # type: ignore
        return MessageBoxResultsEnum(int(box.execute()))
