# coding: utf-8
from __future__ import annotations
import uno  # pylint: disable=unused-import

from com.sun.star.awt import XToolkit2

# pylint: disable=useless-import-alias
from ooo.dyn.awt.message_box_results import MessageBoxResultsEnum as MessageBoxResultsEnum
from ooo.dyn.awt.message_box_buttons import MessageBoxButtonsEnum as MessageBoxButtonsEnum
from ooo.dyn.awt.message_box_type import MessageBoxType as MessageBoxType

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

            * Button press ``Abort`` return ``MessageBoxResultsEnum.CANCEL``
            * Button press ``Cancel`` return ``MessageBoxResultsEnum.CANCEL``
            * Button press ``Ignore`` returns ``MessageBoxResultsEnum.IGNORE``
            * Button press ``No`` returns ``MessageBoxResultsEnum.NO``
            * Button press ``OK`` returns ``MessageBoxResultsEnum.OK``
            * Button press ``Retry`` returns ``MessageBoxResultsEnum.RETRY``
            * Button press ``Yes`` returns ``MessageBoxResultsEnum.YES``
        """
        # sourcery skip: remove-unnecessary-cast
        if boxtype == MessageBoxType.INFOBOX:
            # this is the default behavior anyways. So assigning ok to make it official here
            _buttons = MessageBoxButtonsEnum.BUTTONS_OK.value
        else:
            _buttons = buttons

        tk = mLo.Lo.create_instance_mcf(XToolkit2, "com.sun.star.awt.Toolkit", raise_err=True)
        parent = tk.getDesktopWindow()
        box = tk.createMessageBox(parent, boxtype, int(_buttons), str(title), str(msg))  # type: ignore
        return MessageBoxResultsEnum(int(box.execute()))
