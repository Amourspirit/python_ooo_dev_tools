# coding: utf-8
from __future__ import annotations
from typing import cast
from ..utils.lo import Lo

from com.sun.star.awt import XToolkit2
from com.sun.star.awt import XMessageBox

from ooo.dyn.awt.message_box_results import MessageBoxResultsEnum as MessageBoxResultsEnum
from ooo.dyn.awt.message_box_buttons import MessageBoxButtonsEnum as MessageBoxButtonsEnum
from ooo.dyn.awt.message_box_type import MessageBoxType as MessageBoxType


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
            msg (str): the message for display
            title (str, optional): the title of the message box. Defaults to "Message".
            boxtype (MessageBoxType, optional): determines the type of message box to display. Defaults to ``Type.MESSAGEBOX``.
            buttons (MessageBoxButtonsEnum, int, optional): determines what buttons to display. Defaults to ``Buttons.BUTTONS_OK``.

        Returns:
            Results: MessageBoxResultsEnum

            * Button press ``Abort`` return ``MessageBoxResultsEnum.CANCEL``
            * Button press ``Cancel`` return ``MessageBoxResultsEnum.CANCEL``
            * Button press ``Ignore`` returns ``MessageBoxResultsEnum.IGNORE``
            * Button press ``No`` returns ``MessageBoxResultsEnum.NO``
            * Button press ``OK`` returns ``MessageBoxResultsEnum.OK``
            * Button press ``Retry`` returns ``MessageBoxResultsEnum.RETRY``
            * Button press ``Yes`` returns ``MessageBoxResultsEnum.YES``
        """
        if boxtype == MessageBoxType.INFOBOX:
            # this is the default behaviour anyways. So assigning ok to make it official here
            _buttons = MessageBoxButtonsEnum.BUTTONS_OK.value
        else:
            _buttons = buttons

        tk = Lo.create_instance_mcf(XToolkit2, "com.sun.star.awt.Toolkit")
        parent = tk.getDesktopWindow()
        box = cast(XMessageBox, tk.createMessageBox(parent, boxtype, int(_buttons), str(title), str(msg)))
        return MessageBoxResultsEnum(int(box.execute()))
