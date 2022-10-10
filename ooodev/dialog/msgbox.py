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

    Results = MessageBoxResultsEnum
    Buttons = MessageBoxButtonsEnum
    Type = MessageBoxType

    @staticmethod
    def msgbox(
        msg: str,
        title: str = "Message",
        boxtype: MsgBox.Type = Type.MESSAGEBOX,
        buttons: MsgBox.Buttons | int = Buttons.BUTTONS_OK,
    ) -> Results:
        """
        Simple message box.

        Args:
            msg (str): the message for display
            title (str, optional):  the title of the message box. Defaults to "Message".
            boxtype (MessageBoxType, optional): determins the type of message box to display. Defaults to ``Type.MESSAGEBOX``.
            buttons (Buttons, int, optional): determins what buttons to display. Defaults to ``Buttons.BUTTONS_OK``.

        Returns:
            Results: MsgBox.Results Enum

            * Button press ``Abort`` return ``Results.CANCEL``
            * Button press ``Cancel`` return ``Results.CANCEL``
            * Button press ``Ignore`` returns ``Results.IGNORE``
            * Button press ``No`` returns ``Results.NO``
            * Button press ``OK`` returns ``Results.OK``
            * Button press ``Retry`` returns ``Results.RETRY``
            * Button press ``Yes`` returns ``Results.YES``
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
