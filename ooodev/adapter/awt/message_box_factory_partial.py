from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.awt import XMessageBoxFactory
from ooo.dyn.awt.message_box_buttons import MessageBoxButtonsEnum
from ooo.dyn.awt.message_box_type import MessageBoxType

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.awt import XWindowPeer
    from com.sun.star.awt import XMessageBox
    from ooodev.utils.type_var import UnoInterface


class MessageBoxFactoryPartial:
    """
    Partial class for XMessageBoxFactory.
    """

    def __init__(self, component: XMessageBoxFactory, interface: UnoInterface | None = XMessageBoxFactory) -> None:
        """
        Constructor

        Args:
            component (XMessageBoxFactory): UNO Component that implements ``com.sun.star.awt.XMessageBoxFactory`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMessageBoxFactory``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XMessageBoxFactory
    def create_message_box(
        self,
        parent: XWindowPeer,
        title: str,
        msg: str,
        boxtype: MessageBoxType | int = MessageBoxType.MESSAGEBOX,
        buttons: MessageBoxButtonsEnum | int = MessageBoxButtonsEnum.BUTTONS_OK,
    ) -> XMessageBox:
        """
        creates a message box.

        This parameter must not be null.

        A combination of com.sun.star.awt.MessageBoxButtons

        A com.sun.star.awt.MessageBoxType.INFOBOX ignores this parameter, instead it uses a com.sun.star.awt.MessageBoxButtons.BUTTONS_OK.

        Args:
            parent (XWindowPeer): The parent window.
            title (str): The title of the message box.
            msg (str): The message of the message box.
            boxtype (MessageBoxType, int, optional): determines the type of message box to display. Defaults to ``Type.MESSAGEBOX``.
            buttons (MessageBoxButtonsEnum, int, optional): determines what buttons to display. Defaults to ``Buttons.BUTTONS_OK``.

        Returns:
            XMessageBox: the created message box.

        Note:
            If ``boxtype`` is an integer, the following values are valid:

            - 0: ``MESSAGEBOX``
            - 1: ``INFOBOX``
            - 2: ``WARNINGBOX``
            - 3: ``ERRORBOX``
            - 4: ``QUERYBOX``

        Hint:
            Run ``result.execute()`` to show the message box.

            - ``MessageBoxResultsEnum`` can be imported from ``ooo.dyn.awt.message_box_results``.
            - ``MessageBoxButtonsEnum`` can be imported from ``ooo.dyn.awt.message_box_buttons``.
            - ``MessageBoxType`` can be imported from ``ooo.dyn.awt.message_box_type``.

        See Also:
            - :py:class:`ooodev.dialog.msgbox.MsgBox`
        """
        if isinstance(boxtype, int):
            if boxtype == 1:
                boxtype = MessageBoxType.INFOBOX
            elif boxtype == 2:
                boxtype = MessageBoxType.WARNINGBOX
            elif boxtype == 3:
                boxtype = MessageBoxType.ERRORBOX
            elif boxtype == 4:
                boxtype = MessageBoxType.QUERYBOX
            else:
                boxtype = MessageBoxType.MESSAGEBOX
        if boxtype == MessageBoxType.INFOBOX:
            # this is the default behavior anyways. So assigning ok to make it official here
            _buttons = MessageBoxButtonsEnum.BUTTONS_OK.value
        else:
            _buttons = int(buttons)
        return self.__component.createMessageBox(parent, boxtype, _buttons, str(title), str(msg))  # type: ignore

    # endregion XMessageBoxFactory
