from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

import uno
from ooo.dyn.frame.infobar_type import InfobarTypeEnum
from com.sun.star.frame import XInfobarProvider

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.beans import StringPair
    from ooodev.utils.type_var import UnoInterface


class InfobarProviderPartial:
    """
    Partial class for XInfobarProvider.
    """

    def __init__(self, component: XInfobarProvider, interface: UnoInterface | None = XInfobarProvider) -> None:
        """
        Constructor

        Args:
            component (XInfobarProvider): UNO Component that implements ``com.sun.star.frame.XInfobarProvider`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XInfobarProvider``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XInfobarProvider
    def append_infobar(
        self,
        id: str,
        primary_message: str,
        secondary_message: str,
        infobar_type: int | InfobarTypeEnum,
        action_buttons: Tuple[StringPair, ...],
        show_close_button: bool,
    ) -> None:
        """
        Creates and displays a new Infobar.

        The example below adds a new infobar named MyInfoBar with type INFO and close (x) button.

        Args:
            id (str): The unique identifier of the Infobar.
            primary_message (str): The (short) primary message. Will appear at the start of the infobar in bold letters. May be empty.
            secondary_message (str): The (longer) secondary message. Will appear in normal letters after the primaryMessage
            infobar_type (int | InfobarTypeEnum): The type of the Infobar.
            action_buttons (Tuple[StringPair, ...]): A sequence of action buttons. The buttons will be added from Right to Left at the right side of the info bar. Each button is represented by a ``com.sun.star.beans.StringPair``.
                StringPair: First represents the button label, while StringPair: Second represents the button URL which will be called on button click. The URL can be any URL, either external (http://libreoffice.org), or internal (``.uno:Save``), or from your extension (``service:your.example.Extension?anyAction``).
            show_close_button (bool): Whether the Close (x) button is shown at the end of the Infobar. Set to false, when you don't want the user to close the Infobar.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``

        Returns:
            None:

        Hint:
            - ``InfobarTypeEnum`` can be imported from ``ooo.dyn.frame.infobar_type``.

        Example:
            .. code-block:: python

                from ooo.dyn.frame.infobar_type import InfobarTypeEnum
                from ooodev.utils.props import Props

                def add_infobar():
                    # Create a new infobar
                    buttons = (
                        Props.make_sting_pair("Close doc", ".uno:CloseDoc"),
                        Props.make_sting_pair("Paste into doc", ".uno:Paste"),
                    )
                    inst.append_infobar(
                        id="MyInfoBar",
                        primary_message="Hello world",
                        secondary_message="Things happened. What now?",
                        infobar_type=InfobarTypeEnum.INFO,
                        action_buttons=buttons,
                        show_close_button=True,
                    )
        """
        self.__component.appendInfobar(
            id, primary_message, secondary_message, int(infobar_type), action_buttons, show_close_button
        )

    def has_infobar(self, id: str) -> bool:
        """
        Check if Infobar exists.

        **since**

            LibreOffice 7.0
        """
        return self.__component.hasInfobar(id)

    def remove_infobar(self, id: str) -> None:
        """
        Removes an existing Infobar.

        Remove MyInfoBar infobar

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.removeInfobar(id)

    def update_infobar(
        self, id: str, primary_message: str, secondary_message: str, infobar_type: int | InfobarTypeEnum
    ) -> None:
        """
        Updates an existing Infobar.

        Use if you want to update only small parts of the Infobar.

        Update the infobar and change the type to WARNING

        Args:
            id (str): The unique identifier of the Infobar.
            primary_message (str): The (short) primary message. Will appear at the start of the infobar in bold letters. May be empty.
            secondary_message (str): The (longer) secondary message. Will appear in normal letters after the primaryMessage
            infobar_type (int | InfobarTypeEnum): The type of the Infobar.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``

        Returns:
            None:

        Hint:
            - ``InfobarTypeEnum`` can be imported from ``ooo.dyn.frame.infobar_type``.
        """
        self.__component.updateInfobar(id, primary_message, secondary_message, int(infobar_type))

    # endregion XInfobarProvider
