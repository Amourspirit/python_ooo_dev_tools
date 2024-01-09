from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.awt import XUserInputInterception

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XKeyHandler
    from com.sun.star.awt import XMouseClickHandler
    from com.sun.star.awt import XKeyHandler
    from com.sun.star.awt import XMouseClickHandler
    from ooodev.utils.type_var import UnoInterface


class UserInputInterceptionPartial:
    """
    Partial Class for XUserInputInterception.

    .. versionadded:: 0.20.0
    """

    def __init__(
        self, component: XUserInputInterception, interface: UnoInterface | None = XUserInputInterception
    ) -> None:
        """
        Constructor

        Args:
            component (XUserInputInterception): UNO Component that implements ``com.sun.star.awt.XUserInputInterception``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUserInputInterception``.
        """
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XUserInputInterception
    def add_key_handler(self, handler: XKeyHandler) -> None:
        """
        Adds a key handler.

        Args:
            handler (XKeyHandler): The handler to be added.
        """
        self.__component.addKeyHandler(handler)

    def add_mouse_click_handler(self, handler: XMouseClickHandler) -> None:
        """
        Adds a mouse click handler.

        Args:
            handler (XMouseClickHandler): The handler to be added.
        """
        self.__component.addMouseClickHandler(handler)

    def remove_key_handler(self, handler: XKeyHandler) -> None:
        """
        Removes a key handler.

        Args:
            handler (XKeyHandler): The handler to be removed.
        """
        self.__component.removeKeyHandler(handler)

    def remove_mouse_click_handler(self, handler: XMouseClickHandler) -> None:
        """
        Removes a mouse click handler.

        Args:
            handler (XMouseClickHandler): The handler to be removed.
        """
        self.__component.removeMouseClickHandler(handler)

    # endregion XUserInputInterception
