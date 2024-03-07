from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.awt import XUserInputInterception

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

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
