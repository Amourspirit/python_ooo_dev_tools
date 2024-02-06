from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.lang import XInitialization

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class InitializationPartial:
    """
    Partial class for XInitialization.
    """

    def __init__(self, component: XInitialization, interface: UnoInterface | None = XInitialization) -> None:
        """
        Constructor

        Args:
            component (XInitialization): UNO Component that implements ``com.sun.star.lang.XInitialization`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XInitialization``.
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

    # region XInitialization
    def initialize(self, *args: Any) -> None:
        """
        Initializes the object.

        It should be called directly after the object is created.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.initialize(args)

    # endregion XInitialization
