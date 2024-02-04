from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XNamed

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface


class NamedPartial:
    """
    Partial class for XNamed.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNamed, interface: UnoInterface | None = XNamed) -> None:
        """
        Constructor

        Args:
            component (XNamed): UNO Component that implements ``com.sun.star.container.XNamed`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNamed``.
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

    # region XNamed
    def get_name(self) -> str:
        """Returns the name of the object."""
        return self.__component.getName()

    def set_name(self, name: str) -> None:
        """Sets the name of the object."""
        self.__component.setName(name)

    # endregion XNamed
