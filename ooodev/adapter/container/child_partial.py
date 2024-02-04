from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.container import XChild

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.uno import XInterface


class ChildPartial:
    """
    Partial class for XChild.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XChild, interface: UnoInterface | None = XChild) -> None:
        """
        Constructor

        Args:
            component (XChild): UNO Component that implements ``com.sun.star.container.XChild`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XChild``.
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

    # region XChild
    def get_parent(self) -> XInterface:
        """Returns the parent of the object."""
        return self.__component.getParent()

    def set_parent(self, parent: XInterface) -> None:
        """Sets the parent of the object."""
        self.__component.setParent(parent)

    # endregion XChild
