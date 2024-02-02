from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertyAccess


from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue  # struct
    from ooodev.utils.type_var import UnoInterface


class PropertyAccessPartial:
    """
    Partial class for XPropertyAccess.
    """

    def __init__(self, component: XPropertyAccess, interface: UnoInterface | None = XPropertyAccess) -> None:
        """
        Constructor

        Args:
            component (XPropertyAccess): UNO Component that implements ``com.sun.star.container.XPropertyAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertyAccess``.
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

    # region XPropertyAccess

    def get_property_values(self) -> tuple[PropertyValue, ...]:
        """
        Gets of all property values within the object in a single call.

        Returns:
            tuple[PropertyValue, ...]: The property values.
        """
        return self.__component.getPropertyValues()

    def set_property_values(self, values: tuple[PropertyValue, ...]) -> None:
        """
        Sets of all property values within the object in a single call.

        All properties which are not contained in the sequence ``values`` will be left unchanged.

        Args:
            values (tuple[PropertyValue, ...]): The property values.
        """
        self.__component.setPropertyValues(values)

    # endregion XPropertyAccess
