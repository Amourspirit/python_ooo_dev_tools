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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPropertyAccess

    def get_property_values(self) -> tuple[PropertyValue, ...]:
        """
        Gets of all property values within the object in a single call.

        Returns:
            tuple[PropertyValue, ...]: The property values.
        """
        # This is unexpected but getPropertyValues() on a hidden control result is an error.
        # uno.com.sun.star.lang.IllegalArgumentException: incorrect number of parameters passed invoking function getPropertyValues: expected 1, got 0 at ./stoc/source/invocation/invocation.cxx:606
        # This may be a bug in the uno library.
        # This also may happen on other objects that implement XPropertyAccess.
        # A workaround is to check if the object has a PropertyValues attribute.
        if hasattr(self.__component, "PropertyValues"):
            return self.__component.PropertyValues  # type: ignore
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
