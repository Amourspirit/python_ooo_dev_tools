from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.container import XElementAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ElementAccessPartial:
    """
    Partial class for XElementAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XElementAccess, interface: UnoInterface | None = XElementAccess) -> None:
        """
        Constructor

        Args:
            component (XElementAccess): UNO Component that implements ``com.sun.star.container.XElementAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XElementAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XElementAccess
    def get_element_type(self) -> Any:
        """
        Gets the type of the elements contained in the container.

        Returns:
            Any: The type of the elements. ``None``  means that it is a multi-type container and you cannot determine the exact types with this interface.
        """
        return self.__component.getElementType()

    def has_elements(self) -> bool:
        """Determines whether the container has elements."""
        return self.__component.hasElements()

    # endregion XElementAccess
