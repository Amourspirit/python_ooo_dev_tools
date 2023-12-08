from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XElementAccess

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo


class ElementAccessPartial:
    """
    Class for managing ElementAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XElementAccess) -> None:
        """
        Constructor

        Args:
            component (XElementAccess): UNO Component that implements ``com.sun.star.container.XElementAccess`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XElementAccess):
            raise mEx.MissingInterfaceError("XElementAccess")
        self.__component = component

    # region Methods
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

    # endregion Methods
