from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XNameAccess

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from .element_access_partial import ElementAccessPartial


class NameAccessPartial(ElementAccessPartial):
    """
    Class for managing IndexAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameAccess) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that implements ``com.sun.star.container.XNameAccess`` interface.
        """
        if not mLo.Lo.is_uno_interfaces(component, XNameAccess):
            raise mEx.MissingInterfaceError("XNameAccess")
        ElementAccessPartial.__init__(self, component)
        self.__component = component

    # region Methods
    def get_by_name(self, name: str) -> Any:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            Any: The element with the specified name.
        """
        return self.__component.getByName(name)

    def get_element_names(self) -> tuple[str, ...]:
        """
        Gets the names of all elements contained in the container.

        Returns:
            tuple[str, ...]: The names of all elements.
        """
        return tuple(self.__component.getElementNames())

    def has_by_name(self, name: str) -> bool:
        """
        Checks if the container has an element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            bool: ``True`` if the container has an element with the specified name, otherwise ``False``.
        """
        return self.__component.hasByName(name)

    # endregion Methods
