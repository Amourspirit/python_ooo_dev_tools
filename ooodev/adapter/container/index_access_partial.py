from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XIndexAccess

from ooodev.utils import lo as mLo
from ooodev.exceptions import ex as mEx
from .element_access_partial import ElementAccessPartial


class IndexAccessPartial(ElementAccessPartial):
    """
    Class for managing IndexAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XIndexAccess) -> None:
        """
        Constructor

        Args:
            component (XIndexAccess): UNO Component that implements ``com.sun.star.container.XIndexAccess`` interface.
        """
        if not mLo.Lo.is_uno_interfaces(component, XIndexAccess):
            raise mEx.MissingInterfaceError("XIndexAccess")
        ElementAccessPartial.__init__(self, component)
        self.__component = component

    # region Methods
    def get_count(self) -> int:
        """
        Gets the number of elements contained in the container.

        Returns:
            int: The number of elements.
        """
        return self.__component.getCount()

    def get_by_index(self, index: int) -> Any:
        """
        Gets the element at the specified index.

        Args:
            index (int): The Zero-based index of the element.

        Returns:
            Any: The element at the specified index.
        """
        return self.__component.getByIndex(index)

    # endregion Methods
