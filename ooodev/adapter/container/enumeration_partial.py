from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XEnumeration

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo


class enumeration_partial:
    """
    Class for managing Enumeration.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XEnumeration) -> None:
        """
        Constructor

        Args:
            component (XEnumeration): UNO Component that implements ``com.sun.star.container.XEnumeration`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XEnumeration):
            raise mEx.MissingInterfaceError("XEnumeration")
        self.__component = component

    # region XEnumeration
    def has_more_elements(self) -> bool:
        """
        tests whether this enumeration contains more elements.
        """
        return self.__component.hasMoreElements()

    def next_element(self) -> Any:
        """
        Gets the next element of this enumeration.

        Returns:
            Any: the next element of this enumeration.
        """
        return self.__component.nextElement()

    # endregion XEnumeration
