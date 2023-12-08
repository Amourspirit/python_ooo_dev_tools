from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.container import XEnumerationAccess

if TYPE_CHECKING:
    from com.sun.star.container import XEnumeration

from ooodev.utils import lo as mLo
from ooodev.exceptions import ex as mEx
from .element_access_partial import ElementAccessPartial


class EnumerationAccessPartial(ElementAccessPartial):
    """
    Class for managing EnumerationAccess.

    This class can be used to iterate over the elements of a container.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XEnumerationAccess) -> None:
        """
        Constructor

        Args:
            component (XEnumerationAccess): UNO Component that implements ``com.sun.star.container.XEnumerationAccess`` interface.
        """
        if not mLo.Lo.is_uno_interfaces(component, XEnumerationAccess):
            raise mEx.MissingInterfaceError("XEnumerationAccess")
        ElementAccessPartial.__init__(self, component)
        self.__component = component
        self.__enumeration = None

    def _is_next_element_valid(self, element: Any) -> bool:
        """
        Gets if the next element is valid.
        This method is called when iterating over the elements of this class.

        Args:
            element (Any): Element

        Returns:
            bool: True in this class but can be overridden in child classes.
        """
        return True

    def __iter__(self):
        return self

    def __next__(self):
        if self.__enumeration is None:
            if not self.__component.hasElements():
                raise StopIteration
            self.__enumeration = self.create_enumeration()
        if self.__enumeration.hasMoreElements():
            next_element = self.__enumeration.nextElement()
            if self._is_next_element_valid(next_element):
                return next_element
            return self.__next__()
        self.__enumeration = None
        raise StopIteration

    # region Methods
    def create_enumeration(self) -> XEnumeration:
        """Creates an enumeration of the container's elements."""
        return self.__component.createEnumeration()

    # endregion Methods
