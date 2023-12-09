from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.container import XEnumerationAccess

if TYPE_CHECKING:
    from com.sun.star.container import XEnumeration

from ooodev.utils.type_var import UnoInterface
from .element_access_partial import ElementAccessPartial


class EnumerationAccessPartial(ElementAccessPartial):
    """
    Partial Class for XEnumerationAccess.

    This class can be used to iterate over the elements of a container.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XEnumerationAccess, interface: UnoInterface | None = XEnumerationAccess) -> None:
        """
        Constructor

        Args:
            component (XEnumerationAccess): UNO Component that implements ``com.sun.star.container.XEnumerationAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XEnumerationAccess``.
        """
        ElementAccessPartial.__init__(self, component, interface)
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
