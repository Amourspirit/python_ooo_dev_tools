from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from com.sun.star.container import XEnumerationAccess
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

    def __iter__(self):
        return self

    def __next__(self):
        if self.__enumeration is None:
            if not self.__component.hasElements():
                raise StopIteration
            self.__enumeration = self.create_enumeration()
        if self.__enumeration.hasMoreElements():
            return self.__enumeration.nextElement()
        self.__enumeration = None
        raise StopIteration

    # region Methods
    def create_enumeration(self) -> XEnumeration:
        """Creates an enumeration of the container's elements."""
        return self.__component.createEnumeration()

    # endregion Methods
