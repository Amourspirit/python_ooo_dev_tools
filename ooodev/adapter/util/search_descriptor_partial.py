from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.util import XSearchDescriptor

from ooodev.adapter.beans.property_set_partial import PropertySetPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class SearchDescriptorPartial(PropertySetPartial):
    """
    Partial class for XSearchDescriptor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSearchDescriptor, interface: UnoInterface | None = XSearchDescriptor) -> None:
        """
        Constructor

        Args:
            component (XSearchDescriptor): UNO Component that implements ``com.sun.star.util.XSearchDescriptor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSearchDescriptor``.
        """
        PropertySetPartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XSearchDescriptor
    def get_search_string(self) -> str:
        """
        Gets the string of characters to look for.
        """
        return self.__component.getSearchString()

    def set_search_string(self, string: str) -> None:
        """
        Sets the string of characters to look for.
        """
        self.__component.setSearchString(string)

    # endregion XSearchDescriptor
