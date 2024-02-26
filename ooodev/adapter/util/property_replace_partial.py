from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno
from com.sun.star.util import XPropertyReplace

from ooodev.adapter.util.replace_descriptor_partial import ReplaceDescriptorPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import PropertyValue  # struct


class PropertyReplacePartial(ReplaceDescriptorPartial):
    """
    Partial class for XPropertyReplace.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPropertyReplace, interface: UnoInterface | None = XPropertyReplace) -> None:
        """
        Constructor

        Args:
            component (XPropertyReplace): UNO Component that implements ``com.sun.star.util.XPropertyReplace`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertyReplace``.
        """
        ReplaceDescriptorPartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XPropertyReplace
    def get_replace_attributes(self) -> Tuple[PropertyValue, ...]:
        """
        Gets the properties to replace the found occurrences.
        """
        return self.__component.getReplaceAttributes()

    def get_search_attributes(self) -> Tuple[PropertyValue, ...]:
        """
        Get the search attributes.
        """
        return self.__component.getSearchAttributes()

    def get_value_search(self) -> bool:
        """
        provides the information if specific property values are searched, or just the existence of the specified properties.
        """
        return self.__component.getValueSearch()

    def set_replace_attributes(self, *attribs: PropertyValue) -> None:
        """
        Sets the properties to replace the found occurrences.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.setReplaceAttributes(attribs)

    def set_search_attributes(self, *attribs: PropertyValue) -> None:
        """
        Sets the properties to search for.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.setSearchAttributes(attribs)

    def set_value_search(self, value_search: bool) -> None:
        """
        Sets if specific property values are searched, or just the existence of the specified properties.
        """
        self.__component.setValueSearch(value_search)

    # endregion XPropertyReplace
