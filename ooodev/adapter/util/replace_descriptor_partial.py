from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.util import XReplaceDescriptor

from ooodev.adapter.util.search_descriptor_partial import SearchDescriptorPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ReplaceDescriptorPartial(SearchDescriptorPartial):
    """
    Partial class for XReplaceDescriptor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XReplaceDescriptor, interface: UnoInterface | None = XReplaceDescriptor) -> None:
        """
        Constructor

        Args:
            component (XReplaceDescriptor): UNO Component that implements ``com.sun.star.util.XReplaceDescriptor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XReplaceDescriptor``.
        """
        SearchDescriptorPartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XReplaceDescriptor
    def get_replace_string(self) -> str:
        """
        Gets the string which replaces the found occurrences.
        """
        return self.__component.getReplaceString()

    def set_replace_string(self, replacement: str) -> None:
        """
        Sets the string which replaces the found occurrences.
        """
        self.__component.setReplaceString(replacement)

    # endregion XReplaceDescriptor
