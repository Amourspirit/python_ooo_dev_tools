from __future__ import annotations
from typing import TYPE_CHECKING
import uno  # pylint: disable=unused-import
from com.sun.star.util import XReplaceable
from .searchable_partial import SearchablePartial
from .replace_descriptor_comp import ReplaceDescriptorComp

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.util import XSearchDescriptor


class ReplaceablePartial(SearchablePartial):
    """
    Partial Class XReplaceable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XReplaceable, interface: UnoInterface | None = XReplaceable) -> None:
        """
        Constructor

        Args:
            component (XReplaceable): UNO Component that implements ``com.sun.star.util.XReplaceable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XReplaceable``.
        """
        SearchablePartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XReplaceable
    def create_replace_descriptor(self) -> ReplaceDescriptorComp:
        """
        Creates a descriptor which contains properties that specify a search in this container.

        Returns:
            ReplaceDescriptorComp: The replace descriptor.
        """
        return ReplaceDescriptorComp(self.__component.createReplaceDescriptor())  # type: ignore

    def replace_all(self, desc: XSearchDescriptor) -> int:
        """
        Searches and replace all occurrences of whatever is specified.

        Returns:
            int: The number of replacements.
        """
        return self.__component.replaceAll(desc)

    # endregion XReplaceable
