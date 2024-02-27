from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooodev.adapter.util.search_descriptor_comp import SearchDescriptorComp
from ooodev.adapter.util.replace_descriptor_partial import ReplaceDescriptorPartial


if TYPE_CHECKING:
    from com.sun.star.util import ReplaceDescriptor  # service


class ReplaceDescriptorComp(SearchDescriptorComp, ReplaceDescriptorPartial):
    """
    Class for managing ReplaceDescriptor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ReplaceDescriptor) -> None:
        """
        Constructor

        Args:
            component (ReplaceDescriptor): UNO Component that supports ``com.sun.star.util.ReplaceDescriptor`` service.
        """
        # pylint: disable=no-member
        SearchDescriptorComp.__init__(self, component)
        ReplaceDescriptorPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.util.ReplaceDescriptor",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> ReplaceDescriptor:
        """ReplaceDescriptor Component"""
        # pylint: disable=no-member
        return super().component  # type: ignore

    # endregion Properties
