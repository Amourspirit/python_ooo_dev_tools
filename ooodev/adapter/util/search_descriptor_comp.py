from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.util.search_descriptor_partial_props import SearchDescriptorPartialProps
from ooodev.adapter.util.search_descriptor_partial import SearchDescriptorPartial


if TYPE_CHECKING:
    from com.sun.star.util import SearchDescriptor  # service


class SearchDescriptorComp(
    ComponentBase,
    SearchDescriptorPartial,
    SearchDescriptorPartialProps,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing SearchDescriptor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SearchDescriptor) -> None:
        """
        Constructor

        Args:
            component (SearchDescriptor): UNO Component that supports ``com.sun.star.util.SearchDescriptor`` service.
        """
        # pylint: disable=no-member
        ComponentBase.__init__(self, component)
        SearchDescriptorPartial.__init__(self, component=component, interface=None)
        SearchDescriptorPartialProps.__init__(self, component)

        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.util.SearchDescriptor",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> SearchDescriptor:
        """SearchDescriptor Component"""
        # pylint: disable=no-member
        return cast("SearchDescriptor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
