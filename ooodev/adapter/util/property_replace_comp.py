from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.util.search_descriptor_partial_props import SearchDescriptorPartialProps
from ooodev.adapter.util.property_replace_partial import PropertyReplacePartial


if TYPE_CHECKING:
    from com.sun.star.util import ReplaceDescriptor  # service
    from com.sun.star.util import XPropertyReplace


class PropertyReplaceComp(
    ComponentBase,
    PropertyReplacePartial,
    SearchDescriptorPartialProps,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing ReplaceDescriptor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPropertyReplace) -> None:
        """
        Constructor

        Args:
            component (ReplaceDescriptor): UNO Component that supports ``com.sun.star.util.ReplaceDescriptor`` service.
        """
        # pylint: disable=no-member
        ComponentBase.__init__(self, component)
        PropertyReplacePartial.__init__(self, component=component, interface=None)
        SearchDescriptorPartialProps.__init__(self, component)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

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
        return cast("ReplaceDescriptor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
