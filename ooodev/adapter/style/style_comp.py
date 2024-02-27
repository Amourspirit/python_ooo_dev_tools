from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.style.style_partial import StylePartial

if TYPE_CHECKING:
    from com.sun.star.style import Style  # service
    from com.sun.star.style import XStyle


class StyleComp(ComponentBase, PropertyChangeImplement, VetoableChangeImplement, StylePartial):
    """
    Class for managing Style Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XStyle) -> None:
        """
        Constructor

        Args:
            component (XStyle): UNO Component that support ``com.sun.star.style.Style`` service.
        """
        ComponentBase.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        # no need to validate StylePartial will be validated by ComponentBase
        StylePartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.Style",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Style:
        """Style Component"""
        # pylint: disable=no-member
        return cast("Style", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
