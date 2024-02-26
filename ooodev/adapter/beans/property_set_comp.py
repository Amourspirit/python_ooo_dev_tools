from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.beans.property_set_partial import PropertySetPartial

if TYPE_CHECKING:
    from com.sun.star.beans import PropertySet  # service


class PropertySetComp(ComponentBase, PropertySetPartial, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing PropertySet Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: PropertySet) -> None:
        """
        Constructor

        Args:
            component (ChartType): UNO Chart2 ChartType Component.
        """
        ComponentBase.__init__(self, component)
        PropertySetPartial.__init__(self, component=component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # PropertySetPartial will validate
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> PropertySet:
        """PropertySet Component"""
        # pylint: disable=no-member
        return cast("PropertySet", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
