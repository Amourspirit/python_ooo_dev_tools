from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.chart2 import Title  # service


class TitleComp(ComponentBase, PropertiesChangeImplement, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing Chart2 Title Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Title) -> None:
        """
        Constructor

        Args:
            component (Title): UNO Chart2 Title Component.
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.chart2.Title",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Title:
        """Title Component"""
        return cast("Title", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
