from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase
from .shapes_partial import ShapesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import GenericDrawPage  # service


class GenericDrawPageComp(ComponentBase, PropertyChangeImplement, VetoableChangeImplement, ShapesPartial):
    """
    Class for managing GenericDrawPage Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: GenericDrawPage) -> None:
        """
        Constructor

        Args:
            component (GenericDrawPage): UNO table GenericDrawPage Component.
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        ShapesPartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.GenericDrawPage",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> GenericDrawPage:
        """GenericDrawPage Component"""
        return cast("GenericDrawPage", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
