from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from .shape_partial import ShapePartial
from .shape_descriptor_partial import ShapeDescriptorPartial

# from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
# from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement

if TYPE_CHECKING:
    from com.sun.star.drawing import Shape  # service


class ShapeComp(
    ComponentBase,
    ShapePartial,
    ShapeDescriptorPartial,
):
    """
    Class for managing Shape Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that implements ``com.sun.star.drawing.Shape`` service.
        """
        ComponentBase.__init__(self, component)
        ShapePartial.__init__(self, component=component, interface=None)
        ShapeDescriptorPartial.__init__(self, component=component, interface=None)
        # generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        # PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        # VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return (
            "com.sun.star.drawing.Shape",
            "com.sun.star.presentation.Shape",
        )

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Shape:
        """Shape Component"""
        return cast("Shape", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
