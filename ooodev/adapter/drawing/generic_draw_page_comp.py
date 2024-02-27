from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.drawing.shapes_partial import ShapesPartial
from ooodev.adapter.drawing.shape_grouper_partial import ShapeGrouperPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import GenericDrawPage  # service
    from com.sun.star.lang import XComponent


class GenericDrawPageComp(ComponentBase, ShapesPartial, ShapeGrouperPartial):
    """
    Class for managing GenericDrawPage Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XComponent) -> None:
        """
        Constructor

        Args:
            component (XComponent): UNO component that implements ``com.sun.star.drawing.GenericDrawPage`` service.
        """
        ComponentBase.__init__(self, component)
        # generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ShapesPartial.__init__(self, component=self.component, interface=None)
        ShapeGrouperPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.GenericDrawPage",)

    # endregion Overrides
    # region Properties

    @property
    def component(self) -> GenericDrawPage:
        """GenericDrawPage Component"""
        # pylint: disable=no-member
        return cast("GenericDrawPage", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
