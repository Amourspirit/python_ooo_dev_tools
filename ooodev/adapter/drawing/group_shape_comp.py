from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.drawing.shape_comp import ShapeComp
from ooodev.adapter.drawing.shape_group_partial import ShapeGroupPartial
from ooodev.adapter.drawing.shapes_partial import ShapesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import GroupShape  # service


class GroupShapeComp(ShapeComp, ShapeGroupPartial, ShapesPartial):
    """
    Class for managing EllipseShape Component.
    """

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that implements ``com.sun.star.drawing.GroupShape`` service.
        """
        ShapeComp.__init__(self, component)
        ShapeGroupPartial.__init__(self, component=component, interface=None)
        ShapesPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.GroupShape",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> GroupShape:
        """GroupShape Component"""
        return cast("GroupShape", super().component)
        # return cast("GroupShape", super(ShapeComp, self).component)  # type: ignore

    # endregion Properties
