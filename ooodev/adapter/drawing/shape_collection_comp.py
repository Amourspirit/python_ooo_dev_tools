from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from .shapes_partial import ShapesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import ShapeCollection  # service


class ShapeCollectionComp(
    ComponentBase,
    ShapesPartial,
):
    """
    Class for managing ShapeCollection Component.

    .. versionadded:: 0.20.5
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that implements ``com.sun.star.drawing.ShapeCollection`` service.
        """
        ComponentBase.__init__(self, component)
        ShapesPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.ShapeCollection",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> ShapeCollection:
        """ShapeCollection Component"""
        return cast("ShapeCollection", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
