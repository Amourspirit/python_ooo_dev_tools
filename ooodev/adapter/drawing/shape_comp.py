from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.drawing.shape_partial import ShapePartial
from ooodev.adapter.drawing.shape_descriptor_partial import ShapeDescriptorPartial
from ooodev.adapter.drawing.shape_properties_partial import ShapePropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import Shape  # service


class ShapeComp(
    ComponentBase,
    ShapePartial,
    ShapeDescriptorPartial,
    ShapePropertiesPartial,
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
        ShapePropertiesPartial.__init__(self, component=component)

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
        # pylint: disable=no-member
        return cast("Shape", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
