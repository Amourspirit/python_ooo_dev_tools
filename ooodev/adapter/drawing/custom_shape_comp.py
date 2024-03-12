from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.adapter.drawing.generic_shape import GenericShapeComp
from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial
from ooodev.adapter.drawing.line_properties_partial import LinePropertiesPartial
from ooodev.adapter.drawing.shadow_properties_partial import ShadowPropertiesPartial
from ooodev.adapter.drawing.rotation_descriptor_properties_partial import RotationDescriptorPropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import CustomShape  # service
else:
    CustomShape = Any


class CustomShapeComp(
    GenericShapeComp[CustomShape],
    FillPropertiesPartial,
    LinePropertiesPartial,
    ShadowPropertiesPartial,
    RotationDescriptorPropertiesPartial,
):
    """
    Class for managing CustomShape Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that implements ``com.sun.star.drawing.CustomShape`` service.
        """
        GenericShapeComp.__init__(self, component)
        FillPropertiesPartial.__init__(self, component=component)
        LinePropertiesPartial.__init__(self, component=component)
        ShadowPropertiesPartial.__init__(self, component=component)
        RotationDescriptorPropertiesPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.CustomShape",)

    # endregion Overrides
