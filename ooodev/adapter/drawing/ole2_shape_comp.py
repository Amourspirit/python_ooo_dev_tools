from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.drawing.shape_partial import ShapePartial
from ooodev.adapter.drawing.shape_descriptor_partial import ShapeDescriptorPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import OLE2Shape


class OLE2ShapeComp(
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
            component (Any): UNO component that implements ``com.sun.star.drawing.OLE2Shape`` service.
        """
        ComponentBase.__init__(self, component)
        ShapePartial.__init__(self, component=component, interface=None)
        ShapeDescriptorPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.OLE2Shape",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> OLE2Shape:
        """OLE2Shape Component"""
        # pylint: disable=no-member
        return self._ComponentBase__get_component()  # type: ignore

    # endregion Properties
