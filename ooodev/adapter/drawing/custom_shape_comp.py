from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.adapter.drawing.generic_shape import GenericShapeComp

if TYPE_CHECKING:
    from com.sun.star.drawing import CustomShape  # service
else:
    CustomShape = Any


class CustomShapeComp(GenericShapeComp[CustomShape]):
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

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.CustomShape",)

    # endregion Overrides
