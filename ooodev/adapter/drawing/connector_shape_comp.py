from __future__ import annotations
from typing import Any, TYPE_CHECKING

from .generic_shape import GenericShapeComp

if TYPE_CHECKING:
    from com.sun.star.drawing import ConnectorShape  # service
else:
    ConnectorShape = Any


class ConnectorShapeComp(GenericShapeComp[ConnectorShape]):
    """
    Class for managing ConnectorShape Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that implements ``com.sun.star.drawing.ConnectorShape`` service.
        """
        GenericShapeComp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.ConnectorShape",)

    # endregion Overrides
