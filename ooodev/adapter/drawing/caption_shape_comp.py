from __future__ import annotations
from typing import Any, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.drawing.generic_shape import GenericShapeComp

if TYPE_CHECKING:
    from com.sun.star.drawing import CaptionShape  # service
else:
    CaptionShape = Any


class CaptionShapeComp(GenericShapeComp[CaptionShape]):
    """
    Class for managing CaptionShape Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that implements ``com.sun.star.drawing.CaptionShape`` service.
        """
        GenericShapeComp.__init__(self, component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.CaptionShape",)

    # endregion Overrides
