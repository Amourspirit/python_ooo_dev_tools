from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.drawing.generic_drawing_document_comp import GenericDrawingDocumentComp

if TYPE_CHECKING:
    from com.sun.star.drawing import DrawingDocument  # service
    from com.sun.star.lang import XComponent


class DrawingDocumentComp(GenericDrawingDocumentComp):
    """
    Class for managing DrawingDocument Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XComponent) -> None:
        """
        Constructor

        Args:
            component (XComponent): UNO Component that supports ``com.sun.star.drawing.DrawingDocument`` service.
        """

        super().__init__(component)
        # generic_args = self._ComponentBase__get_generic_args()  # type: ignore

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.DrawingDocument",)

    # endregion Overrides
    # region Properties

    @property
    @override
    def component(self) -> DrawingDocument:
        """DrawingDocument Component"""
        # override to satisfy documentation and type
        # return cast("DrawingDocument", self._ComponentBase__get_component())  # type: ignore
        return cast("DrawingDocument", super().component)  # type: ignore

    # endregion Properties
