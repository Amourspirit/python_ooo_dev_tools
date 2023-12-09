from __future__ import annotations
from typing import cast, TYPE_CHECKING
from .generic_drawing_document_comp import GenericDrawingDocumentComp

if TYPE_CHECKING:
    from com.sun.star.drawing import DrawingDocument  # service


class DrawingDocumentComp(GenericDrawingDocumentComp):
    """
    Class for managing DrawingDocument Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: DrawingDocument) -> None:
        """
        Constructor

        Args:
            component (DrawingDocument): UNO DrawingDocument Component
        """

        super().__init__(component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.DrawingDocument",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> DrawingDocument:
            """DrawingDocument Component"""
            return cast("DrawingDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
