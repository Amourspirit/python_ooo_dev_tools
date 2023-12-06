from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.drawing.generic_drawing_document_comp import GenericDrawingDocumentComp

if TYPE_CHECKING:
    from com.sun.star.presentation import PresentationDocument  # service


class PresentationDocumentComp(GenericDrawingDocumentComp):
    """
    Class for managing PresentationDocument Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: PresentationDocument) -> None:
        """
        Constructor

        Args:
            component (PresentationDocument): UNO PresentationDocument Component
        """

        super().__init__(component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.presentation.PresentationDocument",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> PresentationDocument:
            """PresentationDocument Component"""
            return cast("PresentationDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
