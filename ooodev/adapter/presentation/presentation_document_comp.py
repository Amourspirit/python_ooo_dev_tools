from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.drawing.generic_drawing_document_comp import GenericDrawingDocumentComp

if TYPE_CHECKING:
    from com.sun.star.presentation import PresentationDocument  # service
    from com.sun.star.lang import XComponent


class PresentationDocumentComp(GenericDrawingDocumentComp):
    """
    Class for managing PresentationDocument Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XComponent) -> None:
        """
        Constructor

        Args:
            component (XComponent): UNO component that supports ``com.sun.star.presentation.PresentationDocument`` service.
        """

        super().__init__(component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.presentation.PresentationDocument",)

    # endregion Overrides
    # region Properties

    @property
    def component(self) -> PresentationDocument:
        """PresentationDocument Component"""
        # override to satisfy documentation and type
        return cast("PresentationDocument", super().component)
        # return cast("PresentationDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
