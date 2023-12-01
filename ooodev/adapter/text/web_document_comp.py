from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from .generic_text_document_comp import GenericTextDocumentComp


if TYPE_CHECKING:
    from com.sun.star.text import WebDocument  # service


class WebDocumentComp(GenericTextDocumentComp):
    """
    Class for managing WebDocument Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: WebDocument) -> None:
        """
        Constructor

        Args:
            component (WebDocument): UNO Sheet Cell Cursor Component
        """

        super().__init__(component)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.WebDocument",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> WebDocument:
            """Sheet Cell Cursor Component"""
            return cast("WebDocument", self._get_component())

    # endregion Properties
