from __future__ import annotations
from typing import cast, TYPE_CHECKING
from .generic_text_document_comp import GenericTextDocumentComp


if TYPE_CHECKING:
    from com.sun.star.text import TextDocument  # service


class TextDocumentComp(GenericTextDocumentComp):
    """
    Class for managing Sheet Cell Cursor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: TextDocument) -> None:
        """
        Constructor

        Args:
            component (TextDocument): UNO Sheet Cell Cursor Component
        """

        super().__init__(component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextDocument",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> TextDocument:
            """Sheet Cell Cursor Component"""
            return cast("TextDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
