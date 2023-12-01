from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.text.generic_text_document_comp import GenericTextDocumentComp


if TYPE_CHECKING:
    from com.sun.star.text import GlobalDocument  # service


class GlobalDocumentComp(GenericTextDocumentComp):
    """
    Class for managing Sheet Cell Cursor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: GlobalDocument) -> None:
        """
        Constructor

        Args:
            component (GlobalDocument): UNO Sheet Cell Cursor Component
        """

        super().__init__(component)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.GlobalDocument",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> GlobalDocument:
        """Sheet Cell Cursor Component"""
        return cast("GlobalDocument", self._get_component())

    # endregion Properties
