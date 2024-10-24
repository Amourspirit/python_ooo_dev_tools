from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.text.generic_text_document_comp import GenericTextDocumentComp


if TYPE_CHECKING:
    from com.sun.star.text import GlobalDocument  # service


class GlobalDocumentComp(GenericTextDocumentComp):
    """
    Class for managing GlobalDocumentComp Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: GlobalDocument) -> None:
        """
        Constructor

        Args:
            component (GlobalDocument): UNO GlobalDocumentComp Component
        """

        super().__init__(component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.GlobalDocument",)

    # endregion Overrides
    # region Properties
    @property
    @override
    def component(self) -> GlobalDocument:
        """GlobalDocumentComp Component"""
        # pylint: disable=no-member
        return cast("GlobalDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
