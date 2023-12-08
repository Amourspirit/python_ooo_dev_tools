from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from .text_range_partial import TextRangePartial


if TYPE_CHECKING:
    from com.sun.star.text import TextRange  # service
    from com.sun.star.text import XTextRange


class TextRangeComp(ComponentBase, TextRangePartial):
    """
    Class for managing TextRange Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextRange) -> None:
        """
        Constructor

        Args:
            component (TextRange): UNO TextRange Component that supports ``com.sun.star.text.TextRange`` service.
        """

        ComponentBase.__init__(self, component)
        TextRangePartial.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextRange",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TextRange:
        """TextRange Component"""
        return cast("TextRange", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
