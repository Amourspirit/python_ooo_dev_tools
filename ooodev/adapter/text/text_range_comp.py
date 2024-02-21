from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from . import text_range_partial as mTextRangePartial


if TYPE_CHECKING:
    from com.sun.star.text import TextRange  # service
    from com.sun.star.text import XTextRange


class TextRangeComp(ComponentBase, mTextRangePartial.TextRangePartial):
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
        mTextRangePartial.TextRangePartial.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by mTextRangePartial.TextRangePartial
        return ()  # ("com.sun.star.text.TextRange",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TextRange:
        """TextRange Component"""
        # pylint: disable=no-member
        return cast("TextRange", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
