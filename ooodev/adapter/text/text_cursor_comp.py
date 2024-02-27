from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.text.text_range_comp import TextRangeComp
from ooodev.adapter.text.text_cursor_partial import TextCursorPartial


if TYPE_CHECKING:
    from com.sun.star.text import TextCursor  # service
    from com.sun.star.text import XTextCursor


class TextCursorComp(
    TextRangeComp,
    TextCursorPartial,
):
    """
    Class for managing TextCursor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextCursor) -> None:
        """
        Constructor

        Args:
            component (TextCursor): UNO TextCursor Component that supports ``com.sun.star.text.TextCursor`` service.
        """

        TextRangeComp.__init__(self, component)
        TextCursorPartial.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextCursor",)

    # endregion Overrides

    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> TextCursor:
            """Sheet Cell Cursor Component"""
            # pylint: disable=no-member
            return cast("TextCursor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
