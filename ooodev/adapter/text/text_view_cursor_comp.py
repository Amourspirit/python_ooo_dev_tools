from __future__ import annotations
from typing import cast, TYPE_CHECKING

# from .text_cursor_comp import TextCursorComp
from ooodev.adapter.text import text_view_cursor_partial as mTextViewCursorPartial
from ooodev.adapter.text.page_cursor_partial import PageCursorPartial
from ooodev.adapter.view.screen_cursor_partial import ScreenCursorPartial

# from .word_cursor_partial import WordCursorPartial

from ooodev.adapter.text.text_range_comp import TextRangeComp

# from .text_cursor_partial import TextCursorPartial
# from .sentence_cursor_partial import SentenceCursorPartial
# from .paragraph_cursor_partial import ParagraphCursorPartial

if TYPE_CHECKING:
    from com.sun.star.text import TextViewCursor  # service
    from com.sun.star.text import XTextViewCursor


# it seems that XTextViewCursor Documented in the API for TextCursor Service event thought the service support it.
# To be sure, I will implement it here as TextViewCursorPartial.
# TextRangeComp, TextCursorPartial, ParagraphCursorPartial, SentenceCursorPartial, WordCursorPartial
class TextViewCursorComp(
    mTextViewCursorPartial.TextViewCursorPartial, TextRangeComp, PageCursorPartial, ScreenCursorPartial
):
    """
    Class for managing TextCursor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextViewCursor) -> None:
        """
        Constructor

        Args:
            component (TextCursor): UNO TextCursor Component that supports ``com.sun.star.text.TextViewCursor`` service.
        """

        mTextViewCursorPartial.TextViewCursorPartial.__init__(self, component)
        TextRangeComp.__init__(self, component)
        PageCursorPartial.__init__(self, component, None)  # type: ignore
        ScreenCursorPartial.__init__(self, component, None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextViewCursor",)

    # endregion Overrides

    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> TextViewCursor:
            """Sheet Cell Cursor Component"""
            # pylint: disable=no-member
            return cast("TextViewCursor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
