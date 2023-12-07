from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from .text_cursor_comp import TextCursorComp
from . import text_view_cursor_partial as mTextViewCursorPartial

if TYPE_CHECKING:
    from com.sun.star.text import TextViewCursor  # service
    from com.sun.star.text import XTextViewCursor


# it seems that XTextViewCursor Documented in the API for TextCursor Service event thought the service support it.
# To be sure, I will implement it here as TextViewCursorPartial.
class TextViewCursorComp(TextCursorComp, mTextViewCursorPartial.TextViewCursorPartial):
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

        TextCursorComp.__init__(self, component)
        mTextViewCursorPartial.TextViewCursorPartial.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextViewCursor",)

    # endregion Overrides

    # region Methods

    # endregion Methods

    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> TextViewCursor:
            """Sheet Cell Cursor Component"""
            return cast("TextViewCursor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
