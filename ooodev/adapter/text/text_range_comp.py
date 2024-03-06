from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.text import text_range_partial as mTextRangePartial
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial


if TYPE_CHECKING:
    from com.sun.star.text import TextRange  # service
    from com.sun.star.text import XTextRange


class TextRangeComp(
    ComponentBase, mTextRangePartial.TextRangePartial, CharacterPropertiesPartial, ParagraphPropertiesPartial
):
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
        CharacterPropertiesPartial.__init__(self, component=component)  # type: ignore
        ParagraphPropertiesPartial.__init__(self, component=component)  # type: ignore

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
