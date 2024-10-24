from __future__ import annotations
from typing import cast, TYPE_CHECKING
import contextlib

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.text.text_range_comp import TextRangeComp
from ooodev.adapter.text.text_portion_type_kind import TextPortionTypeKind

if TYPE_CHECKING:
    from com.sun.star.text import XTextRange
    from com.sun.star.text import TextPortion


class TextPortionComp(TextRangeComp):
    """
    Class for managing TextPortion Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextRange) -> None:
        """
        Constructor

        Args:
            component (TextPortion): UNO TextContent Component that supports ``com.sun.star.text.TextPortion`` service.
        """

        ComponentBase.__init__(self, component)
        TextRangeComp.__init__(self, component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextPortion",)

    # endregion Overrides

    # region Properties
    @property
    def is_collapsed(self) -> bool:
        """Returns if the cursor is collapsed."""
        return self.component.IsCollapsed

    @property
    def is_start(self) -> bool:
        """Returns if the cursor is at the start of the text."""
        return self.component.IsStart

    @property
    def text_portion_type(self) -> TextPortionTypeKind:
        """Returns the type of the text portion."""
        with contextlib.suppress(Exception):
            return TextPortionTypeKind(self.component.TextPortionType)
        return TextPortionTypeKind.UNKNOWN

    @property
    @override
    def component(self) -> TextPortion:
        """Sheet Cell Cursor Component"""
        # pylint: disable=no-member
        return cast("TextPortion", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
