from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooo.dyn.text.text_content_anchor_type import TextContentAnchorType

from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.text import TextContent
    from com.sun.star.text import XTextContent


class TextContentComp(ComponentBase):
    """
    Class for managing TextContent Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextContent) -> None:
        """
        Constructor

        Args:
            component (TextContent): UNO TextContent Component that supports ``com.sun.star.text.TextContent`` service.
        """

        ComponentBase.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextContent",)

    # endregion Overrides

    # region Methods

    # endregion Methods

    # region Properties
    @property
    def component(self) -> TextContent:
        """Sheet Cell Cursor Component"""
        return cast("TextContent", self._ComponentBase__get_component())  # type: ignore

    @property
    def anchor_type(self) -> TextContentAnchorType:
        """Returns the anchor type of this text content."""
        return TextContentAnchorType(self.component.AnchorType)

    @anchor_type.setter
    def anchor_type(self, value: TextContentAnchorType) -> None:
        """Sets the anchor type of this text content."""
        self.component.AnchorType = value  # type: ignore

    # endregion Properties
