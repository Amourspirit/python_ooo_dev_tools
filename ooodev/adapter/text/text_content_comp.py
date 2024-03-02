from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno


from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.text.text_content_partial import TextContentPartial


if TYPE_CHECKING:
    from com.sun.star.text import TextContent
    from com.sun.star.text import XTextContent
    from ooo.dyn.text.text_content_anchor_type import TextContentAnchorType
    from ooo.dyn.text.wrap_text_mode import WrapTextMode


class TextContentComp(ComponentBase, TextContentPartial):
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
        TextContentPartial.__init__(self, component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextContent",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TextContent:
        """Sheet Cell Cursor Component"""
        # pylint: disable=no-member
        return cast("TextContent", self._ComponentBase__get_component())  # type: ignore

    @property
    def anchor_type(self) -> TextContentAnchorType:
        """
        Gets/Sets the anchor type of this text content.

        Returns:
            TextContentAnchorType: Anchor Type

        Hint:
            - ``TextContentAnchorType``can be imported from ``ooo.dyn.text.text_content_anchor_type``.
        """
        return self.component.AnchorType  # type: ignore

    @anchor_type.setter
    def anchor_type(self, value: TextContentAnchorType) -> None:
        """Sets the anchor type of this text content."""
        self.component.AnchorType = value  # type: ignore

    @property
    def text_wrap(self) -> WrapTextMode:
        """
        Gets/Sets if the text content is a shape and how the text is wrapped around the shape.

        Returns:
            WrapTextMode: Text Wrap Mode

        Hint:
            - ``WrapTextMode`` can be imported from ``ooo.dyn.text.wrap_text_mode``
        """
        return self.component.TextWrap  # type: ignore

    @text_wrap.setter
    def text_wrap(self, value: WrapTextMode) -> None:
        self.component.TextWrap = value  # type: ignore

    # endregion Properties
