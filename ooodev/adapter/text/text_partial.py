from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.text import XText
from ooodev.adapter.text.simple_text_partial import SimpleTextPartial

if TYPE_CHECKING:
    from com.sun.star.text import XTextContent
    from com.sun.star.text import XTextRange
    from ooodev.utils.type_var import UnoInterface


class TextPartial(SimpleTextPartial):
    """
    Partial class for XText.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XText, interface: UnoInterface | None = XText) -> None:
        """
        Constructor

        Args:
            component (XText): UNO Component that implements ``com.sun.star.text.XText`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XText``.
        """

        SimpleTextPartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XText
    def insert_text_content(self, rng: XTextRange, content: XTextContent, absorb: bool) -> None:
        """
        Inserts a content, such as a text table, text frame or text field.

        Args:
            rng (XTextRange): The position at which the content is inserted.
            content (XTextContent): The content to be inserted.
            absorb (bool): Specifies whether the text spanned by xRange will be replaced.
                If ``True`` then the content of range will be replaced by content,
                otherwise content will be inserted at the end of xRange.

        Returns:
            None:
        """
        self.__component.insertTextContent(rng, content, absorb)

    def remove_text_content(self, content: XTextContent) -> None:
        """
        Removes a text content.

        Args:
            content (XTextContent): the content that is to be removed.

        Returns:
            None:
        """
        self.__component.removeTextContent(content)

    # endregion XText
