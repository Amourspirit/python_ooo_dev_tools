from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.text import XTextContent

if TYPE_CHECKING:
    from com.sun.star.text import XTextRange

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo


class TextContentPartial:
    """
    Class for managing TextRange.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextContent) -> None:
        """
        Constructor

        Args:
            component (XTextContent): UNO Component that implements ``com.sun.star.text.XTextContent`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XTextContent):
            raise mEx.MissingInterfaceError("XTextContent")
        self.__component = component

    # region XTextContent
    def attach(self, text_range: XTextRange) -> None:
        """Attaches a text range to this text content."""
        self.__component.attach(text_range)

    def get_anchor(self) -> XTextRange:
        """Returns the anchor of this text content."""
        return self.__component.getAnchor()

    # endregion XTextContent
