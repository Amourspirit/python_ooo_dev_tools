from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from com.sun.star.text import XTextRange

if TYPE_CHECKING:
    from com.sun.star.text import XText

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from . import text_range_comp as mTextRangeComp


class TextRangePartial:
    """
    Class for managing TextRange.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextRange) -> None:
        """
        Constructor

        Args:
            component (XTextRange): UNO Component that implements ``com.sun.star.text.XTextRange`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XTextRange):
            raise mEx.MissingInterfaceError("XTextRange")
        self.__component = component

    # region Methods
    def get_text(self) -> XText:
        """
        Gets the text of the range.

        Returns:
            XText: The text of the range.
        """
        return self.__component.getText()

    def get_start(self) -> mTextRangeComp.TextRangeComp:
        """Returns a text range which contains only the start of this text range."""
        return mTextRangeComp.TextRangeComp(self.__component.getStart())

    def get_end(self) -> mTextRangeComp.TextRangeComp:
        """Returns a text range which contains only the end of this text range."""
        return mTextRangeComp.TextRangeComp(self.__component.getEnd())

    def get_string(self) -> str:
        """Returns the string of this text range."""
        return self.__component.getString()

    def set_string(self, string: str) -> None:
        """
        Sets the string of this text range.

        The whole string of characters of this piece of text is replaced.
        All styles are removed when applying this method.
        """
        self.__component.setString(string)

    # endregion Methods
