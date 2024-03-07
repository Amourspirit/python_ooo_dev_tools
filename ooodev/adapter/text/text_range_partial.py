from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.text import XTextRange

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.text import XText
    from ooodev.utils.type_var import UnoInterface


class TextRangePartial:
    """
    Partial class for XTextRange.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextRange, interface: UnoInterface | None = XTextRange) -> None:
        """
        Constructor

        Args:
            component (XTextRange): UNO Component that implements ``com.sun.star.text.XTextRange`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextRange``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
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


from ooodev.adapter.text import text_range_comp as mTextRangeComp
