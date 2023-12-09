from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

# com.sun.star.text.ControlCharacter
from ooo.dyn.text.control_character import ControlCharacterEnum

from com.sun.star.text import XSimpleText
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.text import XTextCursor
    from com.sun.star.text import XTextRange
    from ooodev.utils.type_var import UnoInterface


class SimpleTextPartial:
    """
    Partial class for XSimpleText.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSimpleText, interface: UnoInterface | None = XSimpleText) -> None:
        """
        Constructor

        Args:
            component (XSimpleText): UNO Component that implements ``com.sun.star.text.XSimpleText`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSimpleText``.
        """

        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XSimpleText
    def create_text_cursor(self) -> XTextCursor:
        """
        Creates a new text cursor.

        Returns:
            XTextCursor: The new text cursor.
        """
        return self.__component.createTextCursor()

    def create_text_cursor_by_range(self, text_position: XTextRange) -> XTextCursor:
        """
        The initial position is set to aTextPosition.

        Args:
            text_position (XTextRange): The initial position of the new text cursor.

        Returns:
            XTextCursor: The new text cursor.
        """
        return self.__component.createTextCursorByRange(text_position)

    def insert_control_character(self, rng: XTextRange, control_character: ControlCharacterEnum, absorb: bool) -> None:
        """
        Inserts a control character (like a paragraph break or a hard space) into the text.

        Args:
            rng (XTextRange): The position of the new control character.
            control_character (ControlCharacterEnum): The control character to be inserted.
            absorb (bool): If TRUE the text range will contain the new inserted control character, otherwise the range (and it's text) will remain unchanged.

        Raises:
            IllegalArgumentException: ``com.sun.star.lang.IllegalArgumentException``
        """
        return self.__component.insertControlCharacter(rng, int(control_character), absorb)

    def insert_string(self, rng: XTextRange, text: str, absorb: bool) -> None:
        """
        Inserts a string of characters into the text.

        The string may contain the following white spaces:

        If the parameter bAbsorb() was TRUE the text range will contain the new inserted string, otherwise the range (and it's text) will remain unchanged.
        """
        self.__component.insertString(rng, text, absorb)

    # endregion XSimpleText
