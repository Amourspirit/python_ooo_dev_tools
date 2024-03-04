from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.view import XLineCursor

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class LineCursorPartial:
    """
    Partial class for XLineCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XLineCursor, interface: UnoInterface | None = XLineCursor) -> None:
        """
        Constructor

        Args:
            component (XLineCursor ): UNO Component that implements ``com.sun.star.view.XLineCursor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLineCursor``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLineCursor
    def goto_end_of_line(self, expand: bool = False) -> None:
        """
        Moves the cursor to the end of the current line.

        Args:
            expand (bool, optional): Determines whether the text range of the cursor is expanded (``True``) or the cursor will be just at the new position after the move (``False``). Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoEndOfLine(expand)

    def goto_start_of_line(self, expand: bool = False) -> None:
        """
        Moves the cursor to the start of the current line.

        Args:
            expand (bool, optional): Determines whether the text range of the cursor is expanded (``True``) or the cursor will be just at the new position after the move (``False``). Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoStartOfLine(expand)

    def is_at_end_of_line(self) -> bool:
        """
        Determines if the cursor is positioned at the end of a line.

        Returns:
            bool: ``True`` if the cursor is positioned at the end of a line, ``False`` otherwise.
        """
        return self.__component.isAtEndOfLine()

    def is_at_start_of_line(self) -> bool:
        """
        Determines if the cursor is positioned at the start of a line.

        Returns:
            bool: ``True`` if the cursor is positioned at the start of a line, ``False`` otherwise.
        """
        return self.__component.isAtStartOfLine()

    # endregion XLineCursor
