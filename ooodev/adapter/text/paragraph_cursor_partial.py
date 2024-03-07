from __future__ import annotations
from typing import Any
import uno

from com.sun.star.text import XParagraphCursor
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface


class ParagraphCursorPartial:
    """
    Partial class for XParagraphCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XParagraphCursor, interface: UnoInterface | None = XParagraphCursor) -> None:
        """
        Constructor

        Args:
            component (XParagraphCursor): UNO Component that implements ``com.sun.star.text.XParagraphCursor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XParagraphCursor``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XParagraphCursor
    def is_start_of_paragraph(self) -> bool:
        """
        Returns true if the cursor is at the start of a paragraph.

        Returns:
            bool: ``True`` if the cursor is at the start of a paragraph.
        """

        return self.__component.isStartOfParagraph()

    def is_end_of_paragraph(self) -> bool:
        """
        Returns true if the cursor is at the end of a paragraph.

        Returns:
            bool: ``True`` if the cursor is at the end of a paragraph.
        """

        return self.__component.isEndOfParagraph()

    def goto_start_of_paragraph(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the start of the current paragraph.

        Args:
            expand (bool, optional): If ``True`` the range of the cursor will be expanded to the start of the paragraph. Default is ``False``.

        Returns:
            bool: ``True`` if the cursor is now at the start of a paragraph, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """

        return self.__component.gotoStartOfParagraph(expand)

    def goto_end_of_paragraph(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the end of the current paragraph.

        Args:
            expand (bool, optional): If ``True`` the range of the cursor will be expanded to the end of the paragraph. Default is ``False``.

        Returns:
            bool: ``True`` if the cursor is now at the end of a paragraph, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """

        return self.__component.gotoEndOfParagraph(expand)

    def goto_next_paragraph(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the next paragraph.

        Args:
            expand (bool, optional): If ``True`` the range of the cursor will be expanded to the next paragraph. Default is ``False``.

        Returns:
            bool: ``True`` if the cursor is now at the start of a paragraph, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """

        return self.__component.gotoNextParagraph(expand)

    def goto_previous_paragraph(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the previous paragraph.

        Args:
            expand (bool, optional): If ``True`` the range of the cursor will be expanded to the previous paragraph. Default is ``False``.

        Returns:
            bool: ``True`` if the cursor is now at the start of a paragraph, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """

        return self.__component.gotoPreviousParagraph(expand)

    # endregion XParagraphCursor
