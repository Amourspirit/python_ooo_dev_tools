from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.text import XTextViewCursor

from ooodev.adapter.text import text_cursor_partial as mTextCursorPartial

if TYPE_CHECKING:
    from com.sun.star.awt import Point
    from ooodev.utils.type_var import UnoInterface


class TextViewCursorPartial(mTextCursorPartial.TextCursorPartial):
    """
    Partial class for XTextViewCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextViewCursor, interface: UnoInterface | None = XTextViewCursor) -> None:
        """
        Constructor

        Args:
            component (XTextViewCursor): UNO Component that implements ``com.sun.star.text.XTextViewCursor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextViewCursor``.
        """

        mTextCursorPartial.TextCursorPartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XTextViewCursor
    def is_visible(self) -> bool:
        """Returns True if the cursor is visible."""
        return self.__component.isVisible()

    def set_visible(self, visible: bool = True) -> None:
        """
        Sets the visibility of the cursor.

        Args:
            visible (bool, optional): ``True`` to set the view cursor visible, False to hide it. Defaults to ``True``.

        Returns:
            None:
        """
        self.__component.setVisible(visible)

    def get_position(self) -> Point:
        """
        Gets the cursor's coordinates relative to the top left position of the first page
        of the document in ``1/100 mm``.

        Returns:
            Point: The cursor's coordinates as ``com.sun.star.awt.Point``. in ``1/100 mm``.

        Note:
            The ``X`` coordinate is the horizontal position, the ``Y`` coordinate is the vertical position.

        Warning:
            The ``X`` is relative to the document window and not the document page.
            This means when the document window size changes the ``X`` coordinate will
            change even if the cursor has not moved.
            This is also the case if the document zoom changes.

            When the document page is zoomed all the way to fill the document window the ``X`` coordinate
            is ``0`` when at the left hand page margin (beginning of a line).

            The ``Y`` coordinate is relative to the top of the document window and not the top of the document page.
            The ``Y`` coordinate seems not to be affected by Document Zoom or scroll position.
        """
        return self.__component.getPosition()

    # endregion XTextViewCursor
