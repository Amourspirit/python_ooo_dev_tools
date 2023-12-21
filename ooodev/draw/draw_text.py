from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.text import XText

from ooodev.adapter.drawing.text_comp import TextComp
from ooodev.office import draw as mDraw
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from .draw_text_cursor import DrawTextCursor

_T = TypeVar("_T", bound="ComponentT")


class DrawText(Generic[_T], TextComp, QiPartial):
    """
    Represents text content.

    Contains Enumeration Access.
    """

    def __init__(self, owner: _T, component: XText) -> None:
        """
        Constructor

        Args:
            owner (_T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.Text`` service.
        """
        self.__owner = owner
        TextComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    def add_bullet(self, level: int, text: str) -> None:
        """
        Add bullet text to the end of the bullets text area, specifying
        the nesting of the bullet using a numbering level value
        (numbering starts at 0).

        Args:
            level (int): Bullet Level
            text (str): Bullet Text

        Raises:
            DrawError: If error adding bullet.

        Returns:
            None:
        """
        mDraw.Draw.add_bullet(self.component, level, text)

    def get_cursor(self) -> DrawTextCursor[_T]:
        """
        Get the cursor for this text.

        Returns:
            DrawTextCursor[_T]: Cursor for this text.
        """
        return DrawTextCursor(self.owner, self.component.createTextCursor())

    # region Properties
    @property
    def owner(self) -> _T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
