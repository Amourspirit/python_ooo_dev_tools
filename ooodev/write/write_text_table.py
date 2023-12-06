from __future__ import annotations
from typing import TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from .write_text_cursor import WriteTextCursor

from ooodev.adapter.text.text_table_comp import TextTableComp
from ooodev.adapter.text.cell_range_comp import CellRangeComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.text import XTextContent


class WriteTextTable(TextTableComp, QiPartial):
    """Represents writer text content."""

    def __init__(self, owner: WriteTextCursor, component: XTextContent) -> None:
        """
        Constructor

        Args:
            owner (WriteTextCursor): Sheet that owns this cell range.
            component (XTextContent): UNO object that supports ``com.sun.star.text.TextContent`` service.
        """
        self.__owner = owner
        TextTableComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def write_text_cursor(self) -> WriteTextCursor:
        """Doc that owns this Cursor."""
        return self.__owner

    # endregion Properties
