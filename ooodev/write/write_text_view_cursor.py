from __future__ import annotations
from typing import Sequence, overload, TYPE_CHECKING, TypeVar, Generic
from com.sun.star.text import TextViewCursor
import uno

if TYPE_CHECKING:
    from com.sun.star.text import XTextViewCursor
    from ooodev.proto.component_proto import ComponentT

    T = TypeVar("T", bound="ComponentT")

from ooodev.adapter.text.text_view_cursor_comp import TextViewCursorComp
from ooodev.office import write as mWrite
from . import write_text_cursor as mWriteTextCursor


class WriteTextViewCursor(mWriteTextCursor.WriteTextCursor, TextViewCursorComp):
    """Represents a writer text view cursor."""

    def __init__(self, owner, component: XTextViewCursor) -> None:
        """
        Constructor

        Args:
            owner (WriteDoc): Doc that owns this component.
            col_obj (Any): Range object.
        """
        self.__owner = owner
        mWriteTextCursor.WriteTextCursor.__init__(owner, component)  # type: ignore
        TextViewCursorComp.__init__(self, component=component)

    def get_coord_str(self) -> str:
        """
        Gets coordinates for cursor in format such as ``"10, 10"``

        Returns:
            str: coordinates as string
        """
        return mWrite.Write.get_coord_str(self.component)  # type: ignore

    def get_current_page_num(self) -> int:
        """
        Gets current page number.

        Returns:
            int: current page number
        """
        return mWrite.Write.get_current_page(self.component)  # type: ignore
