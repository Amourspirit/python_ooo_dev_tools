from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from .write_text_cursor import WriteTextCursor
    from com.sun.star.text import XTextContent
    from ooodev.proto.component_proto import ComponentT

    T = TypeVar("T", bound="ComponentT")

from ooodev.adapter.text.text_table_comp import TextTableComp
from ooodev.adapter.text.cell_range_comp import CellRangeComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils import lo as mLo


class WriteTextTable(Generic[T], TextTableComp, QiPartial):
    """Represents writer text content."""

    def __init__(self, owner: T, component: XTextContent) -> None:
        """
        Constructor

        Args:
            owner (WriteTextCursor): Owner of this component.
            component (XTextContent): UNO object that supports ``com.sun.star.text.TextContent`` service.
        """
        self.__owner = owner
        TextTableComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
