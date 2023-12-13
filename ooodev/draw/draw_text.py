from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.text import XText

from ooodev.adapter.drawing.text_comp import TextComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write import write_paragraphs as mWriteParagraphs
from ooodev.write import write_text_tables as mWriteTextTables

T = TypeVar("T", bound="ComponentT")


class DrawText(Generic[T], TextComp, QiPartial):
    """
    Represents writer text content.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: XText) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.Text`` service.
        """
        self.__owner = owner
        TextComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    def get_paragraphs(self) -> mWriteParagraphs.WriteParagraphs[T]:
        """Returns the paragraphs of this text."""
        return mWriteParagraphs.WriteParagraphs(owner=self.owner, component=self.component)

    def get_text_tables(self) -> mWriteTextTables.WriteTextTables[T]:
        """Returns the text tables of this text."""
        return mWriteTextTables.WriteTextTables(owner=self.owner, component=self.component)

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
