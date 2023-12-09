from __future__ import annotations
from typing import cast, TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.text import XTextContent
    from com.sun.star.container import XEnumerationAccess

from ooodev.adapter.text.text_table_comp import TextTableComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from . import write_text_portions as mWriteTextPortions

T = TypeVar("T", bound="ComponentT")


class WriteTextTable(Generic[T], TextTableComp, QiPartial):
    """Represents writer text content."""

    def __init__(self, owner: T, component: XTextContent) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextContent): UNO object that supports ``com.sun.star.text.TextContent`` service.
        """
        self.__owner = owner
        TextTableComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    def get_text_portions(self) -> mWriteTextPortions.WriteTextPortions[T]:
        """Returns the text portions of this paragraph."""
        return mWriteTextPortions.WriteTextPortions(
            owner=self.owner, component=cast("XEnumerationAccess", self.component)
        )

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
