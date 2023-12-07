from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno

if TYPE_CHECKING:
    from com.sun.star.text import XTextContent


from ooodev.adapter.text.text_content_comp import TextContentComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial

T = TypeVar("T", bound="ComponentT")


class WriteTextContent(Generic[T], TextContentComp, QiPartial):
    """Represents writer text content."""

    def __init__(self, owner: T, component: XTextContent) -> None:
        """
        Constructor

        Args:
            owner (T): Cursor or Doc that owns this component.
            component (XTextContent): UNO object that supports ``com.sun.star.text.TextContent`` service.
        """
        self.__owner = owner
        TextContentComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Component Owner"""
        return self.__owner

    # endregion Properties
