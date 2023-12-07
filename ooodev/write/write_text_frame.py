from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.text import XTextFrame

from ooodev.adapter.text.text_frame_comp import TextFrameComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial

T = TypeVar("T", bound="ComponentT")


class WriteTextFrame(Generic[T], TextFrameComp, QiPartial, PropPartial):
    """Represents writer text content."""

    def __init__(self, owner: T, component: XTextFrame) -> None:
        """
        Constructor

        Args:
            owner (WriteTextCursor): Owner of this component.
            component (XTextFrame): UNO object that supports ``com.sun.star.text.TextFrame`` service.
        """
        self.__owner = owner
        TextFrameComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
