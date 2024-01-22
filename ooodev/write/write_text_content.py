from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno

if TYPE_CHECKING:
    from com.sun.star.text import XTextContent

from ooodev.adapter.text.text_content_comp import TextContentComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.inst.lo.lo_inst import LoInst

T = TypeVar("T", bound="ComponentT")


class WriteTextContent(Generic[T], TextContentComp, QiPartial, StylePartial):
    """Represents writer text content."""

    def __init__(self, owner: T, component: XTextContent, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextContent): UNO object that supports ``com.sun.star.text.TextContent`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        self._owner = owner
        TextContentComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    # region Properties
    @property
    def owner(self) -> T:
        """Component Owner"""
        return self._owner

    # endregion Properties
