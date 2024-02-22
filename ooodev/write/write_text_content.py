from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.text.text_content_comp import TextContentComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.loader import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial

if TYPE_CHECKING:
    from com.sun.star.text import XTextContent

T = TypeVar("T", bound="ComponentT")


class WriteTextContent(Generic[T], LoInstPropsPartial, TextContentComp, WriteDocPropPartial, QiPartial, StylePartial):
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
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextContentComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    # region Properties
    @property
    def owner(self) -> T:
        """Component Owner"""
        return self._owner

    # endregion Properties
