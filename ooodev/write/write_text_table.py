from __future__ import annotations
from typing import cast, TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.text.text_table_comp import TextTableComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from . import write_text_portions as mWriteTextPortions

if TYPE_CHECKING:
    from com.sun.star.text import XTextContent
    from com.sun.star.container import XEnumerationAccess

T = TypeVar("T", bound="ComponentT")


class WriteTextTable(Generic[T], LoInstPropsPartial, WriteDocPropPartial, TextTableComp, QiPartial, StylePartial):
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
        TextTableComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    def get_text_portions(self) -> mWriteTextPortions.WriteTextPortions[T]:
        """Returns the text portions of this paragraph."""
        return mWriteTextPortions.WriteTextPortions(
            owner=self.owner, component=cast("XEnumerationAccess", self.component), lo_inst=self.lo_inst
        )

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
