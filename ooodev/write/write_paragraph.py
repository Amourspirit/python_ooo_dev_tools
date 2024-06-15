from __future__ import annotations
from typing import Any, cast, TypeVar, Generic, TYPE_CHECKING
import uno

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement

from ooodev.adapter.text.text_range_partial import TextRangePartial
from ooodev.adapter.text.paragraph_comp import ParagraphComp
from ooodev.adapter.text.text_content_comp import TextContentComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write import write_text_portions as mWriteTextPortions

if TYPE_CHECKING:
    from com.sun.star.container import XEnumerationAccess
    from ooodev.proto.component_proto import ComponentT

T = TypeVar("T", bound="ComponentT")


class WriteParagraph(
    Generic[T],
    LoInstPropsPartial,
    WriteDocPropPartial,
    TextContentComp,
    ParagraphComp,
    PropertyChangeImplement,
    VetoableChangeImplement,
    TextRangePartial,
    QiPartial,
    TheDictionaryPartial,
    PropPartial,
    StylePartial,
):
    """
    Represents writer paragraph content.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (Any): UNO object that supports ``com.sun.star.text.Paragraph`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextContentComp.__init__(self, component)
        ParagraphComp.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)  # type: ignore
        TextRangePartial.__init__(self, component=self.component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        TheDictionaryPartial.__init__(self)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
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
