from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_range_comp import TextRangeComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial

if TYPE_CHECKING:
    from com.sun.star.text import XTextRange
    from ooodev.write.write_text_cursor import WriteTextCursor

T = TypeVar("T", bound="ComponentT")


class WriteTextRange(
    Generic[T],
    LoInstPropsPartial,
    WriteDocPropPartial,
    TextRangeComp,
    PropertyChangeImplement,
    VetoableChangeImplement,
    QiPartial,
    PropPartial,
    StylePartial,
):
    """Represents writer TextRange."""

    def __init__(self, owner: T, component: XTextRange, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextRange): UNO object that supports ``com.sun.star.text.TextRange`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextRangeComp.__init__(self, component)  # type: ignore
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    # region Properties

    def get_cursor(self) -> WriteTextCursor[WriteTextRange[T]]:
        """
        Gets a cursor for this text range.

        Returns:
            WriteTextCursor[WriteTextRange[T]]: The cursor.
        """
        # pylint: disable=import-outside-toplevel
        # not concerned about compile import for WriteTextCursor
        from .write_text_cursor import WriteTextCursor

        cursor = self.write_doc.component.getText().createTextCursorByRange(self.component)
        return WriteTextCursor(owner=self, component=cursor, lo_inst=self.lo_inst)

    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
