from __future__ import annotations
from typing import TYPE_CHECKING
import uno


from ooodev.adapter.text.text_table_cursor_comp import TextTableCursorComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.utils.data_type.range_obj import RangeObj

if TYPE_CHECKING:
    from com.sun.star.text import XTextTableCursor
    from ooodev.write.table.write_table import WriteTable


class WriteTextTableCursor(
    WriteDocPropPartial,
    WriteTablePropPartial,
    LoInstPropsPartial,
    TextTableCursorComp,
    QiPartial,
    PropPartial,
    ServicePartial,
    StylePartial,
):
    """Represents writer Text table cursor."""

    def __init__(self, owner: WriteTable, cursor: XTextTableCursor) -> None:
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)
        WriteTablePropPartial.__init__(self, obj=owner)
        LoInstPropsPartial.__init__(self, lo_inst=self.write_doc.lo_inst)
        TextTableCursorComp.__init__(self, cursor)  # type: ignore
        QiPartial.__init__(self, component=cursor, lo_inst=self.lo_inst)  # type: ignore
        PropPartial.__init__(self, component=cursor, lo_inst=self.lo_inst)  # type: ignore
        ServicePartial.__init__(self, component=cursor, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=cursor)
        self._owner = owner

    def get_range_obj(self) -> RangeObj:
        """
        The name is the cell name of the top left table cell of the range concatenated by ``:`` with the table cell name of the bottom left table cell of the cell range.
        If the range consists of one table cell only then ``RangeObj`` will have the same name for both the start and end cell.
        """
        name = self.component.getRangeName()
        return self.write_table.range_converter.rng_from_str(name)

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}({self.get_range_name()})"

    @property
    def owner(self) -> WriteTable:
        """Owner of this component."""
        return self._owner
