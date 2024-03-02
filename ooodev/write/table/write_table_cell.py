from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.adapter.text.cell_properties_comp import CellPropertiesComp
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial

if TYPE_CHECKING:
    from com.sun.star.text import TextTableRow  # service
    from ooodev.write.table.write_table_rows import WriteTableRows
    from ooodev.proto.component_proto import ComponentT


class WriteTableCell(
    WriteDocPropPartial,
    WriteTablePropPartial,
    CellPropertiesComp,
    LoInstPropsPartial,
    PropPartial,
    StylePartial,
):
    """Represents writer table rows."""

    def __init__(self, owner: ComponentT, component: TextTableRow) -> None:
        """
        Constructor

        Args:
            component (TextTableRow): UNO object that supports ``om.sun.star.text.TextTableRow`` service.
        """
        if not isinstance(owner, WriteTablePropPartial):
            raise ValueError("owner must be a WriteTablePropPartial instance.")
        WriteTablePropPartial.__init__(self, obj=owner.write_table)
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        LoInstPropsPartial.__init__(self, lo_inst=self.write_doc.lo_inst)
        CellPropertiesComp.__init__(self, component=component)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        self._owner = owner

    @property
    def owner(self) -> ComponentT:
        """Owner of this component."""
        return self._owner
