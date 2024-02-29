from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.text.text_table_row_comp import TextTableRowComp
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial

if TYPE_CHECKING:
    from com.sun.star.text import TextTableRow  # service
    from ooodev.write.table.write_table_rows import WriteTableRows


class WriteTableRow(TextTableRowComp, WriteDocPropPartial, LoInstPropsPartial):
    """Represents writer table rows."""

    def __init__(self, owner: WriteTableRows, component: TextTableRow) -> None:
        """
        Constructor

        Args:
            component (TextTableRow): UNO object that supports ``om.sun.star.text.TextTableRow`` service.
        """
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        LoInstPropsPartial.__init__(self, lo_inst=owner.lo_inst)
        TextTableRowComp.__init__(self, component=component)  # type: ignore
        self._owner = owner

    @property
    def owner(self) -> WriteTableRows:
        """Owner of this component."""
        return self._owner
