from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.adapter.text.text_columns_comp import TextColumnsComp
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial

if TYPE_CHECKING:
    from com.sun.star.table import XTableRows
    from ooodev.write.write_text_table import WriteTextTable


class WriteTextColumns(TextColumnsComp, WriteDocPropPartial, LoInstPropsPartial):
    """Represents writer table rows."""

    def __init__(self, owner: WriteTextTable, component: XTableRows) -> None:
        """
        Constructor

        Args:
            component (XTableRows): UNO object that supports ``com.sun.star.text.TextColumns`` service.
        """
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        LoInstPropsPartial.__init__(self, lo_inst=owner.lo_inst)
        TextColumnsComp.__init__(self, component=component)  # type: ignore
        self._owner = owner

    @property
    def write_text_table(self) -> WriteTextTable:
        """Owner of this component."""
        return self._owner
