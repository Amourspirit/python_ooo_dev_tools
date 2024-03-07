from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.mock import mock_g
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_table_row_comp import TextTableRowComp
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from com.sun.star.text import TextTableRow  # service
    from ooodev.write.table.write_table_rows import WriteTableRows
    from ooodev.write.style.direct.table.row_styler import RowStyler


class WriteTableRow(
    WriteDocPropPartial,
    WriteTablePropPartial,
    EventsPartial,
    TextTableRowComp,
    LoInstPropsPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    StylePartial,
):
    """Represents writer table rows."""

    def __init__(self, owner: WriteTableRows, component: TextTableRow) -> None:
        """
        Constructor

        Args:
            component (TextTableRow): UNO object that supports ``om.sun.star.text.TextTableRow`` service.
        """
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        WriteTablePropPartial.__init__(self, obj=owner.write_table)
        EventsPartial.__init__(self)
        LoInstPropsPartial.__init__(self, lo_inst=owner.lo_inst)
        TextTableRowComp.__init__(self, component=component)  # type: ignore
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)

        self._style_direct_row = None
        self._owner = owner

    @property
    def owner(self) -> WriteTableRows:
        """Owner of this component."""
        return self._owner

    @property
    def style_direct(self) -> RowStyler:
        """
        Direct Cell Styler.

        Returns:
            CellStyler: Character Styler
        """
        if self._style_direct_row is None:
            # pylint: disable=import-outside-toplevel
            from ooodev.write.style.direct.table.row_styler import RowStyler

            self._style_direct_row = RowStyler(owner=self.write_table, component=self.component)
            self._style_direct_row.add_event_observers(self.event_observer)
        return self._style_direct_row


if mock_g.FULL_IMPORT:
    from ooodev.write.style.direct.table.row_styler import RowStyler
